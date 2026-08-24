"""CATIA COM 客户端 —— 只做最薄的一层封装。

设计原则：
1. 只依赖 pywin32，不引入任何 MCP / LLM 概念。
2. 只暴露"探测型"只读方法，方便 Hello World 阶段验证通路。
3. 单 STA 语义由调用方保证（后续 MCP server 里做）。

只能在 Windows 上运行；Mac / Linux 上 import 会失败，这是刻意的 —— 让通路问题早暴露。
"""

from __future__ import annotations

import math
import os
import sys
import time
from dataclasses import dataclass
from typing import Optional

if sys.platform != "win32":
    raise RuntimeError(
        "catia_client 只能在 Windows 上运行。"
        "请把代码同步到 Windows 机器（那台装了 CATIA 的）再执行。"
    )

# pywin32 —— 只在 Windows 上可用
import pythoncom  # type: ignore[import-not-found]
import win32com.client  # type: ignore[import-not-found]


@dataclass
class CatiaSessionInfo:
    """CATIA 会话的最小健康快照。"""

    system_configuration: str  # 例如 "V5-6R2021"
    release_number: int         # 例如 30
    service_pack: int           # 例如 4
    active_document_name: Optional[str]
    document_count: int
    caption: str                # CATIA 主窗口标题，做人眼二次校验


@dataclass
class BoxResult:
    """create_box 的结构化返回 —— 既是操作结果，也是「检查证据」。"""

    document_name: str          # 新建 Part 文档名
    body_name: str              # 承载几何的 Body（PartBody）名
    sketch_name: str            # 底面草图名
    pad_name: str               # 拉伸特征名
    length_mm: float
    width_mm: float
    height_mm: float
    update_ok: bool             # 特征树 Update 是否成功（无红叉）
    expected_volume_mm3: float  # 理论体积 = L×W×H
    measured_volume_mm3: Optional[float]  # 从 CATIA 回读的真实体积（失败则 None）
    volume_match: Optional[bool]          # 回读体积是否与理论值吻合（容差内）
    relative_error: Optional[float]       # 相对误差，便于诊断


@dataclass
class PocketResult:
    """在活动 Part 上挖一个矩形槽（去料特征）的结构化证据。"""

    document_name: str                     # 活动 Part 文档名
    body_name: str                         # PartBody 名
    sketch_name: str                       # 槽轮廓草图名
    pocket_name: str                       # Pocket 特征名
    pocket_length_mm: float
    pocket_width_mm: float
    depth_mm: float
    update_ok: bool                        # 特征树 Update 是否成功（无红叉）
    volume_before_mm3: Optional[float]     # 挖槽前体积
    volume_after_mm3: Optional[float]      # 挖槽后体积
    expected_removed_mm3: float            # 理论去料体积 = pl×pw×depth
    measured_removed_mm3: Optional[float]  # 实测去料体积 = before − after
    volume_match: Optional[bool]           # 实测去料是否与理论吻合（容差内）
    relative_error: Optional[float]        # 相对误差
    strategy: str = ""                     # 实际生效的草图平面策略（见 _make_pocket_sketch）
    pocket_from: str = ""                  # "top"=从顶面向下挖；"bottom"=从底面向上挖（退化路径）
    plane_errors: Optional[list] = None    # 各失败策略的报错，定位环境差异用


@dataclass
class FilletResult:
    """给实体加倒圆角（EdgeFillet）的结构化证据。

    与 Pad/Pocket 的关键区别：倒圆角必须**指定拾取哪些边**，而 CATIA 的边引用是
    出了名的脆（拓扑一变引用就失效）。所以这里记录 `strategy` / `target_errors`，
    让「到底拾到了什么」这件事**可被证据回答**，而不是靠猜。
    """

    document_name: str
    body_name: str
    fillet_name: str
    radius_mm: float
    propagation: str                       # tangency / minimal
    strategy: str                          # 实际生效的拾取策略（见 _add_fillet_on_first_target）
    update_ok: bool
    volume_before_mm3: Optional[float]
    volume_after_mm3: Optional[float]
    measured_removed_mm3: Optional[float]  # 实测去料 = before − after
    expected_removed_mm3: Optional[float]  # 理论去料（只有传了长方体尺寸才算得出）
    volume_match: Optional[bool]           # 实测与理论是否吻合
    relative_error: Optional[float]
    objects_filleted: Optional[int] = None # 实际倒了多少条边（长方体应为 12，本身就是证据）
    edge_candidates: Optional[int] = None  # 搜索到的候选总数（含草图线等不可倒角对象）
    target_errors: Optional[list] = None   # 各失败拾取策略的原始 COM 报错


@dataclass
class ChamferResult:
    """倒斜角（Chamfer）的结构化证据 —— 与 Fillet 同一条拾取路径，换个刀。"""

    document_name: str
    body_name: str
    chamfer_name: str
    length_mm: float
    angle_deg: float
    propagation: str
    strategy: str                          # 含实际生效的 mode 枚举值（枚举靠实测钉，不靠猜）
    update_ok: bool
    volume_before_mm3: Optional[float]
    volume_after_mm3: Optional[float]
    measured_removed_mm3: Optional[float]
    expected_removed_mm3: Optional[float]  # 只有 45° + 给了长方体尺寸才算得出精确解
    volume_match: Optional[bool]
    relative_error: Optional[float]
    objects_chamfered: Optional[int] = None
    edge_candidates: Optional[int] = None
    target_errors: Optional[list] = None


@dataclass
class FaceInfo:
    """一个面的可测量身份证 —— 挑面靠它，不靠索引。"""

    index: int                             # 在搜索结果里的序号（仅供追溯，**不作为身份**）
    area_mm2: Optional[float]
    cog_mm: Optional[tuple]


@dataclass
class ShellResult:
    """抽壳（Shell）的结构化证据。

    这是第一个**必须指名道姓拾取某一个面**的特征：倒角可以「全都要」，
    抽壳必须回答「去掉哪个面」。所以这里额外记录被选中面的面积与重心，
    让「它到底选中了哪个面」可被证据回答。
    """

    document_name: str
    body_name: str
    shell_name: str
    thickness_mm: float
    removed_face: Optional[FaceInfo]       # 被去掉的那个面（自证选对了没有）
    face_candidates: int
    strategy: str
    update_ok: bool
    volume_before_mm3: Optional[float]
    volume_after_mm3: Optional[float]
    measured_removed_mm3: Optional[float]
    expected_removed_mm3: Optional[float]
    volume_match: Optional[bool]
    relative_error: Optional[float]
    target_errors: Optional[list] = None


@dataclass
class DraftResult:
    """拔模（Draft）的结构化证据。

    比抽壳更进一步：要同时指定**一组面**（被拔模的侧面）、**一个中性面**
    （拔模时保持不变的基准）和**一个拔模方向**。三者错一个，结果就不是想要的。
    """

    document_name: str
    body_name: str
    draft_name: str
    angle_deg: float
    faces_drafted: int
    neutral_face: Optional[FaceInfo]       # 中性面（应是底面）
    face_candidates: int
    strategy: str
    update_ok: bool
    volume_before_mm3: Optional[float]
    volume_after_mm3: Optional[float]
    measured_delta_mm3: Optional[float]    # after − before，**带符号**（拔模可能加料也可能去料）
    expected_outward_mm3: Optional[float]  # 上大下小（加料）时的理论增量
    expected_inward_mm3: Optional[float]   # 上小下大（去料）时的理论增量（负值）
    matched_direction: Optional[str] = None  # "outward" / "inward" —— 实测符合哪一个
    volume_match: Optional[bool] = None
    relative_error: Optional[float] = None
    target_errors: Optional[list] = None



@dataclass
class ExportResult:
    """安全保存 + STEP 导出回读的结构化证据（对应验证原则 5）。"""

    catpart_path: Optional[str]           # 保存的 .CATPart 路径（未保存则 None）
    catpart_saved: bool                   # CATPart 是否成功落盘
    step_path: str                        # 导出的 STEP 路径
    step_written: bool                    # STEP 文件是否真实存在且非空
    step_size_bytes: Optional[int]        # STEP 文件大小（实时文件验证）
    source_volume_mm3: Optional[float]    # 导出前原实体体积
    reimported_volume_mm3: Optional[float]  # STEP 回读后重新测得的体积
    volume_match: Optional[bool]          # 回读体积是否与原始吻合（容差内）
    relative_error: Optional[float]       # 相对误差
    export_error: Optional[str] = None    # 导出失败时的真实 COM 错误（诊断用）
    format_used: Optional[str] = None     # 实际成功使用的中性格式（如 stp / igs）


@dataclass
class MeasureResult:
    """对**任意已存在实体**的只读测量证据。

    和 BoxResult / PocketResult / FilletResult 的关键区别：那些是「我造完了，
    我自己证明我造对了」；这个是「这东西在这儿，它到底是什么样」——**不需要
    知道它是谁造的**。手工建的模、同事发来的、上一轮 AI 造的，都能测。
    """

    document_name: str
    body_name: str
    volume_mm3: Optional[float]           # 体积
    area_mm2: Optional[float]             # 表面积 —— 与体积互相独立的第二条证据
    cog_mm: Optional[tuple]               # 重心 (x, y, z)，第三条独立证据
    cog_strategy: Optional[str] = None    # GetCOG 是靠哪种调用姿势成功的（自证）
    cog_attempts: Optional[list] = None   # 成功前试过但不行的姿势（预期内，跨机器排障用）
    errors: Optional[list] = None         # 真正导致某项测不出来的原始报错


@dataclass
class InspectResult:
    """文档结构清单 —— 先看清有什么，才谈得上测量。"""

    document_name: str
    open_documents: list                  # 当前打开的全部文档名（找错文档时的自救线索）
    is_part: bool                         # 不是 Part（如 Product/Drawing）时后续测量无意义
    bodies: list                          # [{"name", "shape_count", "shapes": [...]}]
    body_count: int
    active_document_name: Optional[str] = None
    errors: Optional[list] = None


@dataclass
class VariantResult:
    """族里单个变体的结果。**失败的变体也有一条**，不会被悄悄跳过。"""

    index: int                            # 在请求列表里的位置（对得上号才好追）
    name: Optional[str]
    length_mm: float
    width_mm: float
    height_mm: float
    fillet_radius_mm: Optional[float]
    ok: bool                              # 建模 + 验证是否全部通过
    document_name: Optional[str] = None
    saved_path: Optional[str] = None
    expected_volume_mm3: Optional[float] = None
    measured_volume_mm3: Optional[float] = None
    volume_match: Optional[bool] = None
    relative_error: Optional[float] = None
    objects_filleted: Optional[int] = None
    error: Optional[str] = None           # 失败原因（只有失败时有）


@dataclass
class FamilyResult:
    """批量建族的汇总。**汇总本身才是产物** —— 20 份原始返回没人看得完。"""

    requested: int
    succeeded: int
    failed: int
    all_verified: bool                    # 全部变体都建成且体积对上
    elapsed_s: float
    output_dir: Optional[str] = None
    documents_left_open: int = 0          # 没存盘就留在 CATIA 里的数量（会话污染量）
    variants: Optional[list] = None


class CatiaClient:
    """连接到本机上已经打开的 CATIA 会话。

    先启动 CATIA、再运行本客户端 —— Hello World 阶段不做自动拉起。
    """

    def __init__(self) -> None:
        self._app = None

    # ------------------------------------------------------------------
    # 连接
    # ------------------------------------------------------------------
    def connect(self) -> None:
        """通过 ROT 获取已存在的 CATIA.Application。

        如果 CATIA 未启动，pywin32 会尝试拉起 —— 但 CATIA 的启动很慢，
        Hello World 阶段我们要求用户手动先启动，出错立刻抛，不隐藏问题。
        """
        try:
            self._app = win32com.client.GetActiveObject("CATIA.Application")
        except pythoncom.com_error as exc:  # type: ignore[attr-defined]
            raise RuntimeError(
                "无法连接到 CATIA。请确认：\n"
                "  1) CATIA V5 已经在当前 Windows 会话中启动；\n"
                "  2) 你不是以管理员身份运行 Python，而 CATIA 是普通用户身份（或反之）；\n"
                "     COM 的 ROT 不跨完整性等级。\n"
                f"原始错误：{exc}"
            ) from exc

        # 掐掉一类真实的死锁源：文件相关的模态对话框（覆盖确认、只读提示……）。
        # 它们一旦弹出，当前 COM 调用永远不返回，整条单 STA 链路就此卡死。
        # 关掉后 CATIA 会用默认行为继续，不再等人点按钮。
        try:
            self._app.DisplayFileAlerts = False
        except (AttributeError, pythoncom.com_error):  # type: ignore[attr-defined]
            pass  # 老版本没有该属性，不影响主流程

    def ping(self) -> str:
        """最廉价的一次真实 COM 往返，用来确认链路不只是「没卡」，而是真能通。

        故意只读 Caption：不碰文档、不触发任何计算，代价接近于零。
        """
        return str(self._require_app().Caption)

    # ------------------------------------------------------------------
    # 只读探测
    # ------------------------------------------------------------------
    def session_info(self) -> CatiaSessionInfo:
        """读取会话最小信息 —— 全部是只读属性，绝不改动 CATIA 状态。"""
        app = self._require_app()

        sys_conf = str(app.SystemConfiguration.Version)  # 例如 "V5-6R2021"
        release = int(app.SystemConfiguration.Release)
        sp = int(app.SystemConfiguration.ServicePack)

        docs = app.Documents
        doc_count = int(docs.Count)

        active_name: Optional[str] = None
        try:
            active_name = str(app.ActiveDocument.Name)
        except pythoncom.com_error:  # type: ignore[attr-defined]
            # 没有活动文档 —— 完全正常，Hello World 阶段允许
            active_name = None

        caption = str(app.Caption)

        return CatiaSessionInfo(
            system_configuration=sys_conf,
            release_number=release,
            service_pack=sp,
            active_document_name=active_name,
            document_count=doc_count,
            caption=caption,
        )

    # ------------------------------------------------------------------
    # 只读测量族 —— 唯一一组「不靠自己造」也能给结论的能力
    # ------------------------------------------------------------------
    def inspect_document(self, document_name: Optional[str] = None) -> InspectResult:
        """列出文档里有哪些 Body、每个 Body 里有哪些特征。

        为什么先要有它：`measure_body` 需要一个 body 名字，而 AI 无从猜起。
        没有这一步，测量族就只能测「自己刚造的东西」，等于没有解决问题。

        全程只读属性遍历，不碰 Selection、不碰几何，不会改动 CATIA 任何状态。
        """
        app = self._require_app()
        errors: list = []

        open_docs = self._list_document_names(app)
        active_name = None
        try:
            active_name = str(app.ActiveDocument.Name)
        except pythoncom.com_error:  # type: ignore[attr-defined]
            pass

        doc = self._find_document(app, document_name)

        try:
            part = doc.Part
        except (AttributeError, pythoncom.com_error):  # type: ignore[attr-defined]
            # Product / Drawing / STEP 导入件都可能走到这里。如实报告，不假装能测。
            return InspectResult(
                document_name=str(doc.Name),
                open_documents=open_docs,
                is_part=False,
                bodies=[],
                body_count=0,
                active_document_name=active_name,
                errors=["该文档没有 .Part（可能是 Product / Drawing），无法按零件测量。"],
            )

        bodies: list = []
        try:
            collection = part.Bodies
            for i in range(1, int(collection.Count) + 1):
                body = collection.Item(i)
                shapes: list = []
                try:
                    shape_col = body.Shapes
                    for j in range(1, int(shape_col.Count) + 1):
                        shapes.append(str(shape_col.Item(j).Name))
                except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                    errors.append(f"读取 {body.Name} 的特征列表失败: {exc}")
                bodies.append(
                    {"name": str(body.Name), "shape_count": len(shapes), "shapes": shapes}
                )
        except pythoncom.com_error as exc:  # type: ignore[attr-defined]
            errors.append(f"遍历 Bodies 失败: {exc}")

        return InspectResult(
            document_name=str(doc.Name),
            open_documents=open_docs,
            is_part=True,
            bodies=bodies,
            body_count=len(bodies),
            active_document_name=active_name,
            errors=errors or None,
        )

    def measure_body(
        self,
        document_name: Optional[str] = None,
        body_name: Optional[str] = None,
    ) -> MeasureResult:
        """测量一个实体的体积、表面积、重心。三者互相独立。

        ── 为什么不能只测体积 ──
        前面每个写操作都用体积自证，但**体积单独一项验证力有限**：
            · Pad 方向反了 → 体积一模一样，重心却跑到了 z 的另一侧
            · 尺寸写反（40×30 vs 30×40）→ 体积一样，重心不一样
            · 形状完全不同但体积恰好相同 → 表面积会露馅
        体积 + 表面积 + 重心三项同时对上，才算把几何真正钉死。这是**第二条
        独立证据轴**，而不是把同一个数再读一遍。

        ── 单位 ──
        CATIA 的 Measurable 走 SI：体积 m³、面积 m²、坐标 m。
        所以换算是 ×1e9 / ×1e6 / ×1e3。这个换算是否正确，冒烟测试一跑就知道
        （错了会差整整 6 个数量级，不可能看不出来）。
        """
        app = self._require_app()
        errors: list = []

        doc = self._find_document(app, document_name)
        try:
            part = doc.Part
        except (AttributeError, pythoncom.com_error) as exc:  # type: ignore[attr-defined]
            raise RuntimeError(
                f"文档 {doc.Name} 没有 .Part（可能是 Product / Drawing），无法测量。"
            ) from exc

        body = self._find_body(part, body_name)

        try:
            spa = doc.GetWorkbench("SPAWorkbench")
            ref = part.CreateReferenceFromObject(body)
            measurable = spa.GetMeasurable(ref)
        except pythoncom.com_error as exc:  # type: ignore[attr-defined]
            raise RuntimeError(
                f"拿不到 {body.Name} 的 Measurable，测量无法进行。原始错误：{exc}"
            ) from exc

        volume = self._read_si(measurable, "Volume", 1e9, errors)
        area = self._read_si(measurable, "Area", 1e6, errors)
        cog, cog_strategy, cog_attempts = self._read_cog_mm(measurable, app)
        if cog is None:
            # 没拿到重心，那些尝试记录就不再是背景噪声，而是唯一的线索
            errors.extend(cog_attempts)

        return MeasureResult(
            document_name=str(doc.Name),
            body_name=str(body.Name),
            volume_mm3=volume,
            area_mm2=area,
            cog_mm=cog,
            cog_strategy=cog_strategy,
            cog_attempts=cog_attempts or None,
            errors=errors or None,
        )

    # ------------------------------------------------------------------
    @staticmethod
    def _list_document_names(app) -> list:
        try:
            docs = app.Documents
            return [str(docs.Item(i).Name) for i in range(1, int(docs.Count) + 1)]
        except pythoncom.com_error:  # type: ignore[attr-defined]
            return []

    def _find_document(self, app, document_name: Optional[str]):
        """按名字找文档；不给名字就用活动文档。

        找不到时把「现在开着哪些」一并报出来 —— 光说「没找到」等于让人重猜。
        """
        if not document_name:
            try:
                return app.ActiveDocument
            except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                raise RuntimeError("CATIA 里没有活动文档，请先打开一个 Part。") from exc

        docs = app.Documents
        target = str(document_name).strip().lower()
        for i in range(1, int(docs.Count) + 1):
            doc = docs.Item(i)
            if str(doc.Name).strip().lower() == target:
                return doc
        available = self._list_document_names(app)
        raise RuntimeError(
            f"没找到名为 {document_name!r} 的文档。当前打开的是：{available}"
        )

    @staticmethod
    def _find_body(part, body_name: Optional[str]):
        """按名字找 Body；不给名字就用第 1 个（PartBody）。"""
        bodies = part.Bodies
        count = int(bodies.Count)
        if count == 0:
            raise RuntimeError("该 Part 里一个 Body 都没有。")
        if not body_name:
            return bodies.Item(1)

        target = str(body_name).strip().lower()
        names = []
        for i in range(1, count + 1):
            body = bodies.Item(i)
            name = str(body.Name)
            names.append(name)
            if name.strip().lower() == target:
                return body
        raise RuntimeError(f"没找到名为 {body_name!r} 的 Body。该 Part 里有：{names}")

    @staticmethod
    def _read_si(measurable, prop: str, factor: float, errors: list) -> Optional[float]:
        """读一个 SI 单位的标量属性并换算。失败记错误返回 None，不打断其余项。

        分开读、分开记：体积测不出不该连累面积也没有 —— 三条证据里能拿到几条
        就报几条，剩下的说明为什么拿不到。
        """
        try:
            return float(getattr(measurable, prop)) * factor
        except (AttributeError, pythoncom.com_error) as exc:  # type: ignore[attr-defined]
            errors.append(f"{prop}: {exc}")
            return None

    @staticmethod
    def _read_cog_mm(measurable, app=None):
        """读重心，返回 ((x, y, z) mm, 生效的姿势, 试过但不行的姿势列表)。

        ── 坑一：一个「不报错但说谎」的调用姿势 ──
        `GetCOG` 不是返回值，而是**出参数组**（`CATSafeArrayVariant`）。实测：

            byref VARIANT(VT_VARIANT / VT_R8) → 报错 "Objects for SAFEARRAYS
                                                must be sequences ..."
            直接传 list                        → **不报错，但缓冲区一个字节没变**

        第二种才是危险的那个：pywin32 把 list 按值传了进去，CATIA 写回的是那份
        副本。**它没有失败，它在说谎。** 静默返回错误数据比抛异常危险得多——
        后者会停下来，前者会把假数据一路带进结论。

        对策是**哨兵值**：缓冲区不填 0，填一个真实重心绝无可能取到的量级；调用完
        若三个分量还是哨兵值，就判定「没被写回」。不能图省事写成「见到 (0,0,0)
        就算失败」—— 居中建模的零件重心本来就是原点，那样会误杀正确结果。
        哨兵区分的是「没被写」，而不是「值恰好为零」。

        ── 坑二：CATIA 自己的单位不一致（实测）──
        同一个 Measurable 上：

            Volume → m³      （×1e9 → mm³）
            Area   → m²      （×1e6 → mm²）
            GetCOG → **mm**  （×1，不用换算）

        这不是笔误，是实测结果：40×30×20 的块，GetCOG 直接给回 (20, 15, 10)。
        按 SI 惯例乘 1e3 会得到 (20000, 15000, 10000)——差整整 1000 倍。
        所以别推断单位，逐个用已知精确解钉死。

        ── 真正跑通的是 VBA 跳板 ──
        绕开 pywin32 的出参 marshalling：让 CATIA 内部跑一小段 VBScript 把数组
        接住，再当**返回值**递出来。代价是需要脚本执行权限（企业环境可能被锁），
        所以放在最后试。
        """
        sentinel = -1.2345678901e9  # 真实重心不可能取到这个量级
        attempts: list = []

        def _harvest(vals, label):
            """把出参缓冲区收成 mm；没被写回则返回 None 并记原因。"""
            vals = [float(v) for v in (vals or [])]
            if len(vals) < 3:
                attempts.append(f"GetCOG({label}): 只拿到 {len(vals)} 个分量")
                return None
            if all(v == sentinel for v in vals[:3]):
                attempts.append(f"GetCOG({label}): 出参没被写回（缓冲区仍是哨兵值）")
                return None
            return tuple(vals[:3])  # GetCOG 已经是 mm，不换算

        # 姿势 A/B：byref VARIANT。CATSafeArrayVariant 的元素类型是 VARIANT，
        # 所以先试 VT_VARIANT，再退回 VT_R8。（本机两者都不行，留作跨机器兜底）
        variant_kinds = (
            ("VT_ARRAY|VT_BYREF|VT_VARIANT", pythoncom.VT_ARRAY | pythoncom.VT_BYREF | pythoncom.VT_VARIANT),
            ("VT_ARRAY|VT_BYREF|VT_R8", pythoncom.VT_ARRAY | pythoncom.VT_BYREF | pythoncom.VT_R8),
        )
        for label, vt in variant_kinds:
            try:
                arr = win32com.client.VARIANT(vt, [sentinel, sentinel, sentinel])
                measurable.GetCOG(arr)
                got = _harvest(arr.value, label)
                if got is not None:
                    return got, f"VARIANT({label})", attempts
            except (AttributeError, pythoncom.com_error, ValueError, TypeError) as exc:  # type: ignore[attr-defined]
                attempts.append(f"GetCOG({label}): {exc}")

        # 姿势 C：直接传 list。已知会静默说谎，靠哨兵拦住。
        try:
            buf = [sentinel, sentinel, sentinel]
            out = measurable.GetCOG(buf)
            got = _harvest(out if out is not None else buf, "plain list")
            if got is not None:
                return got, "plain list", attempts
        except (AttributeError, pythoncom.com_error, ValueError, TypeError) as exc:  # type: ignore[attr-defined]
            attempts.append(f"GetCOG(plain list): {exc}")

        # 姿势 D：VBA 跳板 —— 本机实测生效的就是这条（lang=2）。
        if app is not None:
            vba = (
                "Function GetCOGArray(m)\n"
                "    Dim c(2)\n"
                "    m.GetCOG c\n"
                "    GetCOGArray = c\n"
                "End Function"
            )
            # 语言枚举在不同版本/文档里对不上号，按序试，成的记进 strategy
            for lang in (2, 0, 1):
                try:
                    out = app.SystemService.Evaluate(vba, lang, "GetCOGArray", [measurable])
                    got = _harvest(list(out) if out is not None else [], f"VBA(lang={lang})")
                    if got is not None:
                        return got, f"SystemService.Evaluate(lang={lang})", attempts
                except (AttributeError, pythoncom.com_error, ValueError, TypeError) as exc:  # type: ignore[attr-defined]
                    attempts.append(f"GetCOG(VBA lang={lang}): {exc}")

        return None, None, attempts

    # ------------------------------------------------------------------
    # 写操作（原生特征）
    # ------------------------------------------------------------------
    def create_box(
        self,
        length_mm: float,
        width_mm: float,
        height_mm: float,
        part_name: Optional[str] = None,
        volume_tolerance: float = 1e-3,
    ) -> BoxResult:
        """新建一个 Part，用「草图矩形 + Pad」原生特征建一个长方体，并回读体积验证。

        这是第一个写操作，刻意做成完整闭环：
            建模（原生特征）→ Update（更新检查）→ 回读体积（几何证据）→ 比对（合格判定）

        参数：
            length_mm/width_mm/height_mm: 长宽高（毫米）
            part_name: 可选，Part 文档命名
            volume_tolerance: 体积相对误差容差，默认 0.1%

        产出保留可编辑特征树：Sketch + Pad，参数与引用完整。
        """
        if min(length_mm, width_mm, height_mm) <= 0:
            raise ValueError("长宽高必须为正数。")

        app = self._require_app()

        # 1) 新建 Part 文档
        part_doc = app.Documents.Add("Part")
        part = part_doc.Part

        # 2) 取 PartBody 与 XY 基准面
        body = part.Bodies.Item(1)  # 默认的 PartBody
        xy_plane = part.OriginElements.PlaneXY

        # 3) 在 XY 面上建草图，画一个矩形（4 条首尾相接的直线 → 闭合轮廓）
        sketch = body.Sketches.Add(xy_plane)
        factory_2d = sketch.OpenEdition()
        length = float(length_mm)
        width = float(width_mm)
        # 矩形四角：(0,0)-(L,0)-(L,W)-(0,W)
        factory_2d.CreateLine(0.0, 0.0, length, 0.0)
        factory_2d.CreateLine(length, 0.0, length, width)
        factory_2d.CreateLine(length, width, 0.0, width)
        factory_2d.CreateLine(0.0, width, 0.0, 0.0)
        sketch.CloseEdition()

        # 4) 拉伸成 Pad（原生 Part Design 特征）
        part.InWorkObject = body
        shape_factory = part.ShapeFactory
        pad = shape_factory.AddNewPad(sketch, float(height_mm))

        if part_name:
            # Part.Name 在部分 CATIA 版本/上下文中只读，改不动就跳过，不打断建模
            try:
                part.Name = str(part_name)
            except pythoncom.com_error:  # type: ignore[attr-defined]
                pass

        # 5) 更新检查 —— 命令返回 ≠ 几何合格，必须 Update 并捕获失败
        update_ok = True
        try:
            part.Update()
        except pythoncom.com_error:  # type: ignore[attr-defined]
            update_ok = False

        # 6) 回读体积证据
        expected = length * width * float(height_mm)
        measured = self._measure_volume_mm3(part_doc, part, body)
        match: Optional[bool] = None
        rel_err: Optional[float] = None
        if measured is not None and expected > 0:
            rel_err = abs(measured - expected) / expected
            match = rel_err <= volume_tolerance

        return BoxResult(
            document_name=str(part_doc.Name),
            body_name=str(body.Name),
            sketch_name=str(sketch.Name),
            pad_name=str(pad.Name),
            length_mm=length,
            width_mm=width,
            height_mm=float(height_mm),
            update_ok=update_ok,
            expected_volume_mm3=expected,
            measured_volume_mm3=measured,
            volume_match=match,
            relative_error=rel_err,
        )

    # ------------------------------------------------------------------
    # 写操作（去料特征）
    # ------------------------------------------------------------------
    def add_pocket(
        self,
        pocket_length_mm: float,
        pocket_width_mm: float,
        depth_mm: float,
        at_height_mm: float,
        center_x_mm: Optional[float] = None,
        center_y_mm: Optional[float] = None,
        volume_tolerance: float = 1e-3,
    ) -> PocketResult:
        """在当前活动 Part 的顶面挖一个矩形槽（Pocket 去料特征），并用体积差验证。

        机制与 create_box 同源、同样稳健（只用平面 + 草图，不做脆弱的面/边拾取）：
            在 XY 面上方 at_height 处建一个偏移平面（与长方体顶面共面）
            → 在其上画矩形 → Pocket 向下（-Z）挖 depth → Update → 体积差比对

        参数：
            pocket_length_mm/pocket_width_mm: 槽的长宽（毫米）
            depth_mm:      挖深（毫米，应 < 长方体高度，避免挖穿）
            at_height_mm:  顶面所在的 Z 高度（= 目标长方体的高度）
            center_x_mm/center_y_mm: 槽中心（默认取长方体中心，需与 create_box 的 L/2、W/2 对齐）
            volume_tolerance: 去料体积相对误差容差，默认 0.1%

        产出保留可编辑特征树：偏移平面 + Sketch + Pocket。
        """
        if min(pocket_length_mm, pocket_width_mm, depth_mm) <= 0:
            raise ValueError("槽的长宽深必须为正数。")
        if depth_mm >= at_height_mm:
            raise ValueError(
                f"挖深 {depth_mm} 应小于顶面高度 {at_height_mm}，否则会挖穿（本阶段只做盲槽）。"
            )

        app = self._require_app()

        part_doc = app.ActiveDocument
        part = part_doc.Part
        body = part.Bodies.Item(1)

        # 0) 挖槽前体积证据
        vol_before = self._measure_volume_mm3(part_doc, part, body)

        # 1) 拿一个可画草图的平面并建草图（偏移平面优先，拿不到就退到 PlaneXY 从底面挖）
        sketch, strategy, from_top, plane_errors = self._make_pocket_sketch(
            part, body, float(at_height_mm)
        )

        # 2) 画居中矩形（草图 H/V 轴与全局 X/Y 对齐）
        pl = float(pocket_length_mm)
        pw = float(pocket_width_mm)
        cx = float(center_x_mm) if center_x_mm is not None else pl  # 默认与长方体中心对齐由调用方保证
        cy = float(center_y_mm) if center_y_mm is not None else pw
        x0, x1 = cx - pl / 2.0, cx + pl / 2.0
        y0, y1 = cy - pw / 2.0, cy + pw / 2.0

        factory_2d = sketch.OpenEdition()
        factory_2d.CreateLine(x0, y0, x1, y0)
        factory_2d.CreateLine(x1, y0, x1, y1)
        factory_2d.CreateLine(x1, y1, x0, y1)
        factory_2d.CreateLine(x0, y1, x0, y0)
        sketch.CloseEdition()

        # 3) Pocket 去料
        part.InWorkObject = body
        shape_factory = part.ShapeFactory
        pocket = shape_factory.AddNewPocket(sketch, float(depth_mm))

        # 顶面偏移平面：法向 +Z，Pocket 默认沿 -Z 向下挖，方向天然正确。
        # 退化到 PlaneXY 时草图在 Z=0，必须把挖料方向翻成 +Z 才能吃到料。
        if not from_top:
            try:
                pocket.DirectionOrientation = 1  # catInverseOrientation
            except (AttributeError, pythoncom.com_error):  # type: ignore[attr-defined]
                plane_errors.append("DirectionOrientation 反向失败，去料方向可能背离实体")

        # 4) 更新检查
        update_ok = True
        try:
            part.Update()
        except pythoncom.com_error:  # type: ignore[attr-defined]
            update_ok = False

        # 5) 挖槽后体积 + 去料量比对
        vol_after = self._measure_volume_mm3(part_doc, part, body)
        expected_removed = pl * pw * float(depth_mm)
        measured_removed: Optional[float] = None
        match: Optional[bool] = None
        rel_err: Optional[float] = None
        if vol_before is not None and vol_after is not None:
            measured_removed = vol_before - vol_after
            if expected_removed > 0:
                rel_err = abs(measured_removed - expected_removed) / expected_removed
                match = rel_err <= volume_tolerance

        return PocketResult(
            document_name=str(part_doc.Name),
            body_name=str(body.Name),
            sketch_name=str(sketch.Name),
            pocket_name=str(pocket.Name),
            pocket_length_mm=pl,
            pocket_width_mm=pw,
            depth_mm=float(depth_mm),
            update_ok=update_ok,
            volume_before_mm3=vol_before,
            volume_after_mm3=vol_after,
            expected_removed_mm3=expected_removed,
            measured_removed_mm3=measured_removed,
            volume_match=match,
            relative_error=rel_err,
            strategy=strategy,
            pocket_from="top" if from_top else "bottom",
            plane_errors=plane_errors or None,
        )

    def _make_pocket_sketch(self, part, body, at_height_mm: float):
        """在 PartBody 上建出槽轮廓草图，按稳健度依次降级，并记录每次失败原因。

        为什么要降级链：`AddNewPlaneOffset` 造出的平面必须先被挂进某个容器并 Update
        才有真实几何，否则 `Sketches.Add` 直接报 “The method Add failed”。而挂进
        PartBody 只在 Part 开了 Hybrid Design 时才成立——这台机器上就不成立。

        返回 (sketch, strategy, from_top, errors)
        """
        errors: list = []

        # 策略 A：几何图形集（HybridBody）——任何 Part 都有，不依赖 Hybrid Design 开关
        try:
            hsf = part.HybridShapeFactory
            ref_xy = part.CreateReferenceFromObject(part.OriginElements.PlaneXY)
            hybrid_body = part.HybridBodies.Add()
            try:
                hybrid_body.Name = "PocketPlanes"
            except pythoncom.com_error:  # type: ignore[attr-defined]
                pass
            plane = hsf.AddNewPlaneOffset(ref_xy, at_height_mm, False)
            hybrid_body.AppendHybridShape(plane)
            part.InWorkObject = hybrid_body
            part.Update()
            part.InWorkObject = body
            for as_ref in (False, True):
                candidate = part.CreateReferenceFromObject(plane) if as_ref else plane
                try:
                    return (
                        body.Sketches.Add(candidate),
                        f"hybrid_body_offset_plane(ref={as_ref})",
                        True,
                        errors,
                    )
                except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                    errors.append(f"A/ref={as_ref}: {exc}")
        except pythoncom.com_error as exc:  # type: ignore[attr-defined]
            errors.append(f"A/setup: {exc}")

        # 策略 B：退化——直接用 PlaneXY（create_box 已证明可用），改由 Pocket 反向挖
        errors.append("降级到 PlaneXY，改从底面向上挖（体积证据等价，位置在底部）")
        part.InWorkObject = body
        return (
            body.Sketches.Add(part.OriginElements.PlaneXY),
            "origin_plane_xy_reversed",
            False,
            errors,
        )

    # ------------------------------------------------------------------
    # 写操作（修饰特征）—— 第一次涉及「边/面拾取」
    # ------------------------------------------------------------------
    def add_fillet(
        self,
        radius_mm: float,
        box_length_mm: Optional[float] = None,
        box_width_mm: Optional[float] = None,
        box_height_mm: Optional[float] = None,
        propagation: str = "tangency",
        volume_tolerance: float = 1e-3,
    ) -> FilletResult:
        """给当前活动 Part 的实体加恒定半径倒圆角，并用体积差验证。

        与 Pad / Pocket 的本质区别：
            前两者只需要「平面 + 草图」，输入是**参数**；倒圆角必须先**拾取到边**，
            输入是**几何引用**。CATIA 的边引用（BRep 名字如 `REdge:(Edge:(Face:(Brp:(Pad.1;2)...`）
            绑定具体拓扑，改个尺寸就可能失效 —— 这是所有 CAD 自动化最经典的脆点。

        本实现刻意**绕开逐条边的 BRep 名字**：改为把整个特征/实体作为拾取对象，
        让 CATIA 自己展开成它的全部边。这样不写死任何拓扑名字，
        代价是只能「整体倒角」，不能挑单条边（挑边留到确有需求时再按证据推进）。

        参数：
            radius_mm:      圆角半径（毫米，>0，且直径须小于最小边长）
            box_*_mm:       可选。传入被倒角长方体的长宽高，就能算出**理论去料体积**做硬验证；
                            不传则只报体积前后与实测去料量（弱验证）。
            propagation:    "tangency"=沿相切面传播（默认）；"minimal"=最小传播
            volume_tolerance: 体积相对误差容差，默认 0.1%

        返回：既是结果也是证据，重点看 update_ok / volume_match / strategy。
        """
        if radius_mm <= 0:
            raise ValueError("圆角半径必须为正数。")

        dims = [box_length_mm, box_width_mm, box_height_mm]
        if all(d is not None for d in dims):
            min_dim = min(float(d) for d in dims)  # type: ignore[arg-type]
            if 2 * float(radius_mm) >= min_dim:
                raise ValueError(
                    f"圆角直径 {2 * radius_mm} 必须小于最小边长 {min_dim}，否则几何自相交。"
                )

        app = self._require_app()
        part_doc = app.ActiveDocument
        part = part_doc.Part
        body = part.Bodies.Item(1)

        vol_before = self._measure_volume_mm3(part_doc, part, body)

        part.InWorkObject = body
        (
            fillet,
            strategy,
            edge_count,
            candidate_count,
            target_errors,
        ) = self._add_fillet_on_edges(part_doc, part, float(radius_mm), propagation)

        update_ok = True
        try:
            part.Update()
        except pythoncom.com_error:  # type: ignore[attr-defined]
            update_ok = False

        vol_after = self._measure_volume_mm3(part_doc, part, body)
        measured_removed: Optional[float] = None
        if vol_before is not None and vol_after is not None:
            measured_removed = vol_before - vol_after

        expected_removed: Optional[float] = None
        if all(d is not None for d in dims):
            expected_removed = self._box_fillet_removed_mm3(
                float(box_length_mm),  # type: ignore[arg-type]
                float(box_width_mm),  # type: ignore[arg-type]
                float(box_height_mm),  # type: ignore[arg-type]
                float(radius_mm),
            )

        match: Optional[bool] = None
        rel_err: Optional[float] = None
        if measured_removed is not None and expected_removed:
            rel_err = abs(measured_removed - expected_removed) / expected_removed
            match = rel_err <= volume_tolerance

        return FilletResult(
            document_name=str(part_doc.Name),
            body_name=str(body.Name),
            fillet_name=str(fillet.Name),
            radius_mm=float(radius_mm),
            propagation=propagation,
            strategy=strategy,
            update_ok=update_ok,
            volume_before_mm3=vol_before,
            volume_after_mm3=vol_after,
            measured_removed_mm3=measured_removed,
            expected_removed_mm3=expected_removed,
            volume_match=match,
            relative_error=rel_err,
            objects_filleted=edge_count,
            edge_candidates=candidate_count,
            target_errors=target_errors or None,
        )

    # ------------------------------------------------------------------
    # 修饰特征三兄弟：Chamfer / Shell / Draft
    #
    # 这三个是「抄作业」——最难的一关（怎么拿到合法的几何引用）已经在 Fillet
    # 里趟平了。但抄作业也分三档，难度是递增的，而且每一档新增的东西都不一样：
    #
    #   Chamfer  与 Fillet 同一条边拾取路径，只是换了个 ShapeFactory 方法。
    #            **唯一的新问题是枚举值**（CatChamferMode 的取值各处说法不一）。
    #   Shell    第一次必须「指名道姓挑某一个面」——倒角可以全都要，抽壳不行。
    #            新增的是**挑面的判据**。
    #   Draft    要同时给一组面、一个中性面、一个方向。新增的是**多引用协同**。
    #
    # 三者都沿用同一条纪律：不确定的地方用策略链试，把实际生效的写进 strategy，
    # 再用精确解体积把结果钉死 —— 猜错枚举不会静默通过，因为体积会对不上。
    # ------------------------------------------------------------------
    def add_chamfer(
        self,
        length_mm: float,
        angle_deg: float = 45.0,
        box_length_mm: Optional[float] = None,
        box_width_mm: Optional[float] = None,
        box_height_mm: Optional[float] = None,
        propagation: str = "tangency",
        volume_tolerance: float = 1e-3,
    ) -> ChamferResult:
        """给当前活动 Part 的实体全部边倒斜角，并用体积差验证。

        ── 为什么它比 Fillet 便宜得多 ──
        边拾取那一关（BRep 引用极脆、手写必失效）已经在 `_search_edge_references`
        里解决了，这里直接复用。剩下的唯一新问题是 **CatChamferMode 的枚举值**：
        「长度+角度」模式在不同资料里被写成 0 / 1 / 2 都有。

        ── 枚举猜错会不会静默通过 ──
        不会。如果误用成「两个长度」模式，第二个参数 45 会被当成 45mm 的第二条边长，
        在 40×30×20 的体上直接吃穿 —— 要么 CATIA 报错，要么体积对不上。
        **精确解验证在这里的作用不是"锦上添花"，而是枚举的判据本身。**

        参数：
            length_mm:  斜角第一条边长（毫米，>0）
            angle_deg:  斜角角度（默认 45°）。**只有 45° 有精确解**，其余角度
                        的八个角块是不规则多面体，暂不做硬验证（宁可不验，
                        也不用近似值假装验过了）。
            box_*_mm:   可选。给了才算得出理论去料体积。
            propagation: "tangency"（默认）/ "minimal"
        """
        if length_mm <= 0:
            raise ValueError("斜角边长必须为正数。")
        if not 0 < angle_deg < 90:
            raise ValueError(f"斜角角度必须在 (0, 90) 之间，收到 {angle_deg}。")

        dims = [box_length_mm, box_width_mm, box_height_mm]
        leg2 = float(length_mm) * math.tan(math.radians(angle_deg))
        if all(d is not None for d in dims):
            min_dim = min(float(d) for d in dims)  # type: ignore[arg-type]
            if 2 * max(float(length_mm), leg2) >= min_dim:
                raise ValueError(
                    f"斜角尺寸过大 —— 两条边长 {length_mm:.3f} / {leg2:.3f} 中较大者的 2 倍"
                    f"必须小于最短边 {min_dim}，否则几何自相交。"
                )

        app = self._require_app()
        part_doc = app.ActiveDocument
        part = part_doc.Part
        body = part.Bodies.Item(1)

        vol_before = self._measure_volume_mm3(part_doc, part, body)

        # 理论去料先算 —— 它不只是「事后复核」，下面的策略链要拿它当判据。
        expected_removed: Optional[float] = None
        if all(d is not None for d in dims) and abs(angle_deg - 45.0) < 1e-9:
            expected_removed = self._box_chamfer_removed_mm3(
                float(box_length_mm),  # type: ignore[arg-type]
                float(box_width_mm),   # type: ignore[arg-type]
                float(box_height_mm),  # type: ignore[arg-type]
                float(length_mm),
            )

        part.InWorkObject = body
        # catTangencyChamfer = 1 / catMinimalChamfer = 2
        prop = 1 if propagation == "tangency" else 2
        errors: list = []
        chosen = None

        for cand in self._chamfer_candidates(part.ShapeFactory, errors):
            built = self._try_build_chamfer(
                part_doc, part, cand, float(length_mm), float(angle_deg), prop, errors
            )
            if built is None:
                continue
            chamfer, strategy, added, candidate_count = built

            update_ok = True
            try:
                part.Update()
            except pythoncom.com_error:  # type: ignore[attr-defined]
                update_ok = False

            vol_after = self._measure_volume_mm3(part_doc, part, body)
            removed: Optional[float] = None
            if vol_before is not None and vol_after is not None:
                removed = vol_before - vol_after

            match: Optional[bool] = None
            rel_err: Optional[float] = None
            if removed is not None and expected_removed:
                rel_err = abs(removed - expected_removed) / expected_removed
                match = rel_err <= volume_tolerance

            if expected_removed is None or (update_ok and match):
                chosen = (
                    chamfer, strategy, added, candidate_count,
                    update_ok, vol_after, removed, match, rel_err,
                )
                break

            # 调用被接受了，但做出来的不是想要的东西 —— 撤销，换下一个组合。
            # 这一步是整条链的关键：**枚举猜错在这里被体积当场抓住**，
            # 而不是留一个"没报错但形状不对"的特征混过去。
            errors.append(
                f"{strategy}: update_ok={update_ok}，去料 {removed} 与理论 "
                f"{expected_removed} 不符，已撤销该特征"
            )
            self._delete_feature(part_doc, part, chamfer)

        if chosen is None:
            raise RuntimeError(
                "倒斜角的所有方法/参数个数/枚举组合都没能做出正确结果。\n"
                "各次尝试：\n  " + "\n  ".join(str(e) for e in errors)
            )

        (
            chamfer, strategy, added, candidate_count,
            update_ok, vol_after, measured_removed, match, rel_err,
        ) = chosen

        return ChamferResult(
            document_name=str(part_doc.Name),
            body_name=str(body.Name),
            chamfer_name=str(chamfer.Name),
            length_mm=float(length_mm),
            angle_deg=float(angle_deg),
            propagation=propagation,
            strategy=strategy,
            update_ok=update_ok,
            volume_before_mm3=vol_before,
            volume_after_mm3=vol_after,
            measured_removed_mm3=measured_removed,
            expected_removed_mm3=expected_removed,
            volume_match=match,
            relative_error=rel_err,
            objects_chamfered=added,
            edge_candidates=candidate_count,
            target_errors=errors or None,
        )

    def add_shell(
        self,
        thickness_mm: float,
        box_length_mm: Optional[float] = None,
        box_width_mm: Optional[float] = None,
        box_height_mm: Optional[float] = None,
        volume_tolerance: float = 1e-3,
    ) -> ShellResult:
        """把当前活动 Part 的实体抽成薄壳（去掉顶面开口），并用体积差验证。

        ── 这一步真正新增的东西：挑面的判据 ──
        倒角可以「把搜到的边全都要」，抽壳不行 —— 必须回答「去掉哪一个面」。
        而 `Selection.Search` 回来的顺序**不保证**：换个模型、换个 CATIA 版本，
        第 3 个面就可能不再是顶面。所以这里**不按索引挑面，按测量挑面**：
        逐个面读出重心，取重心 Z 最大的那个 —— 这条判据换任何模型都成立。

        代价是每个面要多一次 COM 往返（重心还得走 VBA 蹦床）。六个面而已，值。

        ── 自证 ──
        返回里带上被选中面的**面积和重心**。如果哪天它挑错了面，
        看一眼 `removed_face` 就知道，不用去 CATIA 里肉眼找。

        参数：
            thickness_mm: 壁厚（毫米，>0，向内偏移）
            box_*_mm:     可选。给了才算得出理论去料体积
                          = 内腔 (L−2t)(W−2t)(H−t)。
        """
        if thickness_mm <= 0:
            raise ValueError("壁厚必须为正数。")

        t = float(thickness_mm)
        dims = [box_length_mm, box_width_mm, box_height_mm]
        if all(d is not None for d in dims):
            L, W, H = (float(d) for d in dims)  # type: ignore[arg-type]
            if 2 * t >= min(L, W) or t >= H:
                raise ValueError(
                    f"壁厚 {t} 过大 —— 2t 必须小于 min(长,宽)={min(L, W)}，"
                    f"且 t 必须小于高 {H}，否则抽不出内腔。"
                )

        app = self._require_app()
        part_doc = app.ActiveDocument
        part = part_doc.Part
        body = part.Bodies.Item(1)

        vol_before = self._measure_volume_mm3(part_doc, part, body)

        errors: list = []
        faces, query = self._face_table(part_doc, app, errors)
        if not faces:
            raise RuntimeError(
                "没能拾取到任何可测量的面，抽壳无法进行。\n"
                "各次尝试的原始报错：\n  " + "\n  ".join(str(e) for e in errors)
            )
        top = max(faces, key=lambda f: f["cog_mm"][2])

        part.InWorkObject = body
        shell, strategy, shell_errors = self._add_shell_on_face(part, top["ref"], t)
        errors.extend(shell_errors)

        update_ok = True
        try:
            part.Update()
        except pythoncom.com_error:  # type: ignore[attr-defined]
            update_ok = False

        vol_after = self._measure_volume_mm3(part_doc, part, body)
        measured_removed: Optional[float] = None
        if vol_before is not None and vol_after is not None:
            measured_removed = vol_before - vol_after

        expected_removed: Optional[float] = None
        if all(d is not None for d in dims):
            L, W, H = (float(d) for d in dims)  # type: ignore[arg-type]
            expected_removed = (L - 2 * t) * (W - 2 * t) * (H - t)

        match: Optional[bool] = None
        rel_err: Optional[float] = None
        if measured_removed is not None and expected_removed:
            rel_err = abs(measured_removed - expected_removed) / expected_removed
            match = rel_err <= volume_tolerance

        return ShellResult(
            document_name=str(part_doc.Name),
            body_name=str(body.Name),
            shell_name=str(shell.Name),
            thickness_mm=t,
            removed_face=FaceInfo(
                index=top["index"], area_mm2=top["area_mm2"], cog_mm=top["cog_mm"]
            ),
            face_candidates=len(faces),
            strategy=f"{query}/{strategy}",
            update_ok=update_ok,
            volume_before_mm3=vol_before,
            volume_after_mm3=vol_after,
            measured_removed_mm3=measured_removed,
            expected_removed_mm3=expected_removed,
            volume_match=match,
            relative_error=rel_err,
            target_errors=errors or None,
        )

    def add_draft(
        self,
        angle_deg: float,
        box_length_mm: Optional[float] = None,
        box_width_mm: Optional[float] = None,
        box_height_mm: Optional[float] = None,
        volume_tolerance: float = 1e-3,
    ) -> DraftResult:
        """给当前活动 Part 的四个侧面加拔模斜度（底面为中性面），并用体积差验证。

        ── 这一步真正新增的东西：多引用协同 ──
        前面所有特征最多只要一组同类引用；拔模要**三样东西同时对**：
            被拔模的面（一组侧面）、中性面（拔模时保持不变的基准，这里取底面）、
            拔模方向（这里取 Z 轴，用 PlaneXY 的法向表达）。
        错一个，结果就不是想要的形状 —— 而且**未必报错**，可能安静地做出个别的东西。

        ── 所以这里的验证要比前面更狠 ──
        拔模到底是加料（上大下小）还是去料（上小下大），取决于角度符号和
        CATIA 的方向约定 —— 这是**实测才能定**的事。所以同时算出两种情形的
        精确解（棱台体积，Prismatoid 公式），看实测符合哪一个，把结论写进
        `matched_direction`。两个候选值大小并不相等，所以"二选一"仍是硬验证，
        不是放水。

        参数：
            angle_deg: 拔模角（度，0 < a < 45）
            box_*_mm:  可选。给了才算得出理论体积变化。
        """
        if not 0 < angle_deg < 45:
            raise ValueError(f"拔模角必须在 (0, 45) 之间，收到 {angle_deg}。")

        app = self._require_app()
        part_doc = app.ActiveDocument
        part = part_doc.Part
        body = part.Bodies.Item(1)

        vol_before = self._measure_volume_mm3(part_doc, part, body)

        errors: list = []
        faces, query = self._face_table(part_doc, app, errors)
        if len(faces) < 3:
            raise RuntimeError(
                f"只拾到 {len(faces)} 个可测量的面，不足以区分顶/底/侧面，拔模无法进行。\n"
                "各次尝试的原始报错：\n  " + "\n  ".join(str(e) for e in errors)
            )

        ordered = sorted(faces, key=lambda f: f["cog_mm"][2])
        bottom = ordered[0]
        top = ordered[-1]
        sides = ordered[1:-1]
        if not sides:
            raise RuntimeError("除顶底面外没有侧面可拔模。")

        part.InWorkObject = body
        draft, strategy, draft_errors = self._add_draft_on_faces(
            part_doc, part, [f["ref"] for f in sides], bottom["ref"], float(angle_deg)
        )
        errors.extend(draft_errors)

        update_ok = True
        try:
            part.Update()
        except pythoncom.com_error:  # type: ignore[attr-defined]
            update_ok = False

        vol_after = self._measure_volume_mm3(part_doc, part, body)
        delta: Optional[float] = None
        if vol_before is not None and vol_after is not None:
            delta = vol_after - vol_before

        exp_out: Optional[float] = None
        exp_in: Optional[float] = None
        dims = [box_length_mm, box_width_mm, box_height_mm]
        if all(d is not None for d in dims):
            exp_out = self._box_draft_delta_mm3(
                float(box_length_mm), float(box_width_mm),  # type: ignore[arg-type]
                float(box_height_mm), float(angle_deg), outward=True,  # type: ignore[arg-type]
            )
            exp_in = self._box_draft_delta_mm3(
                float(box_length_mm), float(box_width_mm),  # type: ignore[arg-type]
                float(box_height_mm), float(angle_deg), outward=False,  # type: ignore[arg-type]
            )

        matched: Optional[str] = None
        match: Optional[bool] = None
        rel_err: Optional[float] = None
        if delta is not None and exp_out is not None and exp_in is not None:
            best = None
            for label, expected in (("outward", exp_out), ("inward", exp_in)):
                if not expected:
                    continue
                err = abs(delta - expected) / abs(expected)
                if best is None or err < best[1]:
                    best = (label, err)
            if best is not None:
                matched, rel_err = best
                match = rel_err <= volume_tolerance
                if not match:
                    matched = None  # 两个都对不上，就别声称符合哪一个

        return DraftResult(
            document_name=str(part_doc.Name),
            body_name=str(body.Name),
            draft_name=str(draft.Name),
            angle_deg=float(angle_deg),
            faces_drafted=len(sides),
            neutral_face=FaceInfo(
                index=bottom["index"], area_mm2=bottom["area_mm2"], cog_mm=bottom["cog_mm"]
            ),
            face_candidates=len(faces),
            strategy=f"{query}/{strategy}/sides={len(sides)}of{len(faces)}"
                     f"/top_z={top['cog_mm'][2]:.3f}",
            update_ok=update_ok,
            volume_before_mm3=vol_before,
            volume_after_mm3=vol_after,
            measured_delta_mm3=delta,
            expected_outward_mm3=exp_out,
            expected_inward_mm3=exp_in,
            matched_direction=matched,
            volume_match=match,
            relative_error=rel_err,
            target_errors=errors or None,
        )

    # ------------------------------------------------------------------
    # 批量：一次生成一族变体
    # ------------------------------------------------------------------
    MAX_FAMILY_VARIANTS = 50

    def create_box_family(
        self,
        variants: list,
        output_dir: Optional[str] = None,
        volume_tolerance: float = 1e-3,
    ) -> FamilyResult:
        """按参数表一次建出一族长方体（可选倒角），每个变体各自验证。

        ── 这一步换的是量级，不是功能 ──
        前面九个工具单次调用都可靠了，但「AI 帮我设计」和「AI 帮我点鼠标」的差别，
        恰恰在于能不能**一次生成并验证 20 个变体**。人做 20 遍会累、会漏、会在第
        13 个上手滑；这里做 20 遍和做 1 遍的心智负担一样。

        ── 批量特有的三个设计点 ──
        1. **参数错 → 整批不动；运行时错 → 继续往下**。这两类失败的正确反应相反：
           规格写错（负数尺寸、2r ≥ 最短边）在碰 CATIA 之前就能查出来，那就一个
           都别建 —— 建了 6 个再报错，等于留下 6 份要手工收拾的垃圾，而且**批量
           的每一步都不可回滚**。反过来，第 7 个在 CATIA 里跑挂了（几何求解失败
           之类），不该让第 8~20 个白等，那种失败必须就地记下继续走。
        2. **失败的变体也要有一条记录**。悄悄跳过 = 汇总里 20 变 19，没人会发现
           少了哪个。所以 `variants` 里条数恒等于请求数，靠 `ok` 区分。
        3. **汇总才是产物**。20 份原始返回没人看得完，AI 也会被淹没。所以先给
           `succeeded / failed / all_verified`，明细放后面备查。

        参数：
            variants: [{"length_mm", "width_mm", "height_mm",
                        "fillet_radius_mm"(可选), "name"(可选)}, ...]
            output_dir: 给了就每个变体存盘后**关掉**；不给就全部留在 CATIA 里
            volume_tolerance: 单个变体的体积相对误差容差
        """
        if not isinstance(variants, list) or not variants:
            raise ValueError("variants 必须是非空列表。")
        if len(variants) > self.MAX_FAMILY_VARIANTS:
            raise ValueError(
                f"一次最多 {self.MAX_FAMILY_VARIANTS} 个变体，收到 {len(variants)} 个。\n"
                "这个上限是保护单 STA 链路的：批量期间整条链路被独占，"
                "太长的批次会让健康检查和其它调用一直排队。请分批提交。"
            )

        specs = [self._parse_variant(i, v) for i, v in enumerate(variants)]

        out_dir = None
        if output_dir:
            out_dir = os.path.abspath(os.path.expanduser(str(output_dir)))
            os.makedirs(out_dir, exist_ok=True)

        app = self._require_app()
        started = time.monotonic()
        results: list = []
        left_open = 0

        for spec in specs:
            idx, name, L, W, H, r = spec
            try:
                box = self.create_box(
                    length_mm=L, width_mm=W, height_mm=H,
                    part_name=name, volume_tolerance=volume_tolerance,
                )
                expected = L * W * H
                measured = box.measured_volume_mm3
                ok = bool(box.update_ok and box.volume_match)
                filleted = None

                if r:
                    fil = self.add_fillet(
                        radius_mm=r, box_length_mm=L, box_width_mm=W,
                        box_height_mm=H, volume_tolerance=volume_tolerance,
                    )
                    expected -= self._box_fillet_removed_mm3(L, W, H, r)
                    measured = fil.volume_after_mm3
                    filleted = fil.objects_filleted
                    ok = ok and bool(fil.update_ok and fil.volume_match)

                rel_err = None
                match = None
                if measured is not None and expected > 0:
                    rel_err = abs(measured - expected) / expected
                    match = rel_err <= volume_tolerance
                    ok = ok and match

                saved_path = None
                if out_dir:
                    saved_path = self._save_and_close(app, box.document_name, out_dir, name, idx)
                    if saved_path is None:
                        ok = False
                else:
                    left_open += 1

                results.append(
                    VariantResult(
                        index=idx, name=name,
                        length_mm=L, width_mm=W, height_mm=H, fillet_radius_mm=r,
                        ok=ok,
                        document_name=box.document_name,
                        saved_path=saved_path,
                        expected_volume_mm3=expected,
                        measured_volume_mm3=measured,
                        volume_match=match,
                        relative_error=rel_err,
                        objects_filleted=filleted,
                        error=None if ok else "验证未通过，见 volume_match / relative_error",
                    )
                )
            except Exception as exc:  # noqa: BLE001 —— 单个变体失败不拖垮整批
                results.append(
                    VariantResult(
                        index=idx, name=name,
                        length_mm=L, width_mm=W, height_mm=H, fillet_radius_mm=r,
                        ok=False, error=str(exc),
                    )
                )

        succeeded = sum(1 for v in results if v.ok)
        return FamilyResult(
            requested=len(specs),
            succeeded=succeeded,
            failed=len(results) - succeeded,
            all_verified=succeeded == len(results),
            elapsed_s=round(time.monotonic() - started, 3),
            output_dir=out_dir,
            documents_left_open=left_open,
            variants=results,
        )

    # ------------------------------------------------------------------
    @staticmethod
    def _parse_variant(index: int, spec) -> tuple:
        """把一条变体规格校验成 (index, name, L, W, H, r)。不合法立刻抛。

        错误信息一律带上是第几个 —— 一张 20 行的表里说「尺寸必须为正」而不说
        哪一行，等于没说。
        """
        where = f"第 {index + 1} 个变体"
        if not isinstance(spec, dict):
            raise ValueError(f"{where}：必须是对象，收到 {type(spec).__name__}。")
        try:
            L = float(spec["length_mm"])
            W = float(spec["width_mm"])
            H = float(spec["height_mm"])
        except KeyError as exc:
            raise ValueError(f"{where}：缺少必填项 {exc.args[0]!r}。") from exc
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{where}：长宽高必须是数字。原始错误：{exc}") from exc

        if min(L, W, H) <= 0:
            raise ValueError(f"{where}：长宽高必须为正数，收到 {(L, W, H)}。")

        r = spec.get("fillet_radius_mm")
        if r is not None:
            r = float(r)
            if r <= 0:
                raise ValueError(f"{where}：倒角半径必须为正数，收到 {r}。")
            if 2 * r >= min(L, W, H):
                raise ValueError(
                    f"{where}：倒角半径 {r} 过大 —— 2r 必须小于最短边 {min(L, W, H)}，"
                    "否则圆角会吃穿整个截面。"
                )

        name = spec.get("name")
        return index, (str(name) if name else None), L, W, H, r

    def _save_and_close(self, app, document_name, out_dir: str, name, index: int):
        """把变体存盘并关闭；成功返回路径，失败返回 None。

        存盘不只是留档 —— 它同时解决了批量最现实的副作用：**20 个变体就是 20 个
        文档**，全留在会话里会把 CATIA 塞满，后续任何「按名字找文档」都变得难用。
        所以给了 output_dir 就存完即关，会话保持干净。
        """
        # 文件名来自调用方（可能是 AI 生成的），必须消毒后才能拼进路径：
        # 否则 name="../../x" 就能写到 output_dir 之外去。
        stem = self._safe_stem(name) or f"Variant_{index + 1:02d}"
        path = self._validate_output_path(os.path.join(out_dir, stem + ".CATPart"), {".catpart"})
        if os.path.commonpath([out_dir, path]) != out_dir:
            raise ValueError(f"变体 {index + 1} 的输出路径逃出了 output_dir。")

        doc = self._find_document(app, document_name)
        try:
            doc.SaveAs(path)
        except pythoncom.com_error:  # type: ignore[attr-defined]
            return None
        if not (os.path.isfile(path) and os.path.getsize(path) > 0):
            return None
        try:
            doc.Close()
        except pythoncom.com_error:  # type: ignore[attr-defined]
            pass  # 关不掉不影响「已存盘」这个事实
        return path

    @staticmethod
    def _safe_stem(name) -> str:
        """只保留字母数字、下划线、短横 —— 路径分隔符和 `..` 一律出局。"""
        if not name:
            return ""
        return "".join(ch for ch in str(name) if ch.isalnum() or ch in "_-")[:60]

    def _add_fillet_on_edges(self, part_doc, part, radius_mm: float, propagation: str):
        """拾取实体的全部边并倒角。返回 (fillet, strategy, filleted, candidates, errors)。

        ── 为什么不是「把整个特征丢给 CATIA」──
        那条路已被实测否决（8 种组合全失败），而且失败方式讲清了原因：
            ref=False → "Type mismatch"  ：方法必须收 Reference，不收裸 COM 对象
            ref=True  → "method failed"  ：引用类型没问题，但**整个特征/实体不是合法的倒角对象**
        CATIA 要的是真正的**边引用**，没有捷径。

        ── 为什么不手写 BRep 名字 ──
        边引用长这样：`REdge:(Edge:(Face:(Brp:(Pad.1;2);None:();Cf11:());...`
        手写就等于把「第几个面、第几条边」硬编码进代码，换个模型立刻失效，
        而且拼错了只会得到一句没有信息量的 "method failed"。

        ── 实际做法：让 CATIA 自己把边找出来 ──
        用文档的 `Selection.Search("Topology.CGMEdge,all")`，拿到的 Reference
        与人在界面上点选那条边**完全等价**，由 CATIA 自己生成，因此天然合法、
        不含任何硬编码拓扑名字。然后在第一条边上建圆角，其余边用
        `AddObjectToFillet` 追加进同一个特征 —— 特征树上就是干净的一个 EdgeFillet。

        ── 候选数 ≠ 实体边数（实测）──
        搜索是**全文档**范围的，40×30×20 的长方体实测捞到 16 个候选：12 条实体边
        + 4 条草图线。草图线被 CATIA 拒绝（`AddObjectToFillet failed`）是**预期行为**，
        不是故障。所以这里同时返回「候选数」与「实际倒角数」，让两者的差值有据可查，
        而不是把预期内的拒绝伪装成错误。
        """
        # catTangencyFilletEdgePropagation = 1 / catMinimalFilletEdgePropagation = 2
        mode = 1 if propagation == "tangency" else 2
        shape_factory = part.ShapeFactory
        errors: list = []

        refs, query = self._search_edge_references(part_doc, errors)
        if not refs:
            raise RuntimeError(
                "没能拾取到任何边，倒圆角无法进行。\n"
                "各次尝试的原始报错：\n  " + "\n  ".join(str(e) for e in errors)
            )

        methods = (
            "AddNewSolidEdgeFilletWithConstantRadius",
            "AddNewEdgeFilletWithConstantRadius",
        )
        for method_name in methods:
            method = getattr(shape_factory, method_name, None)
            if method is None:
                continue
            try:
                fillet = method(refs[0], mode, radius_mm)
            except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                errors.append(f"{method_name}(候选 1/{len(refs)}): {exc}")
                continue

            added = 1
            rejected = 0
            for idx, ref in enumerate(refs[1:], start=2):
                try:
                    fillet.AddObjectToFillet(ref)
                    added += 1
                except pythoncom.com_error:  # type: ignore[attr-defined]
                    # 预期内：搜索是全文档范围的，会捞到草图线等不可倒角的对象。
                    # 只计数、不当错误，避免把正常现象伪装成故障。
                    rejected += 1
            if rejected:
                errors.append(
                    f"{rejected}/{len(refs)} 个候选被 CATIA 拒绝（通常是草图线等非实体边，属预期内）"
                )
            return (
                fillet,
                f"{query}/{method_name}/edges={added}of{len(refs)}",
                added,
                len(refs),
                errors,
            )

        raise RuntimeError(
            f"拾到了 {len(refs)} 个候选，但 ShapeFactory 的倒圆角方法都调用失败。\n"
            "各次尝试的原始报错：\n  " + "\n  ".join(str(e) for e in errors)
        )

    @staticmethod
    def _search_edge_references(part_doc, errors: list):
        """用 CATIA 的选择集搜索语法把实体上的边全部捞出来，返回 (refs, 生效的查询串)。

        查询串在不同版本/语言环境下写法略有差异，所以按序试几种，谁成谁上。
        """
        queries = (
            "Topology.CGMEdge,all",
            "Topology.CGMEdge,sel",
            "'Topology'.CGMEdge,all",
        )
        selection = part_doc.Selection
        for query in queries:
            try:
                selection.Clear()
                selection.Search(query)
                count = int(getattr(selection, "Count2", 0) or selection.Count)
                if count <= 0:
                    errors.append(f"Search({query!r}): 命中 0 条")
                    continue
                refs = [selection.Item(i).Reference for i in range(1, count + 1)]
                selection.Clear()
                return refs, f"Search({query})"
            except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                errors.append(f"Search({query!r}): {exc}")
            finally:
                try:
                    selection.Clear()
                except pythoncom.com_error:  # type: ignore[attr-defined]
                    pass
        return [], ""

    @staticmethod
    def _box_fillet_removed_mm3(
        length: float, width: float, height: float, r: float
    ) -> float:
        """长方体 12 条边全部倒 r 圆角后，被切掉的体积（精确解）。

        把倒圆后的实体拆成互不重叠的四部分来算，而不是去减那些形状怪异的角料：
            内芯长方体 + 6 块面板 + 12 段四分之一圆柱 + 8 个球面八分之一（合成一整球）
        这样每一项都是初等体积，结果是精确值而非近似 —— 验证才立得住。
        """
        a, b, c = length - 2 * r, width - 2 * r, height - 2 * r
        core = a * b * c
        slabs = 2 * r * (a * b + a * c + b * c)
        quarter_cylinders = math.pi * r * r * (a + b + c)
        corners = 4.0 / 3.0 * math.pi * r ** 3
        return length * width * height - (core + slabs + quarter_cylinders + corners)

    # ------------------------------------------------------------------
    # 修饰特征三兄弟的底层实现
    # ------------------------------------------------------------------
    @staticmethod
    def _chamfer_candidates(shape_factory, errors: list) -> list:
        """列出要试的 (方法名, 参数个数, mode, orientation, 第二个尺寸参数的含义)。

        ── 首轮真机实测告诉我们两件事 ──
        1. 这台 V5R34 上**没有** `AddNewSolidEdgeChamfer`（getattr 直接是 None）。
        2. `AddNewChamfer` 收 5 个参数时报 **"Invalid number of parameters."**
           —— 说明真实签名比我押的多一个，多出来的是 `CatChamferOrientation`：
           `AddNewChamfer(边, 传播, 模式, 朝向, 长度1, 角度或长度2)`。

        ── 于是这里要同时枚举三个不确定量 ──
        参数个数（6 / 5）、`CatChamferMode`（1/2/0）、以及**第二个尺寸参数到底是
        角度还是第二条边长** —— 后者才是最阴的：模式枚举猜反时，把 45 当成 45mm
        的边长，调用照样合法，只是做出来的东西完全不对。

        所以候选里成对地放「角度」和「第二边长」两种解释，让上层用体积去裁决。
        对 45° 斜角这两种解释的正确结果**恰好是同一个形状**（两条腿都等于 d），
        所以谁对谁错只能靠算出来的体积分辨，靠读文档分辨不了。

        朝向（orientation）对 45° 对称斜角不影响体积，所以排在最后当兜底。
        """
        names = []
        for name in ("AddNewChamfer", "AddNewSolidEdgeChamfer"):
            if getattr(shape_factory, name, None) is not None:
                names.append(name)
            else:
                errors.append(f"{name}: 这台机器的 ShapeFactory 上不存在，跳过")
        if not names:
            return []

        cands = []
        for name in names:
            for orient in (1, 2):
                for mode in (1, 2, 0):
                    for kind in ("angle", "length2"):
                        cands.append((name, 6, mode, orient, kind))
            for mode in (1, 2, 0):
                for kind in ("angle", "length2"):
                    cands.append((name, 5, mode, 0, kind))
        return cands

    def _try_build_chamfer(self, part_doc, part, cand, length_mm, angle_deg, prop, errors):
        """按一组候选参数建一次斜角。成功返回 (chamfer, strategy, added, candidates)。

        每次都重新搜一遍边：上一轮失败的特征被删掉之后，BRep 名字可能已经变了，
        沿用旧引用是自找麻烦 —— 重搜很便宜，别省这一下。
        """
        name, arity, mode, orient, kind = cand
        method = getattr(part.ShapeFactory, name, None)
        if method is None:
            return None

        refs, query = self._search_edge_references(part_doc, errors)
        if not refs:
            return None

        second = angle_deg if kind == "angle" else length_mm * math.tan(math.radians(angle_deg))
        args = (
            (refs[0], prop, mode, orient, length_mm, second)
            if arity == 6
            else (refs[0], prop, mode, length_mm, second)
        )
        tag = f"{name}(argc={arity},mode={mode},orient={orient},2nd={kind})"
        try:
            chamfer = method(*args)
        except pythoncom.com_error as exc:  # type: ignore[attr-defined]
            errors.append(f"{tag}: {exc}")
            return None

        added = 1
        rejected = 0
        for ref in refs[1:]:
            try:
                chamfer.AddObjectToChamfer(ref)
                added += 1
            except pythoncom.com_error:  # type: ignore[attr-defined]
                # 与 Fillet 同理：搜索是全文档范围的，会捞到草图线等非实体边。
                rejected += 1
        if rejected:
            errors.append(
                f"{tag}: {rejected}/{len(refs)} 个候选被拒（通常是草图线等非实体边，属预期内）"
            )
        return chamfer, f"{query}/{tag}/edges={added}of{len(refs)}", added, len(refs)

    @staticmethod
    def _delete_feature(part_doc, part, feature) -> bool:
        """把一个刚建错的特征从树上删掉，让下一次尝试从干净状态开始。

        没有这一步，策略链就只能「试一次」——第一个被接受但做错的组合会永久
        赖在树上，后面所有尝试都建在一个已经错了的模型上。
        """
        try:
            selection = part_doc.Selection
            selection.Clear()
            selection.Add(feature)
            selection.Delete()
            selection.Clear()
        except pythoncom.com_error:  # type: ignore[attr-defined]
            return False
        try:
            part.Update()
        except pythoncom.com_error:  # type: ignore[attr-defined]
            pass
        return True

    def _face_table(self, part_doc, app, errors: list):
        """把实体的每个面连同面积、重心列成表。返回 (table, 生效的查询串)。

        ── 为什么非得测一遍才能挑面 ──
        `Selection.Search` 回来的顺序**不保证**。按索引挑面（"第 3 个是顶面"）
        在自己这台机器上能跑，换个模型、换个 R 版就悄悄挑错 —— 而且不报错，
        只是做出个别的东西。按「重心 Z 最大」挑顶面则是几何事实，永远成立。

        代价是每个面一次 GetMeasurable + 一次重心读取（重心还得走 VBA 蹦床）。
        长方体六个面而已，这笔开销买的是**换模型不会错**。
        """
        refs, query = self._search_face_references(part_doc, errors)
        if not refs:
            return [], query

        try:
            spa = part_doc.GetWorkbench("SPAWorkbench")
        except pythoncom.com_error as exc:  # type: ignore[attr-defined]
            errors.append(f"GetWorkbench(SPAWorkbench): {exc}")
            return [], query

        table: list = []
        for i, ref in enumerate(refs):
            try:
                measurable = spa.GetMeasurable(ref)
            except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                errors.append(f"GetMeasurable(候选面 {i + 1}): {exc}")
                continue
            area = self._read_si(measurable, "Area", 1e6, errors)
            cog, _strategy, attempts = self._read_cog_mm(measurable, app)
            if cog is None:
                errors.append(
                    f"候选面 {i + 1} 读不到重心，已跳过。首个原因："
                    f"{attempts[0] if attempts else '未知'}"
                )
                continue
            table.append({"ref": ref, "index": i, "area_mm2": area, "cog_mm": cog})
        return table, query

    @staticmethod
    def _search_face_references(part_doc, errors: list):
        """把实体上的面全部捞出来，返回 (refs, 生效的查询串)。与找边同一套路。"""
        queries = (
            "Topology.CGMFace,all",
            "Topology.CGMFace,sel",
            "'Topology'.CGMFace,all",
        )
        selection = part_doc.Selection
        for query in queries:
            try:
                selection.Clear()
                selection.Search(query)
                count = int(getattr(selection, "Count2", 0) or selection.Count)
                if count <= 0:
                    errors.append(f"Search({query!r}): 命中 0 条")
                    continue
                refs = [selection.Item(i).Reference for i in range(1, count + 1)]
                selection.Clear()
                return refs, f"Search({query})"
            except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                errors.append(f"Search({query!r}): {exc}")
            finally:
                try:
                    selection.Clear()
                except pythoncom.com_error:  # type: ignore[attr-defined]
                    pass
        return [], ""

    @staticmethod
    def _add_shell_on_face(part, face_ref, thickness_mm: float):
        """在指定面上开口抽壳。返回 (shell, strategy, errors)。

        `AddNewShell(要去掉的面, 内侧厚度, 外侧厚度)`：内侧厚度是**向内**偏移的
        壁厚，外侧给 0 —— 这样外轮廓不变，体积差才等于内腔体积，精确解才立得住。
        """
        shape_factory = part.ShapeFactory
        errors: list = []
        for method_name in ("AddNewShell", "AddNewSolidShell"):
            method = getattr(shape_factory, method_name, None)
            if method is None:
                continue
            try:
                shell = method(face_ref, float(thickness_mm), 0.0)
            except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                errors.append(f"{method_name}: {exc}")
                continue
            return shell, f"{method_name}(inner={thickness_mm},outer=0)", errors
        raise RuntimeError(
            "ShapeFactory 的抽壳方法都调用失败。\n"
            "各次尝试的原始报错：\n  " + "\n  ".join(str(e) for e in errors)
        )

    @staticmethod
    def _make_references(part, refs: list, errors: list):
        """尽力造出一个 `References` 集合。返回 (集合或 None, 走通的路子)。

        首轮真机实测：这台 V5R34 的 `Part` 上**根本没有** `CreateReferences`
        （`getattr` 直接是 None）。而 `AddNewDraft` 的第一个参数偏偏要它。

        所以这里把「集合从哪来」当成一个未知量按序试，而不是押一个写法。
        全试不出来也不算死路 —— 上层还有「逐面各建一个 Draft」的退化路径。
        """
        routes = (
            ("part.CreateReferences()", lambda: part.CreateReferences()),
            ("part.Application.CreateReferences()",
             lambda: part.Application.CreateReferences()),
            ("part.Parent.CreateReferences()", lambda: part.Parent.CreateReferences()),
        )
        for label, make in routes:
            try:
                col = make()
            except (AttributeError, pythoncom.com_error) as exc:  # type: ignore[attr-defined]
                errors.append(f"{label}: {exc}")
                continue
            if col is None:
                errors.append(f"{label}: 返回 None")
                continue
            try:
                for ref in refs:
                    col.Add(ref)
            except (AttributeError, pythoncom.com_error) as exc:  # type: ignore[attr-defined]
                errors.append(f"{label} 之后 .Add 失败: {exc}")
                continue
            return col, label
        return None, ""

    def _add_draft_on_faces(self, part_doc, part, side_refs: list, neutral_ref, angle_deg: float):
        """给一组侧面加拔模。返回 (draft, strategy, errors)。

        ── 两条路，先走整的，走不通就走碎的 ──
        `AddNewDraft` 的第一个参数要 `References` 集合。首轮实测这台机器造不出来，
        所以准备了退化路径：**一个面建一个 Draft 特征**。

        退化路径为什么在几何上等价：四个侧面用的是同一个中性面、同一个方向、
        同一个角度，分四次做和一次做出来的形状一样，只是特征树上多三个节点。
        **能接受的降级要说清代价** —— 这里的代价就是树变长了，仅此而已。

        ── 枚举照旧按序试 ──
        中性面传播模式和拔模模式的取值资料不一，做法与斜角一致：按序试、
        把生效的写进 strategy、最后用体积把结果钉死。
        """
        shape_factory = part.ShapeFactory
        errors: list = []

        method = getattr(shape_factory, "AddNewDraft", None)
        if method is None:
            raise RuntimeError(
                "这台机器的 ShapeFactory 上没有 AddNewDraft。"
                "请跑 `python scripts\\probe_draft_api.py`，它会列出真实存在的方法名和参数表。"
            )

        try:
            direction = part.CreateReferenceFromObject(part.OriginElements.PlaneXY)
        except (AttributeError, pythoncom.com_error) as exc:  # type: ignore[attr-defined]
            raise RuntimeError(f"取不到拔模方向（PlaneXY 的引用）：{exc}") from exc

        combos = [(n, m) for n in (1, 2, 0) for m in (0, 1)]
        refs_col, route = self._make_references(part, side_refs, errors)

        if refs_col is not None:
            for neutral_mode, draft_mode in combos:
                try:
                    draft = method(
                        refs_col, neutral_ref, neutral_mode,
                        direction, float(angle_deg), draft_mode,
                    )
                except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                    errors.append(
                        f"AddNewDraft({route}, neutral={neutral_mode}, mode={draft_mode}): {exc}"
                    )
                    continue
                return (
                    draft,
                    f"AddNewDraft({route},neutral={neutral_mode},mode={draft_mode})",
                    errors,
                )

        # 退化路径：一个面一个 Draft
        for neutral_mode, draft_mode in combos:
            made: list = []
            failure = None
            for i, ref in enumerate(side_refs):
                try:
                    made.append(
                        method(ref, neutral_ref, neutral_mode,
                               direction, float(angle_deg), draft_mode)
                    )
                except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                    failure = (
                        f"AddNewDraft(逐面 {i + 1}/{len(side_refs)}, "
                        f"neutral={neutral_mode}, mode={draft_mode}): {exc}"
                    )
                    break
            if failure is None and made:
                return (
                    made[-1],
                    f"AddNewDraft(逐面×{len(made)},neutral={neutral_mode},mode={draft_mode})",
                    errors,
                )
            errors.append(failure or "逐面路径没建出任何特征")
            # 半途失败要把已经建出来的清掉，否则下一轮尝试是在脏模型上做的
            for feature in reversed(made):
                self._delete_feature(part_doc, part, feature)

        raise RuntimeError(
            "AddNewDraft 的所有路子（References 集合 / 逐面）和枚举组合都失败了。\n"
            "请跑 `python scripts\\probe_draft_api.py` 看真实签名。\n"
            "各次尝试的原始报错：\n  " + "\n  ".join(str(e) for e in errors)
        )

    @staticmethod
    def _box_chamfer_removed_mm3(
        length: float, width: float, height: float, d: float
    ) -> float:
        """长方体 12 条边全部倒 45° 斜角（边长 d）后被切掉的体积（精确解）。

        与倒圆角同一套拆法，只是把圆的零件换成直的：
            内芯长方体 + 6 块面板 + 12 段三棱柱（截面 d²/2）+ 8 个角块
        角块是三个斜面围出来的立体，体积恰为 d³/4 —— 八个拼起来正好是
        棱长 d 的**菱形十二面体**（体积 2d³），这不是巧合：
        把边长 2d 的正方体十二条边各切 d，剩下的就是它。

        化简后 removed = 2d²(a+b+c) + 6d³，但这里保留拆解写法，
        因为**每一项都能对上一个看得见的零件**，出错时好查。
        """
        a, b, c = length - 2 * d, width - 2 * d, height - 2 * d
        core = a * b * c
        slabs = 2 * d * (a * b + a * c + b * c)
        prisms = 2 * d * d * (a + b + c)          # 12 段三棱柱
        corners = 2 * d ** 3                      # 8 个角块，每个 d³/4
        return length * width * height - (core + slabs + prisms + corners)

    @staticmethod
    def _box_draft_delta_mm3(
        length: float, width: float, height: float, angle_deg: float, outward: bool
    ) -> float:
        """四个侧面拔模后的体积变化量（精确解，带符号）。

        底面为中性面，所以底面仍是 length × width，顶面每边缩放 k = H·tan(a)。
        结果是个棱台，用 Prismatoid 公式（对棱台是精确的，不是近似）：
            V = H/6 · (A_bottom + 4·A_mid + A_top)
        化简后：V = H·[LW ± k(L+W) + 4k²/3]，所以
            ΔV = H·[±k(L+W) + 4k²/3]

        注意两个方向的 |ΔV| **并不相等**（加料那项和去料那项差了 2k(L+W)H），
        所以"看实测符合哪一个"仍然是硬判据，不是二选一的放水。
        """
        k = height * math.tan(math.radians(angle_deg))
        sign = 1.0 if outward else -1.0
        return height * (sign * k * (length + width) + 4.0 * k * k / 3.0)

    # ------------------------------------------------------------------
    # API 探针 —— 不再靠猜签名，直接问类型库
    # ------------------------------------------------------------------
    @staticmethod
    def _com_signatures(obj, keywords: tuple = ()) -> list:
        """把 COM 对象上**真实存在**的方法连同参数名列出来。

        这比「按名字 getattr 试探」强一个量级：
            试探只能回答"我猜的这个在不在"；
            类型信息能回答"到底有哪些、每个要几个参数、参数叫什么"。

        首轮真机踩的两个坑（`AddNewSolidEdgeChamfer` 不存在、`AddNewChamfer`
        参数个数不对、`Part.CreateReferences` 不存在）本质上是同一个问题：
        **拿二手文档当一手事实**。类型库就在那台机器上，问它就好了。
        """
        try:
            type_info = obj._oleobj_.GetTypeInfo()
            attr = type_info.GetTypeAttr()
        except Exception as exc:  # noqa: BLE001 —— 拿不到类型信息不该让探针整个挂掉
            return [f"<取不到类型信息：{exc}>"]

        out: list = []
        for i in range(attr.cFuncs):
            try:
                desc = type_info.GetFuncDesc(i)
                names = type_info.GetNames(desc.memid)
            except Exception:  # noqa: BLE001
                continue
            if not names:
                continue
            method, params = str(names[0]), [str(n) for n in names[1:]]
            if keywords and not any(k.lower() in method.lower() for k in keywords):
                continue
            out.append(f"{method}({', '.join(params)})")
        return sorted(set(out))

    def probe_shape_api(self) -> dict:
        """列出当前 Part 上与修饰特征相关的真实 API 签名。

        用法：先随便建个 Part（或打开一个），再跑 `scripts\\probe_draft_api.py`。
        这是个**诊断工具**，没有固定合格标准，所以不进回归 —— 和 probe_export_formats 同理。
        """
        app = self._require_app()
        part_doc = app.ActiveDocument
        part = part_doc.Part
        shape_factory = part.ShapeFactory

        report: dict = {
            "document_name": str(part_doc.Name),
            "shapefactory_chamfer": self._com_signatures(shape_factory, ("chamfer",)),
            "shapefactory_draft": self._com_signatures(shape_factory, ("draft",)),
            "shapefactory_shell": self._com_signatures(shape_factory, ("shell",)),
            "part_create_methods": self._com_signatures(part, ("create",)),
        }

        # References 集合到底能不能造出来 —— 直接试一遍，把每条路的结果都记下来
        errors: list = []
        col, route = self._make_references(part, [], errors)
        report["references_route"] = route or None
        report["references_errors"] = errors or None
        report["references_type"] = type(col).__name__ if col is not None else None
        return report

    # ------------------------------------------------------------------
    # 导出能力探针 —— 定位「ExportData failed」到底是许可证还是 STEP 单独没授权
    # ------------------------------------------------------------------
    def probe_export_formats(self, out_dir: str) -> dict:
        """对当前活动 Part 逐个尝试多种导出格式，报告每种成功/失败。

        决定性诊断：
            - 若 stp/step/igs/stl 全失败 → translator/许可证整体缺失。
            - 若仅 stp/step 失败、stl/igs 成功 → STEP 单独未授权。
        返回 {fmt: {"ok": bool, "file": path|None, "error": msg|None}}。
        """
        app = self._require_app()
        part_doc = app.ActiveDocument

        os.makedirs(out_dir, exist_ok=True)
        base = os.path.join(out_dir, "probe_export")

        # (格式串, 文件扩展名)
        trials = [
            ("stp", ".stp"),
            ("step", ".step"),
            ("igs", ".igs"),
            ("stl", ".stl"),
            ("wrl", ".wrl"),
            ("model", ".model"),
            ("cgr", ".cgr"),
        ]
        results: dict = {}
        for fmt, ext in trials:
            target = f"{base}_{fmt}{ext}"
            try:
                if os.path.exists(target):
                    os.remove(target)
                part_doc.ExportData(target, fmt)
                resolved = self._find_recent_step_file(target) if ext in (".stp", ".step") else (
                    target if os.path.isfile(target) and os.path.getsize(target) > 0 else None
                )
                results[fmt] = {
                    "ok": resolved is not None,
                    "file": resolved,
                    "error": None if resolved is not None else "调用未抛错但无非空文件（多半是缺许可证）",
                }
            except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                results[fmt] = {"ok": False, "file": None, "error": f"COM: {exc}"}
            except OSError as exc:
                results[fmt] = {"ok": False, "file": None, "error": f"OS: {exc}"}
        return results

    # ------------------------------------------------------------------
    # 安全保存 + STEP 导出回读（验证原则 5）
    # ------------------------------------------------------------------
    def export_step_and_verify(
        self,
        step_path: str,
        catpart_path: Optional[str] = None,
        volume_tolerance: float = 1e-2,
        preferred_formats: Optional[list[tuple[str, str]]] = None,
    ) -> ExportResult:
        """把当前活动 Part 安全保存为 CATPart（可选）并导出中性格式，再回读验证体积。

        对应工程验证原则 5「回读」：导出中性格式后重新导入，比对体积，
        防止导出过程中的几何丢失/退化。容差默认 1%（格式转换会有微小误差）。

        格式策略：默认优先 STEP，STEP 未授权时自动降级到 IGES（B-rep 精确交换格式）。
        实际使用的格式写入 result.format_used。

        参数：
            step_path:    首选导出路径（.stp / .step / .igs）；降级时自动换成对应扩展名
            catpart_path: 可选，同时安全保存 .CATPart
            volume_tolerance: 体积相对误差容差
            preferred_formats: 可选，[(格式串, 扩展名)] 优先级列表；默认 STEP→IGES
        """
        app = self._require_app()

        if preferred_formats is None:
            preferred_formats = [("stp", ".stp"), ("igs", ".igs")]

        # 0) 校验输出路径（实时文件验证 = 安全边界）
        step_path = self._validate_output_path(step_path, {".stp", ".step", ".igs"})
        if catpart_path is not None:
            catpart_path = self._validate_output_path(catpart_path, {".catpart"})

        # 1) 取活动 Part 文档
        part_doc = app.ActiveDocument
        part = part_doc.Part
        body = part.Bodies.Item(1)

        # 2) 导出前测原始体积
        source_vol = self._measure_volume_mm3(part_doc, part, body)

        # 3) 安全保存 CATPart（可选）
        catpart_saved = False
        if catpart_path is not None:
            final_catpart_path = catpart_path
            try:
                part_doc.SaveAs(final_catpart_path)
            except pythoncom.com_error:  # type: ignore[attr-defined]
                # 处理同名冲突：如果当前会话中已有同名项，则生成唯一文件名后重试一次
                head, tail = os.path.split(catpart_path)
                name, ext = os.path.splitext(tail)
                unique_path = os.path.join(head, f"{name}_retry{ext}")
                try:
                    part_doc.SaveAs(unique_path)
                    final_catpart_path = unique_path
                    catpart_saved = os.path.isfile(unique_path) and os.path.getsize(unique_path) > 0
                except pythoncom.com_error:
                    catpart_saved = False
                    final_catpart_path = None
            else:
                catpart_saved = os.path.isfile(final_catpart_path) and os.path.getsize(final_catpart_path) > 0

            if final_catpart_path is not None and final_catpart_path != catpart_path:
                catpart_path = final_catpart_path

        # 4) 导出中性格式：按优先级依次尝试（STEP 不可用时自动降级到 IGES）
        base_no_ext = os.path.splitext(step_path)[0]
        export_file, format_used, export_error = self._export_first_available(
            part_doc, base_no_ext, preferred_formats
        )

        if export_file is None:
            return ExportResult(
                catpart_path=catpart_path,
                catpart_saved=catpart_saved,
                step_path=step_path,
                step_written=False,
                step_size_bytes=None,
                source_volume_mm3=source_vol,
                reimported_volume_mm3=None,
                volume_match=None,
                relative_error=None,
                export_error=export_error,
                format_used=None,
            )

        final_step_path = export_file
        step_written = os.path.isfile(final_step_path) and os.path.getsize(final_step_path) > 0
        step_size = os.path.getsize(final_step_path) if step_written else None

        # 5) 中性格式回读：重新打开导入并测体积（尽力而为，表面型导入可能测不出固体体积）
        reimported_vol: Optional[float] = None
        if step_written:
            reimported_vol = self._reimport_step_volume(app, final_step_path)

        # 6) 比对
        match: Optional[bool] = None
        rel_err: Optional[float] = None
        if source_vol is not None and reimported_vol is not None:
            rel_err = abs(reimported_vol - source_vol) / source_vol
            match = rel_err <= volume_tolerance

        return ExportResult(
            catpart_path=catpart_path,
            catpart_saved=catpart_saved,
            step_path=final_step_path,
            step_written=step_written,
            step_size_bytes=step_size,
            source_volume_mm3=source_vol,
            reimported_volume_mm3=reimported_vol,
            volume_match=match,
            relative_error=rel_err,
            export_error=None,
            format_used=format_used,
        )

    # ------------------------------------------------------------------
    def _export_first_available(self, part_doc, base_no_ext: str, formats: list[tuple[str, str]]):
        """按优先级依次尝试导出，返回第一个成功的。

        返回 (resolved_path, format_used, error_msg)：
            - 成功：(实际写出的文件路径, 格式串, None)
            - 全部失败：(None, None, 汇总错误文本)  ← 不吞错误，便于诊断
        """
        errors: list[str] = []
        for fmt, ext in formats:
            target = f"{base_no_ext}{ext}"
            try:
                if os.path.exists(target):
                    os.remove(target)
                part_doc.ExportData(target, fmt)
                resolved = self._find_recent_step_file(target)
                if resolved is not None:
                    return resolved, fmt, None
                errors.append(f"{fmt}: 调用未报错但无非空文件（多半未授权）")
            except pythoncom.com_error as exc:  # type: ignore[attr-defined]
                errors.append(f"{fmt}: COM {exc}")
            except OSError as exc:
                errors.append(f"{fmt}: OS {exc}")
        return None, None, " | ".join(errors) if errors else "无可用导出格式"

    def _find_recent_step_file(self, preferred_path: str) -> Optional[str]:
        """确认导出文件真实写出；兼容 CATIA 偶发的扩展名改写。

        优先返回精确路径；找不到时回落到同目录中同名（同 stem）最新的中性文件。
        """
        if preferred_path and os.path.isfile(preferred_path) and os.path.getsize(preferred_path) > 0:
            return preferred_path

        directory = os.path.dirname(preferred_path) or "."
        stem = os.path.splitext(os.path.basename(preferred_path))[0]
        want_ext = os.path.splitext(preferred_path)[1].lower()
        # 只在这些中性格式扩展名里回落，避免误取到无关文件
        neutral_ext = {".stp", ".step", ".igs", ".iges", ".stl", ".wrl", ".model", ".cgr"}
        allowed = {want_ext} if want_ext in neutral_ext else neutral_ext
        matches: list[str] = []
        try:
            for name in os.listdir(directory):
                lower = name.lower()
                if os.path.splitext(lower)[1] not in allowed:
                    continue
                full_path = os.path.join(directory, name)
                if not os.path.isfile(full_path):
                    continue
                if os.path.getsize(full_path) <= 0:
                    continue
                if name.lower().startswith(stem.lower()):
                    matches.append(full_path)
        except OSError:
            return None

        if not matches:
            return None
        return max(matches, key=os.path.getmtime)

    # ------------------------------------------------------------------
    def _validate_output_path(self, path: str, allowed_ext: set[str]) -> str:
        """校验并规范化输出路径：绝对路径 + 扩展名白名单 + 父目录存在。

        这是「实时文件验证」安全边界的落点，父目录不存在则创建。
        """
        if not path or not str(path).strip():
            raise ValueError("输出路径不能为空。")
        abspath = os.path.abspath(os.path.expanduser(str(path)))
        ext = os.path.splitext(abspath)[1].lower()
        if ext not in allowed_ext:
            raise ValueError(f"扩展名 {ext!r} 不在允许列表 {sorted(allowed_ext)} 中。")
        parent = os.path.dirname(abspath)
        os.makedirs(parent, exist_ok=True)
        return abspath

    # ------------------------------------------------------------------
    def _reimport_step_volume(self, app, step_path: str) -> Optional[float]:
        """重新打开 STEP 文件并测量体积（mm³）。失败返回 None。

        STEP 打开后可能是 Part 或 Product，做防御式处理，测不出就返回 None。
        """
        doc = None
        try:
            doc = app.Documents.Open(step_path)
            try:
                part = doc.Part
                body = part.Bodies.Item(1)
                return self._measure_volume_mm3(doc, part, body)
            except (AttributeError, pythoncom.com_error):  # type: ignore[attr-defined]
                # 打开成了 Product 或结构不同 —— 本阶段不深挖，返回 None
                return None
        except pythoncom.com_error:  # type: ignore[attr-defined]
            return None
        except Exception:  # noqa: BLE001
            return None
        finally:
            # 回读用的临时文档关掉，不污染会话（不保存）
            if doc is not None:
                try:
                    doc.Close()
                except Exception:  # noqa: BLE001
                    pass

    # ------------------------------------------------------------------
    def _measure_volume_mm3(self, part_doc, part, body) -> Optional[float]:
        """用 SPAWorkbench 回读实体体积，返回 mm³。失败返回 None（不阻断建模）。

        CATIA 测量 API 返回 SI 单位（m³），换算成 mm³ 需乘 1e9。
        """
        try:
            spa = part_doc.GetWorkbench("SPAWorkbench")
            ref = part.CreateReferenceFromObject(body)
            measurable = spa.GetMeasurable(ref)
            volume_m3 = float(measurable.Volume)  # SI: m³
            return volume_m3 * 1e9  # → mm³
        except pythoncom.com_error:  # type: ignore[attr-defined]
            return None
        except Exception:  # noqa: BLE001 —— 测量失败不应打断主流程
            return None

    # ------------------------------------------------------------------
    def _require_app(self):
        if self._app is None:
            raise RuntimeError("尚未 connect()，先调用 connect() 再使用其它方法。")
        return self._app
