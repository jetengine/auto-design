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
    errors: Optional[list] = None         # 各项测量失败时的原始 COM 报错


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
        cog, cog_strategy = self._read_cog_mm(measurable, errors, app)

        return MeasureResult(
            document_name=str(doc.Name),
            body_name=str(body.Name),
            volume_mm3=volume,
            area_mm2=area,
            cog_mm=cog,
            cog_strategy=cog_strategy,
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
    def _read_cog_mm(measurable, errors: list, app=None):
        """读重心，返回 ((x, y, z) mm, 生效的调用姿势)。

        ── 这里踩到的坑，比「调用失败」更值得记 ──
        `GetCOG` 不是返回值，而是**出参数组**（`CATSafeArrayVariant`）。第一版
        试了两种姿势，实测结果是：

            byref VARIANT(VT_R8) → 直接报错 "Objects for SAFEARRAYS must be
                                   sequences (of sequences), or a buffer object"
            直接传 list          → **不报错，但返回全零**

        第二种才是真正危险的那个：pywin32 把 list 按值传了进去，CATIA 写回的是
        那份副本，我们读到的原 list 一个字节都没变。**它没有失败，它在说谎。**
        一个静默返回错误数据的策略，比一个抛异常的策略危险得多 —— 后者会停下来，
        前者会一路把错误数据带进结论里。

        ── 对策：哨兵值 ──
        缓冲区不填 0，填一个真实重心绝无可能取到的magic number。调用完如果三个
        分量**还是哨兵值**，就证明 CATIA 根本没写回来，判为失败换下一种姿势。
        不能简单地「见到 (0,0,0) 就当失败」—— 一个居中建模的零件重心本来就是原点，
        那样会把正确结果误杀。哨兵值区分的是「没被写」而不是「值恰好是零」。
        """
        sentinel = -1.2345678901e9  # 真实重心（米）不可能取到这个量级

        def _harvest(vals, label):
            """把出参缓冲区换算成 mm；没被写回则返回 None 并记原因。"""
            vals = [float(v) for v in (vals or [])]
            if len(vals) < 3:
                errors.append(f"GetCOG({label}): 只拿到 {len(vals)} 个分量")
                return None
            if all(v == sentinel for v in vals[:3]):
                errors.append(f"GetCOG({label}): 出参没被写回（缓冲区仍是哨兵值）")
                return None
            return tuple(v * 1e3 for v in vals[:3])  # SI 米 → mm

        # 姿势 A/B：byref VARIANT。CATSafeArrayVariant 的元素类型是 VARIANT，
        # 所以先试 VT_VARIANT，再退回 VT_R8。
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
                    return got, f"VARIANT({label})"
            except (AttributeError, pythoncom.com_error, ValueError, TypeError) as exc:  # type: ignore[attr-defined]
                errors.append(f"GetCOG({label}): {exc}")

        # 姿势 C：直接传 list。已知会静默说谎，靠哨兵拦住；留着是为了换机器时
        # 万一 pywin32 行为不同，它能顶上。
        try:
            buf = [sentinel, sentinel, sentinel]
            out = measurable.GetCOG(buf)
            got = _harvest(out if out is not None else buf, "plain list")
            if got is not None:
                return got, "plain list"
        except (AttributeError, pythoncom.com_error, ValueError, TypeError) as exc:  # type: ignore[attr-defined]
            errors.append(f"GetCOG(plain list): {exc}")

        # 姿势 D：VBA 跳板。绕开 pywin32 的出参 marshalling —— 让 CATIA 自己在
        # 内部跑一小段 VBScript 把数组接住再当返回值递出来。这是 pycatia 对同类
        # 出参方法的通行解法，代价是需要脚本执行权限（企业环境可能被锁）。
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
                        return got, f"SystemService.Evaluate(lang={lang})"
                except (AttributeError, pythoncom.com_error, ValueError, TypeError) as exc:  # type: ignore[attr-defined]
                    errors.append(f"GetCOG(VBA lang={lang}): {exc}")

        return None, None

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
