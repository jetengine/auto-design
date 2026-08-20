"""CATIA COM 客户端 —— 只做最薄的一层封装。

设计原则：
1. 只依赖 pywin32，不引入任何 MCP / LLM 概念。
2. 只暴露"探测型"只读方法，方便 Hello World 阶段验证通路。
3. 单 STA 语义由调用方保证（后续 MCP server 里做）。

只能在 Windows 上运行；Mac / Linux 上 import 会失败，这是刻意的 —— 让通路问题早暴露。
"""

from __future__ import annotations

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
            part.Name = str(part_name)

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
