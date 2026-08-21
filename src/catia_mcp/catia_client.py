"""CATIA COM 客户端 —— 只做最薄的一层封装。

设计原则：
1. 只依赖 pywin32，不引入任何 MCP / LLM 概念。
2. 只暴露"探测型"只读方法，方便 Hello World 阶段验证通路。
3. 单 STA 语义由调用方保证（后续 MCP server 里做）。

只能在 Windows 上运行；Mac / Linux 上 import 会失败，这是刻意的 —— 让通路问题早暴露。
"""

from __future__ import annotations

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
