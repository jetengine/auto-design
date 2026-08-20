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
    def _require_app(self):
        if self._app is None:
            raise RuntimeError("尚未 connect()，先调用 connect() 再使用其它方法。")
        return self._app
