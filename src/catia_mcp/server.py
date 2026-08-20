"""CATIA MCP Server —— 第一个工具：get_catia_session。

架构：
    MCP client (Claude / VS Code)  ── stdio/JSON-RPC ──▶  本 server
                                                              │
                                                              ▼
                                                     ComWorker（单 STA 线程）
                                                              │
                                                              ▼
                                                       CATIA V5 (COM)

现在只暴露一个只读探测工具，验证「AI 能通过 MCP 问到 CATIA 真实状态」。
后续的 Sketch / Pad / Measure ... 都按同样的模式：定义 schema → submit 到 ComWorker。
"""

from __future__ import annotations

import sys
from dataclasses import asdict

if sys.platform != "win32":
    raise RuntimeError("catia_mcp.server 只能在 Windows 上运行（需要 CATIA COM）。")

from mcp.server.fastmcp import FastMCP  # type: ignore[import-not-found]

from .catia_client import CatiaClient, CatiaSessionInfo
from .com_worker import ComWorker

mcp = FastMCP("catia-mcp")

# 单一 STA worker，进程内唯一
_worker = ComWorker()


@mcp.tool()
def get_catia_session() -> dict:
    """获取当前 CATIA 会话的只读状态快照。

    返回字段：
        system_configuration: CATIA 主版本标识
        release_number:       Release 号（如 34）
        service_pack:         Service Pack 号
        active_document_name: 当前活动文档名（无则为 null）
        document_count:       已打开文档数量
        caption:              CATIA 主窗口标题（供人眼二次核对）

    这是「能力与版本探测」的最小实现，也是所有后续建模操作前的健康检查入口。
    """

    def _job(client: CatiaClient) -> CatiaSessionInfo:
        return client.session_info()

    info = _worker.call(_job, timeout=30.0)
    return asdict(info)


def main() -> None:
    """stdio 模式启动 MCP server。"""
    _worker.start()  # 先把 STA 线程 + CATIA 连接拉起，连不上直接报错退出
    try:
        mcp.run(transport="stdio")
    finally:
        _worker.stop()


if __name__ == "__main__":
    main()
