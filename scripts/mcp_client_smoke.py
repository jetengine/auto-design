"""端到端 MCP 客户端冒烟测试 —— 用真实 MCP 协议驱动 CATIA。

这验证的是「AI 端」到「CATIA」的完整链路，走的和 Claude / VS Code Copilot
完全相同的路径：

    本脚本(MCP client) ── stdio/JSON-RPC ──▶ catia_mcp.server ──▶ ComWorker(STA) ──▶ CATIA

区别只是把「LLM 决策调哪个工具」换成脚本里写死的调用顺序，
从而得到一个**可复现、不依赖人在环**的端到端证明。

运行方式（在 **Windows**、CATIA 已启动、已 `pip install -e .`）：

    python scripts/mcp_client_smoke.py

它会：
    1. 以子进程方式拉起 MCP server（server 自己去连 CATIA）；
    2. initialize + 列出工具（证明协议握手成功）；
    3. 依次调用 get_catia_session → create_box → export_step_and_verify；
    4. 打印每步返回的结构化证据。
"""

from __future__ import annotations

import asyncio
import os
import sys
import time


async def main() -> int:
    try:
        from mcp import ClientSession, StdioServerParameters
        from mcp.client.stdio import stdio_client
    except ImportError as exc:
        print(f"❌ 缺少 mcp 客户端库：{exc}\n   先在 Windows 上 `pip install -e .`。", file=sys.stderr)
        return 2

    out_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "out"))
    os.makedirs(out_dir, exist_ok=True)
    tag = time.strftime("%Y%m%d_%H%M%S")
    catpart_path = os.path.join(out_dir, f"AiBox_{tag}.CATPart")
    step_path = os.path.join(out_dir, f"AiBox_{tag}.step")

    # 用当前解释器以模块方式启动 server —— 不依赖 catia-mcp.exe 的绝对路径
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "catia_mcp.server"],
    )

    def _dump(title: str, result) -> None:
        print(f"\n▶ {title}")
        # CallToolResult：结构化内容优先，否则回退到文本块
        payload = getattr(result, "structuredContent", None)
        if payload:
            for k, v in payload.items():
                print(f"    {k:<22}: {v}")
        else:
            for block in getattr(result, "content", []) or []:
                text = getattr(block, "text", None)
                if text:
                    print(f"    {text}")

    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            print("✅ MCP 协议握手成功（initialize 完成）")

            tools = await session.list_tools()
            names = [t.name for t in tools.tools]
            print(f"✅ 发现工具：{names}")

            r1 = await session.call_tool("get_catia_session", {})
            _dump("get_catia_session（只读健康检查）", r1)

            r2 = await session.call_tool(
                "create_box",
                {"length_mm": 100.0, "width_mm": 60.0, "height_mm": 20.0, "part_name": f"AiBox_{tag}"},
            )
            _dump("create_box（建模 + 体积回读）", r2)

            r3 = await session.call_tool(
                "export_step_and_verify",
                {"step_path": step_path, "catpart_path": catpart_path},
            )
            _dump("export_step_and_verify（保存 + 中性格式导出 + 回读）", r3)

    print("\n判定： ✅ 端到端链路跑通（MCP client → server → ComWorker → CATIA）")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
