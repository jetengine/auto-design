"""Hello World —— 验证 Python 能否与本机 CATIA COM 通话。

运行方式（在 **Windows** 机器上）：

    python scripts/hello_catia.py

前置条件：
    * CATIA V5 已启动（可以只是启动界面，也可以打开了一个 Part）。
    * 已经 `pip install -e .`（会装 pywin32）。

期望输出示例：

    ✅ 已连接到 CATIA。
    版本            : V5-6R2021
    Release / SP    : 30 / 4
    文档数量        : 1
    活动文档        : Part1
    主窗口标题      : CATIA V5 - [Part1]

任何一个字段异常，都是通路问题，先修好这里再往上走。
"""

from __future__ import annotations

import sys
import traceback


def main() -> int:
    try:
        # 延迟导入 —— 让"我在 Mac 上误跑了"这种错误信息更清晰
        from catia_mcp.catia_client import CatiaClient
    except RuntimeError as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 2
    except ImportError as exc:
        print(
            "❌ 无法 import catia_mcp。请先在项目根目录执行：\n"
            "     pip install -e .\n"
            f"原始错误：{exc}",
            file=sys.stderr,
        )
        return 2

    client = CatiaClient()
    try:
        client.connect()
        info = client.session_info()
    except Exception:  # noqa: BLE001 —— 这里就是要把栈完整打出来
        print("❌ 与 CATIA 通话失败。完整栈：", file=sys.stderr)
        traceback.print_exc()
        return 1

    print("✅ 已连接到 CATIA。")
    print(f"  版本         : {info.system_configuration}")
    print(f"  Release / SP : {info.release_number} / {info.service_pack}")
    print(f"  文档数量     : {info.document_count}")
    print(f"  活动文档     : {info.active_document_name or '（无）'}")
    print(f"  主窗口标题   : {info.caption}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
