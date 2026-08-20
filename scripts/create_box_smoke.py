"""长方体建模冒烟测试 —— 不接 AI，先验证写操作 + 体积回读闭环。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/create_box_smoke.py

它会：
    1. 通过单 STA worker 连接 CATIA；
    2. 新建 Part，建 100×60×20 的长方体（Sketch + Pad 原生特征）；
    3. Update 检查 + 回读体积，打印证据。

期望：update_ok=True，volume_match=True，measured≈120000 mm³。
"""

from __future__ import annotations

import sys
import traceback
from dataclasses import asdict


def main() -> int:
    try:
        from catia_mcp.com_worker import ComWorker
    except RuntimeError as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 2

    worker = ComWorker()
    try:
        worker.start()
        result = worker.call(
            lambda c: c.create_box(100.0, 60.0, 20.0, part_name="SmokeBox"),
            timeout=120.0,
        )
    except Exception:  # noqa: BLE001
        print("❌ 建模失败。完整栈：", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        worker.stop()

    print("✅ 长方体建模完成，证据如下：")
    for k, v in asdict(result).items():
        print(f"  {k:<22}: {v}")

    ok = result.update_ok and (result.volume_match is True)
    print("\n判定：", "✅ 几何合格" if ok else "⚠️ 需检查（见上方 update_ok / volume_match）")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
