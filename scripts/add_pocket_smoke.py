"""建模 + 挖槽（Pocket 去料）冒烟测试 —— 不接 AI，验证第二种特征类型。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/add_pocket_smoke.py

它会：
    1. 新建长方体 100×60×20（Sketch + Pad）；
    2. 在顶面（Z=20）居中挖一个 40×30、深 10 的盲槽（偏移平面 + Sketch + Pocket）；
    3. 用体积差验证去料量：理论去料 = 40×30×10 = 12000 mm³，
       终体积应为 120000 − 12000 = 108000 mm³；
    4. 打印全部证据。
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

    # 长方体尺寸
    L, W, H = 100.0, 60.0, 20.0
    # 槽尺寸（居中）
    PL, PW, PD = 40.0, 30.0, 10.0

    worker = ComWorker()
    try:
        worker.start()

        box = worker.call(
            lambda c: c.create_box(L, W, H, part_name="PocketBox"),
            timeout=120.0,
        )
        print("① 建模证据：")
        for k, v in asdict(box).items():
            print(f"    {k:<22}: {v}")

        pocket = worker.call(
            lambda c: c.add_pocket(
                pocket_length_mm=PL,
                pocket_width_mm=PW,
                depth_mm=PD,
                at_height_mm=H,
                center_x_mm=L / 2.0,
                center_y_mm=W / 2.0,
            ),
            timeout=120.0,
        )
        print("\n② 挖槽证据：")
        for k, v in asdict(pocket).items():
            print(f"    {k:<22}: {v}")
    except Exception:  # noqa: BLE001
        print("❌ 失败。完整栈：", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        worker.stop()

    ok = (
        box.update_ok
        and box.volume_match is True
        and pocket.update_ok
        and pocket.volume_match is True
    )
    print("\n判定：", "✅ 去料合格（建模 + 挖槽体积差均通过）" if ok else "⚠️ 需检查（见上方证据）")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
