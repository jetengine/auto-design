"""建模 + STEP 导出回读 冒烟测试 —— 不接 AI，验证完整交付闭环。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/export_step_smoke.py

它会：
    1. 新建长方体 100×60×20（Sketch + Pad）；
    2. 安全保存 .CATPart；
    3. 导出 .STEP，重新打开回读体积并与原始比对；
    4. 打印全部证据。

输出目录：脚本所在项目下的 out/ 。
"""

from __future__ import annotations

import os
import sys
import time
import traceback
from dataclasses import asdict


def main() -> int:
    try:
        from catia_mcp.com_worker import ComWorker
    except RuntimeError as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 2

    out_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "out"))
    # 用时间戳生成唯一名：避免与 CATIA 会话中已打开的同名文档撞名
    tag = time.strftime("%Y%m%d_%H%M%S")
    catpart_path = os.path.join(out_dir, f"SmokeBox_{tag}.CATPart")
    step_path = os.path.join(out_dir, f"SmokeBox_{tag}.step")

    worker = ComWorker()
    try:
        worker.start()

        box = worker.call(
            lambda c: c.create_box(100.0, 60.0, 20.0, part_name=f"SmokeBox_{tag}"),
            timeout=120.0,
        )
        print("① 建模证据：")
        for k, v in asdict(box).items():
            print(f"    {k:<22}: {v}")

        exp = worker.call(
            lambda c: c.export_step_and_verify(step_path, catpart_path=catpart_path),
            timeout=180.0,
        )
        print("\n② 导出/回读证据：")
        for k, v in asdict(exp).items():
            print(f"    {k:<22}: {v}")
    except Exception:  # noqa: BLE001
        print("❌ 失败。完整栈：", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        worker.stop()

    # 交付合格判定：
    #   - 建模合格（体积回读吻合）
    #   - CATPart（可编辑特征树）落盘  ← 核心交付物
    #   - 中性格式导出成功落盘（STEP 未授权时自动降级为 IGES）
    # 中性格式的体积回读是「尽力而为」：IGES 常以曲面导入，可能测不出固体体积，
    # 不作为硬性合格条件。
    ok = (
        box.update_ok
        and box.volume_match is True
        and exp.catpart_saved
        and exp.step_written
    )
    fmt = exp.format_used or "(无)"
    if ok:
        vol_note = "，中性格式体积回读吻合" if exp.volume_match is True else "（中性格式体积回读尽力而为）"
        print(f"\n判定： ✅ 交付合格（建模 + CATPart + {fmt} 导出均通过{vol_note}）")
    else:
        print("\n判定： ⚠️ 需检查（见上方证据）")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
