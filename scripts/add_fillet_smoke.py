"""建模 + 倒圆角（EdgeFillet）冒烟测试 —— 第三种特征，也是第一次碰「边拾取」。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/add_fillet_smoke.py

它会：
    1. 新建长方体 40×30×20（Sketch + Pad）；
    2. 给它的全部 12 条边加 R5 恒定半径圆角；
    3. 用**精确解**验证去料体积（不是估算）：
       倒圆后的实体 = 内芯长方体 + 6 块面板 + 12 段四分之一圆柱 + 1 整球
         内芯   30×20×10                       = 6000.000
         面板   2·5·(30·20 + 30·10 + 20·10)    = 11000.000
         圆柱   π·5²·(30+20+10)                ≈ 4712.389
         球     4π·5³/3                        ≈ 523.599
                                          合计 ≈ 22235.988
       原体积 40×30×20 = 24000 → 理论去料 ≈ 1764.012 mm³
    4. 打印全部证据，含 `strategy`（自证到底拾到了什么）。

为什么这个测试比前两个更有价值：
    Pad / Pocket 的输入是**参数**，倒圆角的输入是**几何引用**。CATIA 的边引用
    绑定具体拓扑、极易失效，是所有 CAD 自动化最经典的脆点。体积能对上，
    说明「拾取到了全部 12 条边且都倒成功了」—— 这是引用层面真正生效的硬证据。
"""

from __future__ import annotations

import sys
import traceback
from dataclasses import asdict

# 长方体尺寸（比前两个测试小，保证 2R < 最小边长）
L, W, H = 40.0, 30.0, 20.0
R = 5.0


def main() -> int:
    try:
        from catia_mcp.com_worker import ComWorker
    except RuntimeError as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 2

    worker = ComWorker()
    try:
        worker.start()
    except Exception as exc:  # noqa: BLE001
        print(f"❌ 连接 CATIA 失败：{exc}", file=sys.stderr)
        return 2

    try:
        box = worker.call(
            lambda c: c.create_box(
                length_mm=L, width_mm=W, height_mm=H, part_name="FilletDemo"
            ),
            timeout=120.0,
            label="create_box",
        )
        print("① 建模证据：")
        for k, v in asdict(box).items():
            print(f"    {k:<22}: {v}")

        fillet = worker.call(
            lambda c: c.add_fillet(
                radius_mm=R,
                box_length_mm=L,
                box_width_mm=W,
                box_height_mm=H,
            ),
            timeout=120.0,
            label="add_fillet",
        )
        print("\n② 倒圆角证据：")
        for k, v in asdict(fillet).items():
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
        and fillet.update_ok
        and fillet.volume_match is True
    )
    print(
        "\n判定：",
        "✅ 倒圆角合格（边拾取生效，去料体积与精确解吻合）"
        if ok
        else "⚠️ 需检查（见上方证据，重点看 strategy 与 target_errors）",
    )
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
