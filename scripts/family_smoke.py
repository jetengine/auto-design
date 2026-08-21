"""批量建族冒烟测试 —— 把单件能力放大到规模。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/family_smoke.py

它会：
    1. 一次提交 6 个变体（3 个纯长方体 + 3 个带倒角），全部存盘后关闭；
    2. 逐个用**精确解**核对最终体积；
    3. **故意混入 2 个非法变体**，验证「单个失败不拖垮整批」。

── 为什么这个测试和前面几个不一样 ──
前面每个冒烟测的是「这个功能对不对」。这个测的是「**批量语义对不对**」：

    · 第 3 个失败了，第 4~8 个还跑不跑？
    · 失败的那个在汇总里还看得见吗？（悄悄跳过 = 8 变 6，没人会发现）
    · 汇总数字和明细对不对得上？

功能正确 ≠ 批量正确。一个 for 循环里 try 写错位置，就能把"20 个全成功"变成
"第 1 个成功、剩下 19 个被静默吞掉"，而单件测试一个都抓不到。
"""

from __future__ import annotations

import math
import os
import sys
import traceback
from dataclasses import asdict

OUT_DIR = os.path.join(os.getcwd(), "out", "family")

# 合法变体：一族由薄到厚的板，后三个带倒角
GOOD = [
    {"name": "Plate_A", "length_mm": 60.0, "width_mm": 40.0, "height_mm": 8.0},
    {"name": "Plate_B", "length_mm": 60.0, "width_mm": 40.0, "height_mm": 12.0},
    {"name": "Plate_C", "length_mm": 60.0, "width_mm": 40.0, "height_mm": 16.0},
    {"name": "Round_A", "length_mm": 50.0, "width_mm": 40.0, "height_mm": 30.0, "fillet_radius_mm": 3.0},
    {"name": "Round_B", "length_mm": 50.0, "width_mm": 40.0, "height_mm": 30.0, "fillet_radius_mm": 5.0},
    {"name": "Round_C", "length_mm": 50.0, "width_mm": 40.0, "height_mm": 30.0, "fillet_radius_mm": 8.0},
]

# 非法变体：都应在**碰 CATIA 之前**就被参数校验拦下
BAD = [
    {"name": "NegDim", "length_mm": -10.0, "width_mm": 20.0, "height_mm": 5.0},
    {"name": "FatFillet", "length_mm": 20.0, "width_mm": 20.0, "height_mm": 10.0,
     "fillet_radius_mm": 6.0},  # 2r=12 ≥ 最短边 10
]


def _expected_volume(v: dict) -> float:
    L, W, H = v["length_mm"], v["width_mm"], v["height_mm"]
    r = v.get("fillet_radius_mm")
    if not r:
        return L * W * H
    a, b, c = L - 2 * r, W - 2 * r, H - 2 * r
    return (
        a * b * c
        + 2 * r * (a * b + a * c + b * c)
        + math.pi * r * r * (a + b + c)
        + 4.0 / 3.0 * math.pi * r ** 3
    )


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

    all_ok = True
    try:
        # ── ① 正常批次：6 个变体全部应通过 ────────────────────────
        fam = worker.call(
            lambda c: c.create_box_family(variants=GOOD, output_dir=OUT_DIR),
            timeout=300.0,
            label="create_box_family(good)",
        )
        print("① 批量汇总：")
        for k, v in asdict(fam).items():
            if k != "variants":
                print(f"    {k:<22}: {v}")

        print("\n   逐个变体：")
        for var in fam.variants:
            exp = _expected_volume(GOOD[var.index])
            got = var.measured_volume_mm3
            rel = abs(got - exp) / exp if got else None
            good = bool(var.ok and rel is not None and rel <= 1e-3)
            all_ok &= good
            saved = os.path.basename(var.saved_path) if var.saved_path else "—"
            print(
                f"    {'✅' if good else '❌'} [{var.index}] {var.name:<9}"
                f" V={got if got is None else round(got, 3)}"
                f"  理论={round(exp, 3)}"
                f"  rel={'—' if rel is None else f'{rel:.2e}'}"
                f"  边={var.objects_filleted}  文件={saved}"
            )

        summary_ok = (
            fam.requested == len(GOOD)
            and fam.succeeded == len(GOOD)
            and fam.all_verified
            and len(fam.variants) == len(GOOD)
            and fam.documents_left_open == 0
        )
        print(f"\n    {'✅' if summary_ok else '❌'} 汇总自洽（请求={len(GOOD)}、全成、明细条数相符、会话未残留文档）")
        all_ok &= summary_ok

        files_ok = all(
            v.saved_path and os.path.isfile(v.saved_path) and os.path.getsize(v.saved_path) > 0
            for v in fam.variants
        )
        print(f"    {'✅' if files_ok else '❌'} {len(fam.variants)} 个 CATPart 真实落盘（实时文件验证）")
        all_ok &= files_ok

        # ── ② 混合批次：非法变体应被前置校验整批拒绝 ──────────────
        print("\n② 前置校验（混入 2 个非法变体，应在碰 CATIA 之前就整批拒绝）：")
        for bad in BAD:
            try:
                worker.call(
                    lambda c, b=bad: c.create_box_family(variants=[b]),
                    timeout=60.0,
                    label="create_box_family(bad)",
                )
                print(f"    ❌ {bad['name']}: 竟然没被拦下")
                all_ok = False
            except Exception as exc:  # noqa: BLE001
                msg = str(exc).strip().splitlines()[0]
                named = "第 1 个变体" in msg
                print(f"    {'✅' if named else '⚠️'} {bad['name']}: {msg}")
                all_ok &= named

        # ── ③ 上限保护 ────────────────────────────────────────────
        print("\n③ 批次上限保护（51 个变体应被拒）：")
        try:
            worker.call(
                lambda c: c.create_box_family(
                    variants=[{"length_mm": 10, "width_mm": 10, "height_mm": 10}] * 51
                ),
                timeout=60.0,
                label="create_box_family(oversize)",
            )
            print("    ❌ 竟然没被拦下")
            all_ok = False
        except Exception as exc:  # noqa: BLE001
            print(f"    ✅ {str(exc).strip().splitlines()[0]}")

    except Exception:  # noqa: BLE001
        print("❌ 失败。完整栈：", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        worker.stop()

    print(
        "\n判定：",
        "✅ 批量合格（逐个变体经精确解验证，汇总自洽，非法输入被前置拦下）"
        if all_ok
        else "⚠️ 需检查 —— 重点看是「某个变体几何不对」还是「批量语义不对」（汇总与明细对不上）",
    )
    return 0 if all_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
