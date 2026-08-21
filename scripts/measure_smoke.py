"""只读测量族冒烟测试 —— 第一个「不靠自己造」也能下结论的能力。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/measure_smoke.py

它会：
    1. 建一个 40×30×20 的长方体；
    2. `inspect_document` 看清结构（有哪些 Body、哪些特征）；
    3. `measure_body` 测体积 / 表面积 / 重心，与**精确解**逐项比对；
    4. 倒 R5 圆角后再测一次 —— 验证三条证据在几何变化后依然各自说得通。

── 为什么值得单独做一族 ──
前面每个写操作都用「体积吻合」自证，但体积单独一项的验证力是有限的：

    · Pad 拉反方向  → 体积一模一样，重心却在 z 的另一侧
    · 长宽写反      → 体积一模一样，重心不一样
    · 形状完全不同  → 体积可以恰好相同，表面积会露馅

体积 + 表面积 + 重心，三条**互相独立**的证据同时对上，才算把几何真正钉死。

── 精确解 ──
长方体 40×30×20，草图矩形角点 (0,0)-(L,0)-(L,W)-(0,W)，沿 +Z 拉伸：

    体积 = L·W·H                 = 24000 mm³
    表面积 = 2(LW + LH + WH)     = 5200 mm²
    重心 = (L/2, W/2, H/2)       = (20, 15, 10) mm

全部 12 条边倒 R5 后（a=L-2r, b=W-2r, c=H-2r）：

    体积 = abc + 2r(ab+ac+bc) + πr²(a+b+c) + 4πr³/3   ≈ 22235.988 mm³
    表面积 = 2(ab+ac+bc) + 2πr(a+b+c) + 4πr²          ≈ 4399.115 mm²
    重心 —— **必须原地不动**。倒圆角对三个中心面都是对称的，重心一旦漂移，
            说明有边没倒到（或倒错了边）。这是体积测不出来的那类错误。
"""

from __future__ import annotations

import math
import sys
import traceback
from dataclasses import asdict

L, W, H = 40.0, 30.0, 20.0
R = 5.0

VOL_TOL = 1e-3   # 体积 / 面积相对误差容差
COG_TOL = 1e-3   # 重心绝对误差容差（mm）


def _box_exact():
    return L * W * H, 2 * (L * W + L * H + W * H), (L / 2, W / 2, H / 2)


def _fillet_exact():
    a, b, c = L - 2 * R, W - 2 * R, H - 2 * R
    vol = (
        a * b * c
        + 2 * R * (a * b + a * c + b * c)
        + math.pi * R * R * (a + b + c)
        + 4.0 / 3.0 * math.pi * R ** 3
    )
    area = 2 * (a * b + a * c + b * c) + 2 * math.pi * R * (a + b + c) + 4 * math.pi * R * R
    return vol, area, (L / 2, W / 2, H / 2)  # 重心不变


def _check_scalar(label: str, got, expected: float, checks: list) -> None:
    if got is None:
        checks.append((False, f"{label}: 没测出来（见 errors）"))
        return
    rel = abs(got - expected) / expected
    ok = rel <= VOL_TOL
    checks.append((ok, f"{label}: {got:.4f}（理论 {expected:.4f}，相对误差 {rel:.2e}）"))


def _check_cog(got, expected, checks: list) -> None:
    if got is None:
        checks.append((False, "重心: 没测出来（见 errors / cog_strategy）"))
        return
    dev = max(abs(g - e) for g, e in zip(got, expected))
    ok = dev <= COG_TOL
    shown = "(" + ", ".join(f"{v:.4f}" for v in got) + ")"
    checks.append((ok, f"重心: {shown}（理论 {expected}，最大偏差 {dev:.2e} mm）"))


def _report(title: str, checks: list) -> bool:
    print(f"\n{title}")
    for ok, msg in checks:
        print(f"    {'✅' if ok else '❌'} {msg}")
    return all(ok for ok, _ in checks)


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
        box = worker.call(
            lambda c: c.create_box(
                length_mm=L, width_mm=W, height_mm=H, part_name="MeasureDemo"
            ),
            timeout=120.0,
            label="create_box",
        )
        doc_name = box.document_name
        print(f"① 已建模：{doc_name}（{L}×{W}×{H}）")

        # ── 结构清单：先看清有什么 ──────────────────────────────────
        info = worker.call(
            lambda c: c.inspect_document(document_name=doc_name),
            timeout=60.0,
            label="inspect_document",
        )
        print("\n② 结构清单：")
        for k, v in asdict(info).items():
            print(f"    {k:<22}: {v}")
        struct_ok = info.is_part and info.body_count >= 1
        print(f"    {'✅' if struct_ok else '❌'} 是零件文档且至少有一个 Body")
        all_ok &= struct_ok

        body_name = info.bodies[0]["name"] if info.bodies else None

        # ── 测量 1：原始长方体 ─────────────────────────────────────
        m1 = worker.call(
            lambda c: c.measure_body(document_name=doc_name, body_name=body_name),
            timeout=60.0,
            label="measure_body(box)",
        )
        print("\n③ 长方体测量原始返回：")
        for k, v in asdict(m1).items():
            print(f"    {k:<22}: {v}")

        v_e, a_e, c_e = _box_exact()
        checks: list = []
        _check_scalar("体积", m1.volume_mm3, v_e, checks)
        _check_scalar("表面积", m1.area_mm2, a_e, checks)
        _check_cog(m1.cog_mm, c_e, checks)
        all_ok &= _report("   逐项比对（长方体）：", checks)

        # ── 测量 2：倒圆角之后 ─────────────────────────────────────
        fillet = worker.call(
            lambda c: c.add_fillet(
                radius_mm=R, box_length_mm=L, box_width_mm=W, box_height_mm=H
            ),
            timeout=120.0,
            label="add_fillet",
        )
        print(f"\n④ 已倒圆角：{fillet.fillet_name}（{fillet.objects_filleted} 条边）")

        m2 = worker.call(
            lambda c: c.measure_body(document_name=doc_name, body_name=body_name),
            timeout=60.0,
            label="measure_body(fillet)",
        )
        print("\n⑤ 倒角后测量原始返回：")
        for k, v in asdict(m2).items():
            print(f"    {k:<22}: {v}")

        v_e2, a_e2, c_e2 = _fillet_exact()
        checks2: list = []
        _check_scalar("体积", m2.volume_mm3, v_e2, checks2)
        _check_scalar("表面积", m2.area_mm2, a_e2, checks2)
        _check_cog(m2.cog_mm, c_e2, checks2)
        all_ok &= _report("   逐项比对（倒角后，重心应原地不动）：", checks2)

    except Exception:  # noqa: BLE001
        print("❌ 失败。完整栈：", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        worker.stop()

    print(
        "\n判定：",
        "✅ 测量族合格（体积 / 表面积 / 重心三条独立证据均与精确解吻合）"
        if all_ok
        else "⚠️ 需检查（重点看单位换算是否差 10^3/10^6，以及 cog_strategy）",
    )
    return 0 if all_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
