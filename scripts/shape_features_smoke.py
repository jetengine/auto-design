"""修饰特征三兄弟冒烟测试 —— Chamfer / Shell / Draft。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/shape_features_smoke.py

三个特征各自新建一个 40×30×20 的长方体，互不干扰：

    ① Chamfer  12 条边倒 45°、边长 4 的斜角
    ② Shell    去掉顶面、壁厚 3 抽壳
    ③ Draft    四个侧面拔模 5°，底面为中性面

── 为什么三段必须互相隔离 ──
它们的难度是递增的（换刀 → 挑一个面 → 多引用协同），失败概率也递增。
如果串在一个模型上，第一个失败就看不到后两个的证据了，一轮真机往返只能修一个 bug。
分开建模，一次跑完就能拿到三份独立结论。

── 期望值在本文件里独立算一遍 ──
不 import 实现里的公式。测试和被测代码共用同一个错误公式，就等于什么都没测。
"""

from __future__ import annotations

import math
import sys
import traceback
from dataclasses import asdict

L, W, H = 40.0, 30.0, 20.0
CHAMFER_D = 4.0        # 45° 斜角边长
SHELL_T = 3.0          # 抽壳壁厚
DRAFT_A = 5.0          # 拔模角（度）
TOL = 1e-3


def chamfer_removed() -> float:
    """12 条边倒 45°、边长 d 的斜角后被切掉的体积（独立推一遍）。

    剩下的实体 = 内芯 + 6 块面板 + 12 段三棱柱 + 8 个角块（合成菱形十二面体 2d³）。
    """
    d = CHAMFER_D
    a, b, c = L - 2 * d, W - 2 * d, H - 2 * d
    kept = (
        a * b * c
        + 2 * d * (a * b + a * c + b * c)
        + 2 * d * d * (a + b + c)
        + 2 * d ** 3
    )
    return L * W * H - kept


def shell_removed() -> float:
    """去顶面、壁厚 t 抽壳后被挖掉的内腔体积。

    内腔四周各让出 t，底部让出 t，顶部开口不让 —— 所以是 (L−2t)(W−2t)(H−t)。
    """
    t = SHELL_T
    return (L - 2 * t) * (W - 2 * t) * (H - t)


def draft_delta(outward: bool) -> float:
    """四侧面拔模 a° 后的体积变化（带符号）。底面中性，顶面每边缩放 k = H·tan(a)。

    棱台用 Prismatoid 公式 V = H/6·(A_bot + 4A_mid + A_top)，对棱台是精确的。
    """
    k = H * math.tan(math.radians(DRAFT_A))
    sign = 1.0 if outward else -1.0
    return H * (sign * k * (L + W) + 4.0 * k * k / 3.0)


def _dump(title: str, obj) -> None:
    print(f"\n{title}")
    for key, value in asdict(obj).items():
        print(f"    {key:<24}: {value}")


def _check(label: str, measured, expected, tol: float = TOL) -> bool:
    if measured is None or expected is None:
        print(f"    ❌ {label}: 测不到（measured={measured}, expected={expected}）")
        return False
    ref = abs(expected) if expected else 1.0
    rel = abs(measured - expected) / ref
    ok = rel <= tol
    print(
        f"    {'✅' if ok else '❌'} {label}: 实测 {measured:.6f}  "
        f"理论 {expected:.6f}  rel={rel:.2e}"
    )
    return ok


def _run_chamfer(worker) -> bool:
    box = worker.call(
        lambda c: c.create_box(length_mm=L, width_mm=W, height_mm=H, part_name="ChamferDemo"),
        timeout=120.0, label="create_box(chamfer)",
    )
    res = worker.call(
        lambda c: c.add_chamfer(
            length_mm=CHAMFER_D, angle_deg=45.0,
            box_length_mm=L, box_width_mm=W, box_height_mm=H,
        ),
        timeout=120.0, label="add_chamfer",
    )
    _dump("① Chamfer 证据：", res)
    ok = bool(box.update_ok and res.update_ok)
    ok &= _check("去料体积", res.measured_removed_mm3, chamfer_removed())
    # 12 条实体边全部倒到，才说明边拾取真的生效了 —— 少一条体积就对不上，
    # 但把条数单独打出来，是为了体积对不上时能一眼分清「拾少了」还是「刀不对」。
    if res.objects_chamfered != 12:
        print(f"    ⚠️ 实际倒角边数 {res.objects_chamfered}，长方体应为 12")
    print(f"    ℹ️ 生效的方法/枚举：{res.strategy}")
    return ok


def _run_shell(worker) -> bool:
    box = worker.call(
        lambda c: c.create_box(length_mm=L, width_mm=W, height_mm=H, part_name="ShellDemo"),
        timeout=120.0, label="create_box(shell)",
    )
    res = worker.call(
        lambda c: c.add_shell(
            thickness_mm=SHELL_T,
            box_length_mm=L, box_width_mm=W, box_height_mm=H,
        ),
        timeout=180.0, label="add_shell",
    )
    _dump("② Shell 证据：", res)
    ok = bool(box.update_ok and res.update_ok)
    ok &= _check("内腔体积", res.measured_removed_mm3, shell_removed())

    # 挑面挑对了没有：顶面重心必须在 z=H，面积必须是 L×W。
    # 这一条是体积验证之外的**独立证据** —— 万一某个别的面碰巧也能凑出相近体积，
    # 它能立刻把「挑错面」和「壁厚算错」区分开。
    face = res.removed_face
    if face and face.cog_mm:
        ok &= _check("去掉面的重心 Z", face.cog_mm[2], H)
        ok &= _check("去掉面的面积", face.area_mm2, L * W)
    else:
        print("    ❌ 没拿到被去掉面的信息")
        ok = False
    if res.face_candidates != 6:
        print(f"    ⚠️ 拾到 {res.face_candidates} 个面，长方体应为 6")
    return ok


def _run_draft(worker) -> bool:
    box = worker.call(
        lambda c: c.create_box(length_mm=L, width_mm=W, height_mm=H, part_name="DraftDemo"),
        timeout=120.0, label="create_box(draft)",
    )
    res = worker.call(
        lambda c: c.add_draft(
            angle_deg=DRAFT_A,
            box_length_mm=L, box_width_mm=W, box_height_mm=H,
        ),
        timeout=180.0, label="add_draft",
    )
    _dump("③ Draft 证据：", res)
    ok = bool(box.update_ok and res.update_ok)

    out_v, in_v = draft_delta(True), draft_delta(False)
    print(f"    ℹ️ 两种情形的理论增量：上大下小 {out_v:+.3f} / 上小下大 {in_v:+.3f}")
    delta = res.measured_delta_mm3
    if delta is None:
        print("    ❌ 测不到体积变化")
        return False
    best = min((("outward", out_v), ("inward", in_v)),
               key=lambda p: abs(delta - p[1]) / abs(p[1]))
    ok &= _check(f"体积变化（判定为 {best[0]}）", delta, best[1])
    if res.matched_direction != best[0]:
        print(f"    ⚠️ 实现判定 {res.matched_direction}，本测试独立判定 {best[0]}")
        ok = False
    if res.faces_drafted != 4:
        print(f"    ⚠️ 实际拔模面数 {res.faces_drafted}，长方体应为 4")
    neutral = res.neutral_face
    if neutral and neutral.cog_mm:
        # 底面重心 Z 的理论值就是 0，相对误差没法算，这里直接用绝对判据。
        z = neutral.cog_mm[2]
        good = abs(z) <= 1e-6
        print(f"    {'✅' if good else '❌'} 中性面重心 Z: {z:.9f}（应为底面 0）")
        ok &= good
    return ok


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

    verdicts: dict = {}
    try:
        # 三段各自 try —— 前一个炸了不该埋掉后两个的证据。
        # 一轮真机往返很贵，要一次拿全。
        for name, runner in (
            ("chamfer", _run_chamfer),
            ("shell", _run_shell),
            ("draft", _run_draft),
        ):
            try:
                verdicts[name] = runner(worker)
            except Exception:  # noqa: BLE001
                print(f"\n❌ {name} 抛异常，完整栈：", file=sys.stderr)
                traceback.print_exc()
                verdicts[name] = False
    finally:
        worker.stop()

    print("\n" + "=" * 60)
    for name in ("chamfer", "shell", "draft"):
        print(f"  {'✅ PASS' if verdicts.get(name) else '❌ FAIL'}  {name}")
    ok = all(verdicts.get(n) for n in ("chamfer", "shell", "draft"))
    print(
        "\n判定：",
        "✅ 三个修饰特征全部合格（几何引用生效，体积与精确解吻合）"
        if ok
        else "⚠️ 有未通过项（看对应段落的 strategy 与 target_errors，"
             "那里写着实际试过哪些方法/枚举）",
    )
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
