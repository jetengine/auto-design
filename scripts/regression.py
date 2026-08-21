"""回归测试：把全部冒烟脚本串成一条命令。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/regression.py              # 跑全部
    python scripts/regression.py --list       # 只看有哪些步骤
    python scripts/regression.py fillet pocket  # 只跑名字含关键字的步骤
    python scripts/regression.py --skip mcp   # 跳过慢的端到端

── 为什么需要它 ──
到目前为止每个能力都有独立冒烟脚本，且都**单独**验证通过过。问题是：
它们共用 `catia_client.py` 和 `com_worker.py`。改倒圆角时顺手动了
`part.InWorkObject`，Pocket 会不会被打断？谁也不知道，除非**每次都全跑一遍**。
一条命令跑不完的回归，等于没有回归。

── 设计取舍 ──
1. **子进程隔离**：每个脚本单独起进程。一个脚本把 CATIA 搞挂了，
   下一个还能跑（且能自证「我是在什么状态下失败的」）。
   同进程串跑的话，第一个泄漏的 COM 引用会污染后面所有结论。
2. **不吞输出**：子进程的 stdout/stderr 直连终端，证据实时滚出来。
   捕获再回放会让 3 分钟的运行看起来像卡死了。
3. **硬超时**：CATIA 卡死时不给「等等看」的余地——超时即杀，记 TIMEOUT。
   回归脚本自己挂住是最糟的失败模式。
4. **前置失败即中止**：连不上 CATIA 时，后面 6 个脚本会吐出 6 份一模一样
   的连接错误，把真正的第一现场淹掉。所以 `hello_catia` 失败直接停。

`probe_export_formats.py` 不在回归里——它是**探测**当前机器支持哪些格式的
诊断工具，没有固定的通过标准，每台机器答案都不同。
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent


@dataclass(frozen=True)
class Step:
    name: str
    script: str
    timeout_s: float
    critical: bool  # 失败则中止后续（前置条件类）
    why: str


# 顺序有讲究：先证明链路通，再证明能力对，最后跑最慢的端到端。
STEPS: tuple[Step, ...] = (
    Step("connect", "hello_catia.py", 60.0, True, "COM 链路 / 版本 / 会话"),
    Step("health", "health_smoke.py", 120.0, False, "超时熔断 + 阻塞自诊断 + 重建恢复"),
    Step("box", "create_box_smoke.py", 180.0, False, "加料特征（Sketch + Pad）"),
    Step("pocket", "add_pocket_smoke.py", 180.0, False, "去料特征（偏移平面 + Pocket）"),
    Step("fillet", "add_fillet_smoke.py", 180.0, False, "几何引用（边拾取 + EdgeFillet）"),
    Step("measure", "measure_smoke.py", 180.0, False, "只读测量（体积/面积/重心 三独立证据）"),
    Step("family", "family_smoke.py", 300.0, False, "批量建族（规模化 + 单个失败不拖垮整批）"),
    Step("export", "export_step_smoke.py", 240.0, False, "交付（STEP，不可用则降级 IGES）"),
    Step("mcp", "mcp_client_smoke.py", 300.0, False, "MCP 端到端（起服务器 + 调工具）"),
)

# 子进程退出码的含义（与各冒烟脚本约定一致）
_VERDICT = {0: "PASS", 1: "FAIL", 2: "ENV"}


def _run(step: Step) -> tuple[str, float]:
    """跑一个步骤，返回 (判定, 耗时秒)。输出直连终端，不捕获。"""
    started = time.monotonic()
    try:
        proc = subprocess.run(
            [sys.executable, str(SCRIPTS_DIR / step.script)],
            timeout=step.timeout_s,
        )
    except subprocess.TimeoutExpired:
        return "TIMEOUT", time.monotonic() - started
    return _VERDICT.get(proc.returncode, f"EXIT{proc.returncode}"), time.monotonic() - started


def main() -> int:
    parser = argparse.ArgumentParser(description="CATIA MCP 回归测试")
    parser.add_argument("only", nargs="*", help="只跑名字含这些关键字的步骤")
    parser.add_argument("--skip", nargs="*", default=[], help="跳过名字含这些关键字的步骤")
    parser.add_argument("--list", action="store_true", help="列出步骤后退出")
    args = parser.parse_args()

    if args.list:
        print(f"{'步骤':<10}{'脚本':<26}{'超时':>6}  验证内容")
        for s in STEPS:
            print(f"{s.name:<10}{s.script:<26}{s.timeout_s:>5.0f}s  {s.why}")
        return 0

    selected = [
        s
        for s in STEPS
        if (not args.only or any(k.lower() in s.name for k in args.only))
        and not any(k.lower() in s.name for k in args.skip)
    ]
    if not selected:
        print("❌ 没有匹配到任何步骤，用 --list 看有哪些。", file=sys.stderr)
        return 2

    print(f"回归测试：{len(selected)} 个步骤，解释器 {sys.executable}")
    print("前置条件：Windows + CATIA 已启动。过程中请勿手动操作 CATIA 窗口。\n")

    results: list[tuple[Step, str, float]] = []
    aborted_at: str | None = None

    for i, step in enumerate(selected, start=1):
        if aborted_at is not None:
            results.append((step, "SKIP", 0.0))
            continue
        print(f"\n{'=' * 72}\n[{i}/{len(selected)}] {step.name} —— {step.why}\n{'=' * 72}")
        verdict, elapsed = _run(step)
        results.append((step, verdict, elapsed))
        if verdict != "PASS" and step.critical:
            aborted_at = step.name

    print(f"\n{'=' * 72}\n回归汇总\n{'=' * 72}")
    icons = {"PASS": "✅", "FAIL": "❌", "ENV": "🚫", "TIMEOUT": "⏱", "SKIP": "⏭"}
    for step, verdict, elapsed in results:
        icon = icons.get(verdict, "❓")
        secs = f"{elapsed:6.1f}s" if verdict != "SKIP" else "     - "
        print(f"  {icon} {verdict:<8}{secs}  {step.name:<10}{step.why}")

    passed = sum(1 for _, v, _ in results if v == "PASS")
    total_s = sum(e for _, _, e in results)
    print(f"\n  {passed}/{len(results)} 通过，合计 {total_s:.1f}s")

    if aborted_at:
        print(
            f"\n⚠️ 前置步骤 `{aborted_at}` 失败，后续已中止——"
            "先解决它，否则后面的报错都只是它的回声。"
        )
    if passed != len(results):
        # 只列真正跑过且没过的：SKIP 是中止的连带结果，重跑它们没有意义。
        broken = [s for s, v, _ in results if v not in ("PASS", "SKIP")]
        if broken:
            print("\n失败的步骤请单独重跑看完整证据：")
            for step in broken:
                print(f"    python scripts/{step.script}")
        return 1

    print("\n✅ 全部通过。每个脚本各自留了一份 CATPart 在 CATIA 里，可手动关闭。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
