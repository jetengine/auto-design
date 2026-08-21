"""链路健壮性冒烟测试 —— 验证「超时熔断 + 阻塞自诊断 + 重建恢复」。

运行方式（在 **Windows**、CATIA 已启动）：

    python scripts/health_smoke.py

为什么需要这个测试：
    单 STA 线程的代价是「一个调用卡死 = 整条链路卡死」。这在真实使用中一定会发生
    （CATIA 弹个模态框就够了）。本脚本**人为制造一次卡死**，验证系统的反应是否正确：

    1. 正常时  health() 应 ping_ok；
    2. 制造阻塞（占住 STA 线程 25s），一次短超时调用应抛 CatiaTimeoutError；
    3. 此后 health() 应如实报告 blocked=true 且指名阻塞源；
    4. 再来一次调用应**立即**失败（CatiaBlockedError，耗时 <1s），而不是又等满超时 —— 这就是熔断；
    5. restart() 后链路应恢复，ping 再次成功。

注意：这里用 time.sleep 占住 STA 线程来模拟卡死。它与真实的 COM 卡死在
「线程被占住、后续任务全部堵在队列里」这一点上完全等价，而且可控、可重复。
"""

from __future__ import annotations

import sys
import time
import traceback

BLOCK_SECONDS = 25.0   # 假卡死持续时间，必须远大于下面的超时预算
SHORT_TIMEOUT = 3.0    # 受害调用的超时预算
FAST_FAIL_BUDGET = 1.0 # 熔断后「立即失败」的耗时上限


def main() -> int:
    try:
        from catia_mcp.com_worker import (
            CatiaBlockedError,
            CatiaTimeoutError,
            ComWorker,
        )
    except RuntimeError as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 2

    checks: list[tuple[str, bool, str]] = []

    def record(name: str, ok: bool, detail: str = "") -> None:
        checks.append((name, ok, detail))
        print(f"  {'✅' if ok else '❌'} {name}" + (f" — {detail}" if detail else ""))

    worker = ComWorker()
    try:
        worker.start()
    except Exception as exc:  # noqa: BLE001
        print(f"❌ 连接 CATIA 失败：{exc}", file=sys.stderr)
        return 2

    try:
        print("① 正常态健康检查")
        h = worker.health()
        caption = worker.call(lambda c: c.ping(), timeout=10.0, label="ping")
        record("链路空闲", h["blocked"] is False and h["current_job"] is None, str(h))
        record("CATIA 可达", bool(caption), f"caption={caption!r}")

        print("\n② 人为制造卡死（占住 STA 线程 %.0fs）" % BLOCK_SECONDS)
        worker.submit(lambda c: time.sleep(BLOCK_SECONDS), label="fake_hang")
        time.sleep(0.5)  # 确保假卡死任务已经真正开跑

        t0 = time.monotonic()
        try:
            worker.call(lambda c: c.ping(), timeout=SHORT_TIMEOUT, label="victim")
            record("受害调用超时", False, "居然成功了，说明阻塞没生效")
        except CatiaTimeoutError as exc:
            record(
                "受害调用超时并给出可执行提示",
                "fake_hang" in str(exc),
                f"耗时 {time.monotonic() - t0:.1f}s，错误指名了阻塞源",
            )

        print("\n③ 卡死态下的健康检查（关键：它自己不能也跟着卡住）")
        t0 = time.monotonic()
        h = worker.health()
        elapsed = time.monotonic() - t0
        record("health 立即返回", elapsed < FAST_FAIL_BUDGET, f"耗时 {elapsed:.3f}s")
        record("如实报告 blocked", h["blocked"] is True, f"blocked_by={h['blocked_by']!r}")
        record(
            "报出卡了多久",
            (h["current_job_age_s"] or 0) > 0,
            f"current_job_age_s={h['current_job_age_s']}",
        )

        print("\n④ 熔断：后续调用应立即失败，而不是再等满一个超时")
        t0 = time.monotonic()
        try:
            worker.call(lambda c: c.ping(), timeout=60.0, label="after_block")
            record("熔断生效", False, "居然成功了")
        except CatiaBlockedError:
            elapsed = time.monotonic() - t0
            record(
                "熔断生效（快速失败）",
                elapsed < FAST_FAIL_BUDGET,
                f"耗时 {elapsed:.3f}s（预算 60s，实际没等）",
            )
        except CatiaTimeoutError:
            record("熔断生效", False, "退化成了普通超时，说明没识别出阻塞")

        print("\n⑤ 恢复：重建链路")
        before = worker.restart()
        caption = worker.call(lambda c: c.ping(), timeout=15.0, label="ping_after_restart")
        h = worker.health()
        record("重建前快照留证", before["blocked"] is True, f"blocked_by={before['blocked_by']!r}")
        record("重建后链路可用", bool(caption), f"caption={caption!r}")
        record("重建后不再阻塞", h["blocked"] is False, str(h))

    except Exception:  # noqa: BLE001
        print("❌ 意外失败。完整栈：", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        worker.stop()

    ok = all(passed for _, passed, _ in checks)
    failed = [name for name, passed, _ in checks if not passed]
    print(
        "\n判定：",
        "✅ 健壮性合格（超时熔断 / 阻塞自诊断 / 重建恢复 全部生效）"
        if ok
        else f"⚠️ 不合格，未通过项：{failed}",
    )
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
