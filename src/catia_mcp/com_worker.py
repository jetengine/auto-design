"""单 STA 线程的 COM 执行器。

为什么需要它：
    CATIA 的 COM 对象是 **单元线程（STA）绑定** 的 —— 谁 CoInitialize 并创建了对象，
    就只能由谁来调用它。MCP server 的工具回调可能被框架调度到任意线程/线程池，
    直接跨线程碰 CATIA 对象会抛 COM 错误甚至崩溃。

设计：
    * 起一个专用线程，在其中 pythoncom.CoInitialize()（默认 STA）。
    * 该线程持有唯一的 CatiaClient 实例。
    * 所有对 CATIA 的调用都以「可调用对象」形式排队，由该线程串行执行。
    * 调用方通过 Future 拿结果 —— 天然实现了「单 STA 串行 COM」。

这样，无论上层有多少并发请求，落到 CATIA 的永远是单线程、串行、有序的。

为什么还要熔断（本文件的第二个职责）：
    单线程串行的代价是 **一个调用卡死 = 整条链路卡死**。CATIA 只要弹一个模态框
    （许可证提示、覆盖确认、错误对话框），COM 调用就永远不返回。这时：
      - 卡住的那次调用会 TimeoutError；
      - 但 STA 线程仍死在里面，后续每次调用都会排队 → 各自等满超时 → 全部失败，
        而且报的都是「超时」这种**没有信息量**的错。
    所以这里做三件事：
      1. 记录「当前正在跑哪个任务、跑了多久」；
      2. 一旦某个任务超时，把它标记为**阻塞源**；此后新调用**立即失败**并指名道姓，
         不再陪着白等（这就是熔断）；
      3. 提供 restart() 重建 STA 线程与 CATIA 连接，作为不重启进程的恢复手段。
"""

from __future__ import annotations

import queue
import sys
import threading
import time
from concurrent.futures import Future
from concurrent.futures import TimeoutError as FutureTimeoutError
from dataclasses import dataclass
from typing import Any, Callable, Optional

if sys.platform != "win32":
    raise RuntimeError("com_worker 只能在 Windows 上运行。")

import pythoncom  # type: ignore[import-not-found]

from .catia_client import CatiaClient


class CatiaLinkError(RuntimeError):
    """CATIA 链路层错误的基类（区别于 CATIA 自身抛的业务/COM 错误）。"""


class CatiaTimeoutError(CatiaLinkError):
    """某次调用超过预算仍未返回 —— 熔断触发点。"""


class CatiaBlockedError(CatiaLinkError):
    """链路已被前一个卡死的任务堵住，新调用直接快速失败，不再排队空等。"""


@dataclass
class RunningJob:
    """当前占用 STA 线程的任务。"""

    seq: int
    label: str
    started_at: float

    @property
    def age_s(self) -> float:
        return time.monotonic() - self.started_at


# 队列里放的元素：(序号, 标签, 要执行的函数, 对应的 Future)
_Job = tuple[int, str, Callable[[CatiaClient], Any], "Future[Any]"]


class ComWorker:
    """拥有一个专用 STA 线程和唯一 CATIA 连接的串行执行器。"""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._seq = 0
        self._current: Optional[RunningJob] = None
        self._blocked_by: Optional[RunningJob] = None  # 已判定卡死的任务
        self._jobs_done = 0
        self._restarts = 0
        self._jobs: "queue.Queue[Optional[_Job]]" = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._started = threading.Event()
        self._start_error: Optional[BaseException] = None
        self._client: Optional[CatiaClient] = None

    # ------------------------------------------------------------------
    # 生命周期
    # ------------------------------------------------------------------
    def start(self) -> None:
        """启动线程并在其中完成 COM 初始化 + 连接 CATIA。

        阻塞直到连接成功或失败，把启动错误如实抛给调用方。
        """
        self._spawn()
        self._started.wait()
        if self._start_error is not None:
            raise self._start_error

    def _spawn(self) -> None:
        self._started = threading.Event()
        self._start_error = None
        self._client = None
        self._jobs = queue.Queue()
        with self._lock:
            self._current = None
            self._blocked_by = None
        self._thread = threading.Thread(
            target=self._run, name=f"catia-sta-worker-{self._restarts}", daemon=True
        )
        self._thread.start()

    def restart(self) -> dict:
        """丢弃当前 STA 线程，重建 COM 连接。

        为什么是「丢弃」而不是「杀死」：卡在 COM 调用里的线程无法被安全终止
        （强杀会破坏 COM 运行时状态）。它是 daemon 线程，会一直挂着直到 CATIA
        那边的模态框被关掉，然后自行结束。我们只是**不再等它**，另起一个干净的
        线程重新连 CATIA —— 让链路恢复可用，代价是短暂多一个僵尸线程。

        返回本次重启前的健康快照，便于事后定位是谁把链路堵死的。
        """
        before = self.health()
        with self._lock:
            self._restarts += 1
        self._jobs.put(None)  # 给旧线程留个毒丸；它若还活着就会自己退出
        self._spawn()
        self._started.wait(timeout=60)
        if self._start_error is not None:
            raise self._start_error
        return before

    def stop(self) -> None:
        """请求线程退出。"""
        self._jobs.put(None)  # 毒丸
        if self._thread is not None:
            self._thread.join(timeout=5)

    # ------------------------------------------------------------------
    # 健康状态
    # ------------------------------------------------------------------
    def is_blocked(self) -> bool:
        """链路是否被一个已判定卡死的任务堵住（且它仍未返回）。"""
        with self._lock:
            return (
                self._blocked_by is not None
                and self._current is not None
                and self._current.seq == self._blocked_by.seq
            )

    def health(self) -> dict:
        """链路健康快照 —— 不碰 CATIA，纯本地状态，因此**卡死时也一定能返回**。

        这点很关键：健康检查如果自己也要走 COM，链路一卡它就跟着卡，
        那就永远问不出「到底出了什么事」。
        """
        with self._lock:
            current = self._current
            blocked_by = self._blocked_by
            jobs_done = self._jobs_done
            restarts = self._restarts
        thread = self._thread
        blocked = (
            blocked_by is not None
            and current is not None
            and current.seq == blocked_by.seq
        )
        return {
            "thread_alive": bool(thread is not None and thread.is_alive()),
            "connected": self._client is not None,
            "blocked": blocked,
            "current_job": current.label if current else None,
            "current_job_age_s": round(current.age_s, 2) if current else None,
            "blocked_by": blocked_by.label if blocked else None,
            "queue_depth": self._jobs.qsize(),
            "jobs_done": jobs_done,
            "restarts": restarts,
            "start_error": str(self._start_error) if self._start_error else None,
        }

    # ------------------------------------------------------------------
    # 提交任务
    # ------------------------------------------------------------------
    def submit(
        self, fn: Callable[[CatiaClient], Any], label: str = "job"
    ) -> "Future[Any]":
        """把一个 `fn(client) -> result` 排到 STA 线程执行，返回 Future。"""
        fut: "Future[Any]" = Future()
        with self._lock:
            self._seq += 1
            seq = self._seq
        self._jobs.put((seq, label, fn, fut))
        return fut

    def call(
        self,
        fn: Callable[[CatiaClient], Any],
        timeout: float = 30.0,
        label: str = "job",
    ) -> Any:
        """提交并同步等待结果 —— 供工具层直接使用。

        timeout 是安全边界：CATIA 卡死时不会无限阻塞。
        超时不只是「这次失败」，还会把链路标记为阻塞，后续调用快速失败（熔断）。
        """
        # 熔断：链路已知被堵，不必再排队等满 timeout
        with self._lock:
            blocker = self._blocked_by
            current = self._current
        if blocker is not None and current is not None and current.seq == blocker.seq:
            raise CatiaBlockedError(
                f"CATIA 链路已被卡死的任务「{blocker.label}」堵住"
                f"（已持续 {blocker.age_s:.0f}s），本次「{label}」不再排队等待。\n"
                "最常见原因：CATIA 前台弹出了模态对话框（许可证/覆盖确认/错误框），"
                "COM 调用永远不会返回。\n"
                "处理：切到 CATIA 窗口关闭对话框；若已关闭仍不恢复，"
                "调用 reconnect_catia 重建链路。"
            )

        fut = self.submit(fn, label=label)
        try:
            return fut.result(timeout=timeout)
        except FutureTimeoutError as exc:
            with self._lock:
                stuck = self._current
                if stuck is not None:
                    self._blocked_by = stuck  # 标记阻塞源，触发后续熔断
            stuck_desc = (
                f"，当前卡在「{stuck.label}」已 {stuck.age_s:.0f}s" if stuck else ""
            )
            raise CatiaTimeoutError(
                f"「{label}」超过 {timeout:.0f}s 未返回{stuck_desc}。\n"
                "CATIA 通常是被模态对话框挡住了（许可证/覆盖确认/错误框）。\n"
                "处理：切到 CATIA 窗口关闭对话框；必要时调用 reconnect_catia 重建链路。"
            ) from exc

    # ------------------------------------------------------------------
    # 线程主体
    # ------------------------------------------------------------------
    def _run(self) -> None:
        jobs = self._jobs  # 绑定到本线程自己的队列，restart 后旧线程不再抢新任务
        started = self._started
        try:
            pythoncom.CoInitialize()  # STA
            client = CatiaClient()
            client.connect()
            self._client = client
        except BaseException as exc:  # noqa: BLE001 —— 启动错误要原样上抛
            self._start_error = exc
            started.set()
            return

        started.set()

        try:
            while True:
                job = jobs.get()
                if job is None:  # 毒丸 —— 退出
                    break
                seq, label, fn, fut = job
                running = RunningJob(seq=seq, label=label, started_at=time.monotonic())
                with self._lock:
                    self._current = running
                try:
                    if fut.set_running_or_notify_cancel():
                        try:
                            fut.set_result(fn(self._client))  # type: ignore[arg-type]
                        except BaseException as exc:  # noqa: BLE001
                            fut.set_exception(exc)
                finally:
                    with self._lock:
                        self._jobs_done += 1
                        if self._current is not None and self._current.seq == seq:
                            self._current = None
                        # 卡死的任务终于回来了 —— 解除熔断
                        if self._blocked_by is not None and self._blocked_by.seq == seq:
                            self._blocked_by = None
        finally:
            pythoncom.CoUninitialize()
