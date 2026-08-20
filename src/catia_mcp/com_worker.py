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
"""

from __future__ import annotations

import queue
import sys
import threading
from concurrent.futures import Future
from typing import Any, Callable, Optional

if sys.platform != "win32":
    raise RuntimeError("com_worker 只能在 Windows 上运行。")

import pythoncom  # type: ignore[import-not-found]

from .catia_client import CatiaClient

# 队列里放的元素：(要执行的函数, 对应的 Future)
_Job = tuple[Callable[[CatiaClient], Any], "Future[Any]"]


class ComWorker:
    """拥有一个专用 STA 线程和唯一 CATIA 连接的串行执行器。"""

    def __init__(self) -> None:
        self._jobs: "queue.Queue[Optional[_Job]]" = queue.Queue()
        self._thread = threading.Thread(
            target=self._run, name="catia-sta-worker", daemon=True
        )
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
        self._thread.start()
        self._started.wait()
        if self._start_error is not None:
            raise self._start_error

    def stop(self) -> None:
        """请求线程退出。"""
        self._jobs.put(None)  # 毒丸
        self._thread.join(timeout=5)

    # ------------------------------------------------------------------
    # 提交任务
    # ------------------------------------------------------------------
    def submit(self, fn: Callable[[CatiaClient], Any]) -> "Future[Any]":
        """把一个 `fn(client) -> result` 排到 STA 线程执行，返回 Future。"""
        fut: "Future[Any]" = Future()
        self._jobs.put((fn, fut))
        return fut

    def call(self, fn: Callable[[CatiaClient], Any], timeout: float = 30.0) -> Any:
        """提交并同步等待结果 —— 供工具层直接使用。

        timeout 是「超时状态检查」安全边界的落点：CATIA 卡死时不会无限阻塞。
        """
        return self.submit(fn).result(timeout=timeout)

    # ------------------------------------------------------------------
    # 线程主体
    # ------------------------------------------------------------------
    def _run(self) -> None:
        try:
            pythoncom.CoInitialize()  # STA
            client = CatiaClient()
            client.connect()
            self._client = client
        except BaseException as exc:  # noqa: BLE001 —— 启动错误要原样上抛
            self._start_error = exc
            self._started.set()
            return

        self._started.set()

        try:
            while True:
                job = self._jobs.get()
                if job is None:  # 毒丸 —— 退出
                    break
                fn, fut = job
                if fut.set_running_or_notify_cancel():
                    try:
                        fut.set_result(fn(self._client))  # type: ignore[arg-type]
                    except BaseException as exc:  # noqa: BLE001
                        fut.set_exception(exc)
        finally:
            pythoncom.CoUninitialize()
