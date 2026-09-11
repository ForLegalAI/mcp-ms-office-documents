"""Async-safe wrapper for blocking document-generation calls.

Every Office-document generator in this project (``markdown_to_word``,
``markdown_to_excel``, ``create_presentation``, ``create_eml``,
``create_xml_file``, and every dynamically-registered template tool) is a
synchronous, blocking function. They perform:

  * file I/O (opening .docx / .pptx templates, which are zip archives)
  * CPU-bound markdown / mustache rendering
  * synchronous network I/O (``requests.get`` for image downloads,
    ``boto3`` S3 uploads)

FastMCP runs on top of an asyncio event loop (uvicorn / Starlette).
Calling a blocking function directly from an ``async def`` MCP tool
handler freezes the loop for the full duration of the call — no other
request can be served, including the Kubernetes liveness / readiness
probes. Repeated probe timeouts cause kubelet to SIGTERM the pod, which
is the failure mode observed in EKS:

    ERROR: ASGI callable returned without completing response.
    ERROR: Cancel 0 running task(s), timeout graceful shutdown exceeded

(The "0 running tasks" line is the giveaway: there were no async tasks
because the work was running synchronously on the event loop itself.)

Whether blocking work is offloaded to a worker thread is controlled by
the ``RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED`` environment variable
(see ``config.Config.run_blocking_by_asyncio_thread_enabled``):

  * Enabled (default) — dispatch the sync function to a bounded,
    process-wide ``ThreadPoolExecutor`` (sized by
    ``RUN_BLOCKING_MAX_WORKERS``, see below). The event loop stays free
    to serve health probes and concurrent requests.
  * Disabled — call the sync function inline on the event loop. The
    loop is blocked until the call returns. Only useful for local
    debugging or to rule threading out of a regression. Dynamic tools
    run on the event loop in this mode too, since their handlers are
    ``async def`` wrappers around ``run_blocking``.

All tool call sites uniformly write ``await run_blocking(func, *args,
**kwargs)`` — the helper internally chooses the dispatch strategy based
on config, so call sites never branch on the flag.

This helper lives in its own module (not inside ``main.py``) so that
dynamically-registered tool modules can import it without creating a
circular dependency back through ``main``.
"""

from __future__ import annotations

import asyncio
import functools
import logging
import threading
from concurrent.futures import ThreadPoolExecutor
from typing import Callable, Optional, TypeVar

from config import get_config

logger = logging.getLogger(__name__)

T = TypeVar("T")


# ---------------------------------------------------------------------------
# Bounded ThreadPoolExecutor singleton
# ---------------------------------------------------------------------------
# Why a bounded executor (instead of `asyncio.to_thread`'s default)?
#
# `asyncio.to_thread` runs on Python's default executor whose
# ``max_workers = min(32, os.cpu_count() + 4)``. On EKS, ``os.cpu_count()``
# returns the **host's** core count (not the pod CPU limit), so the default
# can spawn ~32 worker threads on a 1-vCPU pod. For CPU-bound, pure-Python
# work (such as ``markdown_to_word`` parsing) the GIL serialises all of
# them onto a single core anyway, and each thread receives only a small
# slice of CPU. Concurrent requests pile up and every one of them takes
# 10×+ longer than it should — exceeding the client's timeout while the
# event loop stays "healthy".
#
# A bounded executor with a small ``max_workers`` (default 4 for a 1 vCPU
# pod) keeps concurrent work to a level the GIL can actually progress
# through:
#
#   * 1 thread holds the GIL doing CPU work (markdown parsing)
#   * up to 3 others can be in I/O (S3 upload / image download) — they
#     release the GIL while blocked on syscalls
#   * additional requests queue cleanly in the executor's work queue
#     instead of fanning out and starving each other
#
# Tunable via ``RUN_BLOCKING_MAX_WORKERS`` (see config.Config). Recreate
# the executor by deleting this module from ``sys.modules`` (tests do
# this through their reload helper).
# ---------------------------------------------------------------------------

_EXECUTOR: Optional[ThreadPoolExecutor] = None
_EXECUTOR_LOCK = threading.Lock()


def _get_executor() -> ThreadPoolExecutor:
    """Return the process-wide bounded executor, creating it on first use.

    Thread-safe double-checked locking pattern: the fast path requires no
    lock acquisition once the executor exists.
    """
    global _EXECUTOR
    if _EXECUTOR is None:
        with _EXECUTOR_LOCK:
            if _EXECUTOR is None:
                max_workers = get_config().run_blocking_max_workers
                _EXECUTOR = ThreadPoolExecutor(
                    max_workers=max_workers,
                    thread_name_prefix="run_blocking",
                )
                logger.info(
                    "[run_blocking] bounded ThreadPoolExecutor created max_workers=%d",
                    max_workers,
                )
    return _EXECUTOR


async def run_blocking(func: Callable[..., T], /, *args, **kwargs) -> T:
    """Run a synchronous callable, optionally on a bounded worker thread.

    Dispatch is governed by ``config.run_blocking_by_asyncio_thread_enabled``
    (env var ``RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED``):

    * **Enabled (default)** — the call is submitted to a process-wide
      ``ThreadPoolExecutor`` whose size is controlled by
      ``config.run_blocking_max_workers`` (default 4). The event loop
      remains responsive to health probes and concurrent requests while
      ``func`` runs; when the pool is full, additional calls queue
      cleanly instead of spawning unbounded threads that would all
      contend for the GIL.
    * **Disabled** — the call runs inline on the event loop, the
      original blocking behaviour. ``func``'s return value is returned
      without ever yielding to the loop.

    The signature is identical in both modes — call sites always write
    ``await run_blocking(func, *args, **kwargs)``.

    Args:
        func: The blocking function to execute.
        *args: Positional arguments forwarded to ``func``.
        **kwargs: Keyword arguments forwarded to ``func``.

    Returns:
        Whatever ``func`` returns.

    Raises:
        Any exception ``func`` raises propagates to the caller. In
        threaded mode the exception is marshalled back from the worker
        thread automatically by ``loop.run_in_executor``.
    """
    # Read the flag lazily, on every call, so that tests (or admins)
    # that mutate the config singleton between calls see the new value
    # without having to reload this module.
    if get_config().run_blocking_by_asyncio_thread_enabled:
        logger.debug("%s is running on bounded executor", func.__name__)
        # functools.partial binds kwargs cleanly; run_in_executor only
        # accepts positional args, so partial is required (not optional).
        bound = functools.partial(func, *args, **kwargs)
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(_get_executor(), bound)

    # Offload disabled: call inline on the event loop (blocks until done).
    return func(*args, **kwargs)
