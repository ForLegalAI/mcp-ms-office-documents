"""Lightweight in-process metrics for the admin Status view.

Captures two cheap things with no external dependencies:

* **Per-template usage** — call/error counters with last-used and last-error,
  recorded by the dynamic docx/email tool wrappers (always-on; a couple of dict
  writes under a lock).
* **Recent logs** — a bounded ring buffer of recent log records, populated by a
  logging handler that the admin app installs at startup (only when the admin
  UI is enabled, so there is zero cost otherwise).

This lives at the project root (a sibling of ``template_registry.py``) so the
core tool modules can import it without depending on the optional ``admin``
package.
"""
from __future__ import annotations

import logging
import threading
import time
from collections import deque
from dataclasses import dataclass, field
from typing import Deque, Dict, List, Optional

_LOCK = threading.Lock()

# Process start time, used to show uptime on the status page.
START_TIME = time.time()


@dataclass
class ToolStat:
    """Usage counters for a single tool, static or template-backed."""
    name: str
    kind: str
    calls: int = 0
    errors: int = 0
    last_called: Optional[float] = None
    last_error: Optional[str] = None
    last_error_at: Optional[float] = None
    #: Warnings reported by finished builds, counted per severity. A build
    #: that succeeds while substituting a style or dropping a row is a success
    #: as far as ``calls`` is concerned — this is the only place it shows.
    warnings_by_severity: Dict[str, int] = field(default_factory=dict)
    last_warnings: List[str] = field(default_factory=list)
    last_warning_at: Optional[float] = None

    @property
    def warnings(self) -> int:
        """Every warning reported, across severities."""
        return sum(self.warnings_by_severity.values())

    @property
    def degraded(self) -> int:
        """Warnings that mean the file is not what was asked for.

        ``info`` is a substitution the caller will not mind, so it is counted
        but kept out of the number an operator scans for.
        """
        return sum(count for severity, count in self.warnings_by_severity.items()
                   if severity != "info")


_TOOL_STATS: Dict[str, ToolStat] = {}


def record_call(kind: str, name: str) -> None:
    """Record a successful invocation of template tool *name*."""
    with _LOCK:
        st = _TOOL_STATS.get(name)
        if st is None:
            st = _TOOL_STATS[name] = ToolStat(name=name, kind=kind)
        st.kind = kind
        st.calls += 1
        st.last_called = time.time()


def record_error(kind: str, name: str, message: str) -> None:
    """Record a failed invocation of template tool *name* with its *message*."""
    with _LOCK:
        st = _TOOL_STATS.get(name)
        if st is None:
            st = _TOOL_STATS[name] = ToolStat(name=name, kind=kind)
        st.kind = kind
        st.errors += 1
        st.last_error = (message or "")[:500]
        st.last_error_at = time.time()


#: How many of a tool's most recent warnings to keep for the Status page.
LAST_WARNINGS_KEPT = 10


def record_warnings(kind: str, name: str, warnings) -> None:
    """Record the warnings a finished build reported.

    *warnings* is anything iterable of objects carrying ``severity`` and a
    readable ``str()`` — a :class:`warning_channel.WarningChannel`, its
    ``records()``, or PowerPoint's own list. Nothing is recorded for a build
    that reported none, so a quiet tool stays absent from the counters.

    Called for successful builds: a warning means the file was produced and
    something in it is not as asked. Failures go through :func:`record_error`.
    """
    entries = list(warnings or [])
    if not entries:
        return
    with _LOCK:
        st = _TOOL_STATS.get(name)
        if st is None:
            st = _TOOL_STATS[name] = ToolStat(name=name, kind=kind)
        st.kind = kind
        for entry in entries:
            severity = getattr(entry, "severity", "warning") or "warning"
            st.warnings_by_severity[severity] = (
                st.warnings_by_severity.get(severity, 0) + 1)
        rendered = [str(entry)[:300] for entry in entries]
        # Newest build's warnings first, older ones behind them.
        st.last_warnings = (rendered + st.last_warnings)[:LAST_WARNINGS_KEPT]
        st.last_warning_at = time.time()


def degraded_total() -> int:
    """Warnings across every tool that mean a file is not as asked.

    Built on :attr:`ToolStat.degraded` rather than re-filtering severities, so
    what counts as degraded is decided in exactly one place.
    """
    with _LOCK:
        return sum(st.degraded for st in _TOOL_STATS.values())


def tool_stats() -> List[ToolStat]:
    """Return per-tool stats, most-recently-used first."""
    with _LOCK:
        return sorted(_TOOL_STATS.values(),
                      key=lambda s: (-(s.last_called or 0), s.name))


def get_tool_stat(name: str) -> Optional[ToolStat]:
    with _LOCK:
        return _TOOL_STATS.get(name)


def reset() -> None:
    """Clear all collected metrics (used by tests)."""
    global _LOG_HANDLER
    with _LOCK:
        _TOOL_STATS.clear()
    if _LOG_HANDLER is not None:
        _LOG_HANDLER.records.clear()


class RecentLogHandler(logging.Handler):
    """A logging handler that keeps the most recent records in memory."""

    def __init__(self, capacity: int = 300):
        super().__init__()
        self.records: Deque[dict] = deque(maxlen=capacity)

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.records.append({
                "time": record.created,
                "level": record.levelname,
                "levelno": record.levelno,
                "logger": record.name,
                "message": record.getMessage(),
            })
        except Exception:  # never let logging blow up the app
            pass


_LOG_HANDLER: Optional[RecentLogHandler] = None


def install_log_capture(level: int = logging.INFO, capacity: int = 300) -> RecentLogHandler:
    """Attach the ring-buffer handler to the root logger once. Idempotent."""
    global _LOG_HANDLER
    if _LOG_HANDLER is None:
        _LOG_HANDLER = RecentLogHandler(capacity=capacity)
        _LOG_HANDLER.setLevel(level)
        logging.getLogger().addHandler(_LOG_HANDLER)
    return _LOG_HANDLER


def recent_logs(min_level: int = logging.INFO, limit: int = 200) -> List[dict]:
    """Return recent log records at or above *min_level*, newest first."""
    if _LOG_HANDLER is None:
        return []
    items = [r for r in _LOG_HANDLER.records if r["levelno"] >= min_level]
    return list(reversed(items))[:limit]


def counts_by_level() -> Dict[str, int]:
    """Tally captured log records by level name (for the status summary)."""
    out: Dict[str, int] = {}
    if _LOG_HANDLER is None:
        return out
    for r in list(_LOG_HANDLER.records):
        out[r["level"]] = out.get(r["level"], 0) + 1
    return out
