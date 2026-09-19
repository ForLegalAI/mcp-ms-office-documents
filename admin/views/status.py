"""The Status page: live counts, per-template usage, and a tail of the log."""
from __future__ import annotations

import logging
import time
from typing import Optional

from fasthtml.common import A, Div, H1, P, Span, Table, Tbody, Td, Tr

import metrics
from admin import components as c
from admin.store import KIND_DOCX, KIND_EMAIL
from admin.views.shell import page


def fmt_ts(ts: Optional[float]) -> str:
    if not ts:
        return "—"
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts))


def fmt_uptime(seconds: float) -> str:
    seconds = int(max(0, seconds))
    d, rem = divmod(seconds, 86400)
    h, rem = divmod(rem, 3600)
    m, _ = divmod(rem, 60)
    if d:
        return f"{d}d {h}h {m}m"
    if h:
        return f"{h}h {m}m"
    return f"{m}m"


#: How a warning severity reads in the usage table, worst first.
_SEVERITY_ORDER = ("error", "warning", "info")
_SEVERITY_CLASS = {"error": "num-err", "warning": "warn-text", "info": "muted"}


def _warning_cell(st):
    """A tool's warnings, split by severity.

    A single total would hide the distinction that matters: `info` is a
    substitution the caller will not mind, while `error` means something they
    asked for is not in the file. Lumping them together means a noisy but
    harmless substitution masks a real loss.
    """
    if not st.warnings_by_severity:
        return Span("—", cls="muted")
    parts = []
    for severity in _SEVERITY_ORDER:
        count = st.warnings_by_severity.get(severity)
        if count:
            parts.append(Span(f"{count} {severity}",
                              cls=f"sev {_SEVERITY_CLASS[severity]}"))
    # A severity the table does not know about still has to show up.
    for severity, count in sorted(st.warnings_by_severity.items()):
        if severity not in _SEVERITY_ORDER and count:
            parts.append(Span(f"{count} {severity}", cls="sev"))
    return Span(*parts)


def _usage_table():
    rows = []
    for st in metrics.tool_stats():
        detail = st.last_error or (st.last_warnings[0] if st.last_warnings else "—")
        rows.append(Tr(
            Td(st.name), Td(st.kind), Td(str(st.calls)),
            Td(str(st.errors), cls="num-err" if st.errors else ""),
            Td(_warning_cell(st)),
            Td(fmt_ts(st.last_called)),
            Td(detail, cls="msg"),
        ))
    if not rows:
        return P("No tools have been called yet this session.", cls="muted")
    return c.data_table(
        ["Tool", "Kind", "Calls", "Errors", "Warnings", "Last used",
         "Most recent problem"], rows)


def _warnings_card(ctx):
    """What recent builds worked around, newest first.

    The counters say a tool is degrading; this says how. Without it the only
    way to find out is the server log, which is exactly where the warnings
    channel exists to stop this information living.
    """
    rows = []
    for st in metrics.tool_stats():
        for line in st.last_warnings:
            rows.append(Tr(Td(st.name), Td(line, cls="msg")))
    if not rows:
        return P("No build has reported a warning this session.", cls="muted")
    return c.data_table(["Tool", "Worked around"], rows)


def _log_block(errors_only: bool):
    min_level = logging.WARNING if errors_only else logging.INFO
    rows = [
        Tr(
            Td(fmt_ts(r["time"]), cls="ts"),
            Td(r["level"], cls=f"lvl lvl-{r['level']}"),
            Td(r["logger"]),
            Td(r["message"], cls="msg"),
        )
        for r in metrics.recent_logs(min_level, limit=150)
    ]
    if not rows:
        return P("No log records captured yet.", cls="muted")
    # Already its own scroll container, so not c.data_table.
    return Div(Table(Tbody(*rows)), cls="logs")


def status_page(ctx, level: str = "info"):
    live_docx = ctx.live_names(KIND_DOCX)
    live_email = ctx.live_names(KIND_EMAIL)
    lvl_counts = metrics.counts_by_level()
    err_count = lvl_counts.get("ERROR", 0) + lvl_counts.get("CRITICAL", 0)
    # Only the severities that mean the file is not as asked; `info` is a
    # substitution nobody needs to chase.
    severities = metrics.counts_by_severity()
    degraded = sum(count for severity, count in severities.items()
                   if severity != "info")

    stats = c.stats_row(
        c.stat("Uptime", fmt_uptime(time.time() - metrics.START_TIME)),
        c.stat("Upload backend", ctx.config.storage.strategy.value),
        c.stat("Live Word tools", len(live_docx)),
        c.stat("Live Email tools", len(live_email)),
        c.stat("Errors logged", err_count, num_cls="num-err" if err_count else ""),
        c.stat("Warnings reported", degraded,
               num_cls="warn-text" if degraded else ""),
    )

    errors_only = level == "error"
    toggle = Div(
        Span("Show:", cls="muted"),
        A("All", href=ctx.u("/status"),
          cls="btn btn-sm " + ("btn-secondary" if errors_only else "btn-primary")),
        A("Warnings & errors", href=ctx.u("/status?level=error"),
          cls="btn btn-sm " + ("btn-primary" if errors_only else "btn-secondary")),
        A("Refresh", href=ctx.u(f"/status{'?level=error' if errors_only else ''}"),
          cls="btn btn-sm"),
        cls="toggle-row",
    )

    return page(
        ctx, "Status",
        H1("Status"),
        stats,
        c.card(_usage_table(), title="Tool usage (this session)", level=2),
        c.card(_warnings_card(ctx), title="What builds worked around", level=2),
        c.card(toggle, _log_block(errors_only),
               title="Recent activity & errors", level=2),
    )
