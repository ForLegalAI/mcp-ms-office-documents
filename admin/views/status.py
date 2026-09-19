"""The Status page: live counts, per-template usage, and a tail of the log."""
from __future__ import annotations

import logging
import time
from typing import Optional

from fasthtml.common import (
    A, Button, Div, Form, H1, Input, Option, P, Select, Span, Table, Tbody,
    Td, Tr,
)

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


def _warnings_card():
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


#: The level choices, coarsest last. Replaces a two-state All / Warnings
#: toggle that could not express "errors only" without also showing warnings.
LEVELS = (("debug", logging.DEBUG), ("info", logging.INFO),
          ("warning", logging.WARNING), ("error", logging.ERROR))
_LEVEL_NO = dict(LEVELS)

#: Auto-refresh intervals offered, in seconds. Off by default: a page that
#: reloads under you while you are reading it is worse than one you refresh.
REFRESH_CHOICES = ((0, "off"), (10, "10s"), (30, "30s"), (60, "60s"))

LOG_LIMIT = 200


def level_no(level: str) -> int:
    """The numeric level for a query-string value, defaulting to INFO."""
    return _LEVEL_NO.get((level or "").strip().lower(), logging.INFO)


def available_levels():
    """The level choices that can actually match something.

    The buffer captures at the server's configured level, so below it there is
    nothing to find. Offering "debug and above" on a server running at INFO
    would be a filter that always comes back empty and looks broken.
    """
    floor = metrics.capture_level()
    if floor is None:
        return LEVELS
    return tuple((name, no) for name, no in LEVELS if no >= floor) or LEVELS[-1:]


def _log_filters(ctx, level: str, source: str, search: str, refresh: int):
    """The filter bar. A plain GET form, so every view has a shareable URL."""
    levels = available_levels()
    level_opts = [Option(f"{name} and above", value=name,
                         selected=(name == level)) for name, _no in levels]
    source_opts = [Option("every source", value="", selected=not source)]
    source_opts += [Option(name, value=name, selected=(name == source))
                    for name in metrics.log_sources()]
    refresh_opts = [Option(label, value=str(value),
                           selected=(value == refresh))
                    for value, label in REFRESH_CHOICES]
    return Form(
        Div(
            c.field("Level", Select(*level_opts, name="level")),
            c.field("Source", Select(*source_opts, name="logger")),
            c.field("Search", Input(name="q", value=search,
                                    placeholder="message or logger name"),
                    cls="field grow"),
            c.field("Auto-refresh", Select(*refresh_opts, name="refresh")),
            Div(Button("Apply", type="submit", cls="btn btn-primary"),
                A("Reset", href=ctx.u("/status"), cls="btn"),
                cls="actions"),
            cls="filters",
        ),
        _capture_note(levels),
        action=ctx.u("/status"), method="get",
    )


def _capture_note(levels):
    """Say why `debug` is missing, rather than leaving it unexplained."""
    if any(name == "debug" for name, _no in levels):
        return None
    return P("Capturing at info and above. Set DEBUG=true on the server to "
             "record debug lines as well.", cls="muted")


def _log_block(level: str, source: str, search: str):
    records = metrics.recent_logs(level_no(level), limit=LOG_LIMIT,
                                  search=search, source=source)
    if not records:
        if search or source or level != "info":
            return P("No records match these filters.", cls="muted")
        return P("No log records captured yet.", cls="muted")
    rows = [
        Tr(
            Td(fmt_ts(r["time"]), cls="ts"),
            Td(r["level"], cls=f"lvl lvl-{r['level']}"),
            Td(r["logger"]),
            Td(r["message"], cls="msg"),
        )
        for r in records
    ]
    shown = P(f"{len(rows)} record(s)"
              + (f", newest {LOG_LIMIT}" if len(rows) == LOG_LIMIT else ""),
              cls="muted")
    # Already its own scroll container, so not c.data_table.
    return Div(shown, Div(Table(Tbody(*rows)), cls="logs"))


def status_page(ctx, level: str = "info", source: str = "",
                search: str = "", refresh: int = 0):
    live_docx = ctx.live_names(KIND_DOCX)
    live_email = ctx.live_names(KIND_EMAIL)
    lvl_counts = metrics.counts_by_level()
    err_count = lvl_counts.get("ERROR", 0) + lvl_counts.get("CRITICAL", 0)
    # Warnings that mean a file is not as asked; `info` is a substitution
    # nobody needs to chase, and metrics decides which is which.
    degraded = metrics.degraded_total()

    stats = c.stats_row(
        c.stat("Uptime", fmt_uptime(time.time() - metrics.START_TIME)),
        c.stat("Upload backend", ctx.config.storage.strategy.value),
        c.stat("Live Word tools", len(live_docx)),
        c.stat("Live Email tools", len(live_email)),
        c.stat("Errors logged", err_count, num_cls="num-err" if err_count else ""),
        c.stat("Warnings reported", degraded,
               num_cls="warn-text" if degraded else ""),
    )

    return page(
        ctx, "Status",
        H1("Status"),
        stats,
        c.card(_usage_table(), title="Tool usage (this session)", level=2),
        c.card(_warnings_card(), title="What builds worked around", level=2),
        c.card(_log_filters(ctx, level, source, search, refresh),
               _log_block(level, source, search),
               title="Recent activity", level=2),
        c.auto_refresh(refresh),
    )
