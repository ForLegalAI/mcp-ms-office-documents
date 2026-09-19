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


def _usage_table():
    rows = [
        Tr(
            Td(st.name), Td(st.kind), Td(str(st.calls)),
            Td(str(st.errors), cls="num-err" if st.errors else ""),
            Td(fmt_ts(st.last_called)),
            Td(st.last_error or "—", cls="msg"),
        )
        for st in metrics.tool_stats()
    ]
    if not rows:
        return P("No template tools have been called yet this session.", cls="muted")
    return c.data_table(
        ["Tool", "Kind", "Calls", "Errors", "Last used", "Last error"], rows)


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

    stats = c.stats_row(
        c.stat("Uptime", fmt_uptime(time.time() - metrics.START_TIME)),
        c.stat("Upload backend", ctx.config.storage.strategy.value),
        c.stat("Live Word tools", len(live_docx)),
        c.stat("Live Email tools", len(live_email)),
        c.stat("Errors logged", err_count, num_cls="num-err" if err_count else ""),
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
        c.card(_usage_table(), title="Template usage (this session)", level=2),
        c.card(toggle, _log_block(errors_only),
               title="Recent activity & errors", level=2),
    )
