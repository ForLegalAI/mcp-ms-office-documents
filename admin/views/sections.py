"""The section shell, the Overview tab, and the dashboard that links to them.

:mod:`admin.sections` decides *what* the sections and their tabs are; this
module renders them. Three things live here:

* :func:`section_page` — the wrapper every tabbed page goes through, so the
  title block, the tab bar and the active nav item are decided once rather
  than per view;
* :func:`overview_panel` — the first tab of every document section, which
  answers "what can this server produce here, is it working, and what has it
  been doing" without the admin having to open the Server section; and
* :func:`dashboard_page` — the landing page, one tile per section.

The Overview tab is why Excel and XML have pages at all. Neither has a
template to manage — Excel has one base workbook, XML has nothing — but both
expose a tool, and a tool that no page in the admin UI mentions is a tool
nobody can check on.
"""
from __future__ import annotations

from dataclasses import dataclass, field as dc_field
from typing import Any, List, Optional, Sequence, Tuple

from fasthtml.common import A, Code, P, Span, Td, Tr

from admin import components as c
from admin.sections import Section, tab_items
from admin.views.shell import page


# ---------------------------------------------------------------------------
# The data an Overview renders
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ToolFact:
    """One MCP tool this section exposes, and what it has done this session."""

    name: str
    live: bool = True
    #: "static" for a tool always present, "template" for one a spec created.
    origin: str = "static"
    calls: int = 0
    errors: int = 0
    degraded: int = 0
    last_called: str = "—"


@dataclass(frozen=True)
class SlotFact:
    """One base template slot: what is installed and where it came from."""

    key: str
    label: str
    filename: str
    source: str          # "custom" / "default" / "none"


@dataclass(frozen=True)
class SectionFacts:
    """Everything :func:`overview_panel` needs, loaded by the route."""

    templates: int = 0
    live: int = 0
    disabled: int = 0
    tools: Sequence[ToolFact] = ()
    slots: Sequence[SlotFact] = ()
    notes: Sequence[Tuple[str, str]] = ()   # (message, flash kind)
    extra_stats: Sequence[Tuple[str, Any]] = dc_field(default_factory=tuple)

    @property
    def calls(self) -> int:
        return sum(t.calls for t in self.tools)

    @property
    def errors(self) -> int:
        return sum(t.errors for t in self.tools)

    @property
    def degraded(self) -> int:
        return sum(t.degraded for t in self.tools)


# ---------------------------------------------------------------------------
# The shell
# ---------------------------------------------------------------------------

def section_page(ctx, s: Section, tab: str, *content,
                 actions: Sequence[Any] = (), title: Optional[str] = None,
                 subtitle: Optional[str] = None):
    """The wrapper for every tabbed page.

    The page title is the section's, not the tab's: a tab is a view of one
    thing, and retitling the page per tab would make the browser history read
    as though you had moved between unrelated pages. The tab's own one-liner
    goes under the bar instead, where it explains the view you are looking at.
    """
    t = s.tab(tab)
    blurb = subtitle if subtitle is not None else s.blurb
    return page(
        ctx, f"{title or s.label} · {t.label}",
        c.page_header(s.title, blurb, *actions),
        c.tab_bar(tab_items(s, ctx.u, tab)),
        P(t.blurb, cls="tab-blurb") if t.blurb else None,
        *content,
        active=s.slug,
    )


# ---------------------------------------------------------------------------
# Overview
# ---------------------------------------------------------------------------

def _origin_badge(t: ToolFact):
    """Where a tool comes from — a classification, so no status dot."""
    if t.origin == "template":
        return c.badge("From a template", "off", plain=True)
    return c.badge("Always available", "off", plain=True)


def _tool_row(t: ToolFact):
    return Tr(
        Td(Code(t.name), cls="row-name"),
        Td(c.status_badge(t.live)),
        Td(_origin_badge(t)),
        Td(str(t.calls)),
        Td(str(t.errors), cls="num-err" if t.errors else ""),
        Td(str(t.degraded), cls="warn-text" if t.degraded else ""),
        Td(t.last_called),
    )


def _tools_card(s: Section, facts: SectionFacts):
    """What the AI can actually call in this section, and how it has gone.

    Counted per tool rather than in one total: a template tool erroring while
    the static tool is fine is a different problem from the reverse, and the
    Status page's single table cannot say which section either belongs to.
    """
    if not facts.tools:
        return c.card(
            P("This section exposes no tools of its own.", cls="muted"),
            title="Tools the AI can call", level=2)
    return c.card(
        c.data_table(
            ["Tool", "Status", "Origin", "Calls", "Errors", "Warnings",
             "Last used"],
            [_tool_row(t) for t in facts.tools]),
        P("Counts are for this server run; they reset when it restarts.",
          cls="muted"),
        title="Tools the AI can call", level=2)


def _slot_badge(source: str):
    if source == "custom":
        return c.badge("Custom", "live")
    if source == "default":
        return c.badge("Bundled default", "off")
    return c.badge("None installed", "ro")


def _slots_card(ctx, s: Section, facts: SectionFacts):
    """A summary of the base templates, with the tab that changes them."""
    if not facts.slots:
        return None
    rows = [
        Tr(Td(sl.label, cls="row-name"),
           Td(Code(sl.filename) if sl.filename else Span("—", cls="muted")),
           Td(_slot_badge(sl.source)))
        for sl in facts.slots
    ]
    base_tab = next((t.slug for t in s.tabs if t.slug == "base"), None)
    actions = []
    if base_tab:
        actions.append(A("Manage", href=s.href(ctx.u, base_tab),
                         cls="btn btn-secondary btn-sm"))
    return c.card(
        c.data_table(["Slot", "File in use", "Source"], rows),
        title="Base templates", level=2, actions=actions)


def _templates_card(ctx, s: Section, facts: SectionFacts):
    """The template count, and the way to the list."""
    if s.kind is None:
        return None
    tab_href = s.href(ctx.u, "templates")
    body = [c.stats_row(
        c.stat("Templates", facts.templates),
        c.stat("Live", facts.live),
        c.stat("Disabled", facts.disabled,
               num_cls="warn-text" if facts.disabled else ""),
    )]
    if facts.templates == 0:
        body = [P("None yet. A template turns a document you have already "
                  "designed into something the AI can fill in and return.",
                  cls="muted")]
    return c.card(
        *body,
        title="Templates", level=2,
        actions=[A("Open", href=tab_href, cls="btn btn-secondary btn-sm")])


def overview_panel(ctx, s: Section, facts: SectionFacts):
    """The Overview tab's body."""
    stats = [
        c.stat("Calls this session", facts.calls),
        c.stat("Errors", facts.errors,
               num_cls="num-err" if facts.errors else ""),
        c.stat("Warnings", facts.degraded,
               num_cls="warn-text" if facts.degraded else ""),
    ]
    stats += [c.stat(label, value) for label, value in facts.extra_stats]

    body: List[Any] = [c.stats_row(*stats)]
    body += [c.flash(msg, kind) for msg, kind in facts.notes]
    body.append(_tools_card(s, facts))
    body.append(_templates_card(ctx, s, facts))
    body.append(_slots_card(ctx, s, facts))
    if s.doc:
        body.append(P(
            Span("Reference: ", cls="muted"), Code(s.doc),
            Span(" in this repository.", cls="muted"), cls="muted"))
    return [b for b in body if b is not None]


def overview_page(ctx, s: Section, facts: SectionFacts):
    return section_page(ctx, s, s.default_tab, *overview_panel(ctx, s, facts))


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

def dashboard_page(ctx, tiles: Sequence[Tuple[Section, SectionFacts]]):
    """The landing page: one tile per section, with the numbers that matter.

    Replaces an index that was three tables stacked vertically — which showed
    every Word template before saying that PowerPoint existed, and never
    mentioned Excel at all.
    """
    cards = []
    for s, facts in tiles:
        facts_pairs: List[Tuple[Any, str]] = []
        if s.kind is not None:
            facts_pairs.append((facts.live, "live"))
        if s.tools:
            facts_pairs.append((facts.calls, "calls"))
        if facts.errors:
            facts_pairs.append((facts.errors, "errors"))
        cards.append(c.tile(s.href(ctx.u), s.title, s.blurb, facts_pairs))
    return page(
        ctx, "Dashboard",
        c.page_header(
            "Template Admin",
            "Everything this server can produce, one area at a time. Each "
            "area holds its templates, the base file they are built on, and "
            "how it has been used."),
        c.tile_grid(*cards),
    )
