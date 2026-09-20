"""The Base template tab: the files that style every generated document.

Deliberately not the dynamic-template editor. There is nothing to name, no
arguments to declare and no spec to write — each slot holds exactly one file,
replaced in place, with the bundled default underneath it.

Rendered as a *panel* inside whichever section owns the slot
(:attr:`admin.sections.Section.slot_keys`), not as a page of its own: Word's
base document belongs beside Word's templates, and the Excel style workbook
has nothing to do with either.
"""
from __future__ import annotations

from typing import Any, List, Optional

from fasthtml.common import A, Button, Div, Input, P, Span

from admin import components as c
from admin.analysis import PptxAnalysis
from admin.base_templates import BaseSlot
from admin.views.templates import pptx_analysis_report


def _state_badge(source: str, s: BaseSlot):
    if source == "custom":
        return c.badge("● Custom", "live")
    if source == "default":
        return c.badge("Bundled default", "off")
    return c.badge("None installed", "ro")


def _report(s: BaseSlot, analysis):
    """What the installed file offers, in the terms that matter for its kind."""
    if analysis is None:
        return P("Nothing installed to inspect.", cls="muted")
    if isinstance(analysis, PptxAnalysis):
        return pptx_analysis_report(analysis)

    items: List[Any] = []
    if s.analysis_kind == "docx":
        missing = analysis.missing_required_styles
        if missing:
            items.append(c.flash(
                "⚠ Missing Word styles the renderer uses: " + ", ".join(missing)
                + ". Content using them falls back to the document default.",
                "warn"))
        else:
            items.append(c.flash(
                "✓ Every style the markdown renderer applies is present.", "ok"))
        items.append(c.static_row(
            f"{len(analysis.styles_present)} styles defined",
            c.chips(analysis.styles_present[:40],
                    empty="none readable")))
    elif s.analysis_kind == "xlsx":
        items.append(c.static_row(
            "Referenceable as style:<Name>",
            c.chips(analysis.styles_present, empty="none")))
    else:  # email
        items.append(c.static_row("Mustache variables",
                                  c.chips(analysis.placeholders)))
        items.append(c.static_row("Sections", c.chips(analysis.conditionals)))

    for w in analysis.warnings:
        items.append(c.flash("⚠ " + w, "warn"))
    return Div(*[i for i in items if i is not None])


def _slot_card(ctx, s: BaseSlot, csrf: str, active, source, analysis,
               message: Optional[str] = None, message_kind: str = "ok"):
    body: List[Any] = [
        P(s.controls, cls="muted"),
        c.static_row("In use now", Span(
            f"{active.name} " if active else "nothing — ",
            _state_badge(source, s))),
    ]
    if message:
        body.append(c.flash(message, message_kind))
    body.append(c.details_block(s.report_title, _report(s, analysis)))

    actions = []
    if active is not None:
        actions.append(A(f"⭳ Download {active.name}",
                         href=ctx.u(f"/base/{s.key}/download"),
                         cls="btn btn-secondary btn-sm"))
    if source == "custom":
        # Only ever offered when a custom file exists: with none, there is
        # nothing to revert and the button would be a no-op that implies the
        # bundled default can be removed too.
        actions.append(c.post_form(
            ctx.u(f"/base/{s.key}/revert"),
            Button("Revert to bundled default" if s.has_default
                   else "Remove custom file",
                   type="submit", cls="btn btn-danger btn-sm"),
            csrf=csrf, cls="inline-form",
        ))
    if actions:
        body.append(Div(c.action_bar(*actions), cls="table-actions"))

    body.append(c.post_form(
        ctx.u(f"/base/{s.key}/upload"),
        c.field(f"Replace with a {s.accept} file",
                Input(name="file", type="file", accept=s.accept, required=True)),
        Button("Upload & analyse", type="submit", cls="btn btn-primary"),
        csrf=csrf, enctype="multipart/form-data",
    ))
    return c.card(*body, title=f"{s.icon} {s.label}", level=2)


def base_panel(ctx, csrf: str = "", states=(),
               focus: Optional[str] = None,
               message: Optional[str] = None,
               message_kind: str = "ok"):
    """One card per base slot. *states* is ``(slot, active_path, source, analysis)``.

    A panel rather than a page: the five slots used to be stacked on one
    screen, which meant replacing the Word document and replacing the Excel
    style workbook — two unrelated jobs with very different blast radii — were
    the same page. Each section now shows only its own, and a section route
    passes only the states it asked for.
    """
    return [
        _slot_card(ctx, s, csrf, active, source, analysis,
                   message=message if focus == s.key else None,
                   message_kind=message_kind)
        for s, active, source, analysis in states
    ]
