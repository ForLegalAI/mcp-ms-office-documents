"""The page shell every admin view is rendered into.

Holds the pieces that are the same on every page — the top bar, the themed
wrapper, the "not found" page — plus the two form fragments the docx/email and
pptx edit forms both need. Those two used to be written out once per form and
had already drifted apart; the differences now arrive as arguments.

Views take the :class:`~admin.app.AdminContext` as a plain parameter rather than
importing it, so this package never imports :mod:`admin.app`.
"""
from __future__ import annotations

from typing import Optional


from fasthtml.common import A, Button, Div, Input, P

from admin import components as c
from admin.kinds import descriptor
from admin.sections import nav_items, section_for_kind

BRAND = "📄 Template Admin"


def page(ctx, title: str, *content, authed: bool = True, active: str = ""):
    """Wrap page *content* in the themed shell.

    *active* is the slug of the section the page belongs to, so the top bar
    can mark where you are. A page outside every section — a login screen, a
    delete confirmation reached from two places — passes nothing and simply
    has no item lit, which is honest rather than arbitrarily highlighting one.
    """
    end = [("Log out", ctx.u("/logout"))] if authed else []
    header = c.topbar(BRAND, nav_items(ctx.u, active=active, authed=authed),
                      end_links=end, home=ctx.u("/") if authed else "")
    return c.page(title, header, *content)


def kind_page(ctx, kind: str, title: str, *content, **kwargs):
    """A page belonging to a template *kind*, with its section lit in the nav.

    Every editor, confirmation and report for a template is reached from one
    section's Templates tab, so it keeps that section's nav item active rather
    than dropping the highlight the moment you click into a template.
    """
    kwargs.setdefault("active", section_for_kind(kind).slug)
    return page(ctx, title, *content, **kwargs)


def back_link(ctx, kind: str):
    """The "back to the list" link out of a template page, naming its section."""
    s = section_for_kind(kind)
    return A(f"← Back to {s.label} templates", href=s.href(ctx.u, "templates"))


def not_found_page(ctx, name: str, kind: str = ""):
    """Shown when an edit/reupload route names a template the store does not have."""
    back = (back_link(ctx, kind) if kind
            else A("← Back to the dashboard", href=ctx.u("/")))
    return page(
        ctx, "Not found",
        c.page_header("Not found"),
        c.flash(f"No managed template named '{name}'.", "err"),
        back,
        active=section_for_kind(kind).slug if kind else "",
    )


def name_field(kind: str, value: str, is_new: bool):
    """The template/tool name field, editable in both states.

    The name is the MCP tool name and the spec's filename stem, so changing it
    on an existing template is a rename: the save route writes the spec under
    the new name, unregisters the old tool and registers the new one. It used
    to be disabled here, which made a typo in a tool name permanent short of
    deleting the template and re-entering everything (#165).
    """
    d = descriptor(kind)
    hint = d.name_hint if is_new else (
        "Renaming changes the tool name the AI calls. The source file keeps "
        "its own filename."
    )
    return c.field(
        d.name_label,
        Input(name="name", value=value, required=True,
              placeholder=d.name_placeholder),
        hint=hint,
    )


def form_actions(ctx, kind: str, name: str = ""):
    """The Save / Preview / Cancel card that closes an edit form.

    *name* is empty while creating: there is nothing saved to clone yet.
    """
    d = descriptor(kind)
    controls = [
        Button("Save & make live", type="submit", cls="btn btn-primary"),
        Button(d.preview_label, type="submit",
               formaction=ctx.u(f"/{kind}/preview"),
               formtarget="_blank", cls="btn btn-secondary"),
    ]
    if d.has_args:
        # Not formtarget="_blank": this one goes to a form, not a document.
        controls.append(Button("Preview with my values…", type="submit",
                               formaction=ctx.u(f"/{kind}/preview/values"),
                               cls="btn btn-secondary"))
    if name:
        controls.append(A("Clone", href=ctx.u(f"/{kind}/{name}/clone"),
                          cls="btn btn-secondary",
                          title="Start a new template from this one"))
    controls.append(A("Cancel", href=section_for_kind(kind).href(ctx.u, "templates"),
                      cls="btn"))
    bar = c.action_bar(*controls)
    body = [bar]
    if d.preview_hint:
        body.append(P(d.preview_hint, cls="muted"))
    return c.card(*body)


def replace_card(ctx, kind: str, name: str, csrf: str = "",
                 asset: Optional[str] = None):
    """Download the installed source file, or upload a new one over it."""
    d = descriptor(kind)
    body = [
        P("Upload a new version of the source file. We'll re-scan it for placeholders "
          "and keep the arguments you've already configured.", cls="muted"),
        c.post_form(
            ctx.u(f"/{kind}/{name}/reupload"),
            Div(Input(name="file", type="file", accept=d.accept, required=True),
                cls="field"),
            Button("Upload & re-scan", type="submit", cls="btn btn-secondary"),
            csrf=csrf, enctype="multipart/form-data",
        ),
    ]
    if asset:
        body.append(Div(
            A(f"⭳ Download {asset}", href=ctx.u(f"/{kind}/{name}/download"),
              cls="btn btn-secondary btn-sm"),
            P("The file this template is using right now — useful before "
              "replacing it, or to copy it to another server.", cls="muted"),
            cls="table-actions",
        ))
    return c.card(*body, title="Source file")


def save_failed_page(ctx, message: str, kind: str = ""):
    """Shown when a spec could not be persisted at all."""
    back = (back_link(ctx, kind) if kind
            else A("← Back to the dashboard", href=ctx.u("/")))
    return page(
        ctx, "Save failed",
        c.page_header("Save failed"),
        c.flash(message, "err"),
        back,
        active=section_for_kind(kind).slug if kind else "",
    )


def saved_page(ctx, kind: str, name: str, ok: bool, enabled: bool = True):
    """Confirmation after a save, worded for what the kind actually produces.

    Three outcomes, not two. A template saved while disabled is not live, but
    that is what was asked for — wording it as a failed registration sends the
    admin to read logs about something that was never going to happen (#164).
    """
    d = descriptor(kind)
    s = section_for_kind(kind)
    if not enabled:
        template, heading, tone = d.save_disabled, "✓ Saved", "ok"
    elif ok:
        template, heading, tone = d.save_ok, "✓ Saved", "ok"
    else:
        template, heading, tone = d.save_warn, "Saved with a warning", "warn"
    return kind_page(
        ctx, kind, "Saved",
        c.page_header(heading),
        c.flash(template.format(name=name), tone),
        c.action_bar(
            A(f"Back to {s.label} templates", href=s.href(ctx.u, "templates"),
              cls="btn btn-primary"),
            A("Keep editing", href=ctx.u(f"/{kind}/{name}/edit"), cls="btn btn-secondary"),
        ),
    )
