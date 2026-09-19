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


from fasthtml.common import A, Button, Div, H1, Input, P

from admin import components as c
from admin.kinds import descriptor, nav_links

BRAND = "📄 Template Admin"


def page(ctx, title: str, *content, authed: bool = True):
    """Wrap page *content* in the themed shell."""
    header = c.topbar(BRAND, nav_links(ctx.u, authed=authed))
    return c.page(title, header, *content)


def not_found_page(ctx, name: str):
    """Shown when an edit/reupload route names a template the store does not have."""
    return page(
        ctx, "Not found",
        H1("Not found"),
        c.flash(f"No managed template named '{name}'.", "err"),
        A("← Back to all templates", href=ctx.u("/")),
    )


def name_field(kind: str, value: str, is_new: bool):
    """The template/tool name field.

    Editable only while the template is being created: the name is the MCP tool
    name and the spec's filename stem, so changing it later is a rename, which
    the UI does not support yet (#165). When it is fixed the input is disabled
    for display and the value travels in a hidden field instead.
    """
    d = descriptor(kind)
    if is_new:
        return c.field(
            d.name_label,
            Input(name="name", value=value, required=True,
                  placeholder=d.name_placeholder),
            hint=d.name_hint,
        )
    return c.field(d.name_label, Input(value=value, disabled=True))


def form_actions(ctx, kind: str):
    """The Save / Preview / Cancel card that closes an edit form."""
    d = descriptor(kind)
    bar = c.action_bar(
        Button("Save & make live", type="submit", cls="btn btn-primary"),
        Button(d.preview_label, type="submit",
               formaction=ctx.u(f"/{kind}/preview"),
               formtarget="_blank", cls="btn btn-secondary"),
        A("Cancel", href=ctx.u("/"), cls="btn"),
    )
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


def save_failed_page(ctx, message: str):
    """Shown when a spec could not be persisted at all."""
    return page(
        ctx, "Save failed",
        H1("Save failed"),
        c.flash(message, "err"),
        A("← Back to all templates", href=ctx.u("/")),
    )


def saved_page(ctx, kind: str, name: str, ok: bool):
    """Confirmation after a save, worded for what the kind actually produces."""
    d = descriptor(kind)
    msg = (d.save_ok if ok else d.save_warn).format(name=name)
    return page(
        ctx, "Saved",
        H1("✓ Saved" if ok else "Saved with a warning"),
        c.flash(msg, "ok" if ok else "warn"),
        c.action_bar(
            A("Back to all templates", href=ctx.u("/"), cls="btn btn-primary"),
            A("Keep editing", href=ctx.u(f"/{kind}/{name}/edit"), cls="btn btn-secondary"),
        ),
    )
