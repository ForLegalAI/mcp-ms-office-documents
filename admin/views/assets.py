"""The Source files page: what is in ``custom_templates/`` (#166).

A maintenance view, not an editor. It answers the question nothing else in
the UI could: which uploaded files are still doing something, and which are
left over from a template that was deleted, renamed or replaced.

Only an orphan gets a Delete button. Everything else shows what points at it,
because "why can't I delete this one" is the question the page has to answer
without sending the admin off to read YAML.
"""
from __future__ import annotations

from typing import Any, List
from urllib.parse import quote

from fasthtml.common import A, Button, Code, H1, Li, P, Td, Tr, Ul

from admin import components as c
from admin.assets import AssetFile
from admin.views.shell import page
from admin.views.status import fmt_ts


def _href(ctx, filename: str) -> str:
    """The delete URL for *filename*, percent-encoded.

    `validate_asset_filename` deliberately allows spaces and non-ASCII —
    "Brand Deck.pptx" is an ordinary name — and a `#` or `?` in a filename
    would otherwise end the path early and point the link somewhere else
    entirely. Nothing but the path segment is being built here, so `safe=''`.
    """
    return ctx.u(f"/files/{quote(filename, safe='')}/delete")


def fmt_size(size: int) -> str:
    """Bytes in the unit a human would say them in."""
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.0f} KB"
    return f"{size / (1024 * 1024):.1f} MB"


def _references_cell(f: AssetFile):
    if f.orphaned:
        return c.badge("Unreferenced", "ro")
    if len(f.references) == 1:
        return f.references[0]
    return Ul(*[Li(r) for r in f.references])


def _row(ctx, f: AssetFile):
    return Tr(
        Td(Code(f.name)),
        Td(fmt_size(f.size)),
        Td(fmt_ts(f.mtime)),
        Td(_references_cell(f)),
        Td(A("Delete", href=_href(ctx, f.name),
             cls="btn btn-danger btn-sm") if f.orphaned else "—"),
    )


def assets_page(ctx, files: List[AssetFile], csrf: str = "",
                message: str = "", message_kind: str = "ok"):
    orphans = [f for f in files if f.orphaned]
    body: List[Any] = []
    if message:
        body.append(c.flash(message, message_kind))

    if not files:
        body.append(c.empty_state(
            "Nothing has been uploaded yet.",
            A("← Back to all templates", href=ctx.u("/"), cls="btn"),
        ))
    else:
        body.append(P(
            f"{len(files)} file{'s' if len(files) != 1 else ''} here, "
            f"{len(orphans)} unreferenced. A source file is kept when its "
            "template is deleted, and a rename or a replacement can strand "
            "one — these are what is left behind.",
            cls="muted"))
        body.append(c.data_table(
            ["File", "Size", "Modified", "Referenced by", "Actions"],
            [_row(ctx, f) for f in files],
        ))

    return page(
        ctx, "Source files",
        H1("Source files"),
        P("Every file in the uploads directory, and what still points at it. "
          "Only a file nothing references can be deleted here.", cls="muted"),
        c.card(*body),
    )


def delete_asset_page(ctx, filename: str, csrf: str = ""):
    """Confirm removing one orphan, naming the file that goes."""
    return page(
        ctx, f"Delete {filename}",
        H1("Delete this file?"),
        c.card(
            P("This removes the uploaded file:", cls="muted"),
            Ul(Li(Code(filename))),
            P("Nothing references it, so no template changes. It cannot be "
              "undone — you would have to upload the file again.", cls="muted"),
            c.post_form(
                _href(ctx, filename),
                c.action_bar(
                    Button("Delete this file", type="submit",
                           cls="btn btn-danger"),
                    A("Cancel", href=ctx.u("/files"), cls="btn"),
                ),
                csrf=csrf,
            ),
        ),
    )
