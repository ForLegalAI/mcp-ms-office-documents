"""Shared view primitives and the self-contained theme for the admin UI.

Every admin page is built from the helpers here rather than from hand-written
FastHTML trees, so "what a card looks like" or "how a form field is laid out"
is decided once. Three properties the module exists to keep:

* **No external assets.** :data:`ADMIN_CSS` and :data:`ARG_ROWS_JS` are inlined
  into every page by :func:`head_tags`. Nothing here may emit a ``<link>`` to
  another origin, a ``<script src>`` or a ``url(https://…)`` — the UI has to
  look right offline and in locked-down deployments, which is also why
  :mod:`admin.app` builds its ``FastHTML`` with ``default_hdrs=False``.
  ``tests/test_admin_assets.py`` enforces this on every rendered page.
* **Every control is labelled.** :func:`field` mints an ``id`` and points the
  ``<label>`` at it. That is the only reason the UI has label/control
  associations at all, so build a field with this rather than by hand.
* **Wide tables scroll, the page does not.** :func:`data_table` wraps its table
  in a scroll container; a six-column table must not force the whole document
  sideways on a phone.

The theme is token-based: colours are custom properties on ``:root``,
redefined under ``prefers-color-scheme: dark``. A new rule takes a token, never
a literal, or it will be wrong in one of the two themes.

:mod:`admin.views` builds the pages from these; :mod:`admin.app` only wires
them to routes.
"""
from __future__ import annotations

import re
from typing import Any, Dict, Iterable, Optional, Sequence, Tuple

from fasthtml.common import (
    A, Details, Div, Form, H1, H2, H3, Header, Hidden, Input, Label,
    Main, Meta, Nav, NotStr, P, Script, Span, Style, Summary, Table, Tbody,
    Th, Thead, Title, Tr,
)

# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------

ADMIN_CSS = """
:root{
  color-scheme:light;
  --bg:#f1f5f9; --card:#ffffff; --ink:#0f172a; --muted:#64748b; --line:#e2e8f0;
  --brand:#4f46e5; --brand-d:#4338ca; --on-brand:#ffffff; --link:#4f46e5;
  --topbar-bg:#4f46e5; --topbar-ink:#ffffff; --topbar-link:#e0e7ff;
  --ctl-bg:#ffffff; --ctl-disabled-bg:#f8fafc; --btn-bg:#ffffff;
  --accent-bg:#eef2ff; --accent-ink:#4338ca;
  --neutral-bg:#f1f5f9; --row-line:#f1f5f9;
  --chip-if-bg:#fef9c3; --chip-if-ink:#854d0e;
  --ok:#16a34a; --warn:#b45309; --err:#dc2626; --info:#2563eb;
  --ok-bg:#ecfdf5; --warn-bg:#fffbeb; --err-bg:#fef2f2; --info-bg:#eff6ff;
  --ok-line:#a7f3d0; --warn-line:#fde68a; --err-line:#fecaca; --info-line:#bfdbfe;
  --ok-ink:#065f46; --warn-ink:#854d0e; --err-ink:#991b1b; --info-ink:#1e3a8a;
  --shadow:0 1px 2px rgba(15,23,42,.04); --shadow-bar:0 1px 3px rgba(0,0,0,.15);
}
@media (prefers-color-scheme:dark){
  :root{
    color-scheme:dark;
    --bg:#0f172a; --card:#1e293b; --ink:#e2e8f0; --muted:#94a3b8; --line:#334155;
    --brand:#6366f1; --brand-d:#818cf8; --on-brand:#ffffff; --link:#a5b4fc;
    --topbar-bg:#312e81; --topbar-ink:#ffffff; --topbar-link:#c7d2fe;
    --ctl-bg:#0f172a; --ctl-disabled-bg:#172033; --btn-bg:#1e293b;
    --accent-bg:#312e81; --accent-ink:#c7d2fe;
    --neutral-bg:#334155; --row-line:#273449;
    --chip-if-bg:#422006; --chip-if-ink:#fde68a;
    --ok:#4ade80; --warn:#fbbf24; --err:#f87171; --info:#60a5fa;
    --ok-bg:#14321f; --warn-bg:#3a2a0a; --err-bg:#3b1414; --info-bg:#16233d;
    --ok-line:#166534; --warn-line:#854d0e; --err-line:#991b1b; --info-line:#1e40af;
    --ok-ink:#86efac; --warn-ink:#fcd34d; --err-ink:#fca5a5; --info-ink:#bfdbfe;
    --shadow:0 1px 2px rgba(0,0,0,.3); --shadow-bar:0 1px 3px rgba(0,0,0,.45);
  }
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
  font-size:15px;line-height:1.5}
a{color:var(--link);text-decoration:none}
a:hover{text-decoration:underline}
.topbar{background:var(--topbar-bg);color:var(--topbar-ink);padding:.75rem 1.25rem;
  display:flex;align-items:center;justify-content:space-between;gap:1rem;
  flex-wrap:wrap;box-shadow:var(--shadow-bar)}
.topbar .brand{font-weight:700;font-size:1.05rem;color:var(--topbar-ink);
  display:flex;align-items:center;gap:.5rem}
.topbar nav a{color:var(--topbar-link);margin-left:1rem;font-size:.92rem}
.topbar nav a:hover{color:var(--topbar-ink)}
.container{max-width:980px;margin:1.5rem auto;padding:0 1.25rem}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;
  padding:1.25rem 1.5rem;margin-bottom:1.25rem;box-shadow:var(--shadow)}
h1{font-size:1.5rem;margin:.2rem 0 1rem}
h2{font-size:1.15rem;margin:0 0 .75rem}
h3{font-size:1rem;margin:1.25rem 0 .5rem}
.muted{color:var(--muted);font-size:.9rem}
.field{margin-bottom:1rem}
.field label{display:block;font-weight:600;font-size:.85rem;margin-bottom:.3rem}
input,select,textarea{width:100%;padding:.5rem .6rem;border:1px solid var(--line);
  border-radius:8px;font:inherit;background:var(--ctl-bg);color:var(--ink)}
input:focus,select:focus,textarea:focus{outline:2px solid var(--brand);
  outline-offset:-1px;border-color:var(--brand)}
input[disabled]{background:var(--ctl-disabled-bg);color:var(--muted)}
textarea{resize:vertical}
.btn{display:inline-flex;align-items:center;gap:.4rem;cursor:pointer;border-radius:8px;
  padding:.5rem .9rem;font:inherit;font-weight:600;border:1px solid var(--line);
  background:var(--btn-bg);color:var(--ink);text-decoration:none}
.btn:hover{text-decoration:none}
.btn-primary{background:var(--brand);color:var(--on-brand);border-color:var(--brand)}
.btn-primary:hover{background:var(--brand-d);border-color:var(--brand-d);color:var(--on-brand)}
.btn-secondary{background:var(--btn-bg);color:var(--link);border-color:var(--brand)}
.btn-secondary:hover{background:var(--accent-bg)}
.btn-danger{background:var(--btn-bg);color:var(--err);border-color:var(--err-line)}
.btn-danger:hover{background:var(--err-bg)}
.btn-sm{padding:.3rem .55rem;font-size:.82rem}
.btn-icon{padding:.25rem .5rem;background:var(--btn-bg);border:1px solid var(--line);color:var(--err)}
.btn-icon:hover{background:var(--err-bg);border-color:var(--err-line)}
.actions{display:flex;gap:.6rem;flex-wrap:wrap;align-items:center}
.table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
table{width:100%;border-collapse:collapse}
th,td{text-align:left;padding:.55rem .5rem;border-bottom:1px solid var(--line);vertical-align:middle}
th{font-size:.78rem;text-transform:uppercase;letter-spacing:.03em;color:var(--muted)}
.tpl-table td{font-size:.95rem}
.args-table td{padding:.3rem .35rem;border-bottom:none}
.args-table th{padding-bottom:.25rem}
.args-table input,.args-table select{padding:.4rem .5rem;font-size:.9rem}
.col-name{width:22%}.col-type{width:12%}.col-req{width:13%}.col-def{width:16%}.col-x{width:38px}
.badge{display:inline-block;padding:.12rem .5rem;border-radius:999px;font-size:.74rem;
  font-weight:600;line-height:1.5}
.badge-live{background:var(--ok-bg);color:var(--ok)}
.badge-off{background:var(--neutral-bg);color:var(--muted)}
.badge-ro{background:var(--neutral-bg);color:var(--muted)}
.badge-if{background:var(--accent-bg);color:var(--accent-ink);margin-left:.4rem}
.chip{display:inline-block;padding:.15rem .55rem;margin:.15rem .25rem .15rem 0;border-radius:6px;
  background:var(--accent-bg);color:var(--accent-ink);font-size:.82rem;
  font-family:ui-monospace,Menlo,monospace}
.chip-if{background:var(--chip-if-bg);color:var(--chip-if-ink)}
.flash{padding:.6rem .85rem;border-radius:8px;margin:.5rem 0;border:1px solid transparent}
.flash-info{background:var(--info-bg);border-color:var(--info-line);color:var(--info-ink)}
.flash-ok{background:var(--ok-bg);border-color:var(--ok-line);color:var(--ok-ink)}
.flash-warn{background:var(--warn-bg);border-color:var(--warn-line);color:var(--warn-ink)}
.flash-err{background:var(--err-bg);border-color:var(--err-line);color:var(--err-ink)}
.empty{text-align:center;padding:2rem 1rem;color:var(--muted)}
/* PowerPoint template views */
.role-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:0 1rem}
.facts .field{margin-bottom:.6rem}
.warn-text{color:var(--warn)}
.swatch-wrap{display:inline-flex;align-items:center;gap:.35rem;margin:.15rem .6rem .15rem 0}
.swatch{display:inline-block;width:14px;height:14px;border-radius:3px;border:1px solid var(--line)}
.swatch-label{font-size:.78rem;color:var(--muted);font-family:ui-monospace,Menlo,monospace}
.field label input[type=checkbox]{width:auto;margin-right:.4rem;vertical-align:-2px}
details{margin:1rem 0;border:1px solid var(--line);border-radius:8px;padding:.5rem .85rem}
summary{cursor:pointer;font-weight:600}
.login-wrap{max-width:380px;margin:8vh auto}
.inline-form{display:inline}
hr{border:none;border-top:1px solid var(--line);margin:1.25rem 0}
.stats{display:flex;gap:1rem;flex-wrap:wrap;margin-bottom:1rem}
.stat{flex:1;min-width:150px;background:var(--card);border:1px solid var(--line);
  border-radius:10px;padding:.85rem 1rem}
.stat .num{font-size:1.55rem;font-weight:700;line-height:1.2}
.stat .lbl{font-size:.78rem;text-transform:uppercase;letter-spacing:.03em;color:var(--muted)}
.logs{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:.82rem;
  max-height:460px;overflow:auto;border:1px solid var(--line);border-radius:8px}
.logs table{width:100%}
.logs td{padding:.25rem .5rem;border-bottom:1px solid var(--row-line);white-space:nowrap}
.logs td.msg{white-space:normal;word-break:break-word}
.logs .ts{color:var(--muted)}
.lvl{font-weight:700}
.lvl-ERROR,.lvl-CRITICAL{color:var(--err)}
.lvl-WARNING{color:var(--warn)}
.lvl-INFO{color:var(--info)}
.lvl-DEBUG{color:var(--muted)}
.toggle-row{display:flex;gap:.5rem;align-items:center;margin-bottom:.6rem}
.num-err{color:var(--err)}
"""

# Vanilla JS for dynamic argument rows (add / remove) — avoids a CDN htmx dep.
ARG_ROWS_JS = """
function adminAddArgRow(){
  var tb=document.getElementById('argrows');
  if(!tb) return;
  tb.insertAdjacentHTML('beforeend', window.__ARG_ROW_HTML__);
}
function adminRemoveRow(btn){
  var tr=btn.closest('tr'); if(tr) tr.remove();
}
"""


def head_tags(blank_row_html: str):
    """The full set of ``<head>`` children: charset, viewport, theme and script.

    *blank_row_html* is a JSON-encoded argument row, injected as
    ``window.__ARG_ROW_HTML__`` so "Add argument" can clone it without a
    client-side templating library.

    Everything is inline by design — see the module docstring.
    """
    return (
        Meta(charset="utf-8"),
        Meta(name="viewport", content="width=device-width, initial-scale=1"),
        Style(ADMIN_CSS),
        Script(NotStr(f"window.__ARG_ROW_HTML__ = {blank_row_html};\n{ARG_ROWS_JS}")),
    )


# ---------------------------------------------------------------------------
# Page shell
# ---------------------------------------------------------------------------

def topbar(brand: str, links: Sequence[Tuple[str, str]] = ()):
    """The brand bar, with *links* as ``(label, href)`` pairs."""
    nav = Nav(*[A(label, href=href) for label, href in links]) if links else Span()
    return Header(Span(brand, cls="brand"), nav, cls="topbar")


def page(title: str, header, *content):
    """Wrap page *content* in the themed shell (title + topbar + container)."""
    return (
        Title(f"{title} · Template Admin"),
        header,
        Main(Div(*content, cls="container")),
    )


def flash(msg: Optional[str], kind: str = "info"):
    """A coloured notice, or ``None`` when there is no message to show.

    Returning ``None`` for an empty message lets callers splat the result into
    a page unconditionally; FastHTML drops ``None`` children.
    """
    if not msg:
        return None
    return Div(msg, cls=f"flash flash-{kind}")


def card(*content, title: Optional[str] = None, level: int = 3, cls: str = "card"):
    """A titled panel. *level* picks ``H2`` (section) or ``H3`` (sub-section)."""
    heading = {1: H1, 2: H2, 3: H3}[level]
    children = ([heading(title)] if title else []) + list(content)
    return Div(*children, cls=cls)


def action_bar(*controls):
    """The row of buttons that closes a form."""
    return Div(*controls, cls="actions")


# ---------------------------------------------------------------------------
# Form fields
# ---------------------------------------------------------------------------

_SLUG_RE = re.compile(r"[^a-z0-9]+")


def _slug(text: str) -> str:
    return _SLUG_RE.sub("-", str(text).strip().lower()).strip("-") or "field"


def control_id(control, label: str) -> str:
    """The ``id`` a control should carry, derived from its ``name``.

    Falls back to the label text so a control without a ``name`` (a disabled
    display-only input, say) is still associated with its label rather than
    silently losing the association.
    """
    existing = control.attrs.get("id")
    if existing:
        return str(existing)
    name = control.attrs.get("name")
    return f"f-{_slug(name)}" if name else f"f-{_slug(label)}"


def field(label: str, control, hint: Optional[str] = None, cls: str = "field",
          hint_cls: str = "muted"):
    """A labelled form control, with the label pointing at the control.

    This is the only place that associates a ``<label>`` with its input, so
    every field in the UI goes through it. The ``id`` is derived from the
    control's ``name`` (see :func:`control_id`); pass an explicit ``id=`` on the
    control when two fields on one page would otherwise collide.

    The ``id`` is set on *control* in place, so pass a freshly built element.
    Handing the same instance to two ``field()`` calls silently rewrites the
    first one's ``id`` and leaves its label pointing at nothing.

    *hint_cls* lets a hint be a warning ("no layout matches this role") rather
    than the usual muted aside.
    """
    cid = control_id(control, label)
    control.attrs["id"] = cid
    children = [Label(label, **{"for": cid}), control]
    if hint:
        children.append(P(hint, cls=hint_cls))
    return Div(*children, cls=cls)


def static_row(label: str, value, cls: str = "field"):
    """A read-only labelled row.

    Not a form field: there is no control for a label to point at, so this
    renders a bare ``<label>`` on purpose and must not be used for input.
    """
    return Div(Label(label), value, cls=cls)


def checkbox_field(name: str, label: str, checked: bool = False,
                   hint: Optional[str] = None, value: str = "1"):
    """A checkbox with its label wrapped around it, plus an optional hint.

    Wrapping is the association here — the label is the clickable text beside
    the box, so it needs no ``for``.
    """
    box = Input(name=name, type="checkbox", value=value, checked=checked,
                id=f"f-{_slug(name)}")
    children = [Label(box, f" {label}")]
    if hint:
        children.append(P(hint, cls="muted"))
    return Div(*children, cls="field")


def hidden(name: str, value: str):
    return Hidden(name=name, value=value if value is not None else "")


def csrf_input(token: str):
    return Hidden(name="csrf", value=token or "")


# ---------------------------------------------------------------------------
# Data display
# ---------------------------------------------------------------------------

def data_table(headers: Sequence[Any], rows: Sequence[Any],
               cls: str = "tpl-table", body_id: Optional[str] = None):
    """A table wrapped in a horizontal scroll container.

    The wrapper is the point: a wide table (the status page's six columns, the
    layout report's placeholder lists) scrolls inside its card instead of
    forcing the whole page sideways on a narrow screen.

    *headers* may be plain strings or ready-made ``Th`` cells when a column
    needs a width class. *body_id* puts an id on the ``<tbody>``, which the
    argument editor needs so its "Add argument" script can append to it.
    """
    head = [h if getattr(h, "tag", None) == "th" else Th(h) for h in headers]
    body = Tbody(*rows, id=body_id) if body_id else Tbody(*rows)
    return Div(Table(Thead(Tr(*head)), body, cls=cls), cls="table-wrap")


def badge(text: str, variant: str = "off", **kwargs):
    return Span(text, cls=f"badge badge-{variant}", **kwargs)


def status_badge(live: bool):
    return badge("● Live", "live") if live else badge("Not live", "off")


def chip(text: str, conditional: bool = False):
    """One monospace chip — a placeholder, a role, a literal bit of syntax."""
    return Span(text, cls="chip chip-if" if conditional else "chip")


def chips(items: Iterable[str], highlight: Optional[set] = None,
          empty: str = "none"):
    """Monospace chips for a list of names; *highlight* marks conditionals.

    Wrapped in a ``<span>`` so the group is one node; use :func:`chip` where a
    single chip goes straight into a cell and the wrapper would be noise.
    """
    items = list(items or [])
    if not items:
        return Span(empty, cls="muted")
    highlight = highlight or set()
    return Span(*[chip(it, it in highlight) for it in items])


def swatch(name: str, value: str):
    """One theme colour, shown rather than described."""
    return Span(
        Span(cls="swatch", style=f"background:{value}"),
        Span(f"{name} {value}", cls="swatch-label"),
        cls="swatch-wrap", title=f"{name}: {value}",
    )


def stat(label: str, value: Any, num_cls: str = ""):
    return Div(Div(str(value), cls=f"num {num_cls}".strip()),
               Div(label, cls="lbl"), cls="stat")


def stats_row(*tiles):
    return Div(*tiles, cls="stats")


def details_block(summary: str, *content):
    return Details(Summary(summary), *content)


def empty_state(message: str, *actions):
    return Div(P(message), *actions, cls="empty")


def post_form(action: str, *content, csrf: str = "", cls: str = "",
              enctype: Optional[str] = None):
    """A POST form that always carries the session CSRF token."""
    kwargs: Dict[str, Any] = {"action": action, "method": "post"}
    if cls:
        kwargs["cls"] = cls
    if enctype:
        kwargs["enctype"] = enctype
    return Form(csrf_input(csrf), *content, **kwargs)


__all__ = [
    "ADMIN_CSS", "ARG_ROWS_JS", "head_tags", "topbar", "page", "flash", "card",
    "action_bar", "field", "static_row", "checkbox_field", "hidden", "csrf_input",
    "data_table", "badge", "status_badge", "chip", "chips", "swatch",
    "stat", "stats_row", "details_block", "empty_state", "post_form",
    "control_id",
]
