"""FastHTML admin UI for managing dynamic docx/email templates.

Mounted in the same ASGI process as the FastMCP server (see
:func:`build_combined_app`) so that saving a template registers/updates its MCP
tool immediately — no restart. The UI is gated by a single shared password
(:mod:`admin.auth`) and persists templates through :class:`admin.store.TemplateStore`.

Authoring model (chosen with the maintainer): the user uploads a real ``.docx``
(or email ``.html``); the UI auto-detects placeholders/conditionals
(:mod:`admin.analysis`), pre-builds the argument form, and previews
(:mod:`admin.preview`) without ever hitting the upload backend.
"""
from __future__ import annotations

import hashlib
import json
import logging
from typing import Any, Dict, List, Optional

from fasthtml.common import (
    FastHTML, Title, Main, Header, Div, P, A, H1, H2, H3, Form, Input, Button,
    Label, Span, Table, Thead, Tbody, Tr, Th, Td, Select, Option, Textarea,
    Details, Summary, Meta, Style, Script, NotStr, to_xml,
    RedirectResponse, Response, HTMLResponse, Hidden,
)
from starlette.routing import Mount

from config import Config
from admin import auth
from admin.analysis import analyze, reconcile, Analysis
from admin.preview import sample_values, render_docx_preview, render_email_preview
from admin.store import FileTemplateStore, KIND_DOCX, KIND_EMAIL, TemplateStoreError, validate_name
from template_registry import gather_specs

logger = logging.getLogger(__name__)

KINDS = (KIND_DOCX, KIND_EMAIL)
_KIND_LABEL = {KIND_DOCX: "Word", KIND_EMAIL: "Email"}
_KIND_ICON = {KIND_DOCX: "📝", KIND_EMAIL: "✉️"}
_ARG_TYPES = ["string", "int", "float", "bool", "list"]
# Style-mapping keys surfaced in the UI (subset of the full recognised set).
_STYLE_KEYS = ["heading_1", "list_number", "list_bullet", "quote", "table"]

# Self-contained theme — no CDN dependency, so the UI looks right offline / in
# locked-down deployments (FastHTML's default pico.css is CDN-loaded).
ADMIN_CSS = """
:root{
  --bg:#f1f5f9; --card:#ffffff; --ink:#0f172a; --muted:#64748b; --line:#e2e8f0;
  --brand:#4f46e5; --brand-d:#4338ca; --ok:#16a34a; --warn:#b45309; --err:#dc2626;
  --ok-bg:#ecfdf5; --warn-bg:#fffbeb; --err-bg:#fef2f2; --info-bg:#eff6ff;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
  font-size:15px;line-height:1.5}
a{color:var(--brand);text-decoration:none}
a:hover{text-decoration:underline}
.topbar{background:var(--brand);color:#fff;padding:.75rem 1.25rem;
  display:flex;align-items:center;justify-content:space-between;gap:1rem;
  box-shadow:0 1px 3px rgba(0,0,0,.15)}
.topbar .brand{font-weight:700;font-size:1.05rem;color:#fff;display:flex;align-items:center;gap:.5rem}
.topbar nav a{color:#e0e7ff;margin-left:1rem;font-size:.92rem}
.topbar nav a:hover{color:#fff}
.container{max-width:980px;margin:1.5rem auto;padding:0 1.25rem}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;
  padding:1.25rem 1.5rem;margin-bottom:1.25rem;box-shadow:0 1px 2px rgba(15,23,42,.04)}
h1{font-size:1.5rem;margin:.2rem 0 1rem}
h2{font-size:1.15rem;margin:0 0 .75rem}
h3{font-size:1rem;margin:1.25rem 0 .5rem}
.muted{color:var(--muted);font-size:.9rem}
.field{margin-bottom:1rem}
.field label{display:block;font-weight:600;font-size:.85rem;margin-bottom:.3rem}
input,select,textarea{width:100%;padding:.5rem .6rem;border:1px solid var(--line);
  border-radius:8px;font:inherit;background:#fff;color:var(--ink)}
input:focus,select:focus,textarea:focus{outline:2px solid var(--brand);
  outline-offset:-1px;border-color:var(--brand)}
input[disabled]{background:#f8fafc;color:var(--muted)}
textarea{resize:vertical}
.btn{display:inline-flex;align-items:center;gap:.4rem;cursor:pointer;border-radius:8px;
  padding:.5rem .9rem;font:inherit;font-weight:600;border:1px solid transparent;
  background:#fff;color:var(--ink);text-decoration:none}
.btn:hover{text-decoration:none}
.btn-primary{background:var(--brand);color:#fff}
.btn-primary:hover{background:var(--brand-d);color:#fff}
.btn-secondary{background:#fff;color:var(--brand);border-color:var(--brand)}
.btn-secondary:hover{background:#eef2ff}
.btn-danger{background:#fff;color:var(--err);border-color:#fecaca}
.btn-danger:hover{background:var(--err-bg)}
.btn-sm{padding:.3rem .55rem;font-size:.82rem}
.btn-icon{padding:.25rem .5rem;background:#fff;border:1px solid var(--line);color:var(--err)}
.btn-icon:hover{background:var(--err-bg);border-color:#fecaca}
.actions{display:flex;gap:.6rem;flex-wrap:wrap;align-items:center}
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
.badge-off{background:#f1f5f9;color:var(--muted)}
.badge-ro{background:#f1f5f9;color:var(--muted)}
.badge-if{background:#eef2ff;color:var(--brand);margin-left:.4rem}
.chip{display:inline-block;padding:.15rem .55rem;margin:.15rem .25rem .15rem 0;border-radius:6px;
  background:#eef2ff;color:var(--brand-d);font-size:.82rem;font-family:ui-monospace,Menlo,monospace}
.chip-if{background:#fef9c3;color:#854d0e}
.flash{padding:.6rem .85rem;border-radius:8px;margin:.5rem 0;border:1px solid transparent}
.flash-info{background:var(--info-bg);border-color:#bfdbfe}
.flash-ok{background:var(--ok-bg);border-color:#a7f3d0}
.flash-warn{background:var(--warn-bg);border-color:#fde68a;color:#854d0e}
.flash-err{background:var(--err-bg);border-color:#fecaca;color:#991b1b}
.empty{text-align:center;padding:2rem 1rem;color:var(--muted)}
details{margin:1rem 0;border:1px solid var(--line);border-radius:8px;padding:.5rem .85rem}
summary{cursor:pointer;font-weight:600}
.login-wrap{max-width:380px;margin:8vh auto}
.inline-form{display:inline}
hr{border:none;border-top:1px solid var(--line);margin:1.25rem 0}
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


class AdminContext:
    """Shared services the views depend on."""

    def __init__(self, mcp, config: Config):
        self.mcp = mcp
        self.config = config
        self.path = config.admin.path.rstrip("/")
        self.store = FileTemplateStore.from_config()
        # Global docx style_mapping from the master YAML (for live registration).
        master = self.store.config_dir / "docx_templates.yaml"
        _templates, cfg = gather_specs(master, None)
        self.global_style_mapping = cfg.get("style_mapping") or {}

    def u(self, path: str = "") -> str:
        """Absolute (mount-prefixed) URL for an admin-relative *path*."""
        return f"{self.path}{path}"

    # -- live MCP tool registration ----------------------------------------

    def register(self, kind: str, spec: Dict[str, Any]) -> bool:
        if kind == KIND_DOCX:
            from docx_tools.dynamic_docx_tools import register_docx_template
            return register_docx_template(self.mcp, spec, self.global_style_mapping)
        from email_tools.dynamic_email_tools import register_email_template
        return register_email_template(self.mcp, spec)

    def unregister(self, kind: str, name: str) -> bool:
        if kind == KIND_DOCX:
            from docx_tools.dynamic_docx_tools import unregister_docx_template
            return unregister_docx_template(self.mcp, name)
        from email_tools.dynamic_email_tools import unregister_email_template
        return unregister_email_template(self.mcp, name)

    def live_names(self, kind: str) -> List[str]:
        if kind == KIND_DOCX:
            from docx_tools.dynamic_docx_tools import registered_docx_template_names
            return registered_docx_template_names()
        from email_tools.dynamic_email_tools import registered_email_template_names
        return registered_email_template_names()


# ---------------------------------------------------------------------------
# Form parsing helpers
# ---------------------------------------------------------------------------

def _coerce_default(atype: str, raw: str) -> Any:
    """Coerce a string default from the form into the arg's declared type."""
    raw = (raw or "").strip()
    atype = (atype or "string").lower()
    if atype in ("bool", "boolean"):
        return raw.lower() in ("1", "true", "yes", "on")
    if atype in ("int", "integer") and raw:
        try:
            return int(raw)
        except ValueError:
            return raw
    if atype == "float" and raw:
        try:
            return float(raw)
        except ValueError:
            return raw
    return raw


def _parse_args_from_form(form) -> List[Dict[str, Any]]:
    """Build the ``args`` list from the parallel-indexed form fields."""
    names = form.getlist("arg_name")
    types = form.getlist("arg_type")
    reqs = form.getlist("arg_required")
    defs = form.getlist("arg_default")
    descs = form.getlist("arg_desc")
    args: List[Dict[str, Any]] = []
    for i, raw_name in enumerate(names):
        name = (raw_name or "").strip()
        if not name:
            continue
        atype = types[i] if i < len(types) else "string"
        required = (reqs[i] if i < len(reqs) else "true") == "true"
        default_raw = defs[i] if i < len(defs) else ""
        desc = descs[i] if i < len(descs) else ""
        arg: Dict[str, Any] = {"name": name, "type": atype, "required": required, "description": desc}
        if not required or (default_raw or "").strip():
            arg["default"] = _coerce_default(atype, default_raw)
        args.append(arg)
    return args


def _parse_style_mapping(form) -> Dict[str, str]:
    """Collect non-default style-mapping selections."""
    mapping: Dict[str, str] = {}
    for key in _STYLE_KEYS:
        val = (form.get(f"style_{key}") or "").strip()
        if val and val != "__default__":
            mapping[key] = val
    return mapping


def _build_spec(kind: str, form) -> Dict[str, Any]:
    """Assemble a template spec dict from the submitted edit form."""
    name = validate_name((form.get("name") or "").strip())
    asset_filename = (form.get("asset_filename") or "").strip()
    description = (form.get("description") or "").strip()
    title = (form.get("title") or "").strip()

    spec: Dict[str, Any] = {"name": name, "description": description or f"Generate {name}"}
    if title:
        spec["annotations"] = {"title": title}
    spec[_path_key(kind)] = asset_filename
    spec["args"] = _parse_args_from_form(form)
    if kind == KIND_DOCX:
        style_mapping = _parse_style_mapping(form)
        if style_mapping:
            spec["style_mapping"] = style_mapping
    return spec


def _path_key(kind: str) -> str:
    return "docx_path" if kind == KIND_DOCX else "html_path"


def _asset_ext(kind: str) -> str:
    return ".docx" if kind == KIND_DOCX else ".html"


# ---------------------------------------------------------------------------
# View fragments
# ---------------------------------------------------------------------------

def _topbar(ctx: AdminContext, authed: bool = True):
    from fasthtml.common import Nav
    nav = Nav(
        A("All templates", href=ctx.u("/")),
        A("New Word", href=ctx.u("/new/docx")),
        A("New Email", href=ctx.u("/new/email")),
        A("Log out", href=ctx.u("/logout")),
    ) if authed else Span()
    return Header(Span("📄 Template Admin", cls="brand"), nav, cls="topbar")


def _page(ctx: AdminContext, title: str, *content, authed: bool = True):
    """Wrap page *content* in the themed shell (title + topbar + container)."""
    return (
        Title(f"{title} · Template Admin"),
        _topbar(ctx, authed),
        Main(Div(*content, cls="container")),
    )


def _flash(msg: Optional[str], kind: str = "info"):
    if not msg:
        return None
    return Div(msg, cls=f"flash flash-{kind}")


def _status_badge(live: bool):
    return Span("● Live", cls="badge badge-live") if live else Span("Not live", cls="badge badge-off")


def _template_table(ctx: AdminContext, kind: str):
    managed = ctx.store.list_specs(kind)
    managed_names = {s.get("name") for s in managed}
    live = set(ctx.live_names(kind))

    rows = []
    for spec in managed:
        name = spec.get("name")
        nargs = len(spec.get("args") or [])
        rows.append(Tr(
            Td(A(name, href=ctx.u(f"/{kind}/{name}/edit"))),
            Td(str(nargs)),
            Td(_status_badge(name in live)),
            Td(Div(
                A("Edit", href=ctx.u(f"/{kind}/{name}/edit"), cls="btn btn-secondary btn-sm"),
                Form(Button("Delete", type="submit", cls="btn btn-danger btn-sm",
                            onclick="return confirm('Delete this template? This removes the live tool.')"),
                     action=ctx.u(f"/{kind}/{name}/delete"), method="post", cls="inline-form"),
                cls="actions",
            )),
        ))
    # Read-only templates from the master YAML (live but not UI-managed).
    for name in sorted(live - managed_names):
        rows.append(Tr(
            Td(name), Td("—"),
            Td(Span("● Live", cls="badge badge-live")),
            Td(Span("from master YAML — read-only", cls="badge badge-ro")),
        ))

    if not rows:
        return Div(
            P(f"No {_KIND_LABEL[kind]} templates yet."),
            A(f"{_KIND_ICON[kind]} Create your first {_KIND_LABEL[kind]} template",
              href=ctx.u(f"/new/{kind}"), cls="btn btn-primary"),
            cls="empty",
        )
    return Table(
        Thead(Tr(Th("Name"), Th("Args"), Th("Status"), Th("Actions"))),
        Tbody(*rows),
        cls="tpl-table",
    )


def _arg_row(arg: Dict[str, Any] = None, cond: bool = False):
    """Render one editable argument row (parallel-indexed fields + remove button)."""
    arg = arg or {}
    name = arg.get("name", "")
    atype = str(arg.get("type", "string")).lower()
    required = bool(arg.get("required", True))
    default = arg.get("default", "")
    desc = arg.get("description", "")

    type_opts = [Option(t, value=t, selected=(t == atype)) for t in _ARG_TYPES]
    req_opts = [
        Option("required", value="true", selected=required),
        Option("optional", value="false", selected=not required),
    ]
    name_cell = [Input(name="arg_name", value=name, placeholder="argument_name")]
    if cond:
        name_cell.append(Span("if", cls="badge badge-if", title="Conditional flag"))
    return Tr(
        Td(*name_cell, cls="col-name"),
        Td(Select(*type_opts, name="arg_type"), cls="col-type"),
        Td(Select(*req_opts, name="arg_required"), cls="col-req"),
        Td(Input(name="arg_default", value="" if default in (None,) else str(default),
                 placeholder="optional"), cls="col-def"),
        Td(Input(name="arg_desc", value=desc, placeholder="what this is, for the AI")),
        Td(Button("✕", type="button", cls="btn-icon", title="Remove",
                  onclick="adminRemoveRow(this)"), cls="col-x"),
        cls="arg-row",
    )


def _chips(items, cond_set=None):
    cond_set = cond_set or set()
    if not items:
        return Span("none", cls="muted")
    return Span(*[Span(it, cls="chip chip-if" if it in cond_set else "chip") for it in items])


def _style_mapping_block(analysis: Optional[Analysis], spec: Dict[str, Any]):
    styles = (analysis.styles_present if analysis else []) or []
    current = (spec or {}).get("style_mapping") or {}
    selects = []
    for key in _STYLE_KEYS:
        cur = current.get(key, "")
        opts = [Option("(use built-in)", value="__default__", selected=not cur)]
        opts += [Option(s, value=s, selected=(s == cur)) for s in styles]
        selects.append(Div(Label(key), Select(*opts, name=f"style_{key}"), cls="field"))
    return Details(
        Summary("Advanced — map markdown styles to your template's own style names"),
        P("Only needed if your Word template renames the built-in styles "
          "(Heading 1, List Number, …). Otherwise leave these alone.", cls="muted"),
        *selects,
    )


def _analysis_report(analysis: Analysis, spec: Dict[str, Any]):
    rec = reconcile(analysis, spec.get("args") or [])
    cond_set = set(analysis.conditionals)
    items = [
        Div(Label("Detected placeholders"), _chips(analysis.placeholders), cls="field"),
        Div(Label("Detected conditionals"),
            _chips(analysis.conditionals, cond_set) if analysis.conditionals else Span("none", cls="muted"),
            cls="field"),
    ]
    if analysis.missing_required_styles:
        items.append(_flash(
            "⚠ The document is missing some Word styles the renderer uses: "
            + ", ".join(analysis.missing_required_styles)
            + ". Lists/headings using them will fall back to the default style.", "warn"))
    for w in analysis.warnings:
        items.append(_flash("⚠ " + w, "warn"))
    if rec.orphan_args:
        items.append(_flash("Args with no matching placeholder (they'll be ignored by the document): "
                            + ", ".join(rec.orphan_args), "warn"))
    if rec.non_bool_conditions:
        items.append(_flash("Conditional flags should be type 'bool': "
                            + ", ".join(rec.non_bool_conditions), "warn"))
    return Div(H3("What we found in the document"),
               *[i for i in items if i is not None], cls="card")


def _edit_form(ctx: AdminContext, kind: str, spec: Dict[str, Any],
               analysis: Optional[Analysis], is_new: bool):
    asset_filename = spec.get(_path_key(kind), "")
    annotations = spec.get("annotations") or {}
    cond_set = set(analysis.conditionals) if analysis else set()

    # Build arg rows: declared args first, then detected-but-missing. No blank
    # spares — users add rows with the "Add argument" button.
    existing = spec.get("args") or []
    existing_names = {a.get("name") for a in existing if isinstance(a, dict)}
    rows = [_arg_row(a, cond=a.get("name") in cond_set) for a in existing]
    if analysis:
        for ph in analysis.placeholders:
            if ph not in existing_names:
                is_cond = ph in cond_set
                rows.append(_arg_row(
                    {"name": ph, "type": "bool" if is_cond else "string",
                     "required": not is_cond, "description": ""}, cond=is_cond))
                existing_names.add(ph)
        for c in analysis.conditionals:
            if c not in existing_names:
                rows.append(_arg_row({"name": c, "type": "bool", "required": False,
                                      "description": ""}, cond=True))
                existing_names.add(c)
    if not rows:
        rows = [_arg_row()]

    name_field = (
        Div(Label("Tool name"),
            Input(name="name", value=spec.get("name", ""), required=True,
                  placeholder="e.g. formal_letter"),
            P("Letters, digits and underscores. This is what the AI calls.", cls="muted"),
            cls="field")
        if is_new else
        Div(Label("Tool name"), Input(value=spec.get("name", ""), disabled=True), cls="field")
    )

    details_card = Div(
        H3("Tool details"),
        name_field,
        Div(Label("Title"), Input(name="title", value=annotations.get("title", ""),
            placeholder="Friendly name shown to the user"), cls="field"),
        Div(Label("Description"),
            Textarea(spec.get("description", ""), name="description", rows="3",
                     placeholder="Tell the AI when and how to use this template."),
            P("This is the instruction the AI sees when choosing the tool.", cls="muted"),
            cls="field"),
        cls="card",
    )

    args_card = Div(
        H3("Arguments"),
        P("Every placeholder and conditional in your document needs an argument. "
          "We pre-filled them from the document — adjust types and descriptions, then save.",
          cls="muted"),
        Table(
            Thead(Tr(Th("Name", cls="col-name"), Th("Type", cls="col-type"),
                     Th("Required", cls="col-req"), Th("Default", cls="col-def"),
                     Th("Description"), Th("", cls="col-x"))),
            Tbody(*rows, id="argrows"),
            cls="args-table",
        ),
        Div(Button("+ Add argument", type="button", cls="btn btn-secondary btn-sm",
                   onclick="adminAddArgRow()"), style="margin-top:.6rem"),
        cls="card",
    )

    return Form(
        Hidden(name="kind", value=kind),
        Hidden(name="asset_filename", value=asset_filename),
        Hidden(name="name", value=spec.get("name", "")) if not is_new else None,
        details_card,
        args_card,
        Div(_style_mapping_block(analysis, spec), cls="card") if kind == KIND_DOCX else None,
        Div(
            Div(
                Button("Save & make live", type="submit", cls="btn btn-primary"),
                Button("Preview", type="submit", formaction=ctx.u(f"/{kind}/preview"),
                       formtarget="_blank", cls="btn btn-secondary"),
                A("Cancel", href=ctx.u("/"), cls="btn"),
                cls="actions",
            ),
            cls="card",
        ),
        action=ctx.u(f"/{kind}/save"), method="post",
    )


# ---------------------------------------------------------------------------
# App factory
# ---------------------------------------------------------------------------

def _new_page(ctx: AdminContext, kind: str, error: str = None):
    return _page(
        ctx, f"New {_KIND_LABEL[kind]} template",
        H1(f"{_KIND_ICON[kind]} New {_KIND_LABEL[kind]} template"),
        _flash(error, "err"),
        Div(
            Form(
                Div(Label("Tool name"),
                    Input(name="name", required=True, placeholder="e.g. formal_letter"),
                    P("Letters, digits and underscores — this is what the AI calls.", cls="muted"),
                    cls="field"),
                Div(Label(f"{_KIND_LABEL[kind]} file ({_asset_ext(kind)})"),
                    Input(name="file", type="file", accept=_asset_ext(kind), required=True),
                    cls="field"),
                Button("Upload & analyze", type="submit", cls="btn btn-primary"),
                action=ctx.u(f"/{kind}/draft"), method="post", enctype="multipart/form-data",
            ),
            cls="card",
        ),
        Details(
            Summary("How does this work?"),
            P("Author your document in ",
              "Word" if kind == KIND_DOCX else "any HTML editor",
              " and insert placeholders like ", Span("{{recipient_name}}", cls="chip"),
              " wherever a value should go."),
            P("Wrap optional content in ", Span("{{#if include_clause}}", cls="chip"),
              " … ", Span("{{/if}}", cls="chip"), " to show it only when a flag is set.",
              cls="muted") if kind == KIND_DOCX else None,
            P("When you upload, we scan the file, list every placeholder, and pre-build the "
              "argument form for you. Nothing goes live until you press Save.", cls="muted"),
        ),
    )


def build_admin_app(mcp, config: Config) -> FastHTML:
    ctx = AdminContext(mcp, config)
    expected_pw = config.admin_password_effective
    login_path = ctx.u("/login")

    # Stable session secret derived from the password so cookies survive
    # restarts; falls back to a constant when no password is configured (the
    # login can't succeed in that case anyway).
    secret = hashlib.sha256(f"mcp-office-admin:{expected_pw or ''}".encode()).hexdigest()

    # Self-contained headers (no CDN): meta + theme CSS + the row JS, with the
    # blank-row HTML injected so "Add argument" can clone it client-side.
    blank_row_html = json.dumps(to_xml(_arg_row()))
    hdrs = (
        Meta(charset="utf-8"),
        Meta(name="viewport", content="width=device-width, initial-scale=1"),
        Style(ADMIN_CSS),
        Script(NotStr(f"window.__ARG_ROW_HTML__ = {blank_row_html};\n{ARG_ROWS_JS}")),
    )

    app = FastHTML(secret_key=secret, before=auth.make_before(login_path),
                   default_hdrs=False, htmx=False, surreal=False, hdrs=hdrs)
    rt = app.route

    @rt("/login", methods=["get", "post"])
    async def login(req, sess):
        error = None
        if req.method == "POST":
            form = await req.form()
            if auth.check_password(form.get("password"), expected_pw):
                sess[auth.SESSION_KEY] = True
                return RedirectResponse(ctx.u("/"), status_code=303)
            error = "Incorrect password."
        return _page(
            ctx, "Sign in",
            Div(
                Div(
                    H1("📄 Template Admin"),
                    P("Sign in to manage document templates.", cls="muted"),
                    _flash(error, "err"),
                    Form(
                        Div(Label("Password"),
                            Input(name="password", type="password", required=True, autofocus=True),
                            cls="field"),
                        Button("Sign in", type="submit", cls="btn btn-primary"),
                        action=ctx.u("/login"), method="post",
                    ),
                    cls="card",
                ),
                cls="login-wrap",
            ),
            authed=False,
        )

    @rt("/logout")
    def logout(sess):
        sess.pop(auth.SESSION_KEY, None)
        return RedirectResponse(login_path, status_code=303)

    @rt("/")
    def index(sess):
        return _page(
            ctx, "Templates",
            H1("Templates"),
            P("Create reusable Word and email templates. Each one becomes a tool the AI can call.",
              cls="muted"),
            Div(H2(f"{_KIND_ICON[KIND_DOCX]} Word templates"),
                _template_table(ctx, KIND_DOCX), cls="card"),
            Div(H2(f"{_KIND_ICON[KIND_EMAIL]} Email templates"),
                _template_table(ctx, KIND_EMAIL), cls="card"),
        )

    @rt("/new/{kind}")
    def new(kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        return _new_page(ctx, kind)

    @rt("/{kind}/draft", methods=["post"])
    async def draft(req, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        form = await req.form()
        try:
            name = validate_name((form.get("name") or "").strip())
        except TemplateStoreError as e:
            return _new_page(ctx, kind, error=str(e))
        upload = form.get("file")
        data = await upload.read() if upload is not None else b""
        if not data:
            return _new_page(ctx, kind, error="Please choose a file to upload.")

        analysis = analyze(kind, data)
        if any("Could not open" in w for w in analysis.warnings):
            return _new_page(ctx, kind, error=analysis.warnings[0])

        filename = f"{name}{_asset_ext(kind)}"
        ctx.store.write_asset(kind, filename, data)
        spec = {"name": name, "description": "", _path_key(kind): filename, "args": []}
        return _page(
            ctx, f"Configure {name}",
            H1(f"Configure {name}"),
            _flash(f"Analyzed {filename} — review the arguments below, preview, then save.", "ok"),
            _analysis_report(analysis, spec),
            _edit_form(ctx, kind, spec, analysis, is_new=False),
        )

    @rt("/{kind}/{name}/edit")
    def edit(kind: str, name: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        spec = ctx.store.get_spec(kind, name)
        if spec is None:
            return _page(ctx, "Not found", H1("Not found"),
                         _flash(f"No managed template named '{name}'.", "err"),
                         A("← Back to all templates", href=ctx.u("/")))
        analysis = None
        asset = spec.get(_path_key(kind))
        if asset and ctx.store.asset_exists(kind, asset):
            analysis = analyze(kind, ctx.store.read_asset(kind, asset))
        return _page(
            ctx, f"Edit {name}",
            H1(f"Edit {name}"),
            _analysis_report(analysis, spec) if analysis else
            _flash("The template's source file is missing — arguments can still be edited.", "warn"),
            _edit_form(ctx, kind, spec, analysis, is_new=False),
        )

    @rt("/{kind}/save", methods=["post"])
    async def save(req, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        form = await req.form()
        try:
            spec = _build_spec(kind, form)
            ctx.store.save_spec(kind, spec)
        except TemplateStoreError as e:
            return _page(ctx, "Save failed", H1("Save failed"), _flash(str(e), "err"),
                         A("← Back to all templates", href=ctx.u("/")))
        ok = ctx.register(kind, spec)
        msg = (f"Saved — the tool '{spec['name']}' is now live and ready for the AI to use."
               if ok else f"Saved, but the tool '{spec['name']}' could not be registered (check the logs).")
        return _page(
            ctx, "Saved",
            H1("✓ Saved" if ok else "Saved with a warning"),
            _flash(msg, "ok" if ok else "warn"),
            Div(A("Back to all templates", href=ctx.u("/"), cls="btn btn-primary"),
                A("Keep editing", href=ctx.u(f"/{kind}/{spec['name']}/edit"), cls="btn btn-secondary"),
                cls="actions"),
        )

    @rt("/{kind}/preview", methods=["post"])
    async def preview(req, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        form = await req.form()
        spec = _build_spec(kind, form)
        asset = spec.get(_path_key(kind))
        if not asset or not ctx.store.asset_exists(kind, asset):
            return HTMLResponse("<p>Nothing to preview yet — save the template first.</p>",
                                status_code=400)
        data = ctx.store.read_asset(kind, asset)
        analysis = analyze(kind, data)
        values = sample_values(spec.get("args") or [], analysis.conditionals)
        if kind == KIND_DOCX:
            out = render_docx_preview(data, spec, values, ctx.global_style_mapping)
            return Response(
                content=out,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={"Content-Disposition": f'attachment; filename="{spec["name"]}_preview.docx"'},
            )
        html = render_email_preview(data, spec, values)
        return HTMLResponse(html)

    @rt("/{kind}/{name}/delete", methods=["post"])
    def delete(kind: str, name: str):
        if kind in KINDS:
            ctx.store.delete_spec(kind, name, delete_asset=False)
            ctx.unregister(kind, name)
        return RedirectResponse(ctx.u("/"), status_code=303)

    return app


# ---------------------------------------------------------------------------
# Combined ASGI app (admin + MCP in one process)
# ---------------------------------------------------------------------------

def build_combined_app(mcp, config: Config):
    """Build a single ASGI app serving the MCP endpoint and the admin UI.

    The MCP app (with its required lifespan / session manager) is mounted at the
    root; the admin UI is mounted under ``config.admin.path``. Mount order puts
    the admin prefix first so it wins over the catch-all MCP mount.
    """
    from starlette.applications import Starlette

    mcp_app = mcp.http_app(path="/mcp")
    admin_app = build_admin_app(mcp, config)
    routes = [
        Mount(config.admin.path, app=admin_app),
        Mount("/", app=mcp_app),
    ]
    return Starlette(routes=routes, lifespan=mcp_app.lifespan)
