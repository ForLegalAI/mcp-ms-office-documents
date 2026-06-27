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
import logging
from typing import Any, Dict, List, Optional

from fasthtml.common import (
    FastHTML, Titled, Div, P, A, H2, H3, Form, Input, Button, Label, Span,
    Table, Thead, Tbody, Tr, Th, Td, Select, Option, Textarea, Details, Summary,
    Hr, RedirectResponse, Response, HTMLResponse, Hidden,
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
_KIND_LABEL = {KIND_DOCX: "Word (docx)", KIND_EMAIL: "Email (html)"}
_ARG_TYPES = ["string", "int", "float", "bool", "list"]
# Style-mapping keys surfaced in the UI (subset of the full recognised set).
_STYLE_KEYS = ["heading_1", "list_number", "list_bullet", "quote", "table"]


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

def _nav(ctx: AdminContext):
    return Div(
        A("All templates", href=ctx.u("/")), " · ",
        A("New Word template", href=ctx.u("/new/docx")), " · ",
        A("New Email template", href=ctx.u("/new/email")), " · ",
        A("Log out", href=ctx.u("/logout")),
        style="margin-bottom:1rem",
    )


def _flash(msg: Optional[str], kind: str = "info"):
    if not msg:
        return None
    color = {"info": "#2563eb", "ok": "#16a34a", "warn": "#d97706", "err": "#dc2626"}.get(kind, "#2563eb")
    return Div(msg, style=f"padding:.5rem .75rem;border-left:4px solid {color};background:#f8fafc;margin:.5rem 0")


def _template_table(ctx: AdminContext, kind: str):
    managed = ctx.store.list_specs(kind)
    managed_names = {s.get("name") for s in managed}
    live = set(ctx.live_names(kind))

    rows = []
    for spec in managed:
        name = spec.get("name")
        nargs = len(spec.get("args") or [])
        rows.append(Tr(
            Td(name),
            Td(str(nargs)),
            Td("● live" if name in live else "○ not live"),
            Td(
                A("Edit", href=ctx.u(f"/{kind}/{name}/edit")), " · ",
                Form(Button("Delete", type="submit", cls="secondary",
                            onclick="return confirm('Delete this template?')"),
                     action=ctx.u(f"/{kind}/{name}/delete"), method="post",
                     style="display:inline"),
            ),
        ))
    # Read-only templates from the master YAML (live but not UI-managed).
    for name in sorted(live - managed_names):
        rows.append(Tr(Td(name), Td("—"), Td("● live (master YAML)"), Td(Span("read-only", style="color:#888"))))

    if not rows:
        return P(f"No {_KIND_LABEL[kind]} templates yet.")
    return Table(
        Thead(Tr(Th("Name"), Th("Args"), Th("Status"), Th("Actions"))),
        Tbody(*rows),
    )


def _arg_row(arg: Dict[str, Any] = None, lock_bool: bool = False):
    """Render one editable argument row (parallel-indexed fields)."""
    arg = arg or {}
    name = arg.get("name", "")
    atype = str(arg.get("type", "string")).lower()
    required = bool(arg.get("required", True))
    default = arg.get("default", "")
    desc = arg.get("description", "")

    type_opts = [
        Option(t, value=t, selected=(t == atype)) for t in _ARG_TYPES
    ]
    req_opts = [
        Option("required", value="true", selected=required),
        Option("optional", value="false", selected=not required),
    ]
    return Tr(
        Td(Input(name="arg_name", value=name, placeholder="arg_name")),
        Td(Select(*type_opts, name="arg_type", disabled=lock_bool)),
        Td(Select(*req_opts, name="arg_required")),
        Td(Input(name="arg_default", value="" if default in (None,) else str(default), placeholder="(default)")),
        Td(Input(name="arg_desc", value=desc, placeholder="description shown to the AI")),
    )


def _style_mapping_block(analysis: Optional[Analysis], spec: Dict[str, Any]):
    styles = (analysis.styles_present if analysis else []) or []
    current = (spec or {}).get("style_mapping") or {}
    selects = []
    for key in _STYLE_KEYS:
        cur = current.get(key, "")
        opts = [Option("(use built-in)", value="__default__", selected=not cur)]
        opts += [Option(s, value=s, selected=(s == cur)) for s in styles]
        selects.append(Div(Label(key, style="font-weight:600"), Select(*opts, name=f"style_{key}")))
    return Details(
        Summary("Advanced: map markdown styles to your template's style names"),
        P("Only needed if your template renames the built-in Word styles.",
          style="color:#666;font-size:.9em"),
        *selects,
    )


def _analysis_report(analysis: Analysis, spec: Dict[str, Any]):
    rec = reconcile(analysis, spec.get("args") or [])
    items = [
        P(f"Detected placeholders: {', '.join(analysis.placeholders) or '—'}"),
        P(f"Detected conditionals: {', '.join(analysis.conditionals) or '—'}"),
    ]
    if analysis.missing_required_styles:
        items.append(_flash(
            "Template is missing styles the renderer uses: "
            + ", ".join(analysis.missing_required_styles), "warn"))
    for w in analysis.warnings:
        items.append(_flash(w, "warn"))
    if rec.orphan_args:
        items.append(_flash("Args with no matching placeholder: " + ", ".join(rec.orphan_args), "warn"))
    if rec.non_bool_conditions:
        items.append(_flash(
            "Conditional flags should be type 'bool': " + ", ".join(rec.non_bool_conditions), "warn"))
    return Div(*[i for i in items if i is not None],
               style="background:#f8fafc;padding:.5rem .75rem;margin:.5rem 0")


def _edit_form(ctx: AdminContext, kind: str, spec: Dict[str, Any],
               analysis: Optional[Analysis], is_new: bool):
    asset_filename = spec.get(_path_key(kind), "")
    annotations = spec.get("annotations") or {}

    # Build arg rows: declared args first, then detected-but-missing, then blanks.
    existing = spec.get("args") or []
    existing_names = {a.get("name") for a in existing if isinstance(a, dict)}
    rows = [_arg_row(a) for a in existing]
    if analysis:
        cond_set = set(analysis.conditionals)
        for ph in analysis.placeholders:
            if ph not in existing_names:
                is_cond = ph in cond_set
                rows.append(_arg_row(
                    {"name": ph, "type": "bool" if is_cond else "string",
                     "required": not is_cond, "description": ""}))
                existing_names.add(ph)
        for cond in analysis.conditionals:
            if cond not in existing_names:
                rows.append(_arg_row({"name": cond, "type": "bool", "required": False, "description": ""}))
                existing_names.add(cond)
    rows += [_arg_row() for _ in range(3)]  # spare blank rows for manual additions

    report = _analysis_report(analysis, spec) if analysis else None

    return Form(
        Hidden(name="kind", value=kind),
        Hidden(name="asset_filename", value=asset_filename),
        Hidden(name="name", value=spec.get("name", "")) if not is_new else None,

        report,
        Div(
            Label("Tool name (letters, digits, underscore)"),
            Input(name="name", value=spec.get("name", ""), required=True) if is_new
            else Input(value=spec.get("name", ""), disabled=True),
        ),
        Div(Label("Title (shown to the user)"),
            Input(name="title", value=annotations.get("title", ""))),
        Div(Label("Description (instructions the AI sees)"),
            Textarea(spec.get("description", ""), name="description", rows="3")),

        H3("Arguments"),
        P("Each placeholder/conditional in the document needs an argument. Empty rows are ignored.",
          style="color:#666;font-size:.9em"),
        Table(
            Thead(Tr(Th("Name"), Th("Type"), Th("Required"), Th("Default"), Th("Description"))),
            Tbody(*rows),
        ),

        _style_mapping_block(analysis, spec) if kind == KIND_DOCX else None,

        Hr(),
        Div(
            Button("Save & make live", type="submit"),
            Button("Preview", type="submit", formaction=ctx.u(f"/{kind}/preview"),
                   formtarget="_blank", cls="secondary"),
            style="display:flex;gap:.5rem",
        ),
        action=ctx.u(f"/{kind}/save"), method="post",
    )


# ---------------------------------------------------------------------------
# App factory
# ---------------------------------------------------------------------------

def build_admin_app(mcp, config: Config) -> FastHTML:
    ctx = AdminContext(mcp, config)
    expected_pw = config.admin_password_effective
    login_path = ctx.u("/login")

    # Stable session secret derived from the password so cookies survive
    # restarts; falls back to a constant when no password is configured (the
    # login can't succeed in that case anyway).
    secret = hashlib.sha256(f"mcp-office-admin:{expected_pw or ''}".encode()).hexdigest()

    app = FastHTML(secret_key=secret, before=auth.make_before(login_path))
    rt = app.route

    @rt("/login", methods=["get", "post"])
    async def login(req, sess):
        if req.method == "POST":
            form = await req.form()
            if auth.check_password(form.get("password"), expected_pw):
                sess[auth.SESSION_KEY] = True
                return RedirectResponse(ctx.u("/"), status_code=303)
            return Titled("Admin login",
                          _flash("Incorrect password.", "err"),
                          _login_form(ctx))
        return Titled("Admin login", _login_form(ctx))

    @rt("/logout")
    def logout(sess):
        sess.pop(auth.SESSION_KEY, None)
        return RedirectResponse(login_path, status_code=303)

    @rt("/")
    def index(sess):
        return Titled(
            "Template Admin",
            _nav(ctx),
            H2("Word templates"),
            _template_table(ctx, KIND_DOCX),
            H2("Email templates"),
            _template_table(ctx, KIND_EMAIL),
        )

    @rt("/new/{kind}")
    def new(kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        return Titled(
            f"New {_KIND_LABEL[kind]} template",
            _nav(ctx),
            P(f"Upload a {_asset_ext(kind)} file containing {{{{placeholders}}}}. "
              "We'll detect them and build the argument form."),
            Form(
                Div(Label("Tool name"), Input(name="name", required=True,
                    placeholder="e.g. formal_letter")),
                Div(Label(f"{_asset_ext(kind)} file"),
                    Input(name="file", type="file", accept=_asset_ext(kind), required=True)),
                Button("Upload & analyze", type="submit"),
                action=ctx.u(f"/{kind}/draft"), method="post", enctype="multipart/form-data",
            ),
        )

    @rt("/{kind}/draft", methods=["post"])
    async def draft(req, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        form = await req.form()
        try:
            name = validate_name((form.get("name") or "").strip())
        except TemplateStoreError as e:
            return Titled("New template", _nav(ctx), _flash(str(e), "err"),
                          A("Back", href=ctx.u(f"/new/{kind}")))
        upload = form.get("file")
        data = await upload.read() if upload is not None else b""
        if not data:
            return Titled("New template", _nav(ctx), _flash("No file uploaded.", "err"),
                          A("Back", href=ctx.u(f"/new/{kind}")))

        analysis = analyze(kind, data)
        if any("Could not open" in w for w in analysis.warnings):
            return Titled("New template", _nav(ctx), _flash(analysis.warnings[0], "err"),
                          A("Back", href=ctx.u(f"/new/{kind}")))

        filename = f"{name}{_asset_ext(kind)}"
        ctx.store.write_asset(kind, filename, data)
        spec = {"name": name, "description": "", _path_key(kind): filename, "args": []}
        return Titled(
            f"Configure {name}",
            _nav(ctx),
            _flash(f"Analyzed {filename}. Review the arguments below and save.", "ok"),
            _edit_form(ctx, kind, spec, analysis, is_new=False),
        )

    @rt("/{kind}/{name}/edit")
    def edit(kind: str, name: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        spec = ctx.store.get_spec(kind, name)
        if spec is None:
            return Titled("Not found", _nav(ctx),
                          _flash(f"No managed template named {name}.", "err"))
        analysis = None
        asset = spec.get(_path_key(kind))
        if asset and ctx.store.asset_exists(kind, asset):
            analysis = analyze(kind, ctx.store.read_asset(kind, asset))
        return Titled(f"Edit {name}", _nav(ctx),
                      _edit_form(ctx, kind, spec, analysis, is_new=False))

    @rt("/{kind}/save", methods=["post"])
    async def save(req, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        form = await req.form()
        try:
            spec = _build_spec(kind, form)
            ctx.store.save_spec(kind, spec)
        except TemplateStoreError as e:
            return Titled("Save failed", _nav(ctx), _flash(str(e), "err"),
                          A("Back", href=ctx.u("/")))
        ok = ctx.register(kind, spec)
        msg = (f"Saved — tool '{spec['name']}' is live."
               if ok else f"Saved, but the tool '{spec['name']}' could not be registered (check logs).")
        return Titled("Saved", _nav(ctx), _flash(msg, "ok" if ok else "warn"),
                      A("Back to all templates", href=ctx.u("/")))

    @rt("/{kind}/preview", methods=["post"])
    async def preview(req, kind: str):
        if kind not in KINDS:
            return RedirectResponse(ctx.u("/"), status_code=303)
        form = await req.form()
        spec = _build_spec(kind, form)
        asset = spec.get(_path_key(kind))
        if not asset or not ctx.store.asset_exists(kind, asset):
            return HTMLResponse("<p>No asset to preview yet — save first.</p>", status_code=400)
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


def _login_form(ctx: AdminContext):
    return Form(
        Div(Label("Password"), Input(name="password", type="password", required=True, autofocus=True)),
        Button("Log in", type="submit"),
        action=ctx.u("/login"), method="post",
    )


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
