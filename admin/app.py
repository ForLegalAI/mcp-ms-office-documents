"""FastHTML admin UI for managing dynamic docx / email / pptx templates.

Mounted in the same ASGI process as the FastMCP server (see
:func:`build_combined_app`) so that saving a template takes effect immediately —
no restart. The UI is gated by a single shared password (:mod:`admin.auth`) and
persists templates through :class:`admin.store.TemplateStore`.

Authoring model (chosen with the maintainer): the user uploads a real ``.docx``
(or email ``.html``); the UI auto-detects placeholders/conditionals
(:mod:`admin.analysis`), pre-builds the argument form, and previews
(:mod:`admin.preview`) without ever hitting the upload backend.

This module owns the routes and :class:`AdminContext` — the services a view
needs — and nothing else. The parts it used to hold live next door:

===========================  =================================================
:mod:`admin.components`      markup primitives and the inlined theme
:mod:`admin.kinds`           one descriptor per template kind
:mod:`admin.forms`           reading a submitted form back into a spec
:mod:`admin.views`           the pages
===========================  =================================================

PowerPoint is the odd one out and a few places branch on it. A docx or email
template is a parameterised document that becomes an MCP tool of its own, so it
has arguments to declare and "live" means the tool is registered. A pptx
template is a *design* — layouts, theme, aspect — with no arguments at all; it
becomes one more value for the ``template`` argument of the single
``create_powerpoint_presentation`` tool, and "live" means the template registry
has re-read it. :attr:`admin.kinds.KindDescriptor.has_args` is the flag that
distinguishes them.
"""
from __future__ import annotations

import hashlib
import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

import metrics

from fasthtml.common import FastHTML, RedirectResponse, Response, HTMLResponse, to_xml
from starlette.routing import Mount

from config import Config
from admin import auth, base_templates, views
from admin.analysis import analyze, is_unusable
from admin.components import head_tags
from admin.forms import build_spec, checked
from admin.kinds import (
    KINDS, content_disposition, descriptor, is_kind, media_type,
)
from admin.preview import (
    sample_values, render_docx_preview, render_email_preview, render_pptx_preview,
)
from admin.store import (
    FileTemplateStore, KIND_DOCX, KIND_PPTX, TemplateStoreError, validate_name,
)
from template_registry import gather_specs

logger = logging.getLogger(__name__)

class AdminContext:
    """Shared services the views depend on."""

    def __init__(self, mcp, config: Config):
        self.mcp = mcp
        self.config = config
        self.path = config.admin.path.rstrip("/")
        self.store = FileTemplateStore.from_config()

    @property
    def max_upload_bytes(self) -> int:
        """Largest template file this server accepts (``ADMIN_MAX_UPLOAD_MB``)."""
        return self.config.admin.max_upload_bytes

    def u(self, path: str = "") -> str:
        """Absolute (mount-prefixed) URL for an admin-relative *path*."""
        return f"{self.path}{path}"

    @property
    def global_style_mapping(self) -> Dict[str, Any]:
        """Master-YAML ``style_mapping``, re-read each access so edits on disk
        take effect without a restart (the admin UI is about live editing)."""
        master = self.store.config_dir / "docx_templates.yaml"
        _templates, cfg = gather_specs(master, None)
        return cfg.get("style_mapping") or {}

    # -- live MCP tool registration ----------------------------------------

    # A docx or email template *is* an MCP tool, so saving one registers a tool.
    # A PowerPoint template is not: there is a single
    # ``create_powerpoint_presentation`` tool and a template is one more value
    # for its ``template`` argument. "Going live" therefore means the registry
    # re-reads the spec directory, which is what clear_cache forces. The
    # registry already notices the change on its own via a directory
    # fingerprint; dropping the cache here just makes it immediate rather than
    # on the next call that happens to look.

    def register(self, kind: str, spec: Dict[str, Any]) -> bool:
        if kind == KIND_PPTX:
            return self._refresh_pptx_registry(spec.get("name"))
        if kind == KIND_DOCX:
            from docx_tools.dynamic_docx_tools import register_docx_template
            return register_docx_template(self.mcp, spec, self.global_style_mapping)
        from email_tools.dynamic_email_tools import register_email_template
        return register_email_template(self.mcp, spec)

    def unregister(self, kind: str, name: str) -> bool:
        if kind == KIND_PPTX:
            from pptx_tools.templates import clear_cache
            clear_cache()
            return True
        if kind == KIND_DOCX:
            from docx_tools.dynamic_docx_tools import unregister_docx_template
            return unregister_docx_template(self.mcp, name)
        from email_tools.dynamic_email_tools import unregister_email_template
        return unregister_email_template(self.mcp, name)

    def live_names(self, kind: str) -> List[str]:
        if kind == KIND_PPTX:
            from pptx_tools.templates import template_names
            try:
                return template_names()
            except Exception:
                logger.exception("[admin] Could not read the pptx template registry")
                return []
        if kind == KIND_DOCX:
            from docx_tools.dynamic_docx_tools import registered_docx_template_names
            return registered_docx_template_names()
        from email_tools.dynamic_email_tools import registered_email_template_names
        return registered_email_template_names()

    @staticmethod
    def _refresh_pptx_registry(name: Optional[str]) -> bool:
        """Reload the pptx registry and report whether *name* is now in it.

        Returning False here means the saved spec did not come back out of the
        registry — almost always because its file is missing from the template
        directories — and the admin is told so rather than being shown a
        success message for a template the tool cannot use.
        """
        from pptx_tools.templates import clear_cache, template_names
        clear_cache()
        try:
            return name in template_names()
        except Exception:
            logger.exception("[admin] Could not reload the pptx template registry")
            return False

    # -- asset helpers the routes share -------------------------------------

    def other_specs_using_asset(self, kind: str, name: str,
                                filename: Optional[str]) -> List[str]:
        """Managed templates other than *(kind, name)* whose asset is *filename*.

        Assets share one flat ``custom_templates/`` directory, so two specs can
        point at the same file — usually because one was hand-written. Deleting
        it for one of them would break the other, so the delete page checks
        first and withholds the option rather than offering a destructive
        choice it cannot honour safely.
        """
        if not filename:
            return []
        found: List[str] = []
        for other in KINDS:
            d = descriptor(other)
            for spec in self.store.list_specs(other):
                if other == kind and spec.get("name") == name:
                    continue
                if spec.get(d.path_key) == filename:
                    found.append(f"{d.label} template '{spec.get('name')}'")
            # Hand-written master-YAML templates count too, and a master entry
            # sharing the name being deleted counts most of all: the managed
            # spec was overriding it, so deleting the override brings the master
            # entry back to life — still pointing at this file.
            for spec in self._master_specs(other):
                if spec.get(d.path_key) == filename:
                    found.append(f"{d.label} template '{spec.get('name')}' "
                                 "(from the master YAML)")
        # Preserve order, drop repeats.
        return list(dict.fromkeys(found))

    def _master_specs(self, kind: str) -> List[Dict[str, Any]]:
        """Templates declared in *kind*'s hand-written master YAML.

        Read with no spec directory, so these are the master entries as
        written — not the merged view the registry serves.
        """
        master = self.store.config_dir / descriptor(kind).master_file
        try:
            templates, _cfg = gather_specs(master, None)
        except Exception:
            logger.exception("[admin] Could not read the %s master YAML", kind)
            return []
        return [t for t in templates if isinstance(t, dict)]

    def analyze_asset(self, kind: str, spec: Dict[str, Any]):
        """Analyse a spec's installed source file, or ``None`` when it is gone."""
        asset = spec.get(descriptor(kind).path_key)
        if asset and self.store.asset_exists(kind, asset):
            return analyze(kind, self.store.read_asset(kind, asset))
        return None


def build_admin_app(mcp, config: Config) -> FastHTML:
    ctx = AdminContext(mcp, config)
    expected_pw = config.admin_password_effective
    login_path = ctx.u("/login")

    # Capture recent logs for the Status page (only when the admin UI is on).
    metrics.install_log_capture(level=config.logging.level_no)

    # Stable session secret derived from the password so cookies survive
    # restarts; falls back to a constant when no password is configured (the
    # login can't succeed in that case anyway).
    secret = hashlib.sha256(f"mcp-office-admin:{expected_pw or ''}".encode()).hexdigest()

    # Self-contained headers (no CDN): meta + theme CSS + the row JS, with the
    # blank-row HTML injected so "Add argument" can clone it client-side.
    hdrs = head_tags(json.dumps(to_xml(views.arg_row())))

    app = FastHTML(secret_key=secret, before=auth.make_before(login_path),
                   default_hdrs=False, htmx=False, surreal=False, hdrs=hdrs,
                   htmlkw={"lang": "en"})
    rt = app.route

    def _csrf_guard(sess, form):
        """Return a 403 response when the form's CSRF token is invalid, else None."""
        if not auth.valid_csrf(sess, form.get("csrf")):
            logger.warning("[admin] Rejected POST with invalid CSRF token")
            return Response("CSRF validation failed — reload the page and try again.",
                            status_code=403)
        return None

    def _home():
        return RedirectResponse(ctx.u("/"), status_code=303)

    def _upload_size(upload) -> Optional[int]:
        """The upload's size without reading it, or ``None`` if unknowable.

        Starlette's parser computes ``size`` as it writes each chunk to the
        spooled file, so it counts bytes that actually arrived — it is not a
        client-declared length that could overstate the body, and a truncated
        part fails in ``req.form()`` before this is reached. The seek fallback
        covers a file object that arrived another way, and works either side
        of ``SpooledTemporaryFile``'s rollover. Both leave the stream at the
        start, because the caller still has to read it.
        """
        size = getattr(upload, "size", None)
        if size is not None:
            return int(size)
        stream = getattr(upload, "file", None)
        if stream is None or not hasattr(stream, "seek"):
            return None
        try:
            stream.seek(0, 2)
            size = stream.tell()
            stream.seek(0)
            return int(size)
        except (OSError, ValueError):
            return None

    async def _read_upload(form):
        """The uploaded bytes and an error message, if the upload is unusable.

        The size is checked *before* the body is read. Starlette spools an
        upload to a temp file above 1 MB, so the request itself costs little
        memory — it is this read that materialises it, and an oversized file
        used to be read in full only to be refused (#172).
        """
        limit = ctx.max_upload_bytes
        upload = form.get("file")
        if upload is None:
            return upload, b"", "Please choose a file to upload."

        too_big = f"File too large (max {limit // (1024 * 1024)} MB)."
        size = _upload_size(upload)
        if size is not None and size > limit:
            return upload, b"", too_big

        data = await upload.read()
        if not data:
            return upload, data, "Please choose a file to upload."
        # Belt and braces: an upload whose size could not be read beforehand
        # is still bounded here, just at the cost of having read it.
        if len(data) > limit:
            return upload, b"", too_big
        return upload, data, None

    @rt("/login", methods=["get", "post"])
    async def login(req, sess):
        error = None
        if req.method == "POST":
            form = await req.form()
            if auth.check_password(form.get("password"), expected_pw):
                sess[auth.SESSION_KEY] = True
                auth.ensure_csrf(sess)
                return _home()
            client = req.client.host if req.client else "?"
            logger.warning("[admin] Failed login attempt from %s", client)
            error = "Incorrect password."
        return views.login_page(ctx, login_path, error)

    @rt("/logout", methods=["get", "post"])
    async def logout(req, sess):
        # Only a CSRF-validated POST clears the session, so a cross-origin GET
        # (e.g. <img src="/admin/logout">) can't force-logout an admin.
        if req.method == "POST":
            form = await req.form()
            bad = _csrf_guard(sess, form)
            if bad:
                return bad
            sess.pop(auth.SESSION_KEY, None)
            return RedirectResponse(login_path, status_code=303)
        return views.logout_page(ctx, ctx.u("/logout"), auth.ensure_csrf(sess))

    @rt("/")
    def index(sess):
        return views.index_page(ctx, auth.ensure_csrf(sess))

    @rt("/status")
    def status(level: str = "info"):
        return views.status_page(ctx, level=level)

    # ---- Base templates (#169) -------------------------------------------
    # Registered BEFORE the generic /{kind}/{name}/… routes below: Starlette
    # matches in registration order, and "/base/docx/download" also fits
    # "/{kind}/{name}/download", which would redirect home instead of serving
    # the file. tests/test_admin_base_templates.py pins this ordering.

    def _slot_states():
        """One (slot, active path, source, analysis) per slot, read fresh."""
        states = []
        for s in base_templates.SLOTS:
            active = base_templates.active_path(s)
            source = base_templates.source_of(ctx.store.custom_dir, s, active)
            analysis = None
            if active is not None:
                try:
                    analysis = analyze(s.analysis_kind, active.read_bytes())
                except Exception:
                    logger.exception("[admin] Could not analyse base template %s",
                                     active)
            states.append((s, active, source, analysis))
        return states

    def _base_page(sess, focus=None, message=None, message_kind="ok"):
        return views.base_templates_page(
            ctx, csrf=auth.ensure_csrf(sess), states=_slot_states(),
            focus=focus, message=message, message_kind=message_kind,
        )

    @rt("/base")
    def base_templates_index(sess):
        return _base_page(sess)

    @rt("/base/{key}/download")
    def base_download(key: str):
        if not base_templates.is_slot(key):
            return _home()
        s = base_templates.slot(key)
        active = base_templates.active_path(s)
        if active is None:
            return views.not_found_page(ctx, f"a {s.label} base template")
        return Response(
            content=active.read_bytes(),
            media_type=media_type(active.name),
            headers={"Content-Disposition": content_disposition(active.name)},
        )

    @rt("/base/{key}/upload", methods=["post"])
    async def base_upload(req, sess, key: str):
        if not base_templates.is_slot(key):
            return _home()
        s = base_templates.slot(key)
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad

        upload, data, error = await _read_upload(form)
        if not error:
            suffix = Path(getattr(upload, "filename", "") or "").suffix.lower()
            if suffix not in s.exts:
                error = f"{s.label} expects a {s.accept} file; got {suffix or 'no'} extension."
        analysis = None
        if not error:
            analysis = analyze(s.analysis_kind, data)
            error = is_unusable(analysis)
        if error:
            return _base_page(sess, focus=key, message=error, message_kind="err")

        # The filename comes from the slot table, never from the upload.
        target = base_templates.custom_path(ctx.store.custom_dir, s)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        logger.info("[admin] Installed %s base template (%d bytes)", key, len(data))
        return _base_page(
            sess, focus=key,
            message=f"Installed {target.name}. Every new document uses it from "
                    "now on — no restart needed.")

    @rt("/base/{key}/revert", methods=["post"])
    async def base_revert(req, sess, key: str):
        if not base_templates.is_slot(key):
            return _home()
        s = base_templates.slot(key)
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad

        # unlink() decides, rather than is_file() deciding and unlink()
        # assuming: between the two, a concurrent revert (a second tab) would
        # leave this one raising FileNotFoundError into a 500.
        target = base_templates.custom_path(ctx.store.custom_dir, s)
        try:
            target.unlink()
        except FileNotFoundError:
            return _base_page(sess, focus=key,
                              message="There is no custom file to remove.",
                              message_kind="warn")
        logger.info("[admin] Removed custom %s base template", key)
        if s.has_default:
            message = f"Removed {target.name}; the bundled default is in use again."
        else:
            message = (f"Removed {target.name}. Nothing ships in its place, so no "
                       "named styles are available until you upload one.")
        return _base_page(sess, focus=key, message=message,
                          message_kind="ok" if s.has_default else "warn")

    @rt("/new/{kind}")
    def new(sess, kind: str):
        if not is_kind(kind):
            return _home()
        return views.new_page(ctx, kind, csrf=auth.ensure_csrf(sess))

    @rt("/{kind}/draft", methods=["post"])
    async def draft(req, sess, kind: str):
        if not is_kind(kind):
            return _home()
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        csrf = auth.ensure_csrf(sess)
        d = descriptor(kind)
        try:
            name = validate_name((form.get("name") or "").strip())
        except TemplateStoreError as e:
            return views.new_page(ctx, kind, csrf=csrf, error=str(e))

        upload, data, error = await _read_upload(form)
        if error:
            return views.new_page(ctx, kind, csrf=csrf, error=error)

        analysis = analyze(kind, data)
        unusable = is_unusable(analysis)
        if unusable:
            return views.new_page(ctx, kind, csrf=csrf, error=unusable)

        # Keep the uploaded extension (a .potx stays a .potx) rather than
        # forcing the canonical one onto a file that is not in that format.
        suffix = Path(getattr(upload, "filename", "") or "").suffix.lower()
        if suffix not in d.asset_exts:
            suffix = d.asset_ext
        filename = f"{name}{suffix}"
        try:
            ctx.store.write_asset(kind, filename, data)
        except TemplateStoreError as e:
            return views.new_page(ctx, kind, csrf=csrf, error=str(e))

        spec: Dict[str, Any] = {"name": name, "description": "", d.path_key: filename}
        if d.has_args:
            spec["args"] = []
        else:
            spec["strip_slides"] = True
        return views.configure_page(ctx, kind, name, filename, analysis, spec, csrf)

    @rt("/{kind}/{name}/edit")
    def edit(sess, kind: str, name: str):
        if not is_kind(kind):
            return _home()
        spec = ctx.store.get_spec(kind, name)
        if spec is None:
            return views.not_found_page(ctx, name)
        return views.edit_page(ctx, kind, name, spec, ctx.analyze_asset(kind, spec),
                               csrf=auth.ensure_csrf(sess))

    @rt("/{kind}/{name}/download")
    def download(kind: str, name: str):
        """Serve the source file a template is actually using (#162).

        Sits behind the same gate as every other route — `auth.make_before`
        now exempts nothing but the login endpoint, which matters here because
        this is the first route that serves an admin-uploaded file. The
        filename is resolved through `store.asset_path()`, whose
        `validate_asset_filename` refuses anything but a bare name, so a spec
        cannot point the download out of `custom_templates/`.
        """
        if not is_kind(kind):
            return _home()
        spec = ctx.store.get_spec(kind, name)
        if spec is None:
            return views.not_found_page(ctx, name)
        asset = spec.get(descriptor(kind).path_key)
        if not asset or not ctx.store.asset_exists(kind, asset):
            return views.not_found_page(ctx, f"{name}'s source file")
        try:
            data = ctx.store.read_asset(kind, asset)
        except TemplateStoreError:
            logger.exception("[admin] Could not read %s asset for %r", kind, name)
            return views.not_found_page(ctx, f"{name}'s source file")
        return Response(
            content=data,
            media_type=media_type(asset),
            headers={"Content-Disposition": content_disposition(asset)},
        )

    @rt("/{kind}/{name}/reupload", methods=["post"])
    async def reupload(req, sess, kind: str, name: str):
        if not is_kind(kind):
            return _home()
        spec = ctx.store.get_spec(kind, name)
        if spec is None:
            return views.not_found_page(ctx, name)
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        csrf = auth.ensure_csrf(sess)

        _upload, data, error = await _read_upload(form)
        analysis = analyze(kind, data) if data else None
        if not error and analysis:
            error = is_unusable(analysis)
        if error:
            # Prefill from the file that is still installed, but do not report on
            # it: the admin asked about the file they just submitted, and this is
            # not that file. See views.edit_page's `report`.
            return views.edit_page(ctx, kind, name, spec, ctx.analyze_asset(kind, spec),
                                   csrf=csrf, message=error, message_kind="err",
                                   report=False)

        d = descriptor(kind)
        filename = spec.get(d.path_key) or f"{name}{d.asset_ext}"
        ctx.store.write_asset(kind, filename, data)
        return views.edit_page(
            ctx, kind, name, spec, analysis, csrf=csrf,
            message=f"Re-scanned {filename}. New placeholders (if any) were added "
                    "below — review and save to apply.",
        )

    @rt("/{kind}/save", methods=["post"])
    async def save(req, sess, kind: str):
        if not is_kind(kind):
            return _home()
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        try:
            spec = build_spec(kind, form)
            ctx.store.save_spec(kind, spec)
        except TemplateStoreError as e:
            return views.save_failed_page(ctx, str(e))
        ok = ctx.register(kind, spec)
        return views.saved_page(ctx, kind, spec["name"], ok)

    @rt("/{kind}/preview", methods=["post"])
    async def preview(req, sess, kind: str):
        if not is_kind(kind):
            return _home()
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        spec = build_spec(kind, form)
        asset = spec.get(descriptor(kind).path_key)
        if not asset or not ctx.store.asset_exists(kind, asset):
            return HTMLResponse("<p>Nothing to preview yet — save the template first.</p>",
                                status_code=400)
        if kind == KIND_PPTX:
            # Rendered from the file on disk, not from bytes: open_template needs
            # a path to fall back to its .potx rewrite.
            try:
                out, warnings = render_pptx_preview(ctx.store.asset_path(kind, asset), spec)
            except Exception as e:
                logger.exception("[admin] pptx preview failed")
                return HTMLResponse(f"<p>Could not build a preview: {e}</p>", status_code=400)
            headers = {
                "Content-Disposition": f'attachment; filename="{spec["name"]}_preview.pptx"',
            }
            if warnings:
                # Surfaced in a header because the response body is the deck
                # itself; the same warnings reach the model on a real call.
                headers["X-Preview-Warnings"] = " | ".join(warnings)[:900]
            return Response(
                content=out,
                media_type=(
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                ),
                headers=headers,
            )

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

    @rt("/{kind}/{name}/delete", methods=["get", "post"])
    async def delete(req, sess, kind: str, name: str):
        if not is_kind(kind):
            return _home()
        spec = ctx.store.get_spec(kind, name)
        if spec is None:
            return views.not_found_page(ctx, name)
        asset = spec.get(descriptor(kind).path_key)
        shared_with = ctx.other_specs_using_asset(kind, name, asset)

        if req.method != "POST":
            return views.delete_page(ctx, kind, name, asset, shared_with,
                                     csrf=auth.ensure_csrf(sess))

        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        # Never honour the checkbox for a file another template still points
        # at, whatever was submitted — the page hides the option, but the POST
        # is not where that decision should be trusted from.
        drop_asset = checked(form, "delete_asset") and not shared_with
        ctx.store.delete_spec(kind, name, delete_asset=drop_asset)
        ctx.unregister(kind, name)
        logger.info("[admin] Deleted %s template %r (asset %s)", kind, name,
                    "deleted" if drop_asset else "kept")
        return _home()

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

    mcp_app = mcp.http_app(path="/mcp", stateless_http=config.stateless_http)
    admin_app = build_admin_app(mcp, config)
    routes = [
        Mount(config.admin.path, app=admin_app),
        Mount("/", app=mcp_app),
    ]
    return Starlette(routes=routes, lifespan=mcp_app.lifespan)
