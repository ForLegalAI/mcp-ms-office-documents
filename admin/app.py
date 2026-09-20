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
from admin import assets as asset_index
from admin.analysis import analyze, is_unusable
from admin.components import head_tags
from admin.forms import build_spec, carried_spec, checked, parse_style_mapping
from admin.kinds import (
    KINDS, content_disposition, descriptor, is_kind, media_type,
)
from admin.preview import (
    has_submitted_values, render_docx_preview, render_email_preview,
    render_pptx_preview, sample_values, values_from_form,
)
from admin.store import (
    FileTemplateStore, KIND_DOCX, KIND_PPTX, TemplateStoreError, validate_name,
)
from template_registry import (
    gather_specs, is_enabled, overlay_global, read_master_yaml,
)
from template_utils import find_file_in_template_dirs

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
    def docx_master_yaml(self) -> Path:
        """The hand-written master Word template config. Never rewritten here."""
        return self.store.config_dir / "docx_templates.yaml"

    @property
    def global_style_mapping(self) -> Dict[str, Any]:
        """The ``style_mapping`` actually in force for Word.

        The master YAML under whatever the UI has stored in the ``.d`` layer,
        re-read each access so an edit — here or on the volume — takes effect
        without a restart (the admin UI is about live editing). Goes through
        ``global_config`` so this page, the dynamic loader and the static Word
        tool cannot disagree about which mapping applies (#161).
        """
        cfg = overlay_global(read_master_yaml(self.docx_master_yaml),
                             self.store.global_settings(KIND_DOCX))
        return cfg.get("style_mapping") or {}

    @property
    def master_style_mapping(self) -> Dict[str, Any]:
        """What the master YAML alone declares, ignoring the UI's override.

        Shown beside the editable mapping so a master setting the override has
        switched off is visible rather than merely gone.
        """
        return read_master_yaml(self.docx_master_yaml).get("style_mapping") or {}

    def resync_docx_style_map(self) -> int:
        """Re-register every live Word template tool; returns how many are live.

        The global mapping is *baked into each tool at registration time* —
        ``_register_single_template`` resolves ``build_style_map(global,
        per-template)`` once and closes over the result — so saving a new
        global mapping without this leaves every existing tool rendering on the
        old one, and only the static Word tool (which resolves per document)
        would follow the change. That split is precisely the kind of thing an
        admin would never think to suspect.
        """
        from docx_tools.dynamic_docx_tools import (
            register_docx_template_tools_from_yaml, registered_docx_template_names,
        )

        register_docx_template_tools_from_yaml(self.mcp, self.docx_master_yaml)
        return len(registered_docx_template_names())

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

    def sync(self, kind: str, spec: Dict[str, Any]) -> bool:
        """Bring the live server in line with *spec*'s ``enabled`` flag.

        The one place that decides whether saving a template registers it or
        takes it off: a disabled spec must not come back as a live tool just
        because it was saved (#165).
        """
        if is_enabled(spec):
            return self.register(kind, spec)
        # Taking it off is the outcome asked for, so this succeeded. Returning
        # unregister()'s own bool reported failure whenever there was no live
        # tool to remove — which is every disabled save and every clone of a
        # disabled template — and the page then told the admin to go and read
        # the logs about a registration that was never meant to happen.
        self.unregister(kind, spec.get("name"))
        return True

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
            templates, _cfg = gather_specs(master, None, include_disabled=True)
        except Exception:
            logger.exception("[admin] Could not read the %s master YAML", kind)
            return []
        return [t for t in templates if isinstance(t, dict)]

    def scan_assets(self) -> List[asset_index.AssetFile]:
        """Every file in ``custom_templates/`` and what references it (#166)."""
        refs = asset_index.reference_map(self.store, self._master_specs)
        return asset_index.scan(self.store.custom_dir, refs)

    def orphan(self, filename: str) -> Optional[asset_index.AssetFile]:
        """The scanned, unreferenced file called *filename*, or ``None``.

        Deletion is keyed off this rather than off the URL: the name has to
        match one the scan produced, which is a real basename read from the
        directory, and it has to still be an orphan at the moment of the
        request. A file that gained a reference between the page render and
        the click is no longer deletable, and a crafted path never matches
        anything at all.
        """
        for f in self.scan_assets():
            if f.name == filename and f.orphaned:
                return f
        return None

    def master_spec(self, kind: str, name: str) -> Optional[Dict[str, Any]]:
        """One hand-written master-YAML entry by name, or ``None`` (#167)."""
        for spec in self._master_specs(kind):
            if spec.get("name") == name:
                return spec
        return None

    def unmanaged_master_specs(self, kind: str) -> List[Dict[str, Any]]:
        """Master entries no managed spec overrides.

        Listed from the master YAML rather than from the live tool names, so a
        master template that is disabled — or that failed to register — is
        still shown and can still be inspected and adopted (#167).
        """
        managed = {s.get("name") for s in self.store.list_specs(kind)}
        return [s for s in self._master_specs(kind)
                if s.get("name") and s.get("name") not in managed]

    def master_asset_path(self, kind: str, spec: Dict[str, Any]):
        """Where a master entry's source file actually resolves.

        Through ``template_utils``, the way the loaders resolve it, so a file
        that ships in ``default_templates/`` is found as readily as one in
        ``custom_templates/``.
        """
        filename = spec.get(descriptor(kind).path_key)
        if not filename:
            return None
        return find_file_in_template_dirs(filename)

    def adopt(self, kind: str, name: str) -> Dict[str, Any]:
        """Copy a master-YAML entry into the managed ``*.d`` layer (#167).

        The master file is never touched: the merge in `gather_specs()` lets a
        `.d` entry of the same name win, so adopting is purely additive and
        the admin's hand-written YAML stays as they left it.

        The asset is copied into the writable directory when it is not already
        there — a master template usually points at a file in
        ``default_templates/``, and a template you can edit but whose document
        you cannot replace is a confusing half-state.
        """
        spec = self.master_spec(kind, name)
        if spec is None:
            raise TemplateStoreError(f"No master-YAML {kind} template named {name!r}.")
        if self.store.get_spec(kind, name) is not None:
            raise TemplateStoreError(
                f"{name!r} is already managed here — edit it directly."
            )

        filename = spec.get(descriptor(kind).path_key)
        if not filename:
            raise TemplateStoreError(
                f"{name!r} records no source file, so there is nothing to adopt."
            )
        reserved = filename in base_templates.RESERVED_FILENAMES
        data = None
        if reserved or not self.store.asset_exists(kind, filename):
            found = find_file_in_template_dirs(filename)
            if found is None:
                raise TemplateStoreError(
                    f"Cannot find {filename!r} in the template directories."
                )
            data = found.read_bytes()
        if reserved:
            # Never let an adopted template own a base-template filename.
            # template_utils searches custom_templates/ first, so a copy under
            # this name would shadow the base template for its kind — and
            # replacing this one template's document would then restyle every
            # document the server generates. Give it a private copy instead.
            filename = f"{name}{Path(filename).suffix}"
            logger.info("[admin] Adopting %r under %r: the master entry named a "
                        "base-template file", name, filename)
        return self.store.save_spec(kind, dict(spec), asset_bytes=data,
                                    asset_filename=filename)

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
    def status(level: str = "info", logger: str = "", q: str = "",
               refresh: str = "0"):
        """The Status page. Every filter is a query parameter, so a filtered
        view is a URL that can be bookmarked or pasted to a colleague."""
        return views.status_page(ctx, level=level, source=logger, search=q,
                                 refresh=views.refresh_seconds(refresh))

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

    # ---- Source files (#166) ---------------------------------------------
    # Registered BEFORE the generic /{kind}/{name}/… routes below, or
    # /files/{name}/delete would be matched as a template delete.

    @rt("/files")
    def files_index(sess):
        return views.assets_page(ctx, ctx.scan_assets(),
                                 csrf=auth.ensure_csrf(sess))

    @rt("/files/{filename}/delete", methods=["get", "post"])
    async def delete_file(req, sess, filename: str):
        """Remove one unreferenced file from ``custom_templates/``.

        Only an orphan, and only one the scan just confirmed is still an
        orphan — the referenced files are the base templates and everything a
        spec points at, and deleting one of those would silently change what
        the server produces.
        """
        found = ctx.orphan(filename)
        if found is None:
            return views.not_found_page(ctx, filename)

        if req.method != "POST":
            return views.delete_asset_page(ctx, found.name,
                                           csrf=auth.ensure_csrf(sess))

        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        try:
            # unlink() removes the directory entry, so a symlink placed here
            # by hand takes only the link with it — never what it points at.
            # Together with the name coming from iterdir(), that is why this
            # route cannot reach a file outside custom_templates/.
            (Path(ctx.store.custom_dir) / found.name).unlink()
        except OSError as e:
            logger.exception("[admin] Could not delete %s", found.name)
            return views.assets_page(
                ctx, ctx.scan_assets(), csrf=auth.ensure_csrf(sess),
                message=f"Could not delete {found.name}: {e}",
                message_kind="err")
        logger.info("[admin] Deleted unreferenced source file %r", found.name)
        return views.assets_page(
            ctx, ctx.scan_assets(), csrf=auth.ensure_csrf(sess),
            message=f"Deleted {found.name}.")

    # ---- Global style mapping (#161) -------------------------------------
    # Registered BEFORE the generic /{kind}/… routes below: "/styles/save"
    # also fits "/{kind}/save" with kind="styles", which would 404 through
    # is_kind() instead of reaching this handler. Starlette matches in
    # registration order; tests/test_admin_global_styles.py pins it.

    def _styles_page(sess, message=None, message_kind="ok"):
        return views.global_styles_page(
            ctx, csrf=auth.ensure_csrf(sess),
            mapping=ctx.global_style_mapping,
            master=ctx.master_style_mapping,
            stored=ctx.store.has_global_settings(KIND_DOCX),
            offered=_base_docx_styles(),
            message=message, message_kind=message_kind,
        )

    def _base_docx_styles():
        """Styles the base Word template defines — what this mapping can pick.

        The base template is the right reference here and not a guess: this
        mapping applies to every document, and the static Word tool renders
        onto exactly this file. A style named in the mapping but missing from
        it is still offered by the view, marked, so an existing value is never
        dropped for being unrecognised.
        """
        slot = base_templates.slot("docx")
        active = base_templates.active_path(slot)
        if active is None:
            return []
        try:
            return list(analyze(slot.analysis_kind, active.read_bytes()).styles_present)
        except Exception:
            logger.exception("[admin] Could not read styles from %s", active)
            return []

    @rt("/styles")
    def global_styles(sess):
        return _styles_page(sess)

    @rt("/styles/save", methods=["post"])
    async def save_global_styles(req, sess):
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        mapping = parse_style_mapping(form)
        try:
            ctx.store.save_global_settings(KIND_DOCX, {"style_mapping": mapping})
        except (TemplateStoreError, OSError) as e:
            logger.exception("[admin] Could not save the global style mapping")
            return _styles_page(sess, f"Could not save: {e}", "err")
        live = ctx.resync_docx_style_map()
        logger.info("[admin] Global style mapping set to %r", mapping)
        return _styles_page(
            sess,
            f"Saved. {len(mapping)} key(s) overridden; {live} Word template "
            "tool(s) re-registered on the new mapping.")

    @rt("/styles/revert", methods=["post"])
    async def revert_global_styles(req, sess):
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        try:
            removed = ctx.store.clear_global_settings(KIND_DOCX)
        except (TemplateStoreError, OSError) as e:
            logger.exception("[admin] Could not clear the global style mapping")
            return _styles_page(sess, f"Could not revert: {e}", "err")
        live = ctx.resync_docx_style_map()
        if not removed:
            return _styles_page(sess, "There was nothing stored to revert.", "info")
        logger.info("[admin] Global style mapping reverted to the master YAML")
        return _styles_page(
            sess,
            "Reverted to config/docx_templates.yaml. "
            f"{live} Word template tool(s) re-registered.")

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

    @rt("/{kind}/{name}/master")
    def master_detail(sess, kind: str, name: str):
        """A hand-written master-YAML template, read only (#167)."""
        if not is_kind(kind):
            return _home()
        spec = ctx.master_spec(kind, name)
        if spec is None:
            return views.not_found_page(ctx, name)
        return _master_detail_page(sess, kind, name, spec)

    def _master_detail_page(sess, kind, name, spec, error=None):
        asset = ctx.master_asset_path(kind, spec)
        analysis = None
        if asset is not None:
            try:
                analysis = analyze(kind, asset.read_bytes())
            except Exception:
                logger.exception("[admin] Could not analyse master template %s", name)
        return views.master_page(
            ctx, kind, name, spec, analysis, asset,
            live=name in ctx.live_names(kind),
            csrf=auth.ensure_csrf(sess), error=error,
        )

    @rt("/{kind}/{name}/adopt", methods=["post"])
    async def adopt(req, sess, kind: str, name: str):
        """Copy a master-YAML entry into the managed layer, then edit it."""
        if not is_kind(kind):
            return _home()
        spec = ctx.master_spec(kind, name)
        if spec is None:
            return views.not_found_page(ctx, name)
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        try:
            adopted = ctx.adopt(kind, name)
        except (TemplateStoreError, OSError) as e:
            return _master_detail_page(sess, kind, name, spec, error=str(e))

        ctx.sync(kind, adopted)
        logger.info("[admin] Adopted %s template %r from the master YAML", kind, name)
        return views.edit_page(
            ctx, kind, name, adopted, ctx.analyze_asset(kind, adopted),
            csrf=auth.ensure_csrf(sess),
            message=("Adopted from your master YAML. Edits here take "
                     "precedence; the YAML file itself is untouched."),
        )

    @rt("/{kind}/{name}/clone", methods=["get", "post"])
    async def clone(req, sess, kind: str, name: str):
        """Start a new template from an existing one (#164).

        "Same as the formal letter, but for the Prague office" used to mean
        creating one from scratch and re-typing every argument name, type,
        default and description by hand.
        """
        if not is_kind(kind):
            return _home()
        spec = ctx.store.get_spec(kind, name)
        if spec is None:
            return views.not_found_page(ctx, name)
        csrf = auth.ensure_csrf(sess)

        if req.method != "POST":
            return views.clone_page(ctx, kind, name, spec, csrf=csrf)

        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        try:
            copy = ctx.store.clone_spec(kind, name, (form.get("name") or "").strip())
        except (TemplateStoreError, OSError) as e:
            return views.clone_page(ctx, kind, name, spec, csrf=csrf, error=str(e))

        ok = ctx.sync(kind, copy)
        new_name = copy["name"]
        logger.info("[admin] Cloned %s template %r -> %r", kind, name, new_name)
        # Land on the copy's edit page, not the index: the description almost
        # always needs changing immediately, and the point is to keep going.
        note = f"Copied from {name}. Change what differs, then save."
        if not is_enabled(copy):
            note += (" It is disabled, like the template it was copied from — "
                     "enable it from the template list when you are ready.")
        elif not ok:
            note += " (It is not live yet — save to register it.)"
        return views.edit_page(ctx, kind, new_name, copy,
                               ctx.analyze_asset(kind, copy),
                               csrf=csrf, message=note)

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
            # `original_name` is rendered on the edit form and not on the
            # create form, so its absence means "create". Keying off that
            # rather than off finding a stored spec matters: a create that
            # lands on an existing name must not inherit that template's
            # state, and must not silently overwrite it either (#165).
            previous = (form.get("original_name") or "").strip()
            stored = ctx.store.get_spec(kind, previous) if previous else None

            if stored is None and ctx.store.get_spec(kind, spec["name"]) is not None:
                raise TemplateStoreError(
                    f"A {kind} template named {spec['name']!r} already exists. "
                    "Edit that one, or choose another name."
                )
            # The edit form carries no `enabled` control, so a save rebuilds
            # the spec without it. Carry the stored flag forward or an edit
            # would silently switch a disabled template back on.
            if stored is not None and not is_enabled(stored):
                spec["enabled"] = False

            renamed = stored is not None and previous != spec["name"]
            if renamed:
                # Through the store, which owns the collision check, the name
                # validation and the write-before-unlink ordering. Doing it
                # inline here duplicated that invariant in a copy no test
                # covered.
                ctx.store.rename_spec(kind, previous, spec["name"])
            ctx.store.save_spec(kind, spec)
            if renamed:
                ctx.unregister(kind, previous)
                logger.info("[admin] Renamed %s template %r -> %r",
                            kind, previous, spec["name"])
        except (TemplateStoreError, OSError) as e:
            # OSError too: a failed unlink inside rename_spec is a save
            # failure to report, not an unhandled 500.
            return views.save_failed_page(ctx, str(e))
        ok = ctx.sync(kind, spec)
        return views.saved_page(ctx, kind, spec["name"], ok,
                                enabled=is_enabled(spec))

    @rt("/{kind}/preview/values", methods=["post"])
    async def preview_values(req, sess, kind: str):
        """The form for previewing with your own values (#168).

        Reached from the edit form, so it carries whatever is on screen —
        unsaved edits included — rather than what is saved on disk.
        """
        if not is_kind(kind):
            return _home()
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        try:
            spec = build_spec(kind, form)
        except TemplateStoreError as e:
            return views.save_failed_page(ctx, str(e))
        asset = spec.get(descriptor(kind).path_key)
        conditionals: List[str] = []
        if asset and ctx.store.asset_exists(kind, asset):
            try:
                conditionals = list(analyze(kind, ctx.store.read_asset(kind, asset))
                                    .conditionals)
            except Exception:
                logger.exception("[admin] Could not analyse %s for preview", asset)
        return views.preview_values_page(
            ctx, kind, spec, sample_values(spec.get("args") or [], conditionals),
            conditionals, csrf=auth.ensure_csrf(sess))

    @rt("/{kind}/preview", methods=["post"])
    async def preview(req, sess, kind: str):
        if not is_kind(kind):
            return _home()
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        # The values form carries the spec it was built from, so a preview
        # with your own values reflects the unsaved edits you started from.
        spec = carried_spec(form) or build_spec(kind, form)
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
        values = (values_from_form(form, spec.get("args") or [], analysis.conditionals)
                  if has_submitted_values(form)
                  else sample_values(spec.get("args") or [], analysis.conditionals))
        if kind == KIND_DOCX:
            out = render_docx_preview(data, spec, values, ctx.global_style_mapping)
            return Response(
                content=out,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={"Content-Disposition": f'attachment; filename="{spec["name"]}_preview.docx"'},
            )
        html = render_email_preview(data, spec, values)
        return HTMLResponse(html)

    @rt("/{kind}/{name}/enabled", methods=["post"])
    async def set_enabled(req, sess, kind: str, name: str):
        """Turn a template's tool on or off without touching its configuration.

        Deleting used to be the only off switch, which meant destroying the
        arguments and descriptions to stop the AI reaching for a seasonal or
        under-review template (#165).
        """
        if not is_kind(kind):
            return _home()
        form = await req.form()
        bad = _csrf_guard(sess, form)
        if bad:
            return bad
        if ctx.store.get_spec(kind, name) is None:
            return views.not_found_page(ctx, name)
        want = checked(form, "enabled")
        try:
            spec = ctx.store.set_enabled(kind, name, want)
        except TemplateStoreError as e:
            return views.save_failed_page(ctx, str(e))
        ctx.sync(kind, spec)
        return _home()

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
