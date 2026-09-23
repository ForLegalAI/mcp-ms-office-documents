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
import secrets
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import metrics

from fasthtml.common import FastHTML, RedirectResponse, Response, HTMLResponse, to_xml
from starlette.routing import Mount, Route

from config import Config
from admin import auth, base_templates, views
from admin import assets as asset_index
from admin.analysis import analyze, is_unusable
from admin.components import head_tags
from admin.forms import build_spec, carried_spec, checked, parse_style_mapping
from admin.kinds import (
    KINDS, content_disposition, descriptor, is_kind, media_type,
)
from admin.sections import (
    MOVED_PATHS, SECTIONS, TAB_BASE, TAB_FILES, TAB_LOG, TAB_OVERVIEW,
    TAB_STATUS, TAB_STYLES, TAB_TEMPLATES, assert_slugs_free, section,
    slot_section,
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

    def resync_docx_style_map(self) -> Tuple[int, List[str]]:
        """Re-register every live Word template tool for a changed global map.

        The global mapping is *baked into each tool at registration time* —
        ``_register_single_template`` resolves ``build_style_map(global,
        per-template)`` once and closes over the result — so saving a new
        global mapping without this leaves every existing tool rendering on the
        old one, and only the static Word tool (which resolves per document)
        would follow the change. That split is precisely the kind of thing an
        admin would never think to suspect.

        Returns ``(live, lost)``. ``lost`` is why the before/after comparison
        is here at all: ``register_docx_template()`` removes the existing tool
        *before* rebuilding it, so a template whose source file has gone
        missing since startup does not merely fail to re-register — it is
        taken off the server, and the loop that registers them logs the
        failure and moves on. The caller has to be able to say so, because
        from the admin's side the only thing that happened was saving a style.
        """
        from docx_tools.dynamic_docx_tools import (
            register_docx_template_tools_from_yaml, registered_docx_template_names,
        )
        from docx_tools.style_map import invalidate_global_style_map

        # The static tool's cache is fingerprinted on the config files' mtime
        # and size, which cannot be trusted to notice two writes inside one
        # filesystem timestamp tick. A write we made ourselves we simply know
        # about, so say so rather than hope the fingerprint moved.
        invalidate_global_style_map()
        before = set(registered_docx_template_names())
        register_docx_template_tools_from_yaml(self.mcp, self.docx_master_yaml)
        after = registered_docx_template_names()
        lost = sorted(before - set(after))
        if lost:
            logger.error("[admin] Re-registration dropped Word template(s): %s",
                         ", ".join(lost))
        return len(after), lost

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

    # -- what an Overview tab reports --------------------------------------

    def section_facts(self, s) -> "views.SectionFacts":
        """Counts, tools and base-template state for one section's Overview.

        Loaded here rather than in the view because it reaches into the store,
        the live registries and the metrics counters — the three things a view
        is not allowed to know about. The view gets a plain record.
        """
        # Both halves of what the Templates tab lists. A template written by
        # hand in the master YAML is as live as one made here — it registers
        # the same tool — so counting only the managed specs made this page
        # disagree with the Templates tab beside it and with the Server
        # section's "live Word tools", and left a tool the AI can call out of
        # the table headed "Tools the AI can call".
        managed = self.store.list_specs(s.kind) if s.kind else []
        from_master = self.unmanaged_master_specs(s.kind) if s.kind else []
        specs = [(spec, "template") for spec in managed]
        specs += [(spec, "master") for spec in from_master]
        live_names = set(self.live_names(s.kind)) if s.kind else set()

        tools = [self._tool_fact(name, origin="static") for name in s.tools]
        # A docx or email template is a tool of its own and belongs in the
        # same list; a pptx template is not, so it is counted as a template
        # and nothing else. descriptor().has_args is the distinction.
        if s.kind and descriptor(s.kind).has_args:
            for spec, origin in specs:
                name = spec.get("name")
                tools.append(self._tool_fact(
                    name, origin=origin, live=name in live_names))

        slots = []
        for slot in s.slots:
            active = base_templates.active_path(slot)
            slots.append(views.SlotFact(
                key=slot.key, label=slot.label,
                filename=active.name if active else "",
                source=base_templates.source_of(self.store.custom_dir, slot,
                                                active),
            ))

        notes = []
        for slot, fact in zip(s.slots, slots):
            if fact.source == "none":
                notes.append((
                    f"No {slot.label} file is installed. {slot.controls}",
                    "warn"))
        if from_master:
            notes.append((
                f"{len(from_master)} template(s) come from the hand-written "
                "master YAML. They are counted here and shown read-only on "
                "the Templates tab.", "info"))

        extra = []
        if s.kind == KIND_PPTX:
            extra.append(("Registered designs", len(live_names)))

        return views.SectionFacts(
            templates=len(specs),
            live=len([spec for spec, _origin in specs
                      if spec.get("name") in live_names]),
            disabled=len([spec for spec, _origin in specs
                          if not is_enabled(spec)]),
            tools=tuple(tools), slots=tuple(slots), notes=tuple(notes),
            extra_stats=tuple(extra),
        )

    @staticmethod
    def _tool_fact(name: str, origin: str = "static",
                   live: bool = True) -> "views.ToolFact":
        """One tool's counters, or zeroes when it has not been called yet."""
        st = metrics.get_tool_stat(name)
        if st is None:
            return views.ToolFact(name=name, origin=origin, live=live)
        return views.ToolFact(
            name=name, origin=origin, live=live, calls=st.calls,
            errors=st.errors, degraded=st.degraded,
            last_called=views.fmt_ts(st.last_called),
        )

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
    # restarts. With no password there is no login to survive, and a secret
    # derived from nothing would be a constant anyone can read in this file —
    # enough to sign a session cookie by hand. So it is random per process,
    # and the gate is locked besides (`auth.make_before(locked=True)`).
    if expected_pw:
        secret = hashlib.sha256(f"mcp-office-admin:{expected_pw}".encode()).hexdigest()
    else:
        secret = secrets.token_hex(32)

    # Self-contained headers (no CDN): meta + theme CSS + the row JS, with the
    # blank-row HTML injected so "Add argument" can clone it client-side.
    hdrs = head_tags(json.dumps(to_xml(views.arg_row())))

    app = FastHTML(secret_key=secret, before=auth.make_before(login_path, locked=not expected_pw),
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

    # ---- Sections (the tabbed product pages) -----------------------------
    # Registered BEFORE the /{kind}/... routes below, and GET-only. Both
    # matter: a section slug shares the first path segment with a template
    # kind ("email" is both a section and a kind), and every two-segment
    # /{kind}/... route is a POST, so GET-only is what keeps "/email" the
    # Email section while "/email/save" still reaches the save handler.
    # The slugs are literal rather than a "/{slug}" pattern for the same
    # reason: a catch-all would also swallow "/new/docx".
    assert_slugs_free()

    def _slot_states(keys=None):
        """One (slot, active path, source, analysis) per slot, read fresh.

        *keys* limits it to one section's slots; ``None`` means all of them.
        """
        states = []
        for s in base_templates.SLOTS:
            if keys is not None and s.key not in keys:
                continue
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

    def _panel(sec, tab, sess, *, focus=None, message=None, message_kind="ok",
               query=None):
        """The body of one section tab, as a list of cards.

        The single place that maps a (section, tab) pair to a renderer. Every
        route that has to re-render a tab after a POST - an upload, a revert,
        a style save - comes back through here rather than building the page
        its own way, so a tab looks the same however it was reached.
        """
        csrf = auth.ensure_csrf(sess)
        query = query or {}
        if tab == TAB_OVERVIEW:
            return views.overview_panel(ctx, sec, ctx.section_facts(sec))
        if tab == TAB_TEMPLATES:
            return [views.templates_panel(ctx, sec.kind, csrf)]
        if tab == TAB_BASE:
            return views.base_panel(
                ctx, csrf=csrf, states=_slot_states(sec.slot_keys),
                focus=focus, message=message, message_kind=message_kind)
        if tab == TAB_STYLES:
            return views.style_mapping_panel(
                ctx, csrf=csrf, mapping=ctx.global_style_mapping,
                master=ctx.master_style_mapping,
                stored=ctx.store.has_global_settings(KIND_DOCX),
                offered=_base_docx_styles(),
                message=message, message_kind=message_kind)
        if tab == TAB_STATUS:
            return views.status_panel(ctx)
        if tab == TAB_FILES:
            return views.source_files_panel(
                ctx, ctx.scan_assets(), csrf=csrf,
                message=message or "", message_kind=message_kind)
        if tab == TAB_LOG:
            return views.log_panel(
                ctx, level=query.get("level", "info"),
                source=query.get("logger", ""), search=query.get("q", ""),
                refresh=views.refresh_seconds(query.get("refresh", "0")))
        raise KeyError(f"No renderer for tab {tab!r}")  # pragma: no cover

    def _section_page(sec, tab, sess, **kw):
        """A section tab, wrapped in the section shell."""
        actions = []
        if tab == TAB_TEMPLATES:
            actions.append(views.new_template_button(ctx, sec.kind))
        return views.section_page(ctx, sec, tab, *_panel(sec, tab, sess, **kw),
                                  actions=actions)

    def _section_redirect(sec, tab=None):
        return RedirectResponse(sec.href(ctx.u, tab), status_code=303)

    @rt("/")
    def index(sess):
        auth.ensure_csrf(sess)
        return views.dashboard_page(
            ctx, [(sec, ctx.section_facts(sec)) for sec in SECTIONS])

    def _register_section(sec):
        """One section's two GET routes: its default tab, and a named tab.

        An unknown tab redirects to the section rather than 404ing: a stale
        bookmark to a tab that has been renamed is far likelier than a
        hand-typed URL, and the section it names is still where to land.

        *sec* is captured in the closure and must not become a parameter with
        a default: FastHTML fills every parameter of a handler from the
        request, so a "private" keyword argument arrives as None and the
        handler looks up a section called None.
        """
        def index(sess):
            return _section_page(sec, sec.default_tab, sess)

        def tab(req, sess, tab: str):
            if not sec.has_tab(tab):
                return _section_redirect(sec)
            return _section_page(sec, tab, sess, query=dict(req.query_params))

        index.__name__ = f"section_{sec.slug}"
        tab.__name__ = f"section_{sec.slug}_tab"
        rt(f"/{sec.slug}", methods=["get"])(index)
        rt(f"/{sec.slug}/{{tab}}", methods=["get"])(tab)

    for _sec in SECTIONS:
        _register_section(_sec)

    # ---- Legacy URLs ------------------------------------------------------
    # The pages these named are now tabs. Redirected rather than dropped:
    # they were in the top bar for the UI's whole life, so they are in
    # bookmarks and in links people have pasted to each other.

    def _register_moved(old_path, slug, tab):
        # A factory, not a default argument: FastHTML fills a handler's
        # parameters from the request, so a captured value passed that way
        # arrives as None (see _register_section).
        def moved():
            return _section_redirect(section(slug), tab)
        moved.__name__ = f"moved{old_path.replace('/', '_')}"
        rt(old_path, methods=["get"])(moved)

    for _old, (_slug, _tab) in MOVED_PATHS.items():
        _register_moved(_old, _slug, _tab)

    # ---- Base templates (#169) -------------------------------------------
    # Registered BEFORE the generic /{kind}/{name}/... routes below:
    # Starlette matches in registration order, and "/base/docx/download" also
    # fits "/{kind}/{name}/download", which would redirect home instead of
    # serving the file. tests/test_admin_base_templates.py pins this ordering.

    def _base_page(sess, focus=None, message=None, message_kind="ok"):
        """Re-render the Base tab of whichever section owns *focus*'s slot."""
        sec = slot_section(focus) if focus else section("word")
        return _section_page(sec, TAB_BASE, sess, focus=focus, message=message,
                             message_kind=message_kind)

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

    def _files_page(sess, message=None, message_kind="ok"):
        """Re-render the Server section's Source files tab."""
        return _section_page(section("server"), TAB_FILES, sess,
                             message=message, message_kind=message_kind)

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
            return _files_page(sess, f"Could not delete {found.name}: {e}",
                               "err")
        logger.info("[admin] Deleted unreferenced source file %r", found.name)
        return _files_page(sess, f"Deleted {found.name}.")

    # ---- Global style mapping (#161) -------------------------------------
    # Registered BEFORE the generic /{kind}/… routes below: "/styles/save"
    # also fits "/{kind}/save" with kind="styles", which would 404 through
    # is_kind() instead of reaching this handler. Starlette matches in
    # registration order; tests/test_admin_global_styles.py pins it.

    def _resync_message(ok: str, lost):
        """The flash for a save, worded for what the re-registration did.

        A template that could not be rebuilt is *gone*, not merely stale, and
        the admin was only saving a style — so that outcome must not arrive as
        a success message with a smaller number in it.
        """
        if not lost:
            return ok, "ok"
        return (
            ok + " But " + ", ".join(lost) + " could not be rebuilt and is no "
            "longer registered — its source file is most likely missing. Check "
            "Source files, then re-upload it on the template's page.",
            "warn",
        )

    def _styles_page(sess, message=None, message_kind="ok"):
        """Re-render the Word section's Style mapping tab."""
        return _section_page(section("word"), TAB_STYLES, sess,
                             message=message, message_kind=message_kind)

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
        live, lost = ctx.resync_docx_style_map()
        logger.info("[admin] Global style mapping set to %r", mapping)
        return _styles_page(sess, *_resync_message(
            f"Saved. {len(mapping)} key(s) overridden; {live} Word template "
            "tool(s) re-registered on the new mapping.", lost))

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
        live, lost = ctx.resync_docx_style_map()
        if not removed:
            return _styles_page(sess, "There was nothing stored to revert.", "info")
        logger.info("[admin] Global style mapping reverted to the master YAML")
        return _styles_page(sess, *_resync_message(
            "Reverted to config/docx_templates.yaml. "
            f"{live} Word template tool(s) re-registered.", lost))

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

    The bare mount path needs a redirect of its own. Starlette compiles
    ``Mount("/admin")`` to ``^/admin/(?P<path>.*)$``, which does **not** match
    ``/admin`` — and its router only offers the missing-trailing-slash redirect
    when *nothing* matched at all. The catch-all MCP mount matches everything,
    so a plain ``GET /admin`` reached the MCP app and came back 404: the
    address an admin is given, and the one they type, was the one that did not
    work.
    """
    from starlette.applications import Starlette

    mcp_app = mcp.http_app(path="/mcp", stateless_http=config.stateless_http)
    admin_app = build_admin_app(mcp, config)
    async def _admin_root_redirect(request):
        """Send the bare mount path to the mount root, as Starlette would.

        The target is derived from the **request**, not from
        ``config.admin.path``, and that is the whole subtlety. Starlette's own
        ``redirect_slashes`` builds its ``Location`` from the request scope, so
        it folds in ``root_path`` and keeps the query string. A ``Location``
        built from the configured path instead looks identical in development
        and is wrong behind a path-rewriting proxy or ``uvicorn --root-path``:
        ``GET /office/admin`` would answer ``/admin/``, sending the browser to
        a path that does not exist on that deployment.

        307, also as ``redirect_slashes`` uses: it preserves the method, and
        nothing here wants the caching a permanent redirect invites.
        """
        url = request.url
        return RedirectResponse(str(url.replace(path=url.path + "/")),
                                status_code=307)

    routes = [
        # Before the mounts: a Route for the exact path, so it is matched
        # rather than swallowed by the catch-all below.
        Route(config.admin.path, endpoint=_admin_root_redirect,
              methods=["GET", "HEAD"]),
        Mount(config.admin.path, app=admin_app),
        Mount("/", app=mcp_app),
    ]
    return Starlette(routes=routes, lifespan=mcp_app.lifespan)
