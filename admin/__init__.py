"""Optional FastHTML template-admin UI and its supporting modules.

This package is **opt-in** (gated by ``ADMIN_ENABLED``) and is imported lazily
so that a deployment which never enables the admin UI pays no import cost and
needs none of the admin-only dependencies (e.g. python-fasthtml).

``main.py`` reaches it through :func:`admin.app.build_combined_app`, which
mounts the UI beside the MCP endpoint in one ASGI app.

Module map — see ``docs/development/admin-ui.md`` for the invariants:

======================  ====================================================
:mod:`admin.app`        ``AdminContext`` and the routes
:mod:`admin.components` markup primitives and the inlined theme
:mod:`admin.kinds`      one descriptor per template kind
:mod:`admin.forms`      reading a submitted form back into a spec
:mod:`admin.views`      the pages
:mod:`admin.store`      persistence for UI-managed templates
:mod:`admin.analysis`   what is inside an uploaded document
:mod:`admin.preview`    rendering a template without uploading it
:mod:`admin.auth`       the shared-password gate and CSRF tokens
======================  ====================================================

Only :mod:`admin.app`, :mod:`admin.components` and :mod:`admin.views` require
FastHTML; the rest are importable (and unit-testable) on their own.
"""
from __future__ import annotations

__all__ = ["analysis", "auth", "components", "forms", "kinds", "preview", "store", "views"]
