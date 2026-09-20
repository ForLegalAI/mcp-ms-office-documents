"""Page rendering for the admin UI.

One module per area — :mod:`~admin.views.shell` for the wrapper and the
fragments every editor shares, :mod:`~admin.views.templates` for the template
pages, :mod:`~admin.views.base` for the base-template slots,
:mod:`~admin.views.assets` for the source-file maintenance view,
:mod:`~admin.views.status`, :mod:`~admin.views.login`. Views are pure:
they take the :class:`~admin.app.AdminContext` and already-loaded data and
return a FastHTML tree, so :mod:`admin.app` holds nothing but routes.

Markup primitives live in :mod:`admin.components`; per-kind wording lives in
:mod:`admin.kinds`.
"""
from __future__ import annotations

from admin.views.assets import assets_page, delete_asset_page, fmt_size
from admin.views.base import base_templates_page
from admin.views.login import login_page, logout_page
from admin.views.shell import (
    BRAND, form_actions, name_field, not_found_page, page, replace_card,
    save_failed_page, saved_page,
)
from admin.views.status import (
    fmt_ts, fmt_uptime, refresh_seconds, status_page,
)
from admin.views.templates import (
    analysis_report, arg_row, configure_page, delete_page, edit_form, edit_page,
    index_page,
    builtin_style_names, new_page, pptx_analysis_report, spec_yaml_block,
    style_mapping_block, template_table,
)

__all__ = [
    "BRAND", "assets_page", "base_templates_page", "analysis_report", "arg_row",
    "configure_page", "delete_asset_page", "delete_page", "fmt_size",
    "edit_form",
    "edit_page", "fmt_ts", "fmt_uptime", "form_actions", "index_page",
    "login_page", "logout_page", "name_field", "new_page", "not_found_page",
    "page", "pptx_analysis_report", "refresh_seconds", "replace_card",
    "save_failed_page",
    "saved_page", "status_page",
    "builtin_style_names", "spec_yaml_block", "style_mapping_block",
    "template_table",
]
