"""Page rendering for the admin UI.

One module per area — :mod:`~admin.views.shell` for the wrapper and the
fragments every editor shares, :mod:`~admin.views.sections` for the tabbed
section shell and the dashboard, :mod:`~admin.views.templates` for the
template pages, :mod:`~admin.views.base` for the base-template slots,
:mod:`~admin.views.assets` for the source-file maintenance view,
:mod:`~admin.views.settings` for the Word style mapping,
:mod:`~admin.views.status`, :mod:`~admin.views.login`. Views are pure:
they take the :class:`~admin.app.AdminContext` and already-loaded data and
return a FastHTML tree, so :mod:`admin.app` holds nothing but routes.

A page that lives under a section tab is a **panel** — a list of cards with no
shell of its own — and :func:`~admin.views.sections.section_page` wraps it.
That is what lets one renderer serve a tab without knowing which section it is
being shown under. A page reached from a tab but not itself a tab (an editor, a
delete confirmation) is a full page and says which section it belongs to, so
the nav keeps the right item lit.

Markup primitives live in :mod:`admin.components`; per-kind wording lives in
:mod:`admin.kinds`; what the sections and their tabs are lives in
:mod:`admin.sections`.
"""
from __future__ import annotations

from admin.views.assets import delete_asset_page, fmt_size, source_files_panel
from admin.views.base import base_panel
from admin.views.login import login_page, logout_page
from admin.views.sections import (
    SectionFacts, SlotFact, ToolFact, dashboard_page, overview_panel,
    section_page,
)
from admin.views.settings import style_mapping_panel
from admin.views.shell import (
    BRAND, back_link, form_actions, kind_page, name_field, not_found_page,
    page, replace_card, save_failed_page, saved_page,
)
from admin.views.status import (
    fmt_ts, fmt_uptime, log_panel, refresh_seconds, status_panel,
)
from admin.views.templates import (
    analysis_report, arg_row, clone_page, configure_page, delete_page,
    edit_form, edit_page, master_page, preview_values_page,
    builtin_style_names, new_page, new_template_button, pptx_analysis_report,
    spec_yaml_block, style_mapping_block, template_table, templates_panel,
)

__all__ = [
    "BRAND", "SectionFacts", "SlotFact", "ToolFact", "analysis_report",
    "arg_row", "back_link", "base_panel", "builtin_style_names", "clone_page",
    "configure_page", "dashboard_page", "delete_asset_page", "delete_page",
    "edit_form", "edit_page", "fmt_size", "fmt_ts", "fmt_uptime",
    "form_actions", "kind_page", "log_panel", "login_page", "logout_page",
    "master_page", "name_field", "new_page", "new_template_button",
    "not_found_page", "overview_panel", "page", "pptx_analysis_report",
    "preview_values_page", "refresh_seconds", "replace_card",
    "save_failed_page", "saved_page", "section_page", "source_files_panel",
    "spec_yaml_block", "status_panel", "style_mapping_block",
    "style_mapping_panel", "template_table", "templates_panel",
]
