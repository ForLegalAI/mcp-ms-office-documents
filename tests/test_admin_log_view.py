"""The Status page's log view: filtering, search and auto-refresh.

Before #174 it offered 150 records, a two-state All / Warnings toggle and a
manual Refresh link — usable only if what you wanted was near the top and you
could spot it by eye. These pin the filters, and that a filter which could
never match is not offered at all.
"""
import logging
import re
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from starlette.testclient import TestClient

import admin.store as store_mod
import metrics
import template_utils as tu
from config import Config


@pytest.fixture
def admin_client(tmp_path, monkeypatch):
    custom = tmp_path / "custom_templates"
    cfg = tmp_path / "config"
    custom.mkdir()
    cfg.mkdir()
    monkeypatch.setattr(store_mod, "_APP_CUSTOM_DIR", tmp_path / "nx1")
    monkeypatch.setattr(store_mod, "_APP_CONFIG_DIR", tmp_path / "nx2")
    monkeypatch.setattr(store_mod, "_LOCAL_CUSTOM_DIR", custom)
    monkeypatch.setattr(store_mod, "_LOCAL_CONFIG_DIR", cfg)
    monkeypatch.setattr(tu, "APP_CUSTOM_DIR", tmp_path / "nx1")
    monkeypatch.setattr(tu, "LOCAL_CUSTOM_DIR", custom)
    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.setenv("ADMIN_PASSWORD", "pw")
    monkeypatch.delenv("API_KEY", raising=False)

    from fastmcp import FastMCP
    from admin.app import build_combined_app

    client = TestClient(build_combined_app(FastMCP("t"), Config.from_env()))
    metrics.reset()
    client.__enter__()
    client.post("/admin/login", data={"password": "pw"})
    _seed_logs()
    yield client
    client.__exit__(None, None, None)
    metrics.reset()


def _seed_logs():
    """Records spanning several levels and packages."""
    logging.getLogger("admin.store").info("Wrote docx asset letter.docx")
    logging.getLogger("admin.app").warning("Rejected POST with invalid CSRF token")
    logging.getLogger("docx_tools.style_map").error("Unknown style_mapping key")


def _log_rows(html):
    """The rendered log lines, or the empty-state sentence."""
    block = re.search(r'<div class="logs">(.*?)</div>\s*</div>', html, re.S)
    if not block:
        empty = re.search(r"No records match these filters|"
                          r"No log records captured yet", html)
        return [empty.group(0)] if empty else []
    return [" ".join(re.sub(r"<[^>]+>", " ", row).split())
            for row in re.findall(r"<tr>(.*?)</tr>", block.group(1), re.S)]


def test_level_filter_is_exact_not_a_two_state_toggle(admin_client):
    """`error` used to mean "warnings and errors"; it now means errors."""
    rows = _log_rows(admin_client.get("/admin/server/log?level=error").text)
    assert any("Unknown style_mapping key" in r for r in rows)
    assert not any("CSRF" in r for r in rows), (
        "a WARNING must not appear under an error-only filter"
    )


def test_warning_level_includes_errors(admin_client):
    rows = _log_rows(admin_client.get("/admin/server/log?level=warning").text)
    assert any("CSRF" in r for r in rows)
    assert any("Unknown style_mapping key" in r for r in rows)


def test_source_filter_keeps_one_package(admin_client):
    rows = _log_rows(admin_client.get("/admin/server/log?logger=admin").text)
    assert any("Wrote docx asset" in r for r in rows)
    assert not any("style_map" in r for r in rows)


def test_search_matches_the_message(admin_client):
    rows = _log_rows(admin_client.get("/admin/server/log?q=letter.docx").text)
    assert len(rows) == 1 and "Wrote docx asset" in rows[0]


def test_search_also_matches_the_logger_name(admin_client):
    """"Which module logged this" is as useful a question as the text."""
    rows = _log_rows(admin_client.get("/admin/server/log?q=style_map").text)
    assert rows and all("style_map" in r for r in rows)


@pytest.mark.parametrize("needle", ["CSRF", "csrf", "CsRf"])
def test_search_is_case_insensitive(admin_client, needle):
    """Asserted per-query rather than by comparing two responses.

    The test client logs each request it makes, so a URL containing the
    needle lands in the very buffer the next query searches — comparing two
    responses compares polluted sets.
    """
    rows = _log_rows(admin_client.get(f"/admin/server/log?q={needle}").text)
    assert any("Rejected POST with invalid CSRF token" in r for r in rows)


def test_filters_combine(admin_client):
    rows = _log_rows(
        admin_client.get("/admin/server/log?level=warning&logger=admin&q=csrf").text)
    assert len(rows) == 1 and "CSRF" in rows[0]


def test_no_match_is_distinguished_from_nothing_captured(admin_client):
    html = admin_client.get("/admin/server/log?q=nothing-matches-this").text
    assert "No records match these filters" in html
    assert "No log records captured yet" not in html


def test_filter_state_survives_in_the_form(admin_client):
    """A filtered view has to be a URL you can bookmark or paste."""
    html = admin_client.get("/admin/server/log?level=warning&logger=admin&q=token").text
    assert 'value="token"' in html
    assert 'value="warning" selected' in html
    assert 'value="admin" selected' in html


def test_sources_are_offered_as_packages(admin_client):
    html = admin_client.get("/admin/server/log").text
    options = re.findall(r'<option value="([^"]+)"[^>]*>\1</option>', html)
    assert "admin" in options and "docx_tools" in options
    assert "admin.store" not in options, "grouped by package, not every module"


def test_a_level_that_could_never_match_is_not_offered(admin_client):
    """The buffer captures at the server's level; below it there is nothing.

    Offering `debug` on an INFO server is a filter that always comes back
    empty and reads as broken, so it is withheld and explained instead.
    """
    html = admin_client.get("/admin/server/log").text
    assert metrics.capture_level() == logging.INFO
    assert 'value="debug"' not in html
    assert "Set DEBUG=true" in html
    for level in ("info", "warning", "error"):
        assert f'value="{level}"' in html


def test_auto_refresh_is_off_by_default(admin_client):
    # On the tab that offers it: auto-refresh belongs to the log view, so
    # asserting its absence anywhere else would pass for the wrong reason.
    assert "location.reload" not in admin_client.get("/admin/server/log").text


def test_auto_refresh_emits_an_inline_timer(admin_client):
    html = admin_client.get("/admin/server/log?refresh=30").text
    assert "location.reload" in html
    assert "30000" in html, "seconds must reach the timer as milliseconds"
    assert 'value="30" selected' in html


@pytest.mark.parametrize("raw", ["", "nonsense", "-5", "abc", "1e4",
                                 "999999", "45"])
def test_an_unusable_refresh_value_turns_it_off(admin_client, raw):
    """A bad query parameter must not break the page or start a timer.

    A number outside `REFRESH_CHOICES` counts as bad: accepting it would arm
    a timer no option in the <select> shows as chosen, so the control would
    read "off" while the page reloaded under you.
    """
    r = admin_client.get(f"/admin/server/log?refresh={raw}")
    assert r.status_code == 200
    assert "location.reload" not in r.text
    assert 'value="0" selected' in r.text, "the control must show it is off"


def test_the_limit_is_reported_when_it_bites(admin_client):
    from admin.views.status import LOG_LIMIT

    for i in range(LOG_LIMIT + 20):
        logging.getLogger("admin.flood").info("record %d", i)
    html = admin_client.get("/admin/server/log?logger=admin").text
    assert f"newest {LOG_LIMIT}" in html, (
        "a truncated view must say so, or it reads as the whole story"
    )


def test_the_limit_applies_after_filtering(admin_client):
    """The limit must count matches, not records scanned.

    The needle is logged *first* and then buried behind far more than
    `LOG_LIMIT` records. Filtering the whole buffer still finds it; taking the
    newest `LOG_LIMIT` records and filtering those loses it entirely — and a
    needle logged last would be found either way, so the ordering is the
    whole point of this test.
    """
    from admin.views.status import LOG_LIMIT

    logging.getLogger("admin.needle").info("the one record that matters")
    for i in range(LOG_LIMIT * 2):
        logging.getLogger("admin.flood").info("noise %d", i)

    rows = _log_rows(admin_client.get("/admin/server/log?q=the one record").text)
    assert any("the one record that matters" in r for r in rows), (
        "a match older than the newest LOG_LIMIT records must still be found"
    )
