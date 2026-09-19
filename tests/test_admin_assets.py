"""Every admin page must be self-contained — no request leaves the origin.

The admin UI inlines its theme and its handful of JavaScript rather than
pulling anything from a CDN, so it renders correctly offline, in an air-gapped
deployment, and behind a restrictive CSP. That is a property of the *rendered
output*, not of any one module, so it is tested by walking every page type and
looking at the HTML.

This replaces an earlier spot check that asserted ``"cdn" not in html`` on the
index page alone — which passed happily for ``unpkg.com``, a Google Fonts
``@import``, or anything at all on the other five page types (#145).
"""
import io
import re
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from bs4 import BeautifulSoup
from docx import Document
from starlette.testclient import TestClient

import admin.store as store_mod
import template_utils as tu
import metrics
from config import Config

# Anything with a scheme, or a protocol-relative URL, leaves this origin.
_ABSOLUTE = re.compile(r"^\s*(?:[a-z][a-z0-9+.-]*:)?//", re.I)
# url(...) and @import inside an inline stylesheet.
_CSS_URL = re.compile(r"""url\(\s*['"]?([^'")]+)""", re.I)
_CSS_IMPORT = re.compile(r"""@import\s+(?:url\()?\s*['"]([^'"]+)""", re.I)


def _is_external(value: str) -> bool:
    return bool(value) and bool(_ABSOLUTE.match(value))


@pytest.fixture
def admin_client(tmp_path, monkeypatch):
    """A logged-in TestClient over the combined app, isolated to tmp dirs."""
    custom = tmp_path / "custom"
    cfg = tmp_path / "config"
    custom.mkdir()
    cfg.mkdir()

    monkeypatch.setattr(store_mod, "_APP_CUSTOM_DIR", tmp_path / "noexist_app_custom")
    monkeypatch.setattr(store_mod, "_APP_CONFIG_DIR", tmp_path / "noexist_app_config")
    monkeypatch.setattr(store_mod, "_LOCAL_CUSTOM_DIR", custom)
    monkeypatch.setattr(store_mod, "_LOCAL_CONFIG_DIR", cfg)
    monkeypatch.setattr(tu, "APP_CUSTOM_DIR", tmp_path / "noexist_app_custom")
    monkeypatch.setattr(tu, "LOCAL_CUSTOM_DIR", custom)

    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.setenv("ADMIN_PASSWORD", "pw")
    monkeypatch.delenv("API_KEY", raising=False)

    from fastmcp import FastMCP
    from admin.app import build_combined_app

    mcp = FastMCP("test-admin-assets")
    client = TestClient(build_combined_app(mcp, Config.from_env()))
    metrics.reset()
    client.__enter__()
    yield client
    client.__exit__(None, None, None)
    metrics.reset()


def _csrf(client) -> str:
    html = client.get("/admin/new/docx").text
    tag = re.search(r'<input[^>]*name="csrf"[^>]*>', html)
    assert tag, "no CSRF input found on page"
    return re.search(r'value="([^"]*)"', tag.group(0)).group(1)


def _docx(*paragraphs) -> bytes:
    doc = Document()
    for p in paragraphs:
        doc.add_paragraph(p)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _all_pages(client):
    """Every distinct admin page, as ``(label, html)`` pairs.

    Includes the signed-out pages and the two pages that only exist after an
    upload, since those carry the most markup and are exactly where a future
    convenience import would go unnoticed.
    """
    pages = [("login", client.get("/admin/login").text)]
    client.post("/admin/login", data={"password": "pw"})
    for label, url in [
        ("index", "/admin/"),
        ("status", "/admin/status"),
        ("status-errors", "/admin/status?level=error"),
        ("logout", "/admin/logout"),
        ("new-docx", "/admin/new/docx"),
        ("new-email", "/admin/new/email"),
        ("new-pptx", "/admin/new/pptx"),
    ]:
        pages.append((label, client.get(url).text))

    # The draft (configure) page, and the edit page for the saved template.
    token = _csrf(client)
    draft = client.post(
        "/admin/docx/draft",
        data={"name": "assets_tpl", "csrf": token},
        files={"file": ("assets_tpl.docx",
                        _docx("Dear {{recipient}},", "{{#if ps}}", "{{note}}", "{{/if}}"),
                        "application/octet-stream")},
    )
    pages.append(("draft-docx", draft.text))
    client.post("/admin/docx/save", data={
        "csrf": token, "kind": "docx", "name": "assets_tpl",
        "asset_filename": "assets_tpl.docx", "description": "d",
        "arg_name": "recipient", "arg_type": "string",
        "arg_required": "true", "arg_default": "", "arg_desc": "who",
    })
    pages.append(("edit-docx", client.get("/admin/docx/assets_tpl/edit").text))
    return pages


def test_every_page_is_rendered(admin_client):
    """Guard the guard: the sweep below is worthless if a page came back empty."""
    pages = _all_pages(admin_client)
    assert len(pages) == 10
    for label, html in pages:
        assert "<html" in html, f"{label} did not render a page"


def test_no_external_stylesheets_or_scripts(admin_client):
    """No <link> or <script src> may point off-origin.

    ``rel="canonical"`` is exempt: FastHTML emits it pointing at the page's own
    absolute URL, which is not a fetched asset.
    """
    for label, html in _all_pages(admin_client):
        soup = BeautifulSoup(html, "html.parser")
        for link in soup.find_all("link"):
            rels = [r.lower() for r in (link.get("rel") or [])]
            if "canonical" in rels:
                continue
            assert not _is_external(link.get("href", "")), (
                f"{label}: external <link> {link.get('href')!r}")
        for script in soup.find_all("script"):
            src = script.get("src", "")
            assert not src, f"{label}: <script src={src!r}> — inline the script instead"


def test_no_external_urls_in_inline_css(admin_client):
    """An inlined stylesheet must not @import or url() anything off-origin."""
    for label, html in _all_pages(admin_client):
        soup = BeautifulSoup(html, "html.parser")
        for style in soup.find_all("style"):
            css = style.string or ""
            for target in _CSS_IMPORT.findall(css) + _CSS_URL.findall(css):
                assert not _is_external(target), (
                    f"{label}: inline CSS reaches {target!r}")


def test_no_element_loads_an_external_resource(admin_client):
    """Nothing — img, iframe, source, use — may fetch from another origin."""
    for label, html in _all_pages(admin_client):
        soup = BeautifulSoup(html, "html.parser")
        for element in soup.find_all(True):
            for attr in ("src", "srcset", "poster", "data", "xlink:href"):
                value = element.get(attr)
                if isinstance(value, list):
                    value = " ".join(value)
                assert not _is_external(value or ""), (
                    f"{label}: <{element.name} {attr}={value!r}> is off-origin")


def test_theme_and_scripts_are_actually_inline(admin_client):
    """The flip side: the page must carry its own theme and behaviour.

    Without this, deleting the stylesheet entirely would pass every test above.
    """
    for label, html in _all_pages(admin_client):
        soup = BeautifulSoup(html, "html.parser")
        css = "\n".join(s.string or "" for s in soup.find_all("style"))
        assert "--brand" in css, f"{label}: inline theme missing"
        assert "prefers-color-scheme" in css, f"{label}: no dark-mode block"
        scripts = "\n".join(s.string or "" for s in soup.find_all("script"))
        assert "adminAddArgRow" in scripts, f"{label}: inline row script missing"


def test_page_declares_its_language(admin_client):
    """Assistive tech needs a language to pronounce the page in (#153)."""
    for label, html in _all_pages(admin_client):
        soup = BeautifulSoup(html, "html.parser")
        assert soup.html.get("lang"), f"{label}: <html> has no lang attribute"
