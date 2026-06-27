"""Integration tests for the FastHTML admin UI and combined ASGI app.

The template store and template-resolution directories are redirected to a
temp path so tests never touch the repo's custom_templates/config.
"""
import io
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from docx import Document
from starlette.testclient import TestClient

import admin.store as store_mod
import template_utils as tu
from config import Config


@pytest.fixture
def admin_client(tmp_path, monkeypatch):
    """A logged-in TestClient over the combined app, isolated to tmp dirs."""
    custom = tmp_path / "custom"
    cfg = tmp_path / "config"
    custom.mkdir()
    cfg.mkdir()

    # Redirect the storage layer to the temp directories.
    monkeypatch.setattr(store_mod, "_APP_CUSTOM_DIR", tmp_path / "noexist_app_custom")
    monkeypatch.setattr(store_mod, "_APP_CONFIG_DIR", tmp_path / "noexist_app_config")
    monkeypatch.setattr(store_mod, "_LOCAL_CUSTOM_DIR", custom)
    monkeypatch.setattr(store_mod, "_LOCAL_CONFIG_DIR", cfg)
    # Redirect template resolution so live registration finds the tmp asset.
    monkeypatch.setattr(tu, "APP_CUSTOM_DIR", tmp_path / "noexist_app_custom")
    monkeypatch.setattr(tu, "LOCAL_CUSTOM_DIR", custom)

    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.setenv("ADMIN_PASSWORD", "pw")
    monkeypatch.delenv("API_KEY", raising=False)

    # Import here so the patched module globals are in effect.
    from fastmcp import FastMCP
    from admin.app import build_combined_app

    mcp = FastMCP("test-admin")
    app = build_combined_app(mcp, Config.from_env())
    client = TestClient(app)
    client.__enter__()
    # Authenticate.
    client.post("/admin/login", data={"password": "pw"})
    yield client, mcp
    client.__exit__(None, None, None)


def _docx_with_placeholders(*paragraphs) -> bytes:
    doc = Document()
    for p in paragraphs:
        doc.add_paragraph(p)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


async def _tool_names(mcp):
    return [t.name for t in await mcp.list_tools()]


def test_login_required(tmp_path, monkeypatch):
    monkeypatch.setattr(store_mod, "_LOCAL_CUSTOM_DIR", tmp_path / "c")
    monkeypatch.setattr(store_mod, "_LOCAL_CONFIG_DIR", tmp_path / "cfg")
    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.setenv("ADMIN_PASSWORD", "pw")
    from fastmcp import FastMCP
    from admin.app import build_combined_app
    app = build_combined_app(FastMCP("t"), Config.from_env())
    with TestClient(app) as c:
        r = c.get("/admin/", follow_redirects=False)
        assert r.status_code == 303
        assert r.headers["location"].endswith("/admin/login")
        # Wrong password stays on the login page.
        r = c.post("/admin/login", data={"password": "wrong"}, follow_redirects=False)
        assert r.status_code == 200
        assert "Incorrect" in r.text


def test_index_lists_sections(admin_client):
    client, _ = admin_client
    r = client.get("/admin/")
    assert r.status_code == 200
    assert "Word templates" in r.text
    assert "Email templates" in r.text


def test_draft_detects_placeholders(admin_client):
    client, _ = admin_client
    data = _docx_with_placeholders("Dear {{recipient}},", "{{body}}")
    r = client.post(
        "/admin/docx/draft",
        data={"name": "letter_x"},
        files={"file": ("letter_x.docx", data, "application/octet-stream")},
    )
    assert r.status_code == 200
    assert "recipient" in r.text
    assert "body" in r.text


@pytest.mark.asyncio
async def test_create_makes_tool_live(admin_client):
    client, mcp = admin_client
    data = _docx_with_placeholders("Dear {{recipient}},", "{{body}}")
    client.post(
        "/admin/docx/draft",
        data={"name": "live_letter"},
        files={"file": ("live_letter.docx", data, "application/octet-stream")},
    )
    save = {
        "kind": "docx", "asset_filename": "live_letter.docx", "name": "live_letter",
        "title": "Letter", "description": "desc",
        "arg_name": ["recipient", "body"], "arg_type": ["string", "string"],
        "arg_required": ["true", "true"], "arg_default": ["", ""],
        "arg_desc": ["who", "the body"],
    }
    r = client.post("/admin/docx/save", data=save)
    assert r.status_code == 200
    assert "is live" in r.text
    assert "live_letter" in await _tool_names(mcp)
    # Persisted as a managed spec file.
    assert store_mod.FileTemplateStore.from_config().get_spec("docx", "live_letter")["name"] == "live_letter"


@pytest.mark.asyncio
async def test_edit_then_delete(admin_client):
    client, mcp = admin_client
    data = _docx_with_placeholders("{{body}}")
    client.post("/admin/docx/draft", data={"name": "tmp_tpl"},
                files={"file": ("tmp_tpl.docx", data, "application/octet-stream")})
    save = {
        "kind": "docx", "asset_filename": "tmp_tpl.docx", "name": "tmp_tpl",
        "title": "T", "description": "d",
        "arg_name": ["body"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    }
    client.post("/admin/docx/save", data=save)
    assert "tmp_tpl" in await _tool_names(mcp)

    # Edit page renders with the saved spec.
    r = client.get("/admin/docx/tmp_tpl/edit")
    assert r.status_code == 200 and "tmp_tpl" in r.text

    # Delete unregisters and removes the managed spec.
    r = client.post("/admin/docx/tmp_tpl/delete", follow_redirects=False)
    assert r.status_code == 303
    assert "tmp_tpl" not in await _tool_names(mcp)
    assert store_mod.FileTemplateStore.from_config().get_spec("docx", "tmp_tpl") is None


def test_preview_returns_docx(admin_client):
    client, _ = admin_client
    data = _docx_with_placeholders("Dear {{recipient}},", "{{body}}")
    client.post("/admin/docx/draft", data={"name": "prev_tpl"},
                files={"file": ("prev_tpl.docx", data, "application/octet-stream")})
    pv = {
        "kind": "docx", "asset_filename": "prev_tpl.docx", "name": "prev_tpl",
        "title": "P", "description": "d",
        "arg_name": ["recipient", "body"], "arg_type": ["string", "string"],
        "arg_required": ["true", "true"], "arg_default": ["", ""], "arg_desc": ["", ""],
    }
    r = client.post("/admin/docx/preview", data=pv)
    assert r.status_code == 200
    assert r.content[:2] == b"PK"  # a real .docx (zip)
    assert "attachment" in r.headers.get("content-disposition", "")


def test_invalid_name_rejected(admin_client):
    client, _ = admin_client
    data = _docx_with_placeholders("{{body}}")
    r = client.post("/admin/docx/draft", data={"name": "1 bad name"},
                    files={"file": ("x.docx", data, "application/octet-stream")})
    assert r.status_code == 200
    assert "Invalid template name" in r.text


def test_mcp_endpoint_still_works(admin_client):
    """The MCP endpoint is reachable through the combined app."""
    client, _ = admin_client
    headers = {"Accept": "application/json, text/event-stream", "Content-Type": "application/json"}
    r = client.post("/mcp", json={
        "jsonrpc": "2.0", "id": 1, "method": "initialize",
        "params": {"protocolVersion": "2025-06-18", "capabilities": {},
                   "clientInfo": {"name": "t", "version": "1"}},
    }, headers=headers)
    assert r.status_code == 200
    assert "mcp-session-id" in r.headers
