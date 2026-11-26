"""Re-registering a template never takes a working tool off the server (#192).

``register_docx_template()`` / ``register_email_template()`` used to remove the
live tool *first* and then rebuild it, so any build that failed — an argument
pydantic refuses, a ``docx_path`` that no longer resolves — left nothing under
that name. Re-registration runs on admin actions about something else (saving
the global style mapping rebuilds every Word tool), so a working tool vanished
as a side effect, recorded only in a log line the admin never reads.

Now the replacement is built first and swapped in only once it exists. A failed
build keeps the previous tool — unless that tool's own source file is gone, in
which case it could not run anyway and is removed as before.
"""
import io
import json
import re
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from docx import Document
from fastmcp import FastMCP
from starlette.testclient import TestClient

import admin.store as store_mod
import docx_tools.dynamic_docx_tools as dd
import email_tools.dynamic_email_tools as de
import metrics
import template_utils as tu
from config import Config

DOCX_ASSET = "default_docx_template.docx"     # ships in default_templates/
EMAIL_ASSET = "broadcast_email_style_1.html"  # ships in default_templates/

# An argument pydantic's create_model refuses (leading underscore): a real
# build failure, reached the same way a bad admin edit reaches it.
BAD_ARG = {"name": "_bad", "type": "string", "required": True}


def _docx_spec(name, description="v1", **extra):
    spec = {"name": name, "description": description, "docx_path": DOCX_ASSET,
            "args": [{"name": "body", "type": "string", "required": True,
                      "description": "b"}]}
    spec.update(extra)
    return spec


def _email_spec(name, description="v1", **extra):
    spec = {"name": name, "description": description, "html_path": EMAIL_ASSET,
            "args": [{"name": "headline", "type": "string", "required": True,
                      "description": "h"}]}
    spec.update(extra)
    return spec


async def _tool(mcp, name):
    tools = [t for t in await mcp.list_tools() if t.name == name]
    assert len(tools) <= 1, f"{name!r} registered twice"
    return tools[0] if tools else None


@pytest.fixture(autouse=True)
def _clean_registries():
    yield
    for name in dd.registered_docx_template_names():
        dd._REGISTERED_DOCX.pop(name, None)
        getattr(dd, "_LIVE_SOURCE", {}).pop(name, None)
    for name in de.registered_email_template_names():
        de._REGISTERED_EMAIL.pop(name, None)


# ---- Word ---------------------------------------------------------------

@pytest.mark.asyncio
async def test_docx_a_build_that_raises_keeps_the_previous_tool():
    mcp = FastMCP("t")
    assert dd.register_docx_template(mcp, _docx_spec("keep_doc")) is True

    broken = _docx_spec("keep_doc", "v2", args=[BAD_ARG])
    assert dd.register_docx_template(mcp, broken) is False, \
        "a failed build reports failure rather than raising"

    tool = await _tool(mcp, "keep_doc")
    assert tool is not None, "the working tool must survive a failed rebuild"
    assert tool.description == "v1"
    assert "keep_doc" in dd.registered_docx_template_names()


@pytest.mark.asyncio
async def test_docx_an_edit_to_a_missing_file_keeps_the_previous_tool():
    """The edit names a file that does not exist; the live tool's own file is
    still there, so the live tool still works and stays."""
    mcp = FastMCP("t")
    assert dd.register_docx_template(mcp, _docx_spec("keep_doc2")) is True

    edit = _docx_spec("keep_doc2", "v2", docx_path="nonexistent_192.docx")
    assert dd.register_docx_template(mcp, edit) is False

    tool = await _tool(mcp, "keep_doc2")
    assert tool is not None and tool.description == "v1"


@pytest.mark.asyncio
async def test_docx_a_tool_whose_own_file_is_gone_is_removed(tmp_path, monkeypatch):
    """The one case a failed rebuild still removes: the live tool reads its
    file on every call, so with the file gone it could only fail."""
    src = tmp_path / "gone.docx"
    Document().save(str(src))
    monkeypatch.setattr(dd, "find_docx_template_by_name",
                        lambda f: str(src) if src.exists() else None)

    mcp = FastMCP("t")
    spec = _docx_spec("gone_doc", docx_path="gone.docx")
    assert dd.register_docx_template(mcp, spec) is True

    src.unlink()   # deleted on the volume since startup
    assert dd.register_docx_template(mcp, spec) is False

    assert await _tool(mcp, "gone_doc") is None
    assert "gone_doc" not in dd.registered_docx_template_names()


@pytest.mark.asyncio
async def test_docx_a_successful_edit_replaces_in_place():
    mcp = FastMCP("t")
    assert dd.register_docx_template(mcp, _docx_spec("swap_doc")) is True
    v2 = _docx_spec("swap_doc", "v2",
                    args=[{"name": "greeting", "type": "string", "required": True}])
    assert dd.register_docx_template(mcp, v2) is True

    tool = await _tool(mcp, "swap_doc")
    assert tool.description == "v2"
    params = json.dumps(tool.parameters)
    assert '"greeting"' in params and '"body"' not in params, \
        "the live schema is the new version's"


@pytest.mark.asyncio
async def test_docx_a_kept_tool_still_generates(monkeypatch):
    """Kept means working, not merely listed."""
    monkeypatch.setattr(dd, "upload_file", lambda *a, **k: "https://example/doc.docx")
    mcp = FastMCP("t")
    assert dd.register_docx_template(mcp, _docx_spec("run_doc")) is True
    assert dd.register_docx_template(mcp, _docx_spec("run_doc", "v2", args=[BAD_ARG])) is False

    result = await mcp.call_tool("run_doc", {"data": {"body": "hello"}})
    assert "https://example/doc.docx" in str(result)


# ---- Email --------------------------------------------------------------

@pytest.mark.asyncio
async def test_email_a_build_that_raises_keeps_the_previous_tool():
    mcp = FastMCP("t")
    assert de.register_email_template(mcp, _email_spec("keep_mail")) is True

    broken = _email_spec("keep_mail", "v2", args=[BAD_ARG])
    assert de.register_email_template(mcp, broken) is False

    tool = await _tool(mcp, "keep_mail")
    assert tool is not None and tool.description == "v1"
    assert "keep_mail" in de.registered_email_template_names()


@pytest.mark.asyncio
async def test_email_an_unreadable_file_keeps_the_previous_tool(tmp_path, monkeypatch):
    """An email tool holds its HTML from build time, so it keeps working even
    when the file can no longer be read — it is never removed for that."""
    src = tmp_path / "mail.html"
    src.write_text("<p>{{headline}}</p>", encoding="utf-8")
    monkeypatch.setattr(de, "find_email_template", lambda f: str(src))

    mcp = FastMCP("t")
    spec = _email_spec("bytes_mail", html_path="mail.html")
    assert de.register_email_template(mcp, spec) is True

    src.write_bytes(b"\xff\xfe not utf-8 \x80")
    assert de.register_email_template(mcp, spec) is False

    assert await _tool(mcp, "bytes_mail") is not None


# ---- Admin --------------------------------------------------------------

@pytest.fixture
def admin_client(tmp_path, monkeypatch):
    custom = tmp_path / "custom"
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

    from admin.app import build_combined_app

    mcp = FastMCP("test-192")
    client = TestClient(build_combined_app(mcp, Config.from_env()))
    metrics.reset()
    client.__enter__()
    client.post("/admin/login", data={"password": "pw"})
    yield client, mcp
    client.__exit__(None, None, None)
    metrics.reset()


def _csrf(client) -> str:
    html = client.get("/admin/new/docx").text
    tag = re.search(r'<input[^>]*name="csrf"[^>]*>', html)
    return re.search(r'value="([^"]*)"', tag.group(0)).group(1)


def _post(client, url, data=None, files=None):
    data = dict(data or {})
    data["csrf"] = _csrf(client)
    return client.post(url, data=data, files=files)


def _docx_bytes():
    doc = Document()
    doc.add_paragraph("Dear {{recipient}}")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _save_letter(client, name, arg="recipient", original=None):
    original = original if original is not None else name
    return _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{original}.docx", "name": name,
        "original_name": original,
        "title": "Letter", "description": "d",
        "arg_name": [arg], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": ["who"],
    })


def _create_letter(client, name):
    _post(client, "/admin/docx/draft", data={"name": name},
          files={"file": (f"{name}.docx", _docx_bytes(),
                          "application/octet-stream")})
    r = _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{name}.docx", "name": name,
        "title": "Letter", "description": "d",
        "arg_name": ["recipient"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": ["who"],
    })
    assert "now live" in r.text


@pytest.mark.asyncio
async def test_admin_a_broken_edit_keeps_the_tool_and_says_so(admin_client):
    client, mcp = admin_client
    _create_letter(client, "letter_kept")

    r = _save_letter(client, "letter_kept", arg="_bad")

    assert r.status_code == 200, "a failed build is a warning, not a 500"
    assert await _tool(mcp, "letter_kept") is not None
    assert "previous version" in r.text and "restart" in r.text
    assert 'class="flash flash-warn"' in r.text


@pytest.mark.asyncio
async def test_admin_a_rename_whose_build_fails_is_not_called_kept(admin_client):
    """A rename takes the old name off first; with the new build failing,
    nothing is live under either name, and the page must not claim a previous
    version is still serving."""
    client, mcp = admin_client
    _create_letter(client, "letter_old")

    r = _save_letter(client, "letter_new", arg="_bad", original="letter_old")

    assert r.status_code == 200
    assert await _tool(mcp, "letter_new") is None
    assert await _tool(mcp, "letter_old") is None
    assert "could not be registered" in r.text
    assert "previous version" not in r.text
