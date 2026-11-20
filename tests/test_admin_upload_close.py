"""Every admin request closes the files its form spooled.

Starlette writes an upload to a ``SpooledTemporaryFile`` and closes it only in
``Request.close()``, which neither Starlette's endpoint wrapper nor FastHTML
calls. Each admin upload used to leave that file open until garbage
collection (a ``ResourceWarning`` per upload in the suite). The admin app's
``_CloseRequestForms`` middleware closes it, including for a request the login
gate turns away, whose form FastHTML has parsed all the same.
"""
import re
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from starlette.formparsers import MultiPartParser
from starlette.testclient import TestClient

import admin.store as store_mod
import template_utils as tu
import metrics
from config import Config


@pytest.fixture
def uploads(monkeypatch):
    """Every UploadFile the multipart parser hands out during the test."""
    seen = []
    original = MultiPartParser.parse

    async def recording_parse(self):
        form = await original(self)
        seen.extend(v for _, v in form.multi_items() if hasattr(v, "file"))
        return form

    monkeypatch.setattr(MultiPartParser, "parse", recording_parse)
    return seen


@pytest.fixture
def client(tmp_path, monkeypatch):
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

    c = TestClient(build_combined_app(FastMCP("t"), Config.from_env()))
    metrics.reset()
    with c:
        yield c
    metrics.reset()


def _login(client) -> str:
    client.post("/admin/login", data={"password": "pw"})
    html = client.get("/admin/word/base").text
    tag = re.search(r'<input[^>]*name="csrf"[^>]*>', html)
    return re.search(r'value="([^"]*)"', tag.group(0)).group(1)


def _file():
    # Over Starlette's 1 MB spool threshold, so the part lands on disk.
    return {"file": ("big.docx", b"x" * (2 * 1024 * 1024), "application/octet-stream")}


def test_an_upload_is_closed_after_the_request(client, uploads):
    csrf = _login(client)
    r = client.post("/admin/base/docx/upload", data={"csrf": csrf}, files=_file())
    assert r.status_code == 200
    assert uploads, "the request carried no parsed upload"
    assert all(u.file.closed for u in uploads)


def test_an_upload_the_login_gate_refuses_is_closed_too(client, uploads):
    r = client.post("/admin/base/docx/upload", files=_file(), follow_redirects=False)
    assert r.status_code == 303
    assert uploads, "FastHTML no longer parses the form before the gate"
    assert all(u.file.closed for u in uploads)
