"""The Source files page: what is in custom_templates/, and what points at it (#166).

The directory used to be write-only from the admin's point of view — uploads
went in and were never shown again — so orphans left by a delete, a rename or
a hand-copied file accumulated where nothing reported them.

The property that matters most is the *unreferenced* direction. This page
offers to delete an orphan, so a file wrongly called an orphan is a file
wrongly offered for deletion. Three things count as a reference and each is
easy to forget: a managed spec, a hand-written master-YAML entry, and a base
template — which lives in the same directory under a fixed name and is named
by no spec at all.
"""
import io
import re
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
import yaml
from docx import Document
from starlette.testclient import TestClient

import admin.store as store_mod
import metrics
import template_utils as tu
from admin.assets import AssetFile, scan
from config import Config


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
    import pptx_tools.templates as templates_mod
    monkeypatch.setattr(templates_mod, "_APP_CONFIG_DIR", tmp_path / "nx2")
    monkeypatch.setattr(templates_mod, "_LOCAL_CONFIG_DIR", cfg)
    templates_mod.clear_cache()
    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.setenv("ADMIN_PASSWORD", "pw")
    monkeypatch.delenv("API_KEY", raising=False)

    from fastmcp import FastMCP
    from admin.app import build_combined_app

    client = TestClient(build_combined_app(FastMCP("test-files"), Config.from_env()))
    metrics.reset()
    client.__enter__()
    client.post("/admin/login", data={"password": "pw"})
    yield client, custom, cfg
    client.__exit__(None, None, None)
    metrics.reset()


def _docx_bytes():
    doc = Document()
    doc.add_paragraph("Hi {{who}}")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _csrf(client) -> str:
    html = client.get("/admin/new/docx").text
    tag = re.search(r'<input[^>]*name="csrf"[^>]*>', html)
    return re.search(r'value="([^"]*)"', tag.group(0)).group(1)


def _post(client, url, data=None, files=None):
    payload = dict(data or {})
    payload["csrf"] = _csrf(client)
    return client.post(url, data=payload, files=files)


def _saved(client, name="letter"):
    _post(client, "/admin/docx/draft", data={"name": name},
          files={"file": (f"{name}.docx", _docx_bytes(), "application/octet-stream")})
    _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{name}.docx", "name": name,
        "original_name": name, "title": "T", "description": "d",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })
    return name


def _row(html, filename):
    """The table row for *filename*, as flattened text."""
    for m in re.finditer(r"<tr>(.*?)</tr>", html, re.S):
        if filename in m.group(1):
            return " ".join(re.sub(r"<[^>]+>", " ", m.group(1)).split())
    return ""


# ---------------------------------------------------------------------------
# What counts as referenced — the direction that makes this safe
# ---------------------------------------------------------------------------


def test_a_managed_templates_file_is_referenced(admin_client):
    client, _custom, _cfg = admin_client
    _saved(client, "letter")

    row = _row(client.get("/admin/files").text, "letter.docx")
    assert "Word template 'letter'" in row
    assert "Unreferenced" not in row


def test_a_master_yaml_templates_file_is_not_an_orphan(admin_client):
    """The failure that would make this feature actively dangerous.

    A hand-written template is invisible to the store — it lives only in the
    master YAML — so a scan that reads managed specs alone calls its file an
    orphan and offers to delete it.
    """
    client, custom, cfg = admin_client
    (custom / "handwritten.docx").write_bytes(_docx_bytes())
    (cfg / "docx_templates.yaml").write_text(yaml.safe_dump({"templates": [
        {"name": "legacy", "docx_path": "handwritten.docx", "args": []},
    ]}), encoding="utf-8")

    row = _row(client.get("/admin/files").text, "handwritten.docx")
    assert "master YAML" in row
    assert "Unreferenced" not in row, "a hand-written template's file is in use"


def test_a_disabled_templates_file_is_not_an_orphan(admin_client):
    """Disabling takes the tool off the server; it does not free the file."""
    client, _custom, _cfg = admin_client
    name = _saved(client, "seasonal")
    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": ""})

    row = _row(client.get("/admin/files").text, f"{name}.docx")
    assert "Unreferenced" not in row, "a disabled template still owns its file"


def test_a_base_template_is_not_an_orphan(admin_client):
    """No spec names it, but it styles every document the server makes."""
    client, custom, _cfg = admin_client
    (custom / "custom_docx_template.docx").write_bytes(_docx_bytes())

    row = _row(client.get("/admin/files").text, "custom_docx_template.docx")
    assert "Base template" in row
    assert "Unreferenced" not in row


def test_a_file_two_templates_share_lists_both(admin_client):
    client, _custom, _cfg = admin_client
    _saved(client, "first")
    store = store_mod.FileTemplateStore.from_config()
    store.save_spec("docx", {"name": "second", "description": "d",
                             "docx_path": "first.docx", "args": []})

    row = _row(client.get("/admin/files").text, "first.docx")
    assert "'first'" in row and "'second'" in row


def test_a_leftover_file_is_an_orphan(admin_client):
    client, custom, _cfg = admin_client
    (custom / "leftover.docx").write_bytes(_docx_bytes())

    row = _row(client.get("/admin/files").text, "leftover.docx")
    assert "Unreferenced" in row


def test_deleting_a_template_and_keeping_its_file_surfaces_it(admin_client):
    """The story the page exists for, end to end."""
    client, _custom, _cfg = admin_client
    name = _saved(client, "retired")
    assert "Unreferenced" not in _row(client.get("/admin/files").text, f"{name}.docx")

    _post(client, f"/admin/docx/{name}/delete", data={})  # keep the file

    row = _row(client.get("/admin/files").text, f"{name}.docx")
    assert "Unreferenced" in row, "the kept file must now be findable"


def test_orphans_are_listed_first(admin_client):
    client, custom, _cfg = admin_client
    _saved(client, "aaa_referenced")
    (custom / "zzz_orphan.docx").write_bytes(_docx_bytes())

    html = client.get("/admin/files").text
    assert html.index("zzz_orphan.docx") < html.index("aaa_referenced.docx"), \
        "orphans lead: they are the only rows with anything to do"


# ---------------------------------------------------------------------------
# Deleting — only an orphan, only after confirming
# ---------------------------------------------------------------------------


def test_only_an_orphan_offers_delete(admin_client):
    client, custom, _cfg = admin_client
    _saved(client, "kept")
    (custom / "gone.docx").write_bytes(_docx_bytes())

    html = client.get("/admin/files").text
    assert "/files/gone.docx/delete" in html
    assert "/files/kept.docx/delete" not in html


def test_deleting_an_orphan_removes_it(admin_client):
    client, custom, _cfg = admin_client
    (custom / "gone.docx").write_bytes(_docx_bytes())

    confirm = client.get("/admin/files/gone.docx/delete")
    assert "gone.docx" in confirm.text, "the confirmation must name the file"

    r = _post(client, "/admin/files/gone.docx/delete", data={})
    assert not (custom / "gone.docx").exists()
    assert "Deleted gone.docx" in r.text


def test_a_referenced_file_cannot_be_deleted_by_posting_at_it(admin_client):
    """The page hides the button; the POST is not where that is trusted from."""
    client, custom, _cfg = admin_client
    _saved(client, "inuse")

    r = _post(client, "/admin/files/inuse.docx/delete", data={})
    assert (custom / "inuse.docx").exists(), "a referenced file must survive"
    assert "not found" in r.text.lower()


def test_a_base_template_cannot_be_deleted_by_posting_at_it(admin_client):
    client, custom, _cfg = admin_client
    (custom / "custom_docx_template.docx").write_bytes(_docx_bytes())

    _post(client, "/admin/files/custom_docx_template.docx/delete", data={})
    assert (custom / "custom_docx_template.docx").exists(), \
        "deleting this would silently restyle every document the server makes"


@pytest.mark.parametrize("attempt", [
    "../config/docx_templates.yaml",
    "..%2Fconfig%2Fdocx_templates.yaml",
    "....//secret.docx",
])
def test_a_crafted_filename_matches_nothing(admin_client, attempt):
    """Deletion is keyed off a name the scan produced, so a path never matches."""
    client, _custom, cfg = admin_client
    (cfg / "docx_templates.yaml").write_text("templates: []", encoding="utf-8")

    _post(client, f"/admin/files/{attempt}/delete", data={})
    assert (cfg / "docx_templates.yaml").exists()


def test_a_file_that_gained_a_reference_is_no_longer_deletable(admin_client):
    """The orphan check runs at the moment of the request, not at render."""
    client, custom, _cfg = admin_client
    (custom / "claimed.docx").write_bytes(_docx_bytes())
    assert "Unreferenced" in _row(client.get("/admin/files").text, "claimed.docx")

    # A template starts pointing at it after the page was rendered.
    store = store_mod.FileTemplateStore.from_config()
    store.save_spec("docx", {"name": "claimer", "description": "d",
                             "docx_path": "claimed.docx", "args": []})

    _post(client, "/admin/files/claimed.docx/delete", data={})
    assert (custom / "claimed.docx").exists(), \
        "the file is in use now, whatever the page said a moment ago"


def test_delete_needs_the_csrf_token(admin_client):
    client, custom, _cfg = admin_client
    (custom / "gone.docx").write_bytes(_docx_bytes())

    client.post("/admin/files/gone.docx/delete", data={})
    assert (custom / "gone.docx").exists(), "a POST without the token must not delete"


def test_the_files_route_is_not_shadowed_by_the_template_delete(admin_client):
    """`/files/{name}/delete` and `/{kind}/{name}/delete` have the same shape.

    Registration order decides, so this pins it: a template named `files`
    must not make the page unreachable, and the page must not swallow it.
    """
    client, custom, _cfg = admin_client
    (custom / "gone.docx").write_bytes(_docx_bytes())

    r = client.get("/admin/files/gone.docx/delete")
    assert r.status_code == 200
    assert "Delete this file?" in r.text


# ---------------------------------------------------------------------------
# The scan itself
# ---------------------------------------------------------------------------


def test_scan_skips_directories(tmp_path):
    (tmp_path / "a.docx").write_bytes(b"x")
    (tmp_path / "subdir").mkdir()

    names = [f.name for f in scan(tmp_path, {})]
    assert names == ["a.docx"], "the store writes a flat directory"


def test_scan_of_a_missing_directory_is_empty(tmp_path):
    assert scan(tmp_path / "nope", {}) == []


def test_scan_reports_size_and_mtime(tmp_path):
    (tmp_path / "a.docx").write_bytes(b"12345")
    f = scan(tmp_path, {})[0]
    assert f.size == 5
    assert f.mtime > 0


def test_orphaned_is_just_no_references():
    assert AssetFile("a", 1, 1.0, ()).orphaned is True
    assert AssetFile("a", 1, 1.0, ("something",)).orphaned is False
