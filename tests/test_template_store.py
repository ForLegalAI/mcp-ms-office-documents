"""Tests for the admin template storage layer (admin/store.py)."""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from admin.store import (
    FileTemplateStore,
    TemplateStoreError,
    KIND_DOCX,
    KIND_EMAIL,
    validate_name,
    validate_asset_filename,
)


@pytest.fixture
def store(tmp_path):
    return FileTemplateStore(custom_dir=tmp_path / "custom", config_dir=tmp_path / "config")


def _docx_spec(name="demo_letter"):
    return {
        "name": name,
        "description": "A demo letter",
        "args": [{"name": "body", "type": "string", "required": True, "description": "Body"}],
    }


def test_save_and_get_roundtrip(store):
    saved = store.save_spec(KIND_DOCX, _docx_spec(), asset_bytes=b"PKfake", asset_filename="demo.docx")
    assert saved["docx_path"] == "demo.docx"

    got = store.get_spec(KIND_DOCX, "demo_letter")
    assert got is not None
    assert got["name"] == "demo_letter"
    assert got["docx_path"] == "demo.docx"
    assert got["args"][0]["name"] == "body"


def test_asset_bytes_persisted(store):
    store.save_spec(KIND_DOCX, _docx_spec(), asset_bytes=b"hello-bytes", asset_filename="demo.docx")
    assert store.read_asset(KIND_DOCX, "demo.docx") == b"hello-bytes"
    assert store.asset_exists(KIND_DOCX, "demo.docx")


def test_derives_asset_filename_from_name(store):
    saved = store.save_spec(KIND_DOCX, _docx_spec("auto_named"), asset_bytes=b"x")
    assert saved["docx_path"] == "auto_named.docx"


def test_list_specs_sorted(store):
    store.save_spec(KIND_DOCX, _docx_spec("zeta"), asset_bytes=b"x")
    store.save_spec(KIND_DOCX, _docx_spec("alpha"), asset_bytes=b"x")
    names = [s["name"] for s in store.list_specs(KIND_DOCX)]
    assert names == ["alpha", "zeta"]


def test_delete_spec_and_asset(store):
    store.save_spec(KIND_DOCX, _docx_spec(), asset_bytes=b"x", asset_filename="demo.docx")
    assert store.delete_spec(KIND_DOCX, "demo_letter", delete_asset=True) is True
    assert store.get_spec(KIND_DOCX, "demo_letter") is None
    assert store.asset_exists(KIND_DOCX, "demo.docx") is False
    # Deleting again is a no-op (False).
    assert store.delete_spec(KIND_DOCX, "demo_letter") is False


def test_delete_keeps_asset_by_default(store):
    store.save_spec(KIND_DOCX, _docx_spec(), asset_bytes=b"x", asset_filename="demo.docx")
    store.delete_spec(KIND_DOCX, "demo_letter")
    assert store.asset_exists(KIND_DOCX, "demo.docx") is True


def test_email_kind_uses_html(store):
    spec = {"name": "welcome", "description": "hi", "args": []}
    saved = store.save_spec(KIND_EMAIL, spec, asset_bytes=b"<html></html>", asset_filename="welcome.html")
    assert saved["html_path"] == "welcome.html"
    assert store.read_asset(KIND_EMAIL, "welcome.html") == b"<html></html>"


def test_save_without_bytes_requires_existing_asset(store):
    with pytest.raises(TemplateStoreError):
        store.save_spec(KIND_DOCX, _docx_spec(), asset_filename="missing.docx")


@pytest.mark.parametrize("bad", ["1bad", "has space", "", "bad-dash", "weird!"])
def test_invalid_names_rejected(bad):
    with pytest.raises(TemplateStoreError):
        validate_name(bad)


@pytest.mark.parametrize("good", ["formal_letter", "a", "_x", "Letter1"])
def test_valid_names_accepted(good):
    assert validate_name(good) == good


@pytest.mark.parametrize("bad", ["../evil.docx", "/abs/path.docx", "sub/dir.docx", "noext", "wrong.txt"])
def test_invalid_asset_filenames_rejected(bad):
    with pytest.raises(TemplateStoreError):
        validate_asset_filename(bad, KIND_DOCX)


def test_email_asset_extension_enforced():
    with pytest.raises(TemplateStoreError):
        validate_asset_filename("welcome.docx", KIND_EMAIL)
    assert validate_asset_filename("welcome.html", KIND_EMAIL) == "welcome.html"


def test_unknown_kind_raises(store):
    with pytest.raises(TemplateStoreError):
        store.list_specs("pptx")
