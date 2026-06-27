"""Tests for the shared template registry helpers and live (un)registration.

Covers:
- gather_specs() merging master YAML with a per-template *.d directory
- safe_remove_tool() tolerance
- runtime register / replace / unregister of docx and email tools
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from fastmcp import FastMCP

from template_registry import gather_specs, read_spec_dir, safe_remove_tool
from docx_tools.dynamic_docx_tools import (
    register_docx_template,
    unregister_docx_template,
    registered_docx_template_names,
)
from email_tools.dynamic_email_tools import (
    register_email_template,
    unregister_email_template,
    registered_email_template_names,
)

DOCX_ASSET = "default_docx_template.docx"  # ships in default_templates/
EMAIL_ASSET = "broadcast_email_style_1.html"  # ships in default_templates/


def _write_master(path: Path):
    path.write_text(
        "style_mapping:\n"
        "  heading_1: Brand Title\n"
        "templates:\n"
        "  - name: master_one\n"
        "    description: from master\n"
        f"    docx_path: {DOCX_ASSET}\n"
        "    args:\n"
        "      - {name: body, type: string, required: true, description: b}\n"
        "  - name: shared\n"
        "    description: master version\n"
        f"    docx_path: {DOCX_ASSET}\n"
        "    args: []\n",
        encoding="utf-8",
    )


def _write_override(spec_dir: Path):
    spec_dir.mkdir(parents=True, exist_ok=True)
    (spec_dir / "shared.yaml").write_text(
        "name: shared\ndescription: UI VERSION\n"
        f"docx_path: {DOCX_ASSET}\nargs: []\n",
        encoding="utf-8",
    )
    (spec_dir / "ui_only.yaml").write_text(
        "name: ui_only\ndescription: only in dir\n"
        f"docx_path: {DOCX_ASSET}\nargs: []\n",
        encoding="utf-8",
    )


def test_gather_specs_merges_and_overrides(tmp_path):
    master = tmp_path / "docx_templates.yaml"
    _write_master(master)
    spec_dir = tmp_path / "docx_templates.d"
    _write_override(spec_dir)

    templates, cfg = gather_specs(master, spec_dir)
    by_name = {t["name"]: t for t in templates}

    # Master-only template survives.
    assert by_name["master_one"]["description"] == "from master"
    # Per-template dir wins on a name clash.
    assert by_name["shared"]["description"] == "UI VERSION"
    # Dir-only template is appended.
    assert by_name["ui_only"]["description"] == "only in dir"
    # Top-level master keys still readable.
    assert cfg["style_mapping"]["heading_1"] == "Brand Title"


def test_gather_specs_missing_master(tmp_path):
    spec_dir = tmp_path / "docx_templates.d"
    _write_override(spec_dir)
    templates, cfg = gather_specs(tmp_path / "nope.yaml", spec_dir)
    names = sorted(t["name"] for t in templates)
    assert names == ["shared", "ui_only"]
    assert cfg == {}


def test_gather_specs_missing_dir(tmp_path):
    master = tmp_path / "docx_templates.yaml"
    _write_master(master)
    templates, _ = gather_specs(master, tmp_path / "absent.d")
    assert sorted(t["name"] for t in templates) == ["master_one", "shared"]


def test_read_spec_dir_skips_malformed(tmp_path):
    d = tmp_path / "d"
    d.mkdir()
    (d / "good.yaml").write_text("name: good\ndocx_path: x.docx\n", encoding="utf-8")
    (d / "noname.yaml").write_text("description: missing name\n", encoding="utf-8")
    (d / "broken.yaml").write_text("name: x\n  bad: : :\n", encoding="utf-8")
    specs = read_spec_dir(d)
    assert [s["name"] for s in specs] == ["good"]


def test_safe_remove_tool_absent_is_false():
    mcp = FastMCP("t")
    assert safe_remove_tool(mcp, "does_not_exist") is False


@pytest.mark.asyncio
async def test_docx_live_register_replace_unregister():
    mcp = FastMCP("t")
    spec = {
        "name": "live_doc",
        "description": "v1",
        "docx_path": DOCX_ASSET,
        "args": [{"name": "body", "type": "string", "required": True, "description": "b"}],
    }
    assert register_docx_template(mcp, spec) is True
    assert "live_doc" in registered_docx_template_names()
    tools = [t.name for t in await mcp.list_tools()]
    assert tools.count("live_doc") == 1

    # Re-register (edit) replaces in place — no duplicate.
    assert register_docx_template(mcp, {**spec, "description": "v2"}) is True
    tools = [t.name for t in await mcp.list_tools()]
    assert tools.count("live_doc") == 1

    assert unregister_docx_template(mcp, "live_doc") is True
    assert "live_doc" not in registered_docx_template_names()
    tools = [t.name for t in await mcp.list_tools()]
    assert "live_doc" not in tools


@pytest.mark.asyncio
async def test_docx_register_rejects_missing_asset():
    mcp = FastMCP("t")
    spec = {"name": "bad_doc", "docx_path": "nonexistent_file.docx", "args": []}
    assert register_docx_template(mcp, spec) is False
    assert "bad_doc" not in registered_docx_template_names()


@pytest.mark.asyncio
async def test_email_live_register_unregister():
    mcp = FastMCP("t")
    spec = {
        "name": "live_email",
        "description": "v1",
        "html_path": EMAIL_ASSET,
        "args": [{"name": "headline", "type": "string", "required": True, "description": "h"}],
    }
    assert register_email_template(mcp, spec) is True
    assert "live_email" in registered_email_template_names()
    tools = [t.name for t in await mcp.list_tools()]
    assert "live_email" in tools

    assert unregister_email_template(mcp, "live_email") is True
    tools = [t.name for t in await mcp.list_tools()]
    assert "live_email" not in tools
