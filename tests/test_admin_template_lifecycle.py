"""Turning a template off, and renaming one (#165).

Before this, deleting was the only off switch: to stop the AI reaching for a
seasonal or under-review template you had to destroy its configuration, and a
typo in a tool name was permanent short of deleting and re-entering everything.

The properties that matter: a disabled template keeps its spec and its asset
but is not a live tool, it stays that way across a restart, and a plain save
does not quietly switch it back on. A rename moves the tool, not the asset.
"""
import io
import logging
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
from config import Config
from template_registry import gather_specs, is_enabled


@pytest.fixture
def admin_client(tmp_path, monkeypatch):
    """A logged-in client over the combined app, isolated to temp dirs."""
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
    # The pptx registry resolves its config dir independently of the store.
    import pptx_tools.templates as templates_mod
    monkeypatch.setattr(templates_mod, "_APP_CONFIG_DIR", tmp_path / "nx2")
    monkeypatch.setattr(templates_mod, "_LOCAL_CONFIG_DIR", cfg)
    templates_mod.clear_cache()
    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.setenv("ADMIN_PASSWORD", "pw")
    monkeypatch.delenv("API_KEY", raising=False)

    from fastmcp import FastMCP
    from admin.app import build_combined_app

    mcp = FastMCP("test-lifecycle")
    client = TestClient(build_combined_app(mcp, Config.from_env()))
    metrics.reset()
    client.__enter__()
    client.post("/admin/login", data={"password": "pw"})
    yield client, mcp, cfg
    client.__exit__(None, None, None)
    metrics.reset()


def _docx_bytes(text="Hi {{who}}"):
    doc = Document()
    doc.add_paragraph(text)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _csrf(client) -> str:
    html = client.get("/admin/new/docx").text
    tag = re.search(r'<input[^>]*name="csrf"[^>]*>', html)
    assert tag, "no CSRF input found"
    return re.search(r'value="([^"]*)"', tag.group(0)).group(1)


def _post(client, url, data=None, files=None, **kwargs):
    payload = dict(data or {})
    payload["csrf"] = _csrf(client)
    return client.post(url, data=payload, files=files, **kwargs)


def _saved(client, name="tpl"):
    """A saved, live docx template with one placeholder."""
    _post(client, "/admin/docx/draft", data={"name": name},
          files={"file": (f"{name}.docx", _docx_bytes(), "application/octet-stream")})
    _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{name}.docx", "name": name,
        "original_name": name, "title": "T", "description": "d",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })
    return name


async def _tool_names(mcp):
    return [t.name for t in await mcp.list_tools()]


def _store():
    return store_mod.FileTemplateStore.from_config()


# ---------------------------------------------------------------------------
# is_enabled: the key every registration path reads
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("spec, expected", [
    ({}, True),                       # every spec written before #165
    ({"enabled": True}, True),
    ({"enabled": False}, False),
    ({"enabled": "false"}, False),    # hand-edited YAML, quoted
    ({"enabled": "No"}, False),
    ({"enabled": "off"}, False),
    ({"enabled": ""}, False),
    # PyYAML coerces the full words to bools but leaves these as strings, and
    # they are what someone reaching for "off" actually types — "disable"
    # mirrors the UI's own button label.
    ({"enabled": "n"}, False),
    ({"enabled": "N"}, False),
    ({"enabled": "disable"}, False),
    ({"enabled": "disabled"}, False),
    ({"enabled": "true"}, True),
    ({"enabled": "yes"}, True),
    ({"enabled": "y"}, True),
    ({"enabled": "enable"}, True),
    (None, True),
])
def test_is_enabled(spec, expected):
    assert is_enabled(spec) is expected


def test_a_spec_with_no_enabled_key_stays_live(tmp_path):
    """The upgrade path: nothing written before #165 may go dark on deploy."""
    spec_dir = tmp_path / "docx_templates.d"
    spec_dir.mkdir()
    (spec_dir / "old.yaml").write_text(
        yaml.safe_dump({"name": "old", "docx_path": "old.docx"}), encoding="utf-8")

    specs, _ = gather_specs(None, spec_dir)
    assert [s["name"] for s in specs] == ["old"]


def test_gather_specs_drops_disabled_but_can_include_them(tmp_path):
    spec_dir = tmp_path / "docx_templates.d"
    spec_dir.mkdir()
    for name, enabled in (("on_tpl", True), ("off_tpl", False)):
        (spec_dir / f"{name}.yaml").write_text(
            yaml.safe_dump({"name": name, "docx_path": f"{name}.docx",
                            "enabled": enabled}), encoding="utf-8")

    live, _ = gather_specs(None, spec_dir)
    assert [s["name"] for s in live] == ["on_tpl"]

    everything, _ = gather_specs(None, spec_dir, include_disabled=True)
    assert sorted(s["name"] for s in everything) == ["off_tpl", "on_tpl"]


def test_a_disabled_master_yaml_entry_is_dropped_too(tmp_path):
    """Disabling is a spec property, not a UI one — the master YAML has it."""
    master = tmp_path / "docx_templates.yaml"
    master.write_text(yaml.safe_dump({"templates": [
        {"name": "keep", "docx_path": "a.docx"},
        {"name": "drop", "docx_path": "b.docx", "enabled": False},
    ]}), encoding="utf-8")

    specs, _ = gather_specs(master, None)
    assert [s["name"] for s in specs] == ["keep"]


# ---------------------------------------------------------------------------
# Disable / enable through the UI
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_disabling_takes_the_tool_off_the_server(admin_client):
    client, mcp, _cfg = admin_client
    name = _saved(client)
    assert name in await _tool_names(mcp)

    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": ""})
    assert name not in await _tool_names(mcp), "disabling must unregister the tool"


@pytest.mark.anyio
async def test_disabling_keeps_the_spec_and_the_asset(admin_client):
    """The whole point: an off switch that is not destruction."""
    client, _mcp, _cfg = admin_client
    name = _saved(client)
    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": ""})

    store = _store()
    spec = store.get_spec("docx", name)
    assert spec is not None, "the spec file must survive"
    assert spec["enabled"] is False
    assert spec["args"], "the arguments must survive"
    assert store.asset_exists("docx", f"{name}.docx"), "the source file must survive"


@pytest.mark.anyio
async def test_enabling_puts_it_back(admin_client):
    client, mcp, _cfg = admin_client
    name = _saved(client)
    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": ""})
    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": "true"})

    assert name in await _tool_names(mcp)
    assert _store().get_spec("docx", name)["enabled"] is True


def test_disabled_survives_a_restart(admin_client):
    """The state lives in the spec file, so the startup loader must honour it.

    Registering from YAML is what a restart does, so this runs that loader
    against a fresh server rather than trusting the live unregister.
    """
    import asyncio

    from fastmcp import FastMCP
    from docx_tools.dynamic_docx_tools import register_docx_template_tools_from_yaml

    client, _mcp, cfg = admin_client
    name = _saved(client)

    def restart():
        fresh = FastMCP("restarted")
        register_docx_template_tools_from_yaml(fresh, cfg / "docx_templates.yaml")
        return [t.name for t in asyncio.run(fresh.list_tools())]

    # Positive control first: without it, a loader that registers nothing at
    # all — a wrong config path, say — would make the real assertion vacuous.
    assert name in restart(), "the loader must find the template to begin with"

    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": ""})
    assert name not in restart(), "a disabled template must not come back on restart"


@pytest.mark.anyio
async def test_saving_a_disabled_template_does_not_switch_it_back_on(admin_client):
    """The edit form has no enabled control, so a save rebuilds the spec
    without one. The stored flag has to be carried forward."""
    client, mcp, _cfg = admin_client
    name = _saved(client)
    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": ""})

    _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{name}.docx", "name": name,
        "original_name": name, "title": "T", "description": "edited",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    spec = _store().get_spec("docx", name)
    assert spec["description"] == "edited", "the edit must have landed"
    assert spec["enabled"] is False, "the save must not re-enable it"
    assert name not in await _tool_names(mcp)


def test_the_row_shows_disabled_and_offers_enable(admin_client):
    client, _mcp, _cfg = admin_client
    name = _saved(client)

    html = client.get("/admin/").text
    assert "Disable" in html, "a live template offers Disable"

    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": ""})
    html = client.get("/admin/").text
    assert "Disabled" in html, "the status column must say so"
    assert f"/admin/docx/{name}/enabled" in html
    assert ">Enable<" in html, "a disabled template offers Enable"


def test_enabling_is_a_post_not_a_link(admin_client):
    """A GET would let a prefetch or a crawler take a template off the air."""
    client, _mcp, _cfg = admin_client
    name = _saved(client)
    r = client.get(f"/admin/docx/{name}/enabled", follow_redirects=False)
    assert r.status_code in (404, 405), f"GET must not toggle; got {r.status_code}"


def test_toggling_an_unknown_template_is_not_found(admin_client):
    client, _mcp, _cfg = admin_client
    r = _post(client, "/admin/docx/no_such_tpl/enabled", data={"enabled": ""})
    assert "not found" in r.text.lower()


# ---------------------------------------------------------------------------
# Rename
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_renaming_moves_the_tool(admin_client):
    client, mcp, _cfg = admin_client
    old = _saved(client, "before_tpl")
    assert old in await _tool_names(mcp)

    _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{old}.docx", "name": "after_tpl",
        "original_name": old, "title": "T", "description": "d",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    names = await _tool_names(mcp)
    assert "after_tpl" in names
    assert old not in names, "the old tool must be unregistered"

    store = _store()
    assert store.get_spec("docx", old) is None, "the old spec file must be gone"
    assert store.get_spec("docx", "after_tpl")["name"] == "after_tpl"


def test_renaming_leaves_the_asset_filename_alone(admin_client):
    """Renaming the file would break any master-YAML entry pointing at it."""
    client, _mcp, _cfg = admin_client
    old = _saved(client, "keepfile_tpl")

    _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{old}.docx", "name": "renamed_tpl",
        "original_name": old, "title": "T", "description": "d",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    store = _store()
    assert store.asset_exists("docx", f"{old}.docx"), "the file keeps its name"
    assert store.get_spec("docx", "renamed_tpl")["docx_path"] == f"{old}.docx"


def test_renaming_onto_an_existing_name_is_refused(admin_client):
    client, _mcp, _cfg = admin_client
    first = _saved(client, "occupied_tpl")
    second = _saved(client, "mover_tpl")

    r = _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{second}.docx", "name": first,
        "original_name": second, "title": "T", "description": "clobber",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    assert "already exists" in r.text
    store = _store()
    assert store.get_spec("docx", second) is not None, "the mover must survive"
    assert store.get_spec("docx", first)["description"] == "d", \
        "the occupant must not be overwritten"


def test_renaming_to_an_invalid_name_is_refused(admin_client):
    client, _mcp, _cfg = admin_client
    old = _saved(client, "valid_tpl")

    r = _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{old}.docx", "name": "not a name!",
        "original_name": old, "title": "T", "description": "d",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    assert r.status_code == 200
    assert _store().get_spec("docx", old) is not None, "the original must survive"


def test_renaming_keeps_a_disabled_template_disabled(admin_client):
    client, _mcp, _cfg = admin_client
    old = _saved(client, "off_then_renamed")
    _post(client, f"/admin/docx/{old}/enabled", data={"enabled": ""})

    _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{old}.docx", "name": "still_off_tpl",
        "original_name": old, "title": "T", "description": "d",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    spec = _store().get_spec("docx", "still_off_tpl")
    assert spec is not None
    assert spec["enabled"] is False


def test_the_edit_form_lets_you_change_the_name(admin_client):
    """It used to render a disabled input, which made a typo permanent."""
    client, _mcp, _cfg = admin_client
    name = _saved(client, "editable_tpl")

    html = client.get(f"/admin/docx/{name}/edit").text
    name_input = re.search(r'<input[^>]*name="name"[^>]*>', html)
    assert name_input, "the name must be a submitted input, not display-only"
    assert "disabled" not in name_input.group(0)
    original = re.search(r'<input[^>]*name="original_name"[^>]*>', html)
    assert original and f'value="{name}"' in original.group(0), \
        "the save route needs the current name to tell a rename from an edit"


# ---------------------------------------------------------------------------
# Store-level guarantees
# ---------------------------------------------------------------------------


def test_rename_writes_the_new_spec_before_dropping_the_old(admin_client, monkeypatch):
    """An interrupted rename must leave two templates, never none."""
    client, _mcp, _cfg = admin_client
    _saved(client, "interrupted_tpl")
    store = _store()

    boom = Path.unlink

    def explode(self, *a, **kw):
        if self.name == "interrupted_tpl.yaml":
            raise OSError("interrupted")
        return boom(self, *a, **kw)

    monkeypatch.setattr(Path, "unlink", explode)
    with pytest.raises(OSError):
        store.rename_spec("docx", "interrupted_tpl", "survivor_tpl")
    monkeypatch.undo()

    assert store.get_spec("docx", "survivor_tpl") is not None, \
        "the new spec must already be on disk when the old one is removed"


def test_rename_to_the_same_name_is_a_no_op(admin_client):
    client, _mcp, _cfg = admin_client
    _saved(client, "same_tpl")
    store = _store()

    store.rename_spec("docx", "same_tpl", "same_tpl")
    assert store.get_spec("docx", "same_tpl") is not None, \
        "renaming to itself must not delete the template"


def test_set_enabled_on_a_missing_template_raises(admin_client):
    client, _mcp, _cfg = admin_client
    with pytest.raises(store_mod.TemplateStoreError):
        _store().set_enabled("docx", "ghost_tpl", False)


# ---------------------------------------------------------------------------
# PowerPoint: a template is a value for the `template` argument, not a tool
# ---------------------------------------------------------------------------


@pytest.fixture
def pptx_registry(tmp_path, monkeypatch):
    """A pptx registry rooted at a temp config dir, with one real template."""
    import pptx_tools.templates as templates_mod

    cfg = tmp_path / "config"
    custom = tmp_path / "custom"
    cfg.mkdir()
    custom.mkdir()
    spec_dir = cfg / templates_mod.SPEC_SUBDIR
    spec_dir.mkdir()

    monkeypatch.setattr(templates_mod, "_APP_CONFIG_DIR", tmp_path / "nx")
    monkeypatch.setattr(templates_mod, "_LOCAL_CONFIG_DIR", cfg)
    monkeypatch.setattr(tu, "APP_CUSTOM_DIR", tmp_path / "nx")
    monkeypatch.setattr(tu, "LOCAL_CUSTOM_DIR", custom)

    source = Path(project_root) / "default_templates" / "default_pptx_template_16_9.pptx"
    (custom / "brand.pptx").write_bytes(source.read_bytes())

    def write(name, **extra):
        (spec_dir / f"{name}.yaml").write_text(
            yaml.safe_dump({"name": name, "pptx_path": "brand.pptx", **extra}),
            encoding="utf-8")
        templates_mod.clear_cache()

    templates_mod.clear_cache()
    yield templates_mod, write
    templates_mod.clear_cache()


def test_a_disabled_pptx_template_is_not_offered(pptx_registry):
    """There is no tool to unregister, so the registry has to drop it."""
    templates_mod, write = pptx_registry
    write("brand")
    assert "brand" in templates_mod.template_names()

    write("brand", enabled=False)
    assert "brand" not in templates_mod.template_names()


def test_disabling_every_pptx_template_does_not_resurrect_the_built_ins(pptx_registry):
    """The legacy filename slots are the "nothing configured" safety net.

    Falling back to them here would hand back the very templates the admin
    just took away — and under the built-in names, so the deck would still
    build, silently, off the wrong design.
    """
    templates_mod, write = pptx_registry
    write("brand", enabled=False)

    names = templates_mod.template_names()
    assert "brand" not in names
    assert names == [], f"expected no templates, got {names}"


def test_the_built_ins_still_appear_when_nothing_is_configured(pptx_registry):
    """The guard above must not cost the real fallback its job."""
    templates_mod, _write = pptx_registry
    assert templates_mod.template_names(), \
        "with no specs at all, the historical filename slots must still load"


def test_renaming_a_pptx_template_moves_it_in_the_registry(admin_client):
    """PowerPoint takes the same route but a different path through it.

    A pptx template is not a tool, so there is nothing to unregister — the
    registry has to stop offering the old name and start offering the new one.
    """
    import pptx_tools.templates as templates_mod

    client, _mcp, cfg = admin_client
    source = Path(project_root) / "default_templates" / "default_pptx_template_16_9.pptx"
    _post(client, "/admin/pptx/draft", data={"name": "old_deck"},
          files={"file": ("old_deck.pptx", source.read_bytes(),
                          "application/octet-stream")})
    _post(client, "/admin/pptx/save", data={
        "kind": "pptx", "asset_filename": "old_deck.pptx", "name": "old_deck",
        "original_name": "old_deck", "description": "d",
    })
    templates_mod.clear_cache()
    assert "old_deck" in templates_mod.template_names()

    _post(client, "/admin/pptx/save", data={
        "kind": "pptx", "asset_filename": "old_deck.pptx", "name": "new_deck",
        "original_name": "old_deck", "description": "d",
    })
    templates_mod.clear_cache()

    names = templates_mod.template_names()
    assert "new_deck" in names
    assert "old_deck" not in names, "the old name must stop being offered"


def test_disabling_a_pptx_template_through_the_ui(admin_client):
    import pptx_tools.templates as templates_mod

    client, _mcp, _cfg = admin_client
    source = Path(project_root) / "default_templates" / "default_pptx_template_16_9.pptx"
    _post(client, "/admin/pptx/draft", data={"name": "off_deck"},
          files={"file": ("off_deck.pptx", source.read_bytes(),
                          "application/octet-stream")})
    _post(client, "/admin/pptx/save", data={
        "kind": "pptx", "asset_filename": "off_deck.pptx", "name": "off_deck",
        "original_name": "off_deck", "description": "d",
    })
    templates_mod.clear_cache()
    assert "off_deck" in templates_mod.template_names()

    _post(client, "/admin/pptx/off_deck/enabled", data={"enabled": ""})
    templates_mod.clear_cache()
    assert "off_deck" not in templates_mod.template_names()


def test_an_unrecognised_enabled_value_is_flagged(caplog):
    """Reading someone's intent backwards in silence is the failure mode.

    Enabled is the safe direction — a template the AI cannot call looks like
    a broken server — but it must not be a silent guess.
    """
    with caplog.at_level(logging.WARNING, logger="template_registry"):
        assert is_enabled({"name": "typo_tpl", "enabled": "maybe"}) is True
    assert "typo_tpl" in caplog.text
    assert "maybe" in caplog.text


def test_a_recognised_value_is_not_flagged(caplog):
    """The warning must not cry wolf on a deliberate, spelled-out choice."""
    with caplog.at_level(logging.WARNING, logger="template_registry"):
        is_enabled({"name": "fine_tpl", "enabled": "yes"})
        is_enabled({"name": "fine_tpl", "enabled": "disable"})
        is_enabled({"name": "fine_tpl", "enabled": True})
        is_enabled({"name": "fine_tpl"})
    assert caplog.text == ""


# ---------------------------------------------------------------------------
# Creating is not editing (#182 review)
# ---------------------------------------------------------------------------


def test_creating_over_an_existing_name_is_refused(admin_client):
    """The create form has no `original_name`, so a same-name save used to
    look like an ordinary edit and overwrite the occupant in silence."""
    client, _mcp, _cfg = admin_client
    name = _saved(client, "occupied")

    r = _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{name}.docx", "name": name,
        "title": "T", "description": "clobber",   # no original_name: a create
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    assert "already exists" in r.text
    assert _store().get_spec("docx", name)["description"] == "d", \
        "the existing template must be untouched"


@pytest.mark.anyio
async def test_creating_over_a_disabled_name_does_not_inherit_its_state(admin_client):
    """A fresh upload must not arrive disabled because a stranger with the
    same name was turned off."""
    client, mcp, _cfg = admin_client
    name = _saved(client, "recycled")
    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": ""})

    r = _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{name}.docx", "name": name,
        "title": "T", "description": "fresh",     # no original_name: a create
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    # Refused outright, so there is no half-created template to inherit
    # anything — and the disabled one is still disabled and still off.
    assert "already exists" in r.text
    assert _store().get_spec("docx", name)["enabled"] is False
    assert name not in await _tool_names(mcp)


def test_the_create_form_carries_no_original_name(admin_client):
    """What makes a create distinguishable from an edit at the route."""
    client, _mcp, _cfg = admin_client
    r = _post(client, "/admin/docx/draft", data={"name": "brand_new"},
              files={"file": ("brand_new.docx", _docx_bytes(),
                              "application/octet-stream")})
    assert 'name="original_name"' not in r.text, \
        "the configure form is a create; an original_name makes it look like an edit"

    edit = client.get(f"/admin/docx/{_saved(client, 'existing')}/edit").text
    assert 'name="original_name"' in edit, "the edit form still needs it"


def test_a_failed_rename_reports_instead_of_500ing(admin_client, monkeypatch):
    """rename_spec raises OSError on a failed unlink; the route must catch it."""
    client, _mcp, _cfg = admin_client
    _saved(client, "unlink_fails")

    real = Path.unlink

    def explode(self, *a, **kw):
        if self.name == "unlink_fails.yaml":
            raise OSError("permission denied")
        return real(self, *a, **kw)

    monkeypatch.setattr(Path, "unlink", explode)
    r = _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": "unlink_fails.docx",
        "name": "renamed_ok", "original_name": "unlink_fails",
        "title": "T", "description": "d",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })
    assert r.status_code == 200
    assert "permission denied" in r.text


def test_saving_a_disabled_template_is_not_reported_as_a_failure(admin_client):
    """Staying off is the outcome asked for, not a registration that failed.

    The save page has three states, not two: live, deliberately off, and
    genuinely failed to register. Collapsing the middle one into the last
    told the admin to check the logs about a tool that was never meant to be
    registered.
    """
    client, _mcp, _cfg = admin_client
    name = _saved(client, "quietly_off")
    _post(client, f"/admin/docx/{name}/enabled", data={"enabled": ""})

    r = _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{name}.docx", "name": name,
        "original_name": name, "title": "T", "description": "edited",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    assert "could not be registered" not in r.text
    assert "check the logs" not in r.text
    assert "Saved with a warning" not in r.text
    assert "stays disabled" in r.text


def test_saving_an_enabled_template_still_says_it_is_live(admin_client):
    """The fix must not make every save read as if it were disabled."""
    client, _mcp, _cfg = admin_client
    name = _saved(client, "still_live")

    r = _post(client, "/admin/docx/save", data={
        "kind": "docx", "asset_filename": f"{name}.docx", "name": name,
        "original_name": name, "title": "T", "description": "edited",
        "arg_name": ["who"], "arg_type": ["string"], "arg_required": ["true"],
        "arg_default": [""], "arg_desc": [""],
    })

    assert "is now live" in r.text
    assert "stays disabled" not in r.text


def test_sync_reports_success_when_it_takes_a_template_off(admin_client):
    """`sync()` answers "does the server match the spec now?", not "did I
    remove a live tool?".

    Returning `unregister()`'s own bool meant False whenever there was
    nothing registered to remove — which is every save of a disabled
    template — and callers read that as a failure.
    """
    from admin.app import AdminContext
    from config import Config as Cfg
    from fastmcp import FastMCP

    ctx = AdminContext(FastMCP("sync-probe"), Cfg.from_env())
    spec = {"name": "never_registered", "description": "d",
            "docx_path": "x.docx", "args": [], "enabled": False}

    assert ctx.sync("docx", spec) is True, \
        "taking a template off is the outcome asked for, so it succeeded"
