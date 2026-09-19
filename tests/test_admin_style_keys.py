"""The style-mapping editor must offer exactly what the renderer recognises.

`docx_tools/style_map.py` decides which `style_mapping` keys do anything; the
admin UI decides which ones an admin can set. When those two lists disagree the
UI silently hides part of the feature — which is what it did for a long time,
offering 5 of 16 (#160). This pins them together so the next key added to the
renderer cannot go unreachable in the UI.
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from admin.kinds import STYLE_GROUPS, STYLE_KEYS
from admin.views.templates import builtin_style_names
from docx_tools import style_map


def _recognised_keys():
    """Every key style_map._normalize() acts on, read from its own tables."""
    return (set(style_map._HEADING_KEYS)
            | set(style_map._LIST_NUMBER_KEYS)
            | set(style_map._LIST_BULLET_KEYS)
            | set(style_map._SCALAR_KEYS))


def test_ui_offers_every_recognised_key():
    missing = _recognised_keys() - set(STYLE_KEYS)
    assert not missing, f"the UI cannot set these recognised keys: {sorted(missing)}"


def test_ui_offers_nothing_the_renderer_ignores():
    extra = set(STYLE_KEYS) - _recognised_keys()
    assert not extra, (
        f"the UI offers keys style_map ignores (a silent no-op): {sorted(extra)}")


def test_groups_cover_every_key_exactly_once():
    grouped = [key for _title, keys in STYLE_GROUPS for key in keys]
    assert sorted(grouped) == sorted(set(grouped)), "a key appears in two groups"
    assert set(grouped) == set(STYLE_KEYS)


def test_builtin_names_match_the_renderer_defaults():
    """The labels say what a key resolves to; they must not be a second copy."""
    names = builtin_style_names()
    default = style_map.DEFAULT_STYLE_MAP
    assert names["heading_1"] == default.heading[0]
    assert names["heading_6"] == default.heading[5]
    assert names["list_number"] == default.list_number[0]
    assert names["list_number_3"] == default.list_number[2]
    assert names["list_bullet_2"] == default.list_bullet[1]
    assert names["quote"] == default.quote
    assert names["table"] == default.table
    assert names["normal"] == default.normal
    # `code` has no default paragraph style; the label must cope with None.
    assert names["code"] == default.code
    # Every offered key has an entry, even when that entry is None.
    assert set(names) == set(STYLE_KEYS)


def test_a_key_the_renderer_gains_is_caught():
    """Guard the guard: the comparison is live, not a frozen copy."""
    recognised = _recognised_keys()
    assert "heading_4" in recognised and "code" in recognised
    assert "heading_7" not in recognised
