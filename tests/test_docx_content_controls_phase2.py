"""Content controls, phase 2: combo box and date picker.

A combo box takes an item (by value or display text) or free text. A date
picker takes an ISO date, stores it in ``w:fullDate`` and shows it numerically
in the control's own format — a format with month/day names or a time falls
back to the numeric default for the control's language.
"""
import datetime as dt
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest  # noqa: E402
from docx import Document  # noqa: E402

from docx_tools.content_controls import (  # noqa: E402
    COMBO_BOX, DATE, describe_content_controls, resolve_content_controls,
)
from docx_tools.warnings import channel  # noqa: E402

from . import docx_content_control_helpers as cc  # noqa: E402

W = cc.W


def _doc(*controls):
    doc = Document()
    for control in controls:
        cc.add_inline(doc.add_paragraph(), control)
    return doc


def _pr(control):
    return control.find(f"{{{W}}}sdtPr")


# --- combo box -------------------------------------------------------------------

def test_combobox_is_detected_with_values_and_labels():
    [found] = describe_content_controls(_doc(cc.combobox("country")))
    assert (found["kind"], found["supported"]) == (COMBO_BOX, True)
    assert found["items"] == ["de", "at"] and found["labels"] == ["Germany", "Austria"]


@pytest.mark.parametrize("value", ["at", "AUSTRIA"])
def test_combobox_selects_an_item_by_value_or_label(value):
    box = cc.combobox("country")
    assert resolve_content_controls(_doc(box), {"country": value}) == 1
    assert cc.content_text(box) == "Austria"
    assert _pr(box).find(f"{{{W}}}comboBox").get(f"{{{W}}}lastValue") == "at"
    assert _pr(box).find(f"{{{W}}}showingPlcHdr") is None


def test_combobox_takes_free_text_without_a_warning():
    box, warnings = cc.combobox("country"), channel()
    assert resolve_content_controls(_doc(box), {"country": "Czech Republic"}, warnings) == 1
    assert cc.content_text(box) == "Czech Republic"
    assert _pr(box).find(f"{{{W}}}comboBox").get(f"{{{W}}}lastValue") == "Czech Republic"
    assert warnings.as_dicts() == []


def test_combobox_unset_keeps_its_prompt():
    box = cc.combobox("country")
    resolve_content_controls(_doc(box), {"country": None})
    assert cc.content_text(box) == "Choose or type."


# --- date picker -------------------------------------------------------------------

def _date_el(control):
    return _pr(control).find(f"{{{W}}}date")


@pytest.mark.parametrize("fmt,shown", [
    ("d. M. yyyy", "6. 10. 2026"), ("dd.MM.yyyy", "06.10.2026"),
    ("M/d/yy", "10/6/26"), ("yyyy-MM-dd", "2026-10-06"), ("'Date:' d.M.", "Date: 6.10."),
])
def test_a_date_is_shown_in_the_controls_numeric_format(fmt, shown):
    picker = cc.date_picker("signed", fmt=fmt)
    assert resolve_content_controls(_doc(picker), {"signed": "2026-10-06"}) == 1
    assert cc.content_text(picker) == shown
    assert _date_el(picker).get(f"{{{W}}}fullDate") == "2026-10-06T00:00:00Z"
    assert _pr(picker).find(f"{{{W}}}showingPlcHdr") is None


def test_a_date_object_from_the_tool_schema_is_accepted():
    picker = cc.date_picker("signed")
    resolve_content_controls(_doc(picker), {"signed": dt.date(2026, 1, 2)})
    assert cc.content_text(picker) == "2. 1. 2026"


@pytest.mark.parametrize("fmt,lid,shown", [
    ("d. MMMM yyyy", "cs-CZ", "6. 10. 2026"), ("dddd, MMMM d, yyyy", "en-US", "10/6/2026"),
    ("dd/MM/yyyy HH:mm", "en-GB", "06/10/2026"), ("d MMM yyyy", "de-DE", "06.10.2026"),
    ("MMMM yyyy", "fi-FI", "2026-10-06"),
])
def test_a_format_with_names_or_time_falls_back_to_the_numeric_default(fmt, lid, shown):
    picker, warnings = cc.date_picker("signed", fmt=fmt, lid=lid), channel()
    resolve_content_controls(_doc(picker), {"signed": "2026-10-06"}, warnings)
    assert cc.content_text(picker) == shown
    [w] = warnings.as_dicts()
    assert (w["code"], w["severity"], w["tag"]) == ("control_date_format_simplified", "info", "signed")


def test_a_control_without_a_format_uses_the_language_default_silently():
    picker, warnings = cc.date_picker("signed", fmt=None, lid="cs-CZ"), channel()
    resolve_content_controls(_doc(picker), {"signed": "2026-10-06"}, warnings)
    assert cc.content_text(picker) == "6. 10. 2026"
    assert warnings.as_dicts() == []


@pytest.mark.parametrize("value", ["6. 10. 2026", "tomorrow", "2026-13-01",
                                   "2026-10-06 garbage"])
def test_a_non_iso_date_is_refused_and_reported(value):
    picker, warnings = cc.date_picker("signed"), channel()
    assert resolve_content_controls(_doc(picker), {"signed": value}, warnings) == 0
    assert cc.content_text(picker) == "Enter a date."
    assert _date_el(picker).get(f"{{{W}}}fullDate") is None
    [w] = warnings.as_dicts()
    assert (w["code"], w["severity"]) == ("control_value_invalid", "error")


def test_date_kind_is_detected():
    [found] = describe_content_controls(_doc(cc.date_picker("signed")))
    assert (found["kind"], found["supported"]) == (DATE, True)


# --- through a registered tool --------------------------------------------------------

def test_a_date_argument_is_an_iso_date_in_the_schema_and_fills_end_to_end(tmp_path):
    import asyncio
    import io
    import json
    from unittest.mock import patch

    from fastmcp import Client, FastMCP

    import docx_tools.dynamic_docx_tools as dd

    doc = Document()
    cc.add_inline(doc.add_paragraph("Signed: "), cc.date_picker("signed"))
    doc.add_paragraph("ISO: {{signed}}")
    template = tmp_path / "d.docx"
    doc.save(template)
    out = {}

    def fake_upload(buf, kind, filename=None, add_unique_prefix=None, **kw):
        out["doc"] = Document(io.BytesIO(buf.getvalue()))
        return "ok"

    spec = {"name": "date_probe", "docx_path": "d.docx", "args": [
        {"name": "signed", "type": "date", "required": False, "description": "Signing date"}]}
    mcp = FastMCP("d")
    with patch.object(dd, "find_docx_template_by_name", lambda f: str(template)), \
            patch.object(dd, "upload_file", fake_upload):
        assert dd._register_single_template(mcp, spec)

        async def go():
            async with Client(mcp) as c:
                schema = (await c.list_tools())[0].inputSchema
                await c.call_tool("date_probe", {"data": {"signed": "2026-10-06"}})
                with pytest.raises(Exception):
                    await c.call_tool("date_probe", {"data": {"signed": "6. 10. 2026"}})
                return schema
        schema = asyncio.run(go())

    prop = schema["properties"]["data"]["properties"]["signed"]
    assert (prop["type"], prop["format"], prop["description"]) == ("string", "date", "Signing date")
    assert "anyOf" not in json.dumps(prop)
    picker = next(out["doc"].element.body.iter(f"{{{W}}}sdt"))
    assert cc.content_text(picker) == "6. 10. 2026"
    assert "ISO: 2026-10-06" in [p.text for p in out["doc"].paragraphs]


def test_an_iso_date_time_string_is_accepted_as_its_date():
    picker = cc.date_picker("signed")
    assert resolve_content_controls(_doc(picker), {"signed": "2026-10-06T09:30:00"}) == 1
    assert cc.content_text(picker) == "6. 10. 2026"


@pytest.mark.parametrize("fmt", ["d.M.y", "d.M.yyy"])
def test_a_year_code_word_does_not_define_falls_back(fmt):
    picker, warnings = cc.date_picker("signed", fmt=fmt), channel()
    resolve_content_controls(_doc(picker), {"signed": "2026-10-06"}, warnings)
    assert cc.content_text(picker) == "6. 10. 2026"
    assert [w["code"] for w in warnings.as_dicts()] == ["control_date_format_simplified"]
