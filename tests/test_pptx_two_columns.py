"""A two-column slide keeps both columns, whatever the template numbers them.

The bug (#194): ``_build_two_column_slide`` addressed placeholders by ``idx``,
assuming PowerPoint's built-in numbering — 1/2 for Two Content, 1-4 for
Comparison. A customer template numbered its two cards 4 and 2 (left and right
*in that order*) and its comparison layout 4, 13, 14, 15. The builder wrote the
right column into the left card, dropped the left column and both headings,
left three placeholders empty, and reported nothing.

These tests renumber the shipped template the same way and assert on the
placeholder each column actually landed in, plus the warning every degraded
path now owes the caller.
"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from pptx import Presentation as PptxReader
from pptx.oxml.ns import qn

from pptx_tools import warnings as W
from pptx_tools.placeholder_style import content_columns, content_placeholders
from pptx_tools.slide_builder import PowerpointPresentation
from pptx_tools.templates import TemplateSpec

BASE_16_9 = project_root / "default_templates" / "default_pptx_template_16_9.pptx"

TWO_CONTENT = "Dva obsahy"
COMPARISON = "Porovnání"
SINGLE_CONTENT = "Nadpis a obsah"
BLANK = "Prázdný"

SLIDE = {
    "type": "two_column",
    "title": "Spolupráce",
    "left": {"heading": "Vy dodáváte", "body": "- Odbornost\n- Strategii"},
    "right": {"heading": "AI dodává", "body": "- Rychlost\n- Drafting"},
}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _renumber(layout, mapping):
    """Rewrite the ``idx`` of *layout*'s placeholders, as a customer template may."""
    elements = []
    for placeholder in layout.placeholders:
        ph = (placeholder._element.find(qn("p:nvSpPr"))
              .find(qn("p:nvPr")).find(qn("p:ph")))
        idx = int(ph.get("idx") or 0)
        if idx in mapping:
            elements.append((ph, mapping[idx]))
    for ph, new_idx in elements:
        ph.set("idx", str(new_idx))


def _template(tmp_path, name, mapping) -> TemplateSpec:
    prs = PptxReader(str(BASE_16_9))
    layout = next(one for one in prs.slide_layouts if one.name == name)
    _renumber(layout, mapping)
    path = tmp_path / "renumbered.pptx"
    prs.save(str(path))
    return TemplateSpec(name="t", path=path, description="", is_default=True,
                        layouts={}, defaults={}, strip_slides=True, aspect="16:9")


def _plain(spec_path=BASE_16_9) -> TemplateSpec:
    return TemplateSpec(name="t", path=Path(spec_path), description="",
                        is_default=True, layouts={}, defaults={},
                        strip_slides=True, aspect="16:9")


def _build(slide, spec):
    deck = PowerpointPresentation([slide], "16:9", template_spec=spec)
    return deck.presentation.slides[0], deck.warnings


def _columns_text(slide):
    """Text of each content placeholder, left to right."""
    ordered = sorted(content_placeholders(slide), key=lambda ph: (ph.left, ph.top))
    return [ph.text_frame.text for ph in ordered]


def _codes(warnings):
    return [warning.code for warning in warnings]


# ---------------------------------------------------------------------------
# the regression
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("mapping", [
    {1: 4, 2: 1},        # the customer template: left is 4, right is 1
    {1: 13, 2: 14},      # both renumbered well outside the canonical range
    {1: 2, 2: 1},        # numbering that runs right-to-left
])
def test_two_content_fills_both_columns_whatever_the_idx(tmp_path, mapping):
    spec = _template(tmp_path, TWO_CONTENT, mapping)
    slide, warnings = _build({**SLIDE, "layout": TWO_CONTENT}, spec)

    left, right = _columns_text(slide)
    assert "Odbornost" in left and "Strategii" in left
    assert "Rychlost" in right and "Drafting" in right
    # The left column's words must never appear in the right-hand box.
    assert "Odbornost" not in right
    assert W.COLUMN_DROPPED not in _codes(warnings)


def test_comparison_fills_headings_and_bodies_whatever_the_idx(tmp_path):
    # Exactly the shape of the template in #194.
    spec = _template(tmp_path, COMPARISON, {1: 4, 2: 13, 3: 14, 4: 15})
    slide, warnings = _build({**SLIDE, "layout": COMPARISON}, spec)

    texts = _columns_text(slide)
    assert "Vy dodáváte" in texts[0]
    assert "Odbornost" in texts[1]
    assert "AI dodává" in texts[2]
    assert "Rychlost" in texts[3]
    assert warnings == []


def test_comparison_pairs_each_heading_with_the_body_below_it(tmp_path):
    spec = _template(tmp_path, COMPARISON, {1: 4, 2: 13, 3: 14, 4: 15})
    prs = PptxReader(str(spec.path))
    layout = next(one for one in prs.slide_layouts if one.name == COMPARISON)

    columns = content_columns(layout)
    assert len(columns) == 2
    for heading, body in columns:
        assert heading is not None, "a Comparison column has a heading strip"
        assert heading.top < body.top
        assert heading.height < body.height
        # heading and body belong to the same column
        assert abs(heading.left - body.left) < body.width


def test_two_content_has_no_heading_strip(tmp_path):
    prs = PptxReader(str(BASE_16_9))
    layout = next(one for one in prs.slide_layouts if one.name == TWO_CONTENT)
    columns = content_columns(layout)
    assert len(columns) == 2
    assert all(heading is None for heading, _ in columns)


# ---------------------------------------------------------------------------
# every degraded path reports itself (AGENTS.md: never the log alone)
# ---------------------------------------------------------------------------

def test_heading_is_inlined_when_the_layout_has_no_heading_strip():
    slide, warnings = _build({**SLIDE, "layout": TWO_CONTENT}, _plain())

    left, right = _columns_text(slide)
    assert left.startswith("Vy dodáváte")
    assert right.startswith("AI dodává")
    assert "Odbornost" in left and "Rychlost" in right
    assert _codes(warnings).count(W.HEADING_INLINED) == 2
    assert all(w.severity == "info" for w in warnings if w.code == W.HEADING_INLINED)


def test_columns_merge_when_the_layout_reserves_one_body():
    slide, warnings = _build({**SLIDE, "layout": SINGLE_CONTENT}, _plain())

    body = _columns_text(slide)[0]
    for word in ("Vy dodáváte", "Odbornost", "Strategii",
                 "AI dodává", "Rychlost", "Drafting"):
        assert word in body, f"{word!r} was lost when the columns merged"
    assert W.COLUMNS_MERGED in _codes(warnings)


def test_columns_dropped_is_reported_when_the_layout_has_no_body():
    slide, warnings = _build({**SLIDE, "layout": BLANK}, _plain())

    assert content_placeholders(slide) == []
    dropped = [w for w in warnings if w.code == W.COLUMN_DROPPED]
    assert len(dropped) == 1
    assert dropped[0].severity == "error"


def test_empty_columns_on_a_bodyless_layout_warn_about_nothing():
    empty = {"type": "two_column", "title": "t", "left": {}, "right": {}}
    _, warnings = _build({**empty, "layout": BLANK}, _plain())
    assert W.COLUMN_DROPPED not in _codes(warnings)


def test_nothing_is_lost_without_a_warning_saying_so():
    """The invariant behind #194: content is in the file, or the caller is told."""
    words = ("Vy dodáváte", "Odbornost", "Strategii",
             "AI dodává", "Rychlost", "Drafting")

    for layout in (TWO_CONTENT, COMPARISON, SINGLE_CONTENT, BLANK):
        slide, warnings = _build({**SLIDE, "layout": layout}, _plain())
        rendered = " ".join(
            shape.text_frame.text for shape in slide.shapes if shape.has_text_frame
        )
        missing = [word for word in words if word not in rendered]
        if missing:
            assert warnings, f"{layout}: lost {missing} and said nothing"
