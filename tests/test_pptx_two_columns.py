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


# ---------------------------------------------------------------------------
# classification: four placeholders are not enough to mean "Comparison"
# ---------------------------------------------------------------------------

def _reshape(layout, boxes):
    """Give *layout*'s content placeholders these (left, top, w, h) inches."""
    from pptx.util import Inches
    for placeholder, box in zip(content_placeholders(layout), boxes):
        placeholder.left, placeholder.top = Inches(box[0]), Inches(box[1])
        placeholder.width, placeholder.height = Inches(box[2]), Inches(box[3])


def test_three_cards_and_a_bar_is_not_a_comparison(tmp_path):
    """The layout from #194: three cards side by side plus a caption bar.

    It has four content placeholders, so the old count-only rule called it
    Comparison and every two-column slide in the deck landed on it.
    """
    from pptx_tools.layouts import classify_layout

    prs = PptxReader(str(BASE_16_9))
    layout = next(one for one in prs.slide_layouts if one.name == COMPARISON)
    _reshape(layout, [(1.11, 2.58, 3.29, 2.77),    # card 1
                      (5.12, 2.58, 3.29, 2.77),    # card 2
                      (9.14, 2.58, 3.29, 2.77),    # card 3
                      (3.97, 6.02, 5.59, 0.47)])   # caption bar
    assert classify_layout(layout) is None


def test_a_real_comparison_is_still_recognised():
    from pptx_tools.layouts import ROLE_COMPARISON, classify_layout

    prs = PptxReader(str(BASE_16_9))
    layout = next(one for one in prs.slide_layouts if one.name == COMPARISON)
    assert classify_layout(layout) == ROLE_COMPARISON


def test_four_stacked_bodies_are_not_a_comparison(tmp_path):
    """One column of four boxes is not two columns, whatever the count says."""
    from pptx_tools.layouts import classify_layout

    prs = PptxReader(str(BASE_16_9))
    layout = next(one for one in prs.slide_layouts if one.name == COMPARISON)
    _reshape(layout, [(0.9, 1.6, 11.5, 1.2), (0.9, 2.9, 11.5, 1.2),
                      (0.9, 4.2, 11.5, 1.2), (0.9, 5.5, 11.5, 1.2)])
    assert classify_layout(layout) is None


def _without_layout(tmp_path, name) -> TemplateSpec:
    prs = PptxReader(str(BASE_16_9))
    ids = prs.slide_masters[0]._element.sldLayoutIdLst
    target = next(i for i, one in enumerate(prs.slide_layouts) if one.name == name)
    ids.remove(list(ids)[target])
    path = tmp_path / "trimmed.pptx"
    prs.save(str(path))
    return TemplateSpec(name="t", path=path, description="", is_default=True,
                        layouts={}, defaults={}, strip_slides=True, aspect="16:9")


def test_a_template_with_no_comparison_layout_degrades_into_two_content(tmp_path):
    """Without this the positional fallback picked a layout with no body at all.

    On the #194 template `comparison` is unprovided and position 4 is Title
    Only, so a two-column slide landed there and lost both columns. A role the
    template really has is tried first.
    """
    spec = _without_layout(tmp_path, COMPARISON)
    slide, warnings = _build(SLIDE, spec)

    assert slide.slide_layout.name == TWO_CONTENT
    left, right = _columns_text(slide)
    assert "Odbornost" in left and "Rychlost" in right
    assert W.COLUMN_DROPPED not in _codes(warnings)
    assert W.LAYOUT_SUBSTITUTED in _codes(warnings)


def test_a_template_with_neither_column_layout_merges_rather_than_drops(tmp_path):
    prs = PptxReader(str(BASE_16_9))
    ids = prs.slide_masters[0]._element.sldLayoutIdLst
    for name in (COMPARISON, TWO_CONTENT):
        target = next(i for i, one in enumerate(prs.slide_layouts) if one.name == name)
        ids.remove(list(ids)[target])
    path = tmp_path / "no_columns.pptx"
    prs.save(str(path))
    spec = TemplateSpec(name="t", path=path, description="", is_default=True,
                        layouts={}, defaults={}, strip_slides=True, aspect="16:9")

    slide, warnings = _build(SLIDE, spec)
    rendered = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    for word in ("Vy dodáváte", "Odbornost", "AI dodává", "Rychlost"):
        assert word in rendered
    assert W.COLUMN_DROPPED not in _codes(warnings)


# ---------------------------------------------------------------------------
# the closing role: a contact layout no signature can detect
# ---------------------------------------------------------------------------

def test_closing_defaults_to_the_title_layout_without_complaining():
    slide, warnings = _build(
        {"type": "closing", "title": "Thanks", "contact": ["a@b.c"]}, _plain())
    assert slide.slide_layout.name == "Úvodní snímek"
    assert _codes(warnings) == []


def test_a_template_can_map_closing_to_its_own_contact_layout(tmp_path):
    """The point of the role: no placeholder arrangement says "contact slide"."""
    spec = TemplateSpec(name="t", path=BASE_16_9, description="", is_default=True,
                        layouts={"closing": "Záhlaví oddílu"}, defaults={},
                        strip_slides=True, aspect="16:9")
    slide, warnings = _build(
        {"type": "closing", "title": "Thanks", "contact": ["a@b.c"]}, spec)

    assert slide.slide_layout.name == "Záhlaví oddílu"
    assert _codes(warnings) == []
    rendered = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    assert "Thanks" in rendered and "a@b.c" in rendered


def test_closing_is_not_reported_as_a_missing_role():
    from pptx_tools.layouts import LayoutResolver

    resolver = LayoutResolver(PptxReader(str(BASE_16_9)))
    assert "closing" not in resolver.missing_roles()
