"""The pptx analysis report shows the template's content area (#171).

`content_area` is the rectangle generated content is positioned against. A
`blank` or title-only layout is rarely an empty canvas — templates keep a
logo, a header rule and a footer band on it — so if this rectangle is read
from the wrong layout, or missing, every self-positioned slide lands on the
template's own furniture. `docs/templates.md` treats it as a first-class
diagnostic and `list_presentation_templates` returns it, but the admin UI —
which exists to make exactly this kind of mis-detection visible *before* a
deck is generated — did not show it.
"""
import re
import sys
import tempfile
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from fasthtml.common import to_xml

from admin.analysis import PptxAnalysis, _read_content_area, analyze_pptx
from admin.views.templates import pptx_analysis_report

SHIPPED = project_root / "default_templates" / "default_pptx_template_16_9.pptx"


def _rendered(analysis: PptxAnalysis) -> str:
    return to_xml(pptx_analysis_report(analysis))


def _row(html: str, label: str) -> str:
    m = re.search(re.escape(label) + r".*?</div>", html, re.S)
    return " ".join(re.sub(r"<[^>]+>", " ", m.group(0)).split()) if m else ""


class _StubResolver:
    """Stands in for `LayoutResolver` so a rectangle can be chosen directly."""

    def __init__(self, rect):
        self._rect = rect

    def __call__(self, presentation, overrides):
        return self

    def content_area(self):
        return self._rect


class _Slide:
    def __init__(self, width=12192000, height=6858000):
        self.slide_width = width
        self.slide_height = height


def _analyse_with(monkeypatch, rect, presentation=None):
    import pptx_tools.layouts as layouts_mod

    monkeypatch.setattr(layouts_mod, "LayoutResolver", _StubResolver(rect))
    analysis = PptxAnalysis()
    _read_content_area(analysis, presentation or _Slide())
    return analysis


def _rect(x, y, w, h, source="Content", slide=(12192000, 6858000)):
    from pptx_tools.placeholder_style import Rect

    sw, sh = slide
    return Rect(int(sw * x), int(sh * y), int(sw * w), int(sh * h), source=source)


# ---------------------------------------------------------------------------
# Against the real shipped template
# ---------------------------------------------------------------------------


def test_the_report_agrees_with_the_tools_own_diagnostics():
    """One formatter, so the page and `list_presentation_templates` cannot
    disagree about the rectangle they both describe."""
    from pptx_tools.layouts import LayoutResolver
    from pptx_tools.templates import content_area_summary, open_template

    data = SHIPPED.read_bytes()
    from_admin = analyze_pptx(data).content_area

    with tempfile.TemporaryDirectory() as tmp:
        staged = Path(tmp) / "x.pptx"
        staged.write_bytes(data)
        presentation = open_template(staged)
        from_tool = content_area_summary(LayoutResolver(presentation, {}),
                                         presentation)

    assert from_admin == from_tool
    assert from_admin is not None, "the shipped template does declare one"


def test_the_report_matches_the_documented_example():
    """docs/templates.md shows this exact rectangle for the shipped template.

    If the reading ever drifts, the documented example and the page would
    disagree and there would be nothing to say which was right.
    """
    area = analyze_pptx(SHIPPED.read_bytes()).content_area
    assert area == {"layout": "Nadpis a obsah", "x": "6.9%", "y": "26.6%",
                    "w": "86.2%", "h": "63.4%"}


def test_the_rendered_report_shows_the_percentages_and_the_layout():
    analysis = analyze_pptx(SHIPPED.read_bytes())
    row = _row(_rendered(analysis), "Content area")

    assert "6.9%" in row and "63.4%" in row, "the percentages"
    assert "Nadpis a obsah" in row, "and which layout they were read from"


def test_a_plausible_content_area_raises_no_warning_about_it():
    analysis = analyze_pptx(SHIPPED.read_bytes())
    assert not [w for w in analysis.warnings if "content area" in w.lower()]


# ---------------------------------------------------------------------------
# The cases an admin can act on
# ---------------------------------------------------------------------------


def test_no_content_area_is_reported_and_warned_about(monkeypatch):
    """The case with real consequences: the builder falls back to a fixed
    band that knows nothing about this template's decoration."""
    analysis = _analyse_with(monkeypatch, None)

    assert analysis.content_area is None
    assert any("fixed band" in w for w in analysis.warnings)
    assert "none declared" in _row(_rendered(analysis), "Content area")


def test_an_area_covering_the_whole_slide_is_warned_about(monkeypatch):
    """Then positioned content stays clear of nothing."""
    analysis = _analyse_with(monkeypatch, _rect(0, 0, 1.0, 1.0, "Full Bleed"))

    assert any("almost the whole slide" in w for w in analysis.warnings)
    assert any("Full Bleed" in w for w in analysis.warnings), \
        "the warning must name the layout so it can be checked"


def test_an_unusually_small_area_is_warned_about(monkeypatch):
    analysis = _analyse_with(monkeypatch, _rect(0.4, 0.4, 0.1, 0.1, "Tiny"))

    assert any("unusually small" in w for w in analysis.warnings)
    assert any("Tiny" in w for w in analysis.warnings)


@pytest.mark.parametrize("w, h", [(0.86, 0.63), (0.94, 0.94), (0.26, 0.9)])
def test_an_ordinary_area_is_not_warned_about(monkeypatch, w, h):
    """The thresholds must not cry wolf on the shapes real templates have."""
    analysis = _analyse_with(monkeypatch, _rect(0.05, 0.2, w, h))
    assert analysis.warnings == []


def test_a_zero_sized_slide_does_not_divide_by_zero(monkeypatch):
    """A malformed template must produce a warning, not a traceback."""
    analysis = _analyse_with(monkeypatch, _rect(0, 0, 1.0, 1.0),
                             presentation=_Slide(width=0, height=0))

    assert analysis.content_area is None
    assert analysis.warnings, "it should say something rather than crash"


def test_a_layout_override_does_not_move_the_content_area():
    """Why reading it without overrides is exact, not approximate.

    `LayoutResolver` applies a configured `layouts:` mapping in `provides()`
    and `resolve()`, but `content_area()` reads `_by_role`, which is filled
    from `classify_layout()` alone. So the card showing the un-overridden
    rectangle is showing what the builder will actually use.

    If this ever changes, the card must start taking the spec's overrides
    into account — which is what this test is here to force.
    """
    from pptx_tools.layouts import LayoutResolver
    from pptx_tools.placeholder_style import read_content_rect
    from pptx_tools.templates import open_template

    with tempfile.TemporaryDirectory() as tmp:
        staged = Path(tmp) / "x.pptx"
        staged.write_bytes(SHIPPED.read_bytes())
        presentation = open_template(staged)

        rects = {layout.name: read_content_rect(layout)
                 for layout in presentation.slide_layouts}
        plain = LayoutResolver(presentation, {}).content_area()

        # A layout whose rectangle genuinely differs, so the assertion cannot
        # pass just because two layouts happen to share a geometry.
        other = next(name for name, rect in rects.items()
                     if rect is not None and rect.as_tuple() != plain.as_tuple())
        overridden = LayoutResolver(presentation, {"content": other}).content_area()

    assert overridden.as_tuple() == plain.as_tuple(), (
        f"overriding content->{other!r} moved the content area; the admin card "
        "reads it without overrides and would now be misleading"
    )
