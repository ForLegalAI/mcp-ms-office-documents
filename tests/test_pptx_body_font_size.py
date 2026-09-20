"""The fit estimate measures the size the template really renders body text at.

The bug (#195): ``_fit_text`` always passed ``DEFAULT_BODY_FONT_SIZE`` — 18pt —
to ``estimate_text_fill``, whatever the template said. Both templates this
server ships set 28pt in the master's ``<p:bodyStyle>``, so every fill estimate
was low by the square of 28/18, roughly 2.4x. Text needing 1.9x its placeholder
measured as 0.86x: no ``fontScale`` was written, and PowerPoint renders a bare
``<a:normAutofit/>`` at full size until someone clicks into the box. The
overflow warning is computed from the same number, so it never fired either.
"""

import re
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from pptx import Presentation as PptxReader
from pptx.oxml.ns import qn

from pptx_tools import warnings as W
from pptx_tools.placeholder_style import (
    content_placeholders, read_body_font_size, read_master_body_font_size,
)
from pptx_tools.slide_builder import PowerpointPresentation
from pptx_tools.templates import TemplateSpec

BASE_16_9 = project_root / "default_templates" / "default_pptx_template_16_9.pptx"

# Long enough that 28pt cannot fit the content placeholder but 12pt can.
DENSE = "\n".join(
    f"- Bod {n}: " + "dlouhý text o právních službách a jejich automatizaci " * 3
    for n in range(1, 7)
)


def _spec(path) -> TemplateSpec:
    return TemplateSpec(name="t", path=Path(path), description="", is_default=True,
                        layouts={}, defaults={}, strip_slides=True, aspect="16:9")


def _resize_master_body(tmp_path, points) -> Path:
    """Copy the shipped template with a different master body size."""
    prs = PptxReader(str(BASE_16_9))
    master = prs.slide_masters[0]
    bodyStyle = master._element.find(qn("p:txStyles")).find(qn("p:bodyStyle"))
    for level in range(1, 10):
        lvl = bodyStyle.find(qn(f"a:lvl{level}pPr"))
        if lvl is None:
            continue
        defRPr = lvl.find(qn("a:defRPr"))
        if defRPr is not None:
            defRPr.set("sz", str(int(points * 100)))
    path = tmp_path / f"body{points}.pptx"
    prs.save(str(path))
    return path


def _font_scale(slide):
    """The ``fontScale`` written into the slide's body placeholder, or None."""
    body = content_placeholders(slide)[0]
    match = re.search(r'normAutofit[^/>]*fontScale="(\d+)"',
                      body.text_frame._txBody.xml)
    return int(match.group(1)) / 100000.0 if match else None


def _build(body, path):
    deck = PowerpointPresentation(
        [{"type": "content", "title": "t", "body": body}], "16:9",
        template_spec=_spec(path))
    return deck.presentation.slides[0], deck.warnings


# ---------------------------------------------------------------------------
# reading the size
# ---------------------------------------------------------------------------

def test_shipped_template_body_is_not_the_built_in_default():
    """The assumption the bug rested on, pinned so it cannot drift back."""
    prs = PptxReader(str(BASE_16_9))
    assert read_master_body_font_size(prs.slide_masters[0]) == 28.0


@pytest.mark.parametrize("points", [12, 18, 28, 40])
def test_body_size_is_read_from_the_master(tmp_path, points):
    path = _resize_master_body(tmp_path, points)
    prs = PptxReader(str(path))
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    assert read_body_font_size(content_placeholders(slide)[0]) == float(points)


def test_a_placeholder_overrides_the_master(tmp_path):
    prs = PptxReader(str(BASE_16_9))
    layout = prs.slide_layouts[1]
    body = content_placeholders(layout)[0]
    body._element.find(qn("p:txBody")).append(
        body._element.makeelement(qn("a:lstStyle"), {}))
    lstStyle = body._element.find(qn("p:txBody")).find(qn("a:lstStyle"))
    lvl1 = lstStyle.makeelement(qn("a:lvl1pPr"), {})
    defRPr = lvl1.makeelement(qn("a:defRPr"), {"sz": "1100"})
    lvl1.append(defRPr)
    lstStyle.append(lvl1)
    path = tmp_path / "override.pptx"
    prs.save(str(path))

    prs = PptxReader(str(path))
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    assert read_body_font_size(content_placeholders(slide)[0]) == 11.0


# ---------------------------------------------------------------------------
# what it changes
# ---------------------------------------------------------------------------

def test_dense_text_on_a_large_body_is_shrunk(tmp_path):
    slide, warnings = _build(DENSE, _resize_master_body(tmp_path, 28))
    scale = _font_scale(slide)
    assert scale is not None, "a 28pt body over its box must carry a fontScale"
    assert scale < 1.0
    assert W.TEXT_OVERFLOW in [warning.code for warning in warnings]


def test_the_same_text_on_a_small_body_is_left_alone(tmp_path):
    slide, warnings = _build(DENSE, _resize_master_body(tmp_path, 10))
    assert _font_scale(slide) is None
    assert W.TEXT_OVERFLOW not in [warning.code for warning in warnings]


def test_a_bigger_body_size_means_a_smaller_scale(tmp_path):
    """The estimate must track the template, not a constant."""
    scales = {}
    for points in (14, 20, 28, 40):
        slide, _ = _build(DENSE, _resize_master_body(tmp_path, points))
        scales[points] = _font_scale(slide) or 1.0

    ordered = [scales[points] for points in (14, 20, 28, 40)]
    assert ordered == sorted(ordered, reverse=True), scales
    assert scales[40] < scales[14]


def test_the_scale_brings_the_text_back_inside_the_box(tmp_path):
    """fontScale x body size should land near the size that would have fitted."""
    slide, _ = _build(DENSE, _resize_master_body(tmp_path, 28))
    effective = 28.0 * _font_scale(slide)

    fitted, _ = _build(DENSE, _resize_master_body(tmp_path, int(effective)))
    assert _font_scale(fitted) is None, (
        f"text shrunk to {effective:.1f}pt still does not fit at that size")
