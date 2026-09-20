"""A generated deck carries no "Click to add text" boxes.

``add_slide()`` copies every placeholder the layout defines, so a layout that
offers more than the slide filled left empty prompt boxes behind: the third
card of a three-card layout, the heading strip of a Comparison column given no
heading, the body of a Section Header. They neither print nor show in a
slideshow, but they are the first thing anyone opening the file to edit it
sees — three per slide on the template in #194.
"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from pptx import Presentation as PptxReader
from pptx.enum.shapes import PP_PLACEHOLDER

from pptx_tools.slide_builder import PowerpointPresentation
from pptx_tools.templates import TemplateSpec

BASE_16_9 = project_root / "default_templates" / "default_pptx_template_16_9.pptx"

CHROME = {PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER}

DECK = [
    {"type": "title", "title": "Deck"},                       # no subtitle
    {"type": "section", "title": "Part one"},                 # body unused
    {"type": "content", "title": "Body", "body": "- a\n- b"},
    {"type": "two_column", "title": "Columns",
     "left": {"heading": "L", "body": "- l"},
     "right": {"body": "- r"}},                               # right heading unused
    {"type": "kpi", "title": "Numbers",
     "items": [{"value": "1", "label": "one"}]},
    {"type": "quote", "title": "Q", "text": "words"},
    {"type": "blank", "title": "Free",
     "elements": [{"kind": "text", "x": "10%", "y": "40%", "w": "50%", "text": "hi"}]},
]


def _spec(path=BASE_16_9) -> TemplateSpec:
    return TemplateSpec(name="t", path=Path(path), description="", is_default=True,
                        layouts={}, defaults={}, strip_slides=True, aspect="16:9")


def _empty_placeholders(slide):
    return [ph for ph in slide.placeholders
            if ph.placeholder_format.type not in CHROME
            and ph.has_text_frame
            and not ph.text_frame.text.strip()]


def test_no_slide_keeps_an_empty_placeholder():
    deck = PowerpointPresentation(DECK, "16:9", template_spec=_spec())
    for number, slide in enumerate(deck.presentation.slides, 1):
        assert _empty_placeholders(slide) == [], (
            f"slide {number} ({slide.slide_layout.name}) kept an empty placeholder")


def test_filled_placeholders_are_kept():
    deck = PowerpointPresentation(
        [{"type": "two_column", "title": "t",
          "left": {"heading": "L", "body": "- l"},
          "right": {"heading": "R", "body": "- r"}}],
        "16:9", template_spec=_spec())
    slide = deck.presentation.slides[0]

    rendered = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    for word in ("t", "L", "R", "l", "r"):
        assert word in rendered.split() or word in rendered


def test_a_title_slide_without_a_subtitle_drops_the_subtitle_box():
    deck = PowerpointPresentation([{"type": "title", "title": "Only a title"}],
                                  "16:9", template_spec=_spec())
    slide = deck.presentation.slides[0]
    kinds = {ph.placeholder_format.type for ph in slide.placeholders}
    assert PP_PLACEHOLDER.SUBTITLE not in kinds
    assert slide.shapes.title.text == "Only a title"


def test_an_unfilled_picture_placeholder_is_dropped():
    """An image slide that names no usable picture layout must not leave a frame."""
    deck = PowerpointPresentation(
        [{"type": "content", "title": "t", "body": "- a",
          "layout": "Obrázek s titulkem"}],
        "16:9", template_spec=_spec())
    slide = deck.presentation.slides[0]
    kinds = {ph.placeholder_format.type for ph in slide.placeholders}
    assert PP_PLACEHOLDER.PICTURE not in kinds


def test_a_filled_picture_placeholder_survives(tmp_path):
    """A placeholder holding a picture has no text frame; it must not be swept."""
    import base64
    from PIL import Image

    path = tmp_path / "x.png"
    Image.new("RGB", (400, 300), (10, 20, 30)).save(path)
    source = "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode()

    deck = PowerpointPresentation(
        [{"type": "image", "title": "t", "source": source}],
        "16:9", template_spec=_spec())
    slide = deck.presentation.slides[0]

    pictures = [ph for ph in slide.placeholders
                if ph.placeholder_format.type == PP_PLACEHOLDER.PICTURE]
    assert len(pictures) == 1
    assert pictures[0].image.size == (400, 300)


def test_chrome_placeholders_are_left_for_the_footer_pass():
    deck = PowerpointPresentation(
        [{"type": "content", "title": "t", "body": "- a"}], "16:9",
        template_spec=_spec(), footer_text="ForLegalAI", show_slide_numbers=True)
    slide = deck.presentation.slides[0]
    rendered = " ".join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
    assert "ForLegalAI" in rendered
