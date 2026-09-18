"""Measuring text instead of counting characters (#125).

The fill estimate decided how full a text box was from the number of
characters in it, with one mean glyph width for every face and every letter —
so "WWW WWW" and "iii iii" were the same length. It now measures against a
real font file where one can be loaded, and keeps the old arithmetic only
when none can.
"""
import sys
from pathlib import Path
from unittest.mock import patch

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from pptx import Presentation as PptxReader
from pptx.util import Inches

import pptx_tools.helpers as helpers
from pptx_tools import text_metrics
from pptx_tools.schema import Bullet
from pptx_tools.slide_builder import PowerpointPresentation
from pptx_tools.text_metrics import (
    _candidate_names, font_name, load_font, measure_lines, theme_body_typeface,
)

BASE_16_9 = project_root / "default_templates" / "default_pptx_template_16_9.pptx"

WIDTH, HEIGHT = Inches(6), Inches(3)


def fill(text, typeface="Arial", **kwargs):
    return helpers.estimate_text_fill([Bullet(text=text)], WIDTH, HEIGHT, 18.0,
                                      typeface=typeface, **kwargs)


class TestItMeasuresGlyphs:

    def test_the_same_characters_are_not_the_same_width(self):
        """The bug in one assertion: identical lengths, different text."""
        thin, wide = "iii " * 80, "WWW " * 80
        assert len(thin) == len(wide)

        assert fill(thin) < 0.6
        assert fill(wide) > 1.0

    def test_character_counting_could_not_tell_them_apart(self):
        thin, wide = "iii " * 80, "WWW " * 80
        with patch.object(helpers, "measure_lines", return_value=None):
            assert fill(thin) == fill(wide)

    def test_more_text_is_still_more_full(self):
        assert fill("word " * 20) < fill("word " * 200)

    def test_an_empty_box_is_empty(self):
        assert helpers.estimate_text_fill([], WIDTH, HEIGHT, 18.0) == 0.0

    def test_a_deeper_level_has_less_room(self):
        text = "word " * 60
        shallow = helpers.estimate_text_fill([Bullet(text=text, level=1)],
                                             WIDTH, HEIGHT, 18.0, typeface="Arial")
        deep = helpers.estimate_text_fill([Bullet(text=text, level=3)],
                                          WIDTH, HEIGHT, 18.0, typeface="Arial")
        assert deep > shallow


class TestWrapping:

    def test_lines_are_counted_by_wrapping_at_spaces(self):
        one_line = measure_lines([("short", 400)], "Arial", 18)
        several = measure_lines([("word " * 40, 400)], "Arial", 18)

        assert one_line == 1
        assert several > 1

    def test_a_word_wider_than_the_box_breaks_rather_than_counting_as_one(self):
        """Counting it as one line would under-measure a long unbroken string."""
        lines = measure_lines([("W" * 200, 100)], "Arial", 18)

        assert lines > 1

    def test_an_empty_bullet_still_occupies_a_line(self):
        assert measure_lines([("", 400)], "Arial", 18) == 1

    def test_a_zero_width_box_does_not_loop_forever(self):
        assert measure_lines([("some text", 0)], "Arial", 18) == 1


class TestChoosingAFace:

    def test_a_metric_compatible_substitute_is_preferred_over_a_generic_one(self):
        """Carlito's advances are Calibri's, so measuring it is exact."""
        candidates = _candidate_names("Calibri")

        assert candidates.index("Carlito") < candidates.index("DejaVuSans")

    def test_the_theme_reference_form_is_left_to_the_fallback(self):
        assert _candidate_names("+mj-lt")[0] in text_metrics.GENERIC_SANS

    def test_an_unknown_face_still_measures_with_something(self):
        font = load_font("Aptos", 18)

        assert font is not None
        assert font_name(font).endswith((".ttf", ".otf"))

    def test_fonts_are_cached_per_face_and_size(self):
        assert load_font("Arial", 18) is load_font("Arial", 18)

    def test_no_font_at_all_means_no_measurement(self):
        with patch.object(text_metrics, "GENERIC_SANS", ()):
            load_font.cache_clear()
            assert measure_lines([("text", 400)], "NoSuchFaceAnywhere", 18) is None
        load_font.cache_clear()


class TestTheDecksOwnFace:

    def test_the_theme_body_font_is_read_from_the_template(self):
        assert theme_body_typeface(PptxReader(str(BASE_16_9))) == "Aptos"

    def test_the_builder_measures_with_it(self):
        pres = PowerpointPresentation([{"type": "content", "title": "C", "body": "- x"}], "16:9")

        assert pres._typeface == "Aptos"

    def test_a_presentation_without_a_theme_does_not_fail(self):
        presentation = PptxReader(str(BASE_16_9))
        with patch.object(type(presentation), "slide_masters",
                          property(lambda self: [])):
            assert theme_body_typeface(presentation) is None


class TestItStillDrivesTheDeck:

    def test_overfull_text_is_shrunk_and_reported(self):
        pres = PowerpointPresentation([{
            "type": "content", "title": "C",
            "body": "\n".join(f"- {'WWWW ' * 20}" for _ in range(20)),
        }], "16:9")

        assert any(w.code == "text_overflow" for w in pres.warnings)

    def test_text_that_fits_is_left_alone(self):
        pres = PowerpointPresentation([{
            "type": "content", "title": "C", "body": "- one\n- two",
        }], "16:9")

        assert pres.warnings == []
