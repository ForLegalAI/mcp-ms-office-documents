"""The warnings channel as data (#122).

The channel used to be English sentences: displayable and nothing else. Each
entry now carries a stable `code`, the `slide` it happened on, a `severity`
and the same `message` as before, so a caller can branch, count and decide.
"""
import base64
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from pptx_tools import warnings as W
from pptx_tools.slide_builder import PowerpointPresentation
from pptx_tools.warnings import SlideWarning, make_warning

PNG_DATA_URI = "data:image/png;base64," + base64.b64encode(base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)).decode()


def build(slides, fmt="16:9", **kwargs):
    return PowerpointPresentation(slides, fmt, **kwargs)


def codes(pres):
    return [warning.code for warning in pres.warnings]


class TestTheRecord:

    def test_a_warning_carries_code_slide_and_severity(self):
        pres = build([{"type": "blank", "title": "T",
                       "elements": [{"kind": "text", "text": "x", "x": 40, "y": 1, "w": 1}]}])

        warning = pres.warnings[0]
        assert isinstance(warning, SlideWarning)
        assert warning.code == W.ELEMENT_OFF_SLIDE
        assert warning.slide == 0      # the index in the slides list, as the message says
        assert warning.severity == W.SEVERITY_ERROR
        assert "starts off the slide" in warning.message

    def test_the_message_is_the_sentence_it_always_was(self):
        pres = build([{"type": "image", "title": "I", "source": "data:image/png;base64,nope!"}])

        assert str(pres.warnings[0]) == "slide 0: " + pres.warnings[0].message
        assert pres.warning_messages == [str(w) for w in pres.warnings]

    def test_a_deck_wide_warning_has_no_slide(self):
        pres = build([{"type": "content", "title": "C", "body": "- a"}], "21:9")

        warning = [w for w in pres.warnings if w.code == W.FORMAT_SUBSTITUTED][0]
        assert warning.slide is None
        assert str(warning) == warning.message
        assert "slide" not in warning.as_dict()

    def test_as_dict_is_the_published_shape(self):
        warning = make_warning(W.IMAGE_FAILED, "image could not be loaded", slide=2)

        assert warning.as_dict() == {
            "code": "image_failed",
            "severity": "error",
            "message": "image could not be loaded",
            "slide": 2,
        }


class TestTheTaxonomy:

    def test_severity_comes_from_the_code(self):
        assert make_warning(W.TEXT_OVERFLOW, "m").severity == W.SEVERITY_WARNING
        assert make_warning(W.SUBTITLE_DROPPED, "m").severity == W.SEVERITY_ERROR
        assert make_warning(W.KPI_CROWDED, "m").severity == W.SEVERITY_INFO

    def test_every_code_has_a_severity(self):
        """A code without one would default silently to 'warning'."""
        declared = {
            value for name, value in vars(W).items()
            if name.isupper() and not name.startswith("SEVERITY_") and isinstance(value, str)
        }
        assert declared == set(W.WARNING_SEVERITY)

    def test_every_severity_is_one_of_the_three(self):
        assert set(W.WARNING_SEVERITY.values()) <= {
            W.SEVERITY_ERROR, W.SEVERITY_WARNING, W.SEVERITY_INFO}

    def test_an_unregistered_code_does_not_fail_a_finished_deck(self):
        assert make_warning("invented", "m").severity == W.SEVERITY_WARNING


class TestCodesInPractice:

    @pytest.mark.parametrize("slide_data,expected", [
        ({"type": "image", "title": "I", "source": "data:image/png;base64,nope!"},
         W.IMAGE_FAILED),
        ({"type": "table", "title": "T", "rows": []}, W.TABLE_EMPTY),
        ({"type": "chart", "title": "C", "chart_type": "column", "categories": ["a"],
          "series": [{"name": "s", "values": [1, 2]}]}, W.CHART_SERIES_LENGTH),
        ({"type": "kpi", "title": "K", "items": [{"value": str(n), "label": "l"}
                                                 for n in range(5)]}, W.KPI_CROWDED),
        ({"type": "agenda", "title": "A", "items": []}, W.AGENDA_EMPTY),
    ])
    def test_the_expected_code_is_recorded(self, slide_data, expected):
        assert expected in codes(build([slide_data]))

    def test_a_clean_deck_records_nothing(self):
        pres = build([{"type": "title", "title": "T", "subtitle": "S"},
                      {"type": "content", "title": "C", "body": "- one"}])
        assert pres.warnings == [] and pres.warning_messages == []

    def test_a_caller_can_tell_loss_from_alteration(self):
        """The point of severity: one deck is missing content, the other is not."""
        lost = build([{"type": "blank", "elements": [
            {"kind": "text", "text": "x", "x": 40, "y": 1, "w": 1}]}])
        altered = build([{"type": "content", "title": "C",
                          "body": "\n".join(f"- {'word ' * 14}" for _ in range(30))}])

        assert any(w.severity == W.SEVERITY_ERROR for w in lost.warnings)
        assert altered.warnings and not any(
            w.severity == W.SEVERITY_ERROR for w in altered.warnings)


# =============================================================================
# The tool boundary
# =============================================================================
# The structure only matters if it survives to the caller.

class TestToolBoundary:

    @staticmethod
    async def call(monkeypatch, **arguments):
        import main
        from fastmcp import Client

        async def fake_upload(file_buffer, extension, file_name, user_context, message, **kwargs):
            return "https://example.invalid/deck.pptx"

        monkeypatch.setattr(main, "upload_and_format_response", fake_upload)
        async with Client(main.mcp) as client:
            return await client.call_tool("create_powerpoint_presentation",
                                          arguments, raise_on_error=False)

    async def test_warnings_reach_the_caller_as_objects(self, monkeypatch):
        result = await self.call(monkeypatch, slides=[
            {"type": "image", "title": "I", "source": "data:image/png;base64,nope!"},
        ])

        assert not result.is_error
        warning = result.data["warnings"][0]
        assert warning["code"] == "image_failed"
        assert warning["severity"] == "error"
        assert warning["slide"] == 0
        assert "could not be loaded" in warning["message"]

    async def test_a_clean_deck_still_returns_a_bare_url(self, monkeypatch):
        result = await self.call(monkeypatch, slides=[
            {"type": "title", "title": "T", "subtitle": "S"},
        ])

        assert not result.is_error
        assert result.data == "https://example.invalid/deck.pptx"
