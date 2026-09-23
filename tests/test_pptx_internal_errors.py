"""A bug inside a slide builder is not reported as the caller's bad input.

Regression from the PowerPoint review: ``_build_slides`` turned *any*
exception into a ``ValueError`` — without ``from e`` and without a traceback —
and the handler reports every ``ValueError`` as "Invalid presentation input".
So a KeyError in our own code told the model to fix its slides, and the stack
trace that would locate the bug was never logged. Only a ``ValueError`` (the
builders' checks on caller input) keeps that path now; anything else is a
``RuntimeError`` chained to the original and logged with its traceback.
"""
import logging
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from pptx_tools.slide_builder import PowerpointPresentation

SLIDES = [{"type": "content", "title": "T", "body": "- a"}]


def _break_content_builder(monkeypatch, exc):
    def boom(self, slide_data, index):
        raise exc

    monkeypatch.setattr(PowerpointPresentation, "_build_content_slide", boom)


def test_an_internal_error_is_a_runtime_error_with_its_cause(monkeypatch, caplog):
    original = KeyError("layout_cache")
    _break_content_builder(monkeypatch, original)

    with caplog.at_level(logging.ERROR, logger="pptx_tools.slide_builder"):
        with pytest.raises(RuntimeError, match=r"slide 0 \(content\)") as excinfo:
            PowerpointPresentation(SLIDES, "16:9")

    assert excinfo.value.__cause__ is original
    assert not isinstance(excinfo.value, ValueError)
    logged = [r for r in caplog.records if "Internal error creating slide 0" in r.getMessage()]
    assert logged and logged[0].exc_info is not None


def test_a_value_error_is_still_the_callers_input(monkeypatch):
    original = ValueError("width must be positive")
    _break_content_builder(monkeypatch, original)

    with pytest.raises(ValueError, match="width must be positive") as excinfo:
        PowerpointPresentation(SLIDES, "16:9")

    assert excinfo.value.__cause__ is original


async def test_the_handler_reports_it_as_a_server_error(monkeypatch):
    import main
    from fastmcp import Client

    _break_content_builder(monkeypatch, KeyError("layout_cache"))
    async with Client(main.mcp) as client:
        result = await client.call_tool(
            "create_powerpoint_presentation", {"slides": SLIDES}, raise_on_error=False,
        )

    assert result.is_error
    text = result.content[0].text
    assert text.startswith("Error creating PowerPoint presentation")
    assert "Invalid" not in text
