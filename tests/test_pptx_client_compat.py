"""Tests for what the PowerPoint tool publishes and what it will accept.

The typed union (``tests/test_pptx_schema.py``) is the model of a slide; this
module covers the two places it meets a real MCP client:

* the **published** parameter schema, which must stay inside the keyword subset
  every client dialect can carry — a ``oneOf``/``$ref``/``discriminator`` union
  is silently downgraded by clients that bridge to a provider without them, and
  the model is then shown ``slides: array of string`` while the client keeps
  validating against the union, rejecting every call before it is sent;
* the **text fallback**, which reads a slide out of the JSON or markdown string
  such a client ends up sending.
"""

import json
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from pptx_tools import schema as slide_schema
from pptx_tools.schema import SLIDE_TYPES, coerce_slides, flat_slide_schema, slide_from_text
from pptx_tools.slide_builder import PowerpointPresentation


# =============================================================================
# The published schema
# =============================================================================

@pytest.fixture(scope="module")
def schema():
    return flat_slide_schema()


class TestPublishedSchema:

    def test_only_portable_keywords(self, schema):
        """Every keyword in the published schema must survive the trip.

        Walks schema positions rather than the rendered JSON, so a field *named*
        like a keyword ('type', 'items') is not mistaken for one.
        """
        def keywords(node):
            if isinstance(node, list):
                for item in node:
                    yield from keywords(item)
            elif isinstance(node, dict):
                for key, value in node.items():
                    if key in ("properties", "$defs"):
                        for sub in value.values():
                            yield from keywords(sub)
                        continue
                    yield key
                    yield from keywords(value)

        assert schema["type"] == "object"
        unportable = {
            "oneOf", "$ref", "$defs", "discriminator", "const", "allOf", "not",
            "if", "then", "else", "patternProperties", "prefixItems", "additionalItems",
        }
        found = unportable & set(keywords(schema))
        assert not found, f"{sorted(found)} do not survive every client"

    def test_type_is_an_enum_of_every_slide_type(self, schema):
        assert set(schema["properties"]["type"]["enum"]) == set(SLIDE_TYPES)
        assert schema["required"] == ["type"]

    def test_every_field_of_every_type_is_published(self, schema):
        published = set(schema["properties"])
        for model in slide_schema._SLIDE_MODELS:
            missing = set(model.model_fields) - published
            assert not missing, f"{model.__name__} fields missing from the schema: {missing}"

    def test_each_field_names_the_types_that_accept_it(self, schema):
        assert "[table]" in schema["properties"]["rows"]["description"]
        assert "[image]" in schema["properties"]["source"]["description"]
        # Shared fields are not spelled out fourteen times.
        assert "[all types]" in schema["properties"]["title"]["description"]

    def test_a_type_that_describes_nothing_joins_the_described_group(self, schema):
        """scatter declares 'legend' without a description of its own."""
        description = schema["properties"]["legend"]["description"]
        assert description.startswith("[chart, scatter] ")
        assert not description.endswith("[scatter]")

    def test_fields_two_types_spell_differently_offer_both_shapes(self, schema):
        def item_properties(shape):
            return set(shape.get("items", {}).get("properties", ()))

        # kpi.items is a list of objects, agenda.items a list of strings.
        shapes = schema["properties"]["items"]["anyOf"]
        assert {"value", "label", "delta"} in [item_properties(s) for s in shapes]
        assert {"items": {"type": "string"}, "type": "array"} in shapes

        # chart.series carries values, scatter.series carries points.
        shapes = schema["properties"]["series"]["anyOf"]
        assert {"name", "values"} in [item_properties(s) for s in shapes]
        assert {"name", "points"} in [item_properties(s) for s in shapes]

    def test_per_type_required_fields_are_documented(self, schema):
        description = schema["description"]
        assert "table: rows" in description
        assert "chart: chart_type, categories, series" in description

    def test_unknown_fields_are_still_refused(self, schema):
        assert schema["additionalProperties"] is False


# =============================================================================
# The text fallback
# =============================================================================

class TestTextSlides:

    def test_h1_becomes_a_title_slide(self):
        assert slide_from_text("# Deck\nBy someone") == {
            "type": "title", "title": "Deck", "subtitle": "By someone",
        }

    def test_h2_with_bullets_becomes_a_content_slide(self):
        assert slide_from_text("## Topic\n- one\n- two") == {
            "type": "content", "title": "Topic", "body": "- one\n- two",
        }

    def test_h1_with_bullets_is_content_not_a_title(self):
        assert slide_from_text("# Topic\n- one")["type"] == "content"

    def test_a_bare_line_is_a_titled_slide(self):
        assert slide_from_text("Slide 1") == {"type": "content", "title": "Slide 1"}

    def test_blank_text_does_not_crash(self):
        assert slide_from_text("   \n\n") == {"type": "content"}

    def test_the_payload_that_used_to_be_rejected(self):
        """The report that prompted this: markdown lines, one per slide."""
        slides = coerce_slides([
            "# Úvodní snímek\nPodtitul prezentace",
            "## První obsahový snímek\n- První bod\n- Druhý bod",
        ])
        assert [slide.type for slide in slides] == ["title", "content"]
        assert slides[0].subtitle == "Podtitul prezentace"
        assert slides[1].body == "- První bod\n- Druhý bod"

    def test_a_slide_sent_as_a_json_string(self):
        slides = coerce_slides(['{"type": "quote", "text": "Hello", "attribution": "Someone"}'])
        assert slides[0].type == "quote"
        assert slides[0].attribution == "Someone"

    def test_a_whole_deck_sent_as_one_json_string(self):
        slides = coerce_slides('[{"type": "title", "title": "Deck"}, {"type": "section", "title": "Part 1"}]')
        assert [slide.type for slide in slides] == ["title", "section"]

    def test_legacy_keys_still_apply_to_a_json_string_slide(self):
        slides = coerce_slides(['{"slide_type": "section", "slide_title": "Part 1"}'])
        assert slides[0].type == "section"
        assert slides[0].title == "Part 1"

    def test_a_text_deck_builds(self):
        pres = PowerpointPresentation(
            ["# Deck\nSubtitle", "## Topic\n- one\n- two", "Just a title"], "16:9",
        )
        assert len(pres.slides) == 3
        assert pres.save().getvalue()[:2] == b"PK"

    def test_bullets_with_no_heading_are_all_body(self):
        """The first bullet is not a title: it would lose the item and its marker."""
        assert slide_from_text("- point one\n- point two") == {
            "type": "content", "body": "- point one\n- point two",
        }

    def test_a_slide_that_starts_like_json_must_be_json(self):
        with pytest.raises(ValueError) as excinfo:
            coerce_slides(['{"type": "quote", "text": '])
        assert "slide 0" in str(excinfo.value)
        assert "does not parse" in str(excinfo.value)

    def test_a_truncated_deck_is_reported_not_collapsed(self):
        """A deck cut short by the client must not become one slide of raw JSON."""
        with pytest.raises(ValueError) as excinfo:
            coerce_slides('[{"type": "title", "title": "A"}, {"type": "content"')
        assert "does not parse" in str(excinfo.value)

    def test_the_error_reads_as_one_sentence(self):
        """No pydantic scaffolding ('slide ?: Value error, …') around the message."""
        with pytest.raises(ValueError) as excinfo:
            coerce_slides('[{"type": "title"')
        message = str(excinfo.value)
        assert message.startswith("Invalid slides: the deck was sent as one string")
        assert "Value error" not in message


# =============================================================================
# The tool boundary
# =============================================================================
# The failure this module exists for happened before any of the code above ran,
# in the client's own validation of the tool's declared schema. These tests call
# the registered tool the way a client does, so the declaration and what the
# server accepts are checked together.

class TestToolBoundary:

    @staticmethod
    async def call(monkeypatch, **arguments):
        import main
        from fastmcp import Client

        captured = {}

        async def fake_upload(file_buffer, extension, file_name, user_context, message, **kwargs):
            captured["bytes"] = file_buffer.getvalue()
            return f"https://example.invalid/{file_name or 'deck'}.{extension}"

        monkeypatch.setattr(main, "upload_and_format_response", fake_upload)
        async with Client(main.mcp) as client:
            result = await client.call_tool(
                "create_powerpoint_presentation", arguments, raise_on_error=False,
            )
        return result, captured

    async def test_slides_as_markdown_strings_are_accepted(self, monkeypatch):
        result, captured = await self.call(
            monkeypatch,
            slides=["# Úvodní snímek\nPodtitul prezentace", "## Obsah\n- První bod"],
            file_name="test",
        )
        assert not result.is_error
        assert captured["bytes"][:2] == b"PK"

    async def test_slides_as_objects_are_still_accepted(self, monkeypatch):
        result, captured = await self.call(
            monkeypatch,
            slides=[
                {"type": "title", "title": "Deck", "subtitle": "Subtitle"},
                {"type": "content", "title": "Topic", "body": "- one\n- two"},
            ],
        )
        assert not result.is_error
        assert captured["bytes"][:2] == b"PK"

    async def test_an_invalid_slide_still_gets_a_readable_error(self, monkeypatch):
        result, _ = await self.call(
            monkeypatch, slides=[{"type": "quote", "txt": "hello"}],
        )
        assert result.is_error
        message = result.content[0].text
        assert "slide 0" in message and "txt" in message

    async def test_a_truncated_deck_string_reaches_the_client_as_an_error(self, monkeypatch):
        result, captured = await self.call(
            monkeypatch, slides='[{"type": "title", "title": "A"}, {"type": "content"',
        )
        assert result.is_error
        assert "does not parse" in result.content[0].text
        assert "bytes" not in captured, "a truncated deck must not be built and uploaded"

    async def test_an_unknown_type_names_the_valid_ones(self, monkeypatch):
        result, _ = await self.call(monkeypatch, slides=[{"type": "bullets", "title": "x"}])
        assert result.is_error
        assert "Valid slide types" in result.content[0].text


class TestHeadingParseIsLinear:
    """Regression: the heading pattern "(.*?)\\s*#*$" retried its lazy group
    at every character of a blank run — 20k characters took 4 s, on the event
    loop, inside the argument validator. The closing hashes are now stripped
    in code; the result must be the same."""

    @pytest.mark.parametrize("text, title", [
        ("## Title ##", "Title"), ("# C#", "C"), ("# a # #", "a #"),
        ("###  spaced  ", "spaced"), ("####### seven", "####### seven"),
    ])
    def test_titles_are_unchanged(self, text, title):
        assert slide_from_text(text + "\n- a")["title"] == title

    def test_a_long_blank_run_is_cheap(self):
        import time

        start = time.perf_counter()
        slide = slide_from_text("# a" + " " * 50_000 + "b\n- x")
        assert time.perf_counter() - start < 0.5
        assert slide["title"].startswith("a") and slide["title"].endswith("b")
