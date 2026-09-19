"""The shared warnings channel (#114).

Word and Excel produced documents that quietly lost content: a block that
would not render, a line with nowhere to go, a formula pointing at a table
that does not exist. Each was a log line in a response that said "created".
This module is what they now write to instead, alongside PowerPoint's own
channel — so the tests here are about the contract every tool depends on:
a stable code, a severity that follows from it, a location, and a channel
that survives being called from inside a loop over every cell.
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from warning_channel import (
    DEFAULT_LIMIT,
    SEVERITIES,
    SEVERITY_ERROR,
    SEVERITY_INFO,
    SEVERITY_WARNING,
    WARNINGS_TRUNCATED,
    DocumentWarning,
    WarningChannel,
)

SEVERITIES_BY_CODE = {
    "lost": SEVERITY_ERROR,
    "changed": SEVERITY_WARNING,
    "substituted": SEVERITY_INFO,
}


def channel(limit=DEFAULT_LIMIT):
    return WarningChannel(SEVERITIES_BY_CODE, limit=limit)


class TestTheRecord:

    def test_severity_comes_from_the_code(self):
        c = channel()
        c.add("lost", "a block is missing")
        c.add("changed", "a style was substituted")

        assert [w.severity for w in c] == [SEVERITY_ERROR, SEVERITY_WARNING]

    def test_an_unregistered_code_does_not_fail_a_finished_document(self):
        """A typo in a code is a bug, but not one worth losing the file over."""
        c = channel()
        c.add("invented", "something")

        assert c.records()[0].severity == SEVERITY_WARNING

    def test_as_dict_is_the_published_shape(self):
        c = channel()
        c.add("lost", "the block is missing", line=42)

        assert c.as_dicts() == [{
            "code": "lost",
            "severity": SEVERITY_ERROR,
            "message": "the block is missing",
            "line": 42,
        }]

    def test_the_location_is_spread_in_not_nested(self):
        """A caller reads warning["cell"], not warning["location"]["cell"]."""
        c = channel()
        c.add("changed", "renamed", sheet="Data", cell="C7")

        assert c.as_dicts()[0]["sheet"] == "Data"
        assert c.as_dicts()[0]["cell"] == "C7"
        assert "location" not in c.as_dicts()[0]

    def test_a_warning_with_no_location_carries_none(self):
        c = channel()
        c.add("changed", "the template has no such style")

        assert c.as_dicts() == [{
            "code": "changed",
            "severity": SEVERITY_WARNING,
            "message": "the template has no such style",
        }]

    def test_an_unknown_location_is_simply_absent(self):
        """Call sites pass line=None unconditionally rather than branching."""
        c = channel()
        c.add("lost", "somewhere", line=None, sheet="Data")

        assert c.as_dicts()[0] == {
            "code": "lost", "severity": SEVERITY_ERROR,
            "message": "somewhere", "sheet": "Data",
        }

    def test_str_is_the_log_line(self):
        assert str(DocumentWarning("c", "gone", location={"line": 7})) == "line 7: gone"
        assert str(DocumentWarning("c", "gone")) == "gone"
        assert str(DocumentWarning(
            "c", "gone", location={"sheet": "Data", "cell": "C7"},
        )) == "sheet Data, cell C7: gone"

    def test_every_severity_is_one_of_the_three(self):
        assert set(SEVERITIES) == {SEVERITY_ERROR, SEVERITY_WARNING, SEVERITY_INFO}


class TestTheChannel:

    def test_order_is_the_order_things_happened(self):
        c = channel()
        for i in range(3):
            c.add("lost", f"message {i}")

        assert [w.message for w in c] == ["message 0", "message 1", "message 2"]

    def test_the_same_warning_is_recorded_once(self):
        """A template missing 'List Number' warns once, not once per item."""
        c = channel()
        for _ in range(20):
            c.add("changed", "style 'List Number' is missing")

        assert len(c) == 1

    def test_the_same_message_at_different_places_is_kept(self):
        c = channel()
        c.add("lost", "the cell could not be written", cell="A1")
        c.add("lost", "the cell could not be written", cell="A2")

        assert [w.location["cell"] for w in c] == ["A1", "A2"]

    def test_a_flood_is_capped_and_says_so(self):
        """5000 dropped lines help nobody and cost the caller its context."""
        c = channel(limit=3)
        for i in range(10):
            c.add("lost", f"line {i} is not in the workbook")

        reported = c.records()
        assert len(reported) == 4
        assert [w.message for w in reported[:3]] == [
            "line 0 is not in the workbook",
            "line 1 is not in the workbook",
            "line 2 is not in the workbook",
        ]
        assert reported[-1].code == WARNINGS_TRUNCATED
        assert reported[-1].severity == SEVERITY_INFO
        assert "7 further warning(s)" in reported[-1].message

    def test_duplicates_do_not_count_towards_the_cap(self):
        c = channel(limit=2)
        for _ in range(50):
            c.add("changed", "the same thing")
        c.add("lost", "something else")

        assert [w.code for w in c] == ["changed", "lost"]

    def test_an_empty_channel_is_falsy_and_reports_nothing(self):
        c = channel()

        assert not c
        assert c.records() == []
        assert c.as_dicts() == []
        assert c.messages == []

    def test_a_channel_with_warnings_is_truthy(self):
        c = channel()
        c.add("lost", "gone")

        assert c

    def test_messages_are_the_rendered_lines(self):
        c = channel()
        c.add("lost", "gone", line=3)

        assert c.messages == ["line 3: gone"]

    def test_two_builds_never_share_a_channel(self):
        """Builds run concurrently on worker threads; state must not be global."""
        first, second = channel(), channel()
        first.add("lost", "only mine")

        assert len(second) == 0


@pytest.mark.parametrize("severity", SEVERITIES)
def test_severities_are_plain_strings_a_client_can_compare(severity):
    assert isinstance(severity, str)
