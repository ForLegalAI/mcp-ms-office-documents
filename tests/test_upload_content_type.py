"""The MIME type an uploaded object is given (#116).

`get_content_type()` decided by substring — `"pptx" in file_name` — so a
sanitised name like `notes.pptx_v2.docx` was uploaded to S3 as a PowerPoint,
and any name merely containing `xml` as an XML document. The extension after
the last dot decides now, from the one table the whole upload layer shares.
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from upload_tools.backends import librechat
from upload_tools.utils import MIME_TYPES, file_extension, get_content_type

DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class TestTheExtensionDecides:

    @pytest.mark.parametrize("name,expected", [
        ("report.docx", DOCX),
        ("deck.pptx", PPTX),
        ("book.xlsx", XLSX),
        ("draft.eml", "message/rfc822"),
        ("feed.xml", "application/xml"),
        ("REPORT.DOCX", DOCX),                  # case does not matter
        ("a1b2c3d4_My_Report.docx", DOCX),       # the unique-prefix form
    ])
    def test_known_types(self, name, expected):
        assert get_content_type(name) == expected

    @pytest.mark.parametrize("name,expected", [
        ("notes.pptx_v2.docx", DOCX),            # the reported case
        ("quarterly.xlsx.backup.docx", DOCX),
        ("presentation_about_docx.pptx", PPTX),
    ])
    def test_a_type_named_inside_the_name_is_not_the_type(self, name, expected):
        assert get_content_type(name) == expected

    @pytest.mark.parametrize("name", ["myxmlfile", "report", "notes.rtf", ""])
    def test_an_unknown_type_raises_rather_than_guessing(self, name):
        """The traditional backends upload only what this server generates,
        so anything else is a bug worth surfacing at the call site."""
        with pytest.raises(ValueError):
            get_content_type(name)

    def test_the_error_names_the_file(self):
        with pytest.raises(ValueError, match="myxmlfile"):
            get_content_type("myxmlfile")


class TestOneTableForTheUploadLayer:
    """The two halves disagreed about `.eml` until #116 — one called it
    `application/octet-stream`, the other `message/rfc822`."""

    def test_librechat_uses_the_same_table(self):
        assert librechat.MIME_TYPES is MIME_TYPES

    @pytest.mark.parametrize("name", ["draft.eml", "report.docx", "deck.pptx"])
    def test_both_paths_agree_on_what_this_server_generates(self, name):
        assert librechat.get_mime_type(name) == get_content_type(name)

    def test_librechat_still_guesses_where_the_others_raise(self):
        """LibreChat receives whatever a caller attaches; an attachment typed
        as a guess still reaches the conversation."""
        assert librechat.get_mime_type("notes.txt") == "text/plain"
        assert librechat.get_mime_type("mystery") == "application/octet-stream"
        with pytest.raises(ValueError):
            get_content_type("mystery")


@pytest.mark.parametrize("name,expected", [
    ("a.docx", "docx"), ("a.b.PPTX", "pptx"), ("noext", ""), ("", ""),
])
def test_file_extension(name, expected):
    assert file_extension(name) == expected
