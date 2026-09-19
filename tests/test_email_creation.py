"""The static email tool's draft, end to end (#112).

`email_tools/base_email_tool.py` had no test module of its own: it was
exercised only incidentally by the run_blocking, unique-prefix and LibreChat
tests, none of which look at the draft they produce. What a caller actually
receives — the headers Outlook reads, the base64 HTML body, what is escaped
and what is not — was unasserted.

`tests/test_email_language.py` owns the `EMAIL_DEFAULT_LANGUAGE` setting;
this module takes the language as given and checks where it lands.
"""
import base64
import email
import sys
from email.header import decode_header, make_header
from pathlib import Path
from unittest.mock import patch

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from email_tools.base_email_tool import _create_eml_buffer, create_eml


def draft(**kwargs):
    """Build a draft and return the parsed message."""
    kwargs.setdefault("to", ["someone@example.com"])
    kwargs.setdefault("re", "Subject line")
    kwargs.setdefault("content", "<p>Body</p>")
    buffer = _create_eml_buffer(**kwargs)
    message = email.message_from_bytes(buffer.getvalue())
    buffer.close()
    return message


def body_of(message) -> str:
    """The decoded HTML the recipient sees."""
    return base64.b64decode(message.get_payload()).decode("utf-8")


class TestHeaders:

    def test_recipients_are_joined_with_a_comma(self):
        message = draft(to=["a@x.com", "b@x.com"], cc=["c@x.com"], bcc=["d@x.com"])

        assert message["To"] == "a@x.com, b@x.com"
        assert message["Cc"] == "c@x.com"
        assert message["Bcc"] == "d@x.com"

    def test_absent_recipient_fields_are_absent_headers(self):
        message = draft(to=["a@x.com"], cc=None, bcc=None)

        assert message["Cc"] is None
        assert message["Bcc"] is None

    def test_a_non_ascii_subject_survives_the_round_trip(self):
        message = draft(re="Přehled výsledků")

        assert str(make_header(decode_header(message["Subject"]))) == "Přehled výsledků"

    def test_an_ascii_subject_is_encoded_too(self):
        """`Header(subject, 'utf-8')` is unconditional, so even a plain
        subject travels as an encoded word. Legal, and what clients get."""
        assert draft(re="Hello")["Subject"] == "=?utf-8?q?Hello?="

    def test_the_subject_header_carries_the_raw_text_not_the_escaped_one(self):
        """The escaping is for the `<title>` only. The header is read by a
        mail client, not a browser, so `&amp;` there would be shown
        literally. `Header(re, …)` and `{{subject}}` take deliberately
        different strings and nothing else pinned that they stay different."""
        subject = "Q3 <b>report</b> & more"

        message = draft(re=subject)

        assert str(make_header(decode_header(message["Subject"]))) == subject

    def test_the_draft_carries_a_date(self):
        assert draft()["Date"] is not None

    def test_the_language_reaches_both_language_headers(self):
        message = draft(language="de-DE")

        assert message["Content-Language"] == "de-DE"
        assert message["Accept-Language"] == "de-DE"

    def test_x_unsent_is_what_makes_outlook_open_it_as_a_draft(self):
        assert draft()["X-Unsent"] == "1"


class TestPriority:

    @pytest.mark.parametrize("priority,expected", [
        ("high", ("1 (Highest)", "High", "High")),
        ("low", ("5 (Lowest)", "Low", "Low")),
    ])
    def test_the_three_headers_move_together(self, priority, expected):
        message = draft(priority=priority)

        assert (message["X-Priority"], message["X-MSMail-Priority"],
                message["Importance"]) == expected

    def test_normal_sets_none_of_them(self):
        message = draft(priority="normal")

        assert message["X-Priority"] is None
        assert message["X-MSMail-Priority"] is None
        assert message["Importance"] is None

    @pytest.mark.parametrize("priority", ["HIGH", "Low", "NoRmAl"])
    def test_case_does_not_matter(self, priority):
        draft(priority=priority)      # no raise

    def test_an_unknown_priority_is_refused(self):
        with pytest.raises(ValueError, match="Priority must be"):
            draft(priority="urgent")

    @pytest.mark.parametrize("unset", [None, ""])
    def test_an_unset_priority_is_the_normal_one(self, unset):
        """The MCP parameter defaults to `"normal"`, so an unset priority is
        the absence of a choice, not a bad one. Until #112 `None` reached a
        second `priority.lower()` past the guard and surfaced as a
        RuntimeError; `""` was quietly accepted here but by a different
        route, so the two disagreed about what "unset" meant."""
        message = draft(priority=unset)

        assert message["X-Priority"] is None
        assert message["X-MSMail-Priority"] is None
        assert message["Importance"] is None

    @pytest.mark.parametrize("not_a_priority", [5, 0, ["high"], object()])
    def test_a_priority_that_is_not_a_string_is_refused_not_a_crash(self, not_a_priority):
        """A truthy non-string used to reach `.lower()` in the guard itself
        and escape as a bare AttributeError, outside the try; a falsy one
        got as far as the header block and came back as RuntimeError. Same
        mistake by the caller, three different answers. It is input they can
        fix, so it is a ValueError like any other (#112)."""
        with pytest.raises(ValueError, match="Priority must be"):
            draft(priority=not_a_priority)


class TestTheBody:

    def test_it_is_base64_utf8_html(self):
        message = draft()

        assert message.get_content_type() == "text/html"
        assert message.get_content_charset() == "utf-8"
        assert message["Content-Transfer-Encoding"] == "base64"

    def test_the_callers_html_is_inserted_raw(self):
        """`{{{content}}}`: the caller is trusted with the fragment, which is
        the whole point of the tool taking HTML."""
        message = draft(content="<p>Kept <strong>bold</strong></p>")

        assert "<p>Kept <strong>bold</strong></p>" in body_of(message)

    def test_the_subject_is_escaped_into_the_title(self):
        """`{{subject}}`, not `{{{subject}}}` — a subject is text."""
        body = body_of(draft(re="Q3 <b>report</b> & more"))

        assert "<title>Q3 &lt;b&gt;report&lt;/b&gt; &amp; more</title>" in body
        assert "<title>Q3 <b>report</b>" not in body

    def test_the_language_lands_in_the_lang_attributes(self):
        body = body_of(draft(language="fr-FR"))

        assert '<html lang="fr-FR">' in body
        assert '<body lang="fr-FR">' in body

    def test_a_language_cannot_break_out_of_its_attribute(self):
        """Quotes are stripped, so a crafted tag cannot add an attribute."""
        message = draft(language='en-US" onload="alert(1)')
        body = body_of(message)

        assert 'onload="' not in body
        assert '"' not in message["Content-Language"]


class TestRefusals:

    def test_content_is_required(self):
        with pytest.raises(ValueError, match="content is required"):
            draft(content=None)

    def test_a_subject_is_required(self):
        with pytest.raises(ValueError, match="subject is required"):
            draft(re=None)

    def test_an_empty_subject_is_no_subject(self):
        with pytest.raises(ValueError, match="subject is required"):
            draft(re="")


class TestTemplateResolution:

    def test_the_default_template_is_used_when_no_custom_one_exists(self):
        """The shipped template is what produces the styling the tool promises."""
        body = body_of(draft())

        assert "font-family: Arial" in body
        assert "<!DOCTYPE html>" in body

    def test_a_custom_template_wins(self, tmp_path):
        custom = tmp_path / "custom_email_template.html"
        custom.write_text(
            '<html lang="{{language}}"><title>{{subject}}</title>'
            "<body>CUSTOM {{{content}}}</body></html>",
            encoding="utf-8",
        )

        with patch("email_tools.base_email_tool.find_email_template",
                   return_value=str(custom)):
            body = body_of(draft(content="<p>hi</p>"))

        assert "CUSTOM <p>hi</p>" in body
        assert "font-family: Arial" not in body      # not the default one

    def test_no_template_at_all_is_an_error_not_an_empty_draft(self):
        with patch("email_tools.base_email_tool.find_email_template",
                   return_value=None):
            with pytest.raises(RuntimeError, match="template not found"):
                draft()


class TestTheUploadingWrapper:

    def test_create_eml_uploads_the_buffer_and_returns_the_backend_result(self):
        with patch("email_tools.base_email_tool.upload_file") as upload:
            upload.return_value = "https://example.invalid/draft.eml"
            result = create_eml(to=["a@x.com"], re="S", content="<p>x</p>",
                                file_name="my_draft")

        assert result == "https://example.invalid/draft.eml"
        buffer, extension = upload.call_args.args
        assert extension == "eml"
        assert upload.call_args.kwargs["filename"] == "my_draft"

    def test_the_buffer_is_closed_even_when_the_upload_fails(self):
        with patch("email_tools.base_email_tool.upload_file",
                   side_effect=RuntimeError("backend down")) as upload:
            with pytest.raises(RuntimeError, match="backend down"):
                create_eml(to=["a@x.com"], re="S", content="<p>x</p>")

        assert upload.call_args.args[0].closed


class TestTheDynamicToolDiffersOnPurpose:
    """The dynamic email tools build their own message and set a smaller
    header set than the static tool — no Date, no language, no priority.
    Documented in docs/development/tools/email.md; pinned here because the
    difference is easy to erase by accident when editing either one."""

    SPEC = {
        "name": "header_probe_email",
        "description": "header probe",
        "html_path": "default_email_template.html",
        "args": [{"name": "subject", "type": "string", "required": True,
                  "description": "Subject"}],
    }

    def _draft_from_dynamic_tool(self):
        import asyncio

        from fastmcp import FastMCP

        from email_tools.dynamic_email_tools import register_email_template

        mcp = FastMCP("header-probe")
        assert register_email_template(mcp, self.SPEC)
        tool = asyncio.run(mcp.get_tool(self.SPEC["name"]))

        captured = {}

        def capture(buffer, extension, **kwargs):
            captured["bytes"] = buffer.getvalue()
            return "ok"

        with patch("email_tools.dynamic_email_tools.upload_file", side_effect=capture):
            # The generated handler takes one `data` model, as the other
            # dynamic-tool tests call it.
            asyncio.run(tool.run({"data": {"subject": "hello", "to": ["a@x.com"]}}))
        return email.message_from_bytes(captured["bytes"])

    def test_it_sets_subject_recipients_and_x_unsent(self):
        message = self._draft_from_dynamic_tool()

        assert message["Subject"] == "hello"
        assert message["To"] == "a@x.com"
        assert message["X-Unsent"] == "1"

    def test_it_sets_none_of_the_static_tools_other_headers(self):
        message = self._draft_from_dynamic_tool()

        for header in ("Date", "Content-Language", "Accept-Language",
                       "X-Priority", "Importance"):
            assert message[header] is None, f"{header} is the static tool's"
