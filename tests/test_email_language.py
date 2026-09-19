"""The default proofing language of an email draft (#116).

`cs-CZ` was hard-coded in two places — the tool parameter and
`_create_eml_buffer()` — so a deployment writing in any other language had to
patch the source. `EMAIL_DEFAULT_LANGUAGE` sets it instead. The default is
still `cs-CZ`, deliberately: changing it would alter the drafts every
existing deployment produces.
"""
import base64
import email
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

import config as config_module
from config import Config
from email_tools.base_email_tool import _create_eml_buffer


def draft_language(**kwargs) -> tuple[str, str]:
    """(the <html lang> attribute, the Content-Language header) of a draft."""
    buffer = _create_eml_buffer(to=["a@b.c"], re="Subject", content="<p>x</p>", **kwargs)
    message = email.message_from_bytes(buffer.getvalue())
    buffer.close()
    body = base64.b64decode(message.get_payload()).decode("utf-8")
    lang_attr = next(line for line in body.splitlines() if "<html lang=" in line)
    return lang_attr, message.get("Content-Language")


@pytest.fixture
def env_language(monkeypatch):
    """Set EMAIL_DEFAULT_LANGUAGE and rebuild the config singleton from it."""
    def _set(value):
        if value is None:
            monkeypatch.delenv("EMAIL_DEFAULT_LANGUAGE", raising=False)
        else:
            monkeypatch.setenv("EMAIL_DEFAULT_LANGUAGE", value)
        monkeypatch.setattr(config_module, "_CONFIG", Config.from_env())
    return _set


class TestTheSetting:

    def test_unset_keeps_the_historical_default(self, env_language):
        """Existing deployments must not have their drafts change language."""
        env_language(None)
        assert Config.from_env().email_default_language == "cs-CZ"

    def test_the_environment_sets_it(self, env_language):
        env_language("en-US")
        assert Config.from_env().email_default_language == "en-US"

    @pytest.mark.parametrize("blank", ["", "   "])
    def test_a_blank_value_is_not_a_language(self, blank, env_language):
        """An empty tag would put lang="" on every draft."""
        env_language(blank)
        assert Config.from_env().email_default_language == "cs-CZ"


class TestWhatReachesTheDraft:

    def test_the_default_is_used_when_the_call_gives_none(self, env_language):
        env_language("en-US")
        lang_attr, header = draft_language()
        assert 'lang="en-US"' in lang_attr
        assert header == "en-US"

    def test_the_call_still_wins(self, env_language):
        env_language("en-US")
        lang_attr, header = draft_language(language="de-DE")
        assert 'lang="de-DE"' in lang_attr
        assert header == "de-DE"

    def test_the_shipped_default_is_unchanged(self, env_language):
        env_language(None)
        lang_attr, header = draft_language()
        assert 'lang="cs-CZ"' in lang_attr
        assert header == "cs-CZ"
