"""Which template file wins, when several could (#112).

`template_utils` resolves every tool's template along two axes: the
directory order in `find_file_in_template_dirs` (container before checkout,
custom before default) and the filename order in `_resolve_from_candidates`
(`custom_<x>_template` before `default_<x>_template`). Both are what the
"drop a file in custom_templates/ to override" promise in docs/templates.md
rests on, and neither had a test — `tests/test_email_creation.py` patches
`find_email_template` out, so it proves only that the resolved file is the
one rendered, never which file that is.

The resolution is shared by Word, Excel, PowerPoint and email alike, so it
is tested here rather than in any one tool's module.
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

import template_utils


@pytest.fixture
def template_dirs(tmp_path, monkeypatch):
    """Point every search directory at a temp dir, in the real order.

    Returns the four directories under their production names so a test can
    place a file in exactly one of them.
    """
    dirs = {}
    for name in ("app_custom", "local_custom", "app_default", "local_default"):
        d = tmp_path / name
        d.mkdir()
        dirs[name] = d

    monkeypatch.setattr(template_utils, "APP_CUSTOM_DIR", dirs["app_custom"])
    monkeypatch.setattr(template_utils, "LOCAL_CUSTOM_DIR", dirs["local_custom"])
    monkeypatch.setattr(template_utils, "APP_DEFAULT_DIR", dirs["app_default"])
    monkeypatch.setattr(template_utils, "LOCAL_DEFAULT_DIR", dirs["local_default"])
    return dirs


def write(directory, filename, text="x"):
    path = directory / filename
    path.write_text(text, encoding="utf-8")
    return path


class TestTheCustomFileWins:
    """The override promise in docs/templates.md."""

    def test_a_custom_name_beats_a_default_name_in_the_same_directory(self, template_dirs):
        write(template_dirs["local_custom"], "custom_email_template.html", "CUSTOM")
        write(template_dirs["local_custom"], "default_email_template.html", "DEFAULT")

        resolved = Path(template_utils.find_email_template())

        assert resolved.read_text(encoding="utf-8") == "CUSTOM"

    def test_a_custom_directory_beats_a_default_directory(self, template_dirs):
        """The same filename in both. A shipped dynamic template such as
        `broadcast_email_style_1.html` is overridden by dropping a file of
        that name into custom_templates/, and only a shared name puts the
        directory order — rather than the candidate-name order — in charge.
        """
        write(template_dirs["local_custom"], "broadcast.html", "CUSTOM")
        write(template_dirs["local_default"], "broadcast.html", "DEFAULT")

        resolved = Path(template_utils.find_email_template("broadcast.html"))

        assert resolved.read_text(encoding="utf-8") == "CUSTOM"

    def test_the_container_mount_beats_the_checkout(self, template_dirs):
        """Both exist in a deployed container; the mounted one is the
        operator's and must win over whatever the image was built with."""
        write(template_dirs["app_custom"], "custom_email_template.html", "MOUNTED")
        write(template_dirs["local_custom"], "custom_email_template.html", "IMAGE")

        resolved = Path(template_utils.find_email_template())

        assert resolved.read_text(encoding="utf-8") == "MOUNTED"


class TestTheFallback:

    def test_the_default_is_used_when_no_custom_file_exists(self, template_dirs):
        write(template_dirs["local_default"], "default_email_template.html", "DEFAULT")

        resolved = Path(template_utils.find_email_template())

        assert resolved.read_text(encoding="utf-8") == "DEFAULT"

    def test_nothing_anywhere_resolves_to_none(self, template_dirs):
        assert template_utils.find_email_template() is None

    def test_an_explicit_filename_skips_the_candidate_list(self, template_dirs):
        """A dynamic template names its own file; the custom/default pair
        must not be consulted for it."""
        write(template_dirs["local_custom"], "custom_email_template.html", "CUSTOM")
        write(template_dirs["local_custom"], "broadcast.html", "NAMED")

        resolved = Path(template_utils.find_email_template("broadcast.html"))

        assert resolved.read_text(encoding="utf-8") == "NAMED"

    def test_an_explicit_filename_that_is_missing_resolves_to_none(self, template_dirs):
        write(template_dirs["local_custom"], "custom_email_template.html", "CUSTOM")

        assert template_utils.find_email_template("absent.html") is None


class TestTheSameRulesServeEveryTool:
    """`find_docx_template` and `find_pptx_templates` share the resolution,
    so a change to it that spared email would still be caught here."""

    def test_word_prefers_the_custom_document(self, template_dirs):
        write(template_dirs["local_custom"], "custom_docx_template.docx", "CUSTOM")
        write(template_dirs["local_default"], "default_docx_template.docx", "DEFAULT")

        resolved = Path(template_utils.find_docx_template())

        assert resolved.read_text(encoding="utf-8") == "CUSTOM"

    def test_powerpoint_resolves_each_aspect_independently(self, template_dirs):
        write(template_dirs["local_custom"], "custom_pptx_template_4_3.pptx", "CUSTOM 4:3")
        write(template_dirs["local_default"], "default_pptx_template_4_3.pptx", "DEFAULT 4:3")
        write(template_dirs["local_default"], "default_pptx_template_16_9.pptx", "DEFAULT 16:9")

        four_three, sixteen_nine = template_utils.find_pptx_templates()

        assert Path(four_three).read_text(encoding="utf-8") == "CUSTOM 4:3"
        assert Path(sixteen_nine).read_text(encoding="utf-8") == "DEFAULT 16:9"
