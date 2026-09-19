"""Warnings a finished build reported must reach the Status page.

`warning_channel.py` exists because a build that succeeds while substituting a
style or dropping a row returns a success the caller never questions, and the
detail must not live only in the server log. Until #173, `metrics` counted
calls and errors only — so on the Status page that degraded build looked
exactly like a clean one.
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

import metrics
from warning_channel import DocumentWarning, WarningChannel

SEVERITIES = {"dropped": "error", "renamed": "warning", "swapped": "info"}


@pytest.fixture(autouse=True)
def clean_metrics():
    metrics.reset()
    yield
    metrics.reset()


def _channel(*pairs):
    channel = WarningChannel(SEVERITIES)
    for code, message in pairs:
        channel.add(code, message)
    return channel


def test_a_clean_build_records_no_warnings():
    metrics.record_warnings("docx", "create_word_document", _channel())
    assert metrics.get_tool_stat("create_word_document") is None, (
        "a tool with nothing to report should not appear in the counters"
    )


def test_warnings_are_counted_per_severity():
    metrics.record_warnings("docx", "create_word_document", _channel(
        ("dropped", "a table would not render"),
        ("renamed", "used Heading 2 instead"),
        ("swapped", "substituted a font"),
    ))
    st = metrics.get_tool_stat("create_word_document")
    assert st.warnings_by_severity == {"error": 1, "warning": 1, "info": 1}
    assert st.warnings == 3


def test_degraded_excludes_info():
    """`info` is a substitution nobody needs to chase; it must not inflate the count."""
    metrics.record_warnings("xlsx", "create_excel_document", _channel(
        ("swapped", "substituted a font"),
        ("swapped", "substituted another font"),
    ))
    st = metrics.get_tool_stat("create_excel_document")
    assert st.warnings == 2
    assert st.degraded == 0


def test_counts_accumulate_across_builds():
    for i in range(3):
        metrics.record_warnings("docx", "create_word_document",
                                _channel(("dropped", f"block {i}")))
    st = metrics.get_tool_stat("create_word_document")
    assert st.warnings_by_severity == {"error": 3}


def test_recent_warnings_are_kept_newest_first_and_bounded():
    for i in range(metrics.LAST_WARNINGS_KEPT + 5):
        metrics.record_warnings("docx", "create_word_document",
                                _channel(("dropped", f"block {i}")))
    st = metrics.get_tool_stat("create_word_document")
    assert len(st.last_warnings) == metrics.LAST_WARNINGS_KEPT
    assert "block 14" in st.last_warnings[0], "newest build should be first"


def test_location_is_kept_in_the_rendered_line():
    """A warning without its location is far less actionable."""
    channel = WarningChannel(SEVERITIES)
    channel.add("renamed", "used the default style", sheet="Data", cell="B2")
    metrics.record_warnings("xlsx", "create_excel_document", channel)
    assert "sheet Data, cell B2" in metrics.get_tool_stat(
        "create_excel_document").last_warnings[0]


def test_a_plain_list_of_warnings_is_accepted():
    """PowerPoint keeps its own records rather than a WarningChannel."""
    metrics.record_warnings("pptx", "create_powerpoint_presentation", [
        DocumentWarning(code="dropped", message="slide 2 lost its picture",
                        severity="error"),
    ])
    st = metrics.get_tool_stat("create_powerpoint_presentation")
    assert st.warnings_by_severity == {"error": 1}


def test_an_unknown_severity_is_still_counted():
    """A code with no severity must not vanish from the totals."""
    metrics.record_warnings("docx", "create_word_document", [
        DocumentWarning(code="odd", message="something", severity="curious"),
    ])
    st = metrics.get_tool_stat("create_word_document")
    assert st.warnings_by_severity == {"curious": 1}
    assert st.degraded == 1, "an unrecognised severity should not be treated as info"


def test_degraded_total_spans_tools_and_still_excludes_info():
    metrics.record_warnings("docx", "create_word_document",
                            _channel(("dropped", "a block")))
    metrics.record_warnings("xlsx", "create_excel_document",
                            _channel(("swapped", "a font")))
    assert metrics.degraded_total() == 1, "the info substitution must not count"


def test_warnings_do_not_disturb_call_and_error_counters():
    metrics.record_call("docx", "create_word_document")
    metrics.record_warnings("docx", "create_word_document",
                            _channel(("dropped", "a block")))
    st = metrics.get_tool_stat("create_word_document")
    assert st.calls == 1 and st.errors == 0 and st.warnings == 1


def test_reset_clears_warnings():
    metrics.record_warnings("docx", "create_word_document",
                            _channel(("dropped", "a block")))
    metrics.reset()
    assert metrics.degraded_total() == 0
    assert metrics.get_tool_stat("create_word_document") is None


# ---------------------------------------------------------------------------
# The wiring in main.py
# ---------------------------------------------------------------------------


def test_with_warnings_records_the_call_and_its_warnings():
    """_with_warnings is the one point every channel-carrying tool passes.

    Recording there rather than in each handler is what stops a new tool
    being silently uncounted — so this pins that it actually records.
    """
    import main

    channel = _channel(("dropped", "a block"))
    result = main._with_warnings("https://example/f.docx", channel,
                                 "docx", "create_word_document")
    assert isinstance(result, dict), "warnings widen the bare URL into a dict"
    st = metrics.get_tool_stat("create_word_document")
    assert st.calls == 1
    assert st.warnings_by_severity == {"error": 1}


def test_with_warnings_counts_a_clean_build_as_a_call():
    """A build with nothing to report is still a call, and still returns a bare URL."""
    import main

    result = main._with_warnings("https://example/f.xlsx", _channel(),
                                 "xlsx", "create_excel_document")
    assert result == "https://example/f.xlsx", "an unchanged response shape"
    st = metrics.get_tool_stat("create_excel_document")
    assert st.calls == 1 and st.warnings == 0


def test_with_warnings_requires_the_tool_identity():
    """Omitting kind/name must fail loudly rather than go uncounted."""
    import main

    with pytest.raises(TypeError):
        main._with_warnings("https://example/f.docx", _channel())


def test_every_with_warnings_call_site_passes_an_identity():
    """Guard the guard: a new call site cannot quietly skip the counters.

    Parsed rather than matched. A regex over the source truncates a call at
    the first `)`, which for the PowerPoint site is the one inside
    `len(slides)` — so the identity arguments could fall outside the captured
    text and this would stop enforcing anything without failing.
    """
    import ast

    tree = ast.parse((project_root / "main.py").read_text(encoding="utf-8"))
    sites = [
        node for node in ast.walk(tree)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Name)
        and node.func.id == "_with_warnings"
    ]
    assert sites, "no _with_warnings call sites found; this test is stale"
    for site in sites:
        positional = site.args
        assert len(positional) >= 4, (
            f"line {site.lineno}: _with_warnings needs result, warnings, kind, name"
        )
        kind, name = positional[2], positional[3]
        for arg, label in ((kind, "kind"), (name, "name")):
            assert isinstance(arg, ast.Constant) and isinstance(arg.value, str), (
                f"line {site.lineno}: {label} must be a string literal, "
                f"got {ast.dump(arg)}"
            )
            assert arg.value, f"line {site.lineno}: {label} must not be empty"
