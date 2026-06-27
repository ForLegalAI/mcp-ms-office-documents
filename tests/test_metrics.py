"""Tests for the in-process metrics module."""
import logging
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

import metrics


@pytest.fixture(autouse=True)
def _clean_metrics():
    metrics.install_log_capture()
    metrics.reset()
    yield
    metrics.reset()


def test_record_call_and_error_counters():
    metrics.record_call("docx", "t1")
    metrics.record_call("docx", "t1")
    metrics.record_error("docx", "t1", "boom")
    st = metrics.get_tool_stat("t1")
    assert st.calls == 2
    assert st.errors == 1
    assert st.last_error == "boom"
    assert st.last_called is not None
    assert st.kind == "docx"


def test_tool_stats_sorted_by_recency():
    metrics.record_call("docx", "older")
    metrics.record_call("email", "newer")
    names = [s.name for s in metrics.tool_stats()]
    assert names[0] == "newer"  # most-recently-used first


def test_error_message_truncated():
    metrics.record_error("docx", "t", "x" * 1000)
    assert len(metrics.get_tool_stat("t").last_error) == 500


def test_recent_logs_level_filter():
    log = logging.getLogger("metrics-test")
    log.info("an info line")
    log.error("an error line")
    infos = metrics.recent_logs(logging.INFO)
    errs = metrics.recent_logs(logging.ERROR)
    assert any("an error line" in r["message"] for r in errs)
    assert all(r["levelno"] >= logging.ERROR for r in errs)
    # INFO view is a superset and newest-first.
    assert len(infos) >= len(errs)
    assert infos[0]["message"] in ("an error line", "an info line")


def test_counts_by_level():
    log = logging.getLogger("metrics-test")
    log.warning("w1")
    log.error("e1")
    counts = metrics.counts_by_level()
    assert counts.get("WARNING", 0) >= 1
    assert counts.get("ERROR", 0) >= 1


def test_reset_clears_state():
    metrics.record_call("docx", "t")
    logging.getLogger("metrics-test").error("e")
    metrics.reset()
    assert metrics.tool_stats() == []
    assert metrics.recent_logs(logging.INFO) == []
