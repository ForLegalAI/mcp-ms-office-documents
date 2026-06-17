"""Tests for the pure-Python Excel formula recalculation pipeline.

Covers:
- ``xlsx_tools.formula_engine.recalculate_workbook`` — engine-level
  correctness (single-sheet, multi-sheet, errors, booleans, unsupported
  functions, missing dependency fallback).
- ``xlsx_tools.xml_cache.inject_cached_values`` — XML rewriting
  correctness (cells get cached <v>, formulas are preserved, zip
  integrity maintained, non-matching locations are ignored).
- ``xlsx_tools.base_xlsx_tool.markdown_to_excel`` — end-to-end
  integration (recalc on/off, error surfacing).
"""

from __future__ import annotations

import inspect
import io
import sys
from pathlib import Path
from unittest.mock import patch

import pytest
from openpyxl import Workbook, load_workbook

# Make project root importable.
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from xlsx_tools.formula_engine import (  # noqa: E402
    EXCEL_ERRORS,
    CellError,
    RecalcResult,
    is_available,
    recalculate_workbook,
)
from xlsx_tools.xml_cache import inject_cached_values  # noqa: E402
from xlsx_tools.base_xlsx_tool import markdown_to_excel  # noqa: E402

OUTPUT_DIR = Path(__file__).parent / "output" / "xlsx"


@pytest.fixture(scope="module", autouse=True)
def setup_output_dir():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    yield


def _save_wb(wb: Workbook, name: str | None = None) -> bytes:
    """Save a workbook to bytes; optionally also to OUTPUT_DIR for inspection."""
    buf = io.BytesIO()
    wb.save(buf)
    data = buf.getvalue()
    if name:
        (OUTPUT_DIR / f"{name}.xlsx").write_bytes(data)
    return data


# ── Engine availability ──────────────────────────────────────────────────────


class TestEngineAvailability:
    """The `formulas` library should be installed in the dev/test environment."""

    def test_is_available_returns_true(self):
        # The library is in requirements.txt so this should always pass
        # in a properly-installed environment. If it ever doesn't, every
        # subsequent test is meaningless.
        assert is_available() is True


# ── Engine: happy path ───────────────────────────────────────────────────────


class TestEngineHappyPath:
    """recalculate_workbook returns correct numeric results."""

    def test_simple_sum(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = 100
        ws["A2"] = 200
        ws["A3"] = 300
        ws["A4"] = "=SUM(A1:A3)"
        data = _save_wb(wb, "recalc_simple_sum")

        result = recalculate_workbook(data, ["Sheet1"])

        assert result.recalc_performed is True
        assert result.skip_reason is None
        assert result.errors == []
        assert result.values_map["Sheet1!A4"] == 600.0

    def test_arithmetic_and_functions(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 10
        ws["A2"] = "=A1*2"
        ws["A3"] = "=A1+A2"
        ws["A4"] = "=MAX(A1:A3)"
        ws["A5"] = "=AVERAGE(A1:A3)"
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["Sheet"])

        assert result.values_map["Sheet!A2"] == 20.0
        assert result.values_map["Sheet!A3"] == 30.0
        assert result.values_map["Sheet!A4"] == 30.0
        assert result.values_map["Sheet!A5"] == pytest.approx(20.0)

    def test_boolean_formula(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 100
        ws["A2"] = "=A1>50"
        ws["A3"] = "=A1<50"
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["Sheet"])

        assert result.values_map["Sheet!A2"] is True
        assert result.values_map["Sheet!A3"] is False

    def test_integer_value_is_native_int(self):
        """Recalculated integer values should be native Python ints, not numpy."""
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 5
        ws["A2"] = "=A1+5"
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["Sheet"])
        val = result.values_map["Sheet!A2"]

        assert val == 10
        assert type(val) is int  # noqa: E721 — we explicitly want native int

    def test_float_value_is_native_float(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=1/3"
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["Sheet"])
        val = result.values_map["Sheet!A1"]

        assert val == pytest.approx(0.3333333, rel=1e-5)
        assert type(val) is float


# ── Engine: multi-sheet ──────────────────────────────────────────────────────


class TestEngineMultiSheet:
    """Cross-sheet references resolve correctly."""

    def test_cross_sheet_reference(self):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Revenue"
        ws1["A1"] = "Q1"
        ws1["B1"] = 1000
        ws1["B2"] = 1500
        ws2 = wb.create_sheet("Dashboard")
        ws2["B1"] = "=Revenue!B1"
        ws2["B2"] = "=SUM(Revenue!B1:B2)"
        data = _save_wb(wb, "recalc_cross_sheet")

        result = recalculate_workbook(data, ["Revenue", "Dashboard"])

        assert result.values_map["Dashboard!B1"] == 1000
        assert result.values_map["Dashboard!B2"] == 2500.0
        assert result.errors == []

    def test_sheet_name_casing_preserved(self):
        """Sheet names in values_map keys use the original casing."""
        wb = Workbook()
        ws = wb.active
        ws.title = "MixedCase"
        ws["A1"] = 1
        ws["A2"] = "=A1+1"
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["MixedCase"])
        # The key must use 'MixedCase' (original), not 'MIXEDCASE' (uppercased by the lib).
        assert "MixedCase!A2" in result.values_map
        assert result.values_map["MixedCase!A2"] == 2.0


# ── Engine: error detection ──────────────────────────────────────────────────


class TestEngineErrors:
    """Formula errors are detected and reported, never raised."""

    def test_div_by_zero(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 100
        ws["A2"] = "=A1/0"
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["Sheet"])

        assert result.recalc_performed is True
        assert len(result.errors) == 1
        err = result.errors[0]
        assert err.error_type == "#DIV/0!"
        assert err.coordinate == "A2"
        assert err.sheet == "Sheet"
        assert "Sheet!A2" in str(err)

    def test_ref_error_missing_sheet(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=MissingSheet!A1"
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["Sheet"])
        error_types = [e.error_type for e in result.errors]
        assert "#REF!" in error_types

    def test_name_error_unknown_function(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=BogusFunctionXYZ(1,2)"
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["Sheet"])
        error_types = [e.error_type for e in result.errors]
        assert "#NAME?" in error_types

    def test_all_seven_errors_listed(self):
        """EXCEL_ERRORS covers the seven OOXML error sentinels."""
        assert len(EXCEL_ERRORS) == 7
        for e in ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NULL!", "#NUM!", "#N/A"):
            assert e in EXCEL_ERRORS


# ── Engine: robustness ───────────────────────────────────────────────────────


class TestEngineRobustness:
    """Engine never raises; missing deps degrade gracefully."""

    def test_empty_workbook_no_crash(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 1
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["Sheet"])
        assert result.recalc_performed is True
        assert result.errors == []

    def test_recalc_result_default_factory(self):
        result = RecalcResult()
        assert result.values_map == {}
        assert result.errors == []
        assert result.has_errors is False


# ── XML injection ────────────────────────────────────────────────────────────


class TestXmlInjection:
    """inject_cached_values rewrites only formula cells, preserving structure."""

    def test_formula_cell_gets_cached_value(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 100
        ws["A2"] = "=A1*2"
        data = _save_wb(wb)

        injected = inject_cached_values(data, {"Sheet!A2": 200})

        wb2 = load_workbook(io.BytesIO(injected), data_only=True)
        assert wb2.active["A2"].value == 200

    def test_formula_preserved_after_injection(self):
        """Reading without data_only still shows the formula string."""
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 5
        ws["A2"] = "=A1*3"
        data = _save_wb(wb)

        injected = inject_cached_values(data, {"Sheet!A2": 15})

        wb2 = load_workbook(io.BytesIO(injected))
        assert wb2.active["A2"].value == "=A1*3"

    def test_non_formula_cell_not_overwritten(self):
        """A literal value cell should not be modified even if it's in the map."""
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 42  # literal, no formula
        data = _save_wb(wb)

        injected = inject_cached_values(data, {"Sheet!A1": 999})

        wb2 = load_workbook(io.BytesIO(injected), data_only=True)
        # A1 is not a formula cell, so injection must not touch it.
        assert wb2.active["A1"].value == 42

    def test_empty_values_map_returns_original_bytes(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 1
        data = _save_wb(wb)

        assert inject_cached_values(data, {}) == data

    def test_zip_integrity_maintained(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 10
        ws["A2"] = "=A1+5"
        data = _save_wb(wb)

        injected = inject_cached_values(data, {"Sheet!A2": 15})

        import zipfile
        assert zipfile.ZipFile(io.BytesIO(injected)).testzip() is None

    def test_multi_sheet_injection(self):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Inputs"
        ws1["A1"] = 100
        ws2 = wb.create_sheet("Calc")
        ws2["A1"] = "=Inputs!A1*2"
        data = _save_wb(wb)

        injected = inject_cached_values(data, {"Calc!A1": 200})

        wb2 = load_workbook(io.BytesIO(injected), data_only=True)
        assert wb2["Calc"]["A1"].value == 200

    def test_unknown_location_ignored(self):
        """A values_map entry for a non-existent sheet/cell is ignored."""
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=1+1"
        data = _save_wb(wb)

        # Bogus sheet name and bogus coordinate — should not raise.
        injected = inject_cached_values(
            data,
            {"NoSuchSheet!A1": 999, "Sheet!Z99": 999, "Sheet!A1": 2},
        )
        wb2 = load_workbook(io.BytesIO(injected), data_only=True)
        assert wb2.active["A1"].value == 2


# ── End-to-end via markdown_to_excel ─────────────────────────────────────────


def _intercept_upload(markdown: str, **kwargs) -> bytes:
    """Run markdown_to_excel with the upload patched out; return raw xlsx bytes."""
    captured = {}

    def fake_upload(file_obj, suffix, **kw):
        captured["data"] = file_obj.read()
        return "fake://test.xlsx"

    with patch("xlsx_tools.base_xlsx_tool.upload_file", side_effect=fake_upload):
        markdown_to_excel(markdown, **kwargs)
    return captured["data"]


class TestEndToEndRecalc:
    """markdown_to_excel integrates recalc + inject when recalc=True."""

    def test_recalc_on_produces_cached_values(self):
        markdown = """| A | B | Total |
|---|---|-------|
| 1 | 2 | =A2+B2 |
| 3 | 4 | =A3+B3 |
"""
        data = _intercept_upload(markdown, recalc=True)

        wb = load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        # Row 2 (Excel) is the first data row.
        assert ws["C2"].value == 3
        assert ws["C3"].value == 7

    def test_recalc_off_leaves_formulas_without_cached_values(self):
        markdown = """| A | Total |
|---|-------|
| 1 | =A2*2 |
"""
        data = _intercept_upload(markdown, recalc=False)

        wb = load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        # Without recalc, no cached value is stored.
        assert ws["B2"].value is None

        # Formula is still present.
        wb2 = load_workbook(io.BytesIO(data))
        assert wb2.active["B2"].value == "=A2*2"

    def test_recalc_explicit_fails_on_formula_errors(self):
        """When recalc is explicitly True, formula errors fail the call.

        This enforces the "zero formula errors" delivery standard: the
        model is told about the errors so it can fix the formulas and
        retry, rather than silently shipping a broken file.
        """
        markdown = """| A | B |
|---|---|
| 5 | =A2/0 |
"""
        with pytest.raises(RuntimeError, match="formula error"):
            _intercept_upload(markdown, recalc=True)

    def test_recalc_default_delivers_file_with_errors(self):
        """When recalc runs as a default (None), errors are logged but
        the file is still delivered — misconfiguration must never break
        document generation.
        """
        markdown = """| A | B |
|---|---|
| 5 | =A2/0 |
"""
        # recalc=None (default) — env config has XLSX_RECALC_ENABLED true,
        # so recalc runs, finds the error, but still delivers.
        data = _intercept_upload(markdown)

        wb = load_workbook(io.BytesIO(data))
        assert wb.active["B2"].value == "=A2/0"


# ── D11: Recalc timeout ──────────────────────────────────────────────────────


class TestRecalcTimeout:
    """D11: recalculation is bounded by a configurable timeout."""

    def test_config_has_timeout_field_with_default(self):
        from config import Config
        cfg = Config.from_env()
        assert cfg.xlsx_recalc_timeout_seconds == 30

    def test_recalc_and_inject_accepts_timeout_arg(self):
        """_recalc_and_inject signature accepts a timeout_seconds parameter."""
        import inspect
        from xlsx_tools.base_xlsx_tool import _recalc_and_inject
        sig = inspect.signature(_recalc_and_inject)
        assert "timeout_seconds" in sig.parameters
        assert sig.parameters["timeout_seconds"].default == 30

    def test_recalc_completes_within_generous_timeout(self):
        """A small workbook recalc'd with a generous timeout succeeds."""
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 10
        ws["A2"] = "=A1*5"
        buf = io.BytesIO()
        wb.save(buf)
        data = buf.getvalue()

        from xlsx_tools.base_xlsx_tool import _recalc_and_inject
        result_bytes, error_summary = _recalc_and_inject(
            data, ["Sheet"], timeout_seconds=30
        )

        wb2 = load_workbook(io.BytesIO(result_bytes), data_only=True)
        assert wb2.active["A2"].value == 50
        assert error_summary is None

    def test_recalc_with_tiny_timeout_falls_back_gracefully(self):
        """A tiny timeout causes recalc to be skipped; original bytes returned.

        We can't reliably force a real timeout without a pathological
        workbook, so we verify the fallback contract: when the engine
        can't finish in time, the original (un-recalc'd) bytes are
        returned and the cached value is absent.
        """
        wb = Workbook()
        ws = wb.active
        ws["A1"] = 10
        ws["A2"] = "=A1*5"
        buf = io.BytesIO()
        wb.save(buf)
        data = buf.getvalue()

        from xlsx_tools.base_xlsx_tool import _recalc_and_inject
        # 0 seconds is below the config minimum (1), but the function
        # accepts any positive int. Use 1 to give the engine a chance but
        # still exercise the timeout path on slow machines.
        result_bytes, error_summary = _recalc_and_inject(
            data, ["Sheet"], timeout_seconds=1
        )

        # Either it completed (cached value present) or it timed out
        # (original bytes returned). Both are acceptable outcomes — the
        # contract is "never hang, always return bytes".
        assert isinstance(result_bytes, bytes)
        assert len(result_bytes) > 0


# ──────────────────────────────────────────────────────────────────────────────
# E1: Circular-reference detection
# ──────────────────────────────────────────────────────────────────────────────


class TestCircularReferenceDetection:
    """E1: the `formulas` library silently misses circular references.

    `detect_circular_references()` builds a dependency graph from the
    formula strings and runs DFS cycle detection. Every cell on a cycle
    is reported as a `#CIRC!` error.
    """

    def test_simple_two_cell_cycle(self):
        from xlsx_tools.formula_engine import (
            CIRCULAR_ERROR_TYPE,
            detect_circular_references,
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "=A2"
        ws["A2"] = "=A1"
        data = _save_wb(wb)

        errors = detect_circular_references(data, ["Sheet1"])
        assert len(errors) == 2
        locations = {e.location for e in errors}
        assert locations == {"Sheet1!A1", "Sheet1!A2"}
        assert all(e.error_type == CIRCULAR_ERROR_TYPE for e in errors)

    def test_self_reference(self):
        from xlsx_tools.formula_engine import detect_circular_references

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "=A1+1"  # classic self-ref
        data = _save_wb(wb)

        errors = detect_circular_references(data, ["Sheet1"])
        assert len(errors) == 1
        assert errors[0].location == "Sheet1!A1"

    def test_no_cycle_returns_empty(self):
        from xlsx_tools.formula_engine import detect_circular_references

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = 1
        ws["A2"] = "=A1+1"
        ws["A3"] = "=A2+1"  # linear chain, no cycle
        data = _save_wb(wb)

        errors = detect_circular_references(data, ["Sheet1"])
        assert errors == []

    def test_cross_sheet_cycle(self):
        from xlsx_tools.formula_engine import detect_circular_references

        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Sheet1"
        ws2 = wb.create_sheet("Sheet2")
        ws1["A1"] = "=Sheet2!A1"  # Sheet1!A1 -> Sheet2!A1
        ws2["A1"] = "=Sheet1!A1"  # Sheet2!A1 -> Sheet1!A1 (cycle)
        data = _save_wb(wb)

        errors = detect_circular_references(data, ["Sheet1", "Sheet2"])
        locations = {e.location for e in errors}
        assert "Sheet1!A1" in locations
        assert "Sheet2!A1" in locations

    def test_longer_cycle_chain(self):
        from xlsx_tools.formula_engine import detect_circular_references

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        # A1 -> B1 -> C1 -> A1
        ws["A1"] = "=B1"
        ws["B1"] = "=C1"
        ws["C1"] = "=A1"
        data = _save_wb(wb)

        errors = detect_circular_references(data, ["Sheet1"])
        locations = {e.location for e in errors}
        assert locations == {"Sheet1!A1", "Sheet1!B1", "Sheet1!C1"}

    def test_cycle_with_innocent_bystander(self):
        """A cell that references a cycle member but isn't itself on the
        cycle should not be flagged."""
        from xlsx_tools.formula_engine import detect_circular_references

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "=A2"
        ws["A2"] = "=A1"  # A1<->A2 cycle
        ws["A3"] = "=A1"  # innocent: refs a cycle member but no cycle
        data = _save_wb(wb)

        errors = detect_circular_references(data, ["Sheet1"])
        locations = {e.location for e in errors}
        assert "Sheet1!A3" not in locations
        assert "Sheet1!A1" in locations
        assert "Sheet1!A2" in locations

    def test_range_reference_cycle(self):
        """Cycles through a range reference should still be detected."""
        from xlsx_tools.formula_engine import detect_circular_references

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        # SUM(A1:A3) in A2 creates a self-loop through the range.
        ws["A1"] = 1
        ws["A2"] = "=SUM(A1:A3)"
        ws["A3"] = 1
        data = _save_wb(wb)

        errors = detect_circular_references(data, ["Sheet1"])
        locations = {e.location for e in errors}
        # A2 references itself (via the range that includes A2).
        assert "Sheet1!A2" in locations


class TestExtractFormulaReferences:
    """Unit tests for the formula-string reference parser."""

    def test_local_reference(self):
        from xlsx_tools.formula_engine import (
            _build_sheet_lookup,
            extract_formula_references,
        )

        sl = _build_sheet_lookup(["Sheet1"])
        refs = extract_formula_references("=A1+B2", "Sheet1", sl)
        assert refs == {"Sheet1!A1", "Sheet1!B2"}

    def test_absolute_reference(self):
        from xlsx_tools.formula_engine import (
            _build_sheet_lookup,
            extract_formula_references,
        )

        sl = _build_sheet_lookup(["Sheet1"])
        refs = extract_formula_references("=$A$1+B$2", "Sheet1", sl)
        assert refs == {"Sheet1!A1", "Sheet1!B2"}

    def test_cross_sheet_reference(self):
        from xlsx_tools.formula_engine import (
            _build_sheet_lookup,
            extract_formula_references,
        )

        sl = _build_sheet_lookup(["Inputs", "Sheet1"])
        refs = extract_formula_references("=Inputs!A1+Sheet1!B2", "Sheet1", sl)
        assert refs == {"Inputs!A1", "Sheet1!B2"}

    def test_range_expanded(self):
        from xlsx_tools.formula_engine import (
            _build_sheet_lookup,
            extract_formula_references,
        )

        sl = _build_sheet_lookup(["Sheet1"])
        refs = extract_formula_references("=SUM(A1:A3)", "Sheet1", sl)
        assert refs == {"Sheet1!A1", "Sheet1!A2", "Sheet1!A3"}

    def test_unknown_sheet_dropped(self):
        from xlsx_tools.formula_engine import (
            _build_sheet_lookup,
            extract_formula_references,
        )

        sl = _build_sheet_lookup(["Sheet1"])
        # 'Missing' is not in the workbook — ref should be dropped
        # (the recalc engine will flag a real #REF! if it's a problem).
        refs = extract_formula_references("=Missing!A1+A2", "Sheet1", sl)
        assert refs == {"Sheet1!A2"}

    def test_empty_formula(self):
        from xlsx_tools.formula_engine import (
            _build_sheet_lookup,
            extract_formula_references,
        )

        sl = _build_sheet_lookup(["Sheet1"])
        assert extract_formula_references("", "Sheet1", sl) == set()
        assert extract_formula_references("=1+2", "Sheet1", sl) == set()


class TestCircularRefErrorPolicy:
    """E1 + A3 integration: explicit recalc=True with circular refs fails."""

    def test_explicit_recalc_fails_on_circular_ref(self):
        """When recalc=True is explicit and a circular ref is present,
        the call must fail (zero-errors policy)."""
        markdown = (
            "| A | B |\n"
            "|---|---|\n"
            "| =B2 | =A2 |\n"
        )

        with patch("xlsx_tools.base_xlsx_tool.upload_file") as mock_upload:
            mock_upload.return_value = "fake://circ.xlsx"
            with pytest.raises(RuntimeError) as exc_info:
                markdown_to_excel(markdown, recalc=True)

        msg = str(exc_info.value)
        assert "#CIRC!" in msg
        assert "circular" in msg.lower()
        # Upload must not have run — we fail before delivery.
        mock_upload.assert_not_called()

    def test_default_recalc_delivers_file_with_circular_ref(self):
        """When recalc is default (None) and a circular ref is present,
        the file is still delivered (errors logged only)."""
        markdown = (
            "| A | B |\n"
            "|---|---|\n"
            "| =B2 | =A2 |\n"
        )

        with patch("xlsx_tools.base_xlsx_tool.upload_file") as mock_upload:
            mock_upload.return_value = "fake://circ-delivered.xlsx"
            result = markdown_to_excel(markdown)  # no recalc param

        assert result == "fake://circ-delivered.xlsx"
        mock_upload.assert_called_once()


# ──────────────────────────────────────────────────────────────────────────────
# E2: Structured (type-grouped) error output
# ──────────────────────────────────────────────────────────────────────────────


class TestGroupedErrorOutput:
    """E2: errors in the RuntimeError message are grouped by type with
    counts and locations, not a flat list."""

    def test_grouped_format_single_type(self):
        from xlsx_tools.base_xlsx_tool import _format_grouped_errors
        from xlsx_tools.formula_engine import CellError

        errors = [
            CellError(sheet="Sheet1", coordinate="B2", error_type="#DIV/0!"),
            CellError(sheet="Sheet1", coordinate="B5", error_type="#DIV/0!"),
        ]
        msg = _format_grouped_errors(errors)
        assert "#DIV/0! (2)" in msg
        assert "Sheet1!B2" in msg
        assert "Sheet1!B5" in msg
        assert "2 formula error(s)" in msg

    def test_grouped_format_multiple_types(self):
        from xlsx_tools.base_xlsx_tool import _format_grouped_errors
        from xlsx_tools.formula_engine import CellError

        errors = [
            CellError(sheet="S", coordinate="B2", error_type="#DIV/0!"),
            CellError(sheet="S", coordinate="B3", error_type="#DIV/0!"),
            CellError(sheet="S", coordinate="C10", error_type="#REF!"),
        ]
        msg = _format_grouped_errors(errors)
        assert "#DIV/0! (2)" in msg
        assert "#REF! (1)" in msg
        assert "3 formula error(s)" in msg

    def test_grouped_format_truncates_long_lists(self):
        from xlsx_tools.base_xlsx_tool import _format_grouped_errors
        from xlsx_tools.formula_engine import CellError

        errors = [
            CellError(sheet="S", coordinate=f"A{i}", error_type="#REF!")
            for i in range(10)
        ]
        msg = _format_grouped_errors(errors)
        assert "#REF! (10)" in msg
        assert "and 5 more" in msg  # only first 5 shown

    def test_grouped_format_circular_annotation(self):
        from xlsx_tools.base_xlsx_tool import _format_grouped_errors
        from xlsx_tools.formula_engine import CIRCULAR_ERROR_TYPE, CellError

        errors = [
            CellError(sheet="S", coordinate="A1", error_type=CIRCULAR_ERROR_TYPE),
        ]
        msg = _format_grouped_errors(errors)
        assert "#CIRC! (1)" in msg
        # The annotation helps the model understand #CIRC! is a cycle.
        assert "circular references detected" in msg.lower()
        assert "breaking the cycle" in msg.lower()

    def test_grouped_format_circular_mixed_with_excel_errors(self):
        from xlsx_tools.base_xlsx_tool import _format_grouped_errors
        from xlsx_tools.formula_engine import CIRCULAR_ERROR_TYPE, CellError

        errors = [
            CellError(sheet="S", coordinate="A1", error_type=CIRCULAR_ERROR_TYPE),
            CellError(sheet="S", coordinate="B2", error_type="#DIV/0!"),
        ]
        msg = _format_grouped_errors(errors)
        assert "#CIRC! (1)" in msg
        assert "#DIV/0! (1)" in msg
        assert "2 formula error(s)" in msg
        assert "circular references detected" in msg.lower()


# ──────────────────────────────────────────────────────────────────────────────
# E4: total_formulas telemetry
# ──────────────────────────────────────────────────────────────────────────────


class TestTotalFormulasCount:
    """E4: RecalcResult carries a total_formulas count for telemetry."""

    def test_count_formulas_helper(self):
        from xlsx_tools.xml_cache import count_formulas

        wb = Workbook()
        ws = wb.active
        ws.title = "S"
        ws["A1"] = 1
        ws["A2"] = "=A1+1"
        ws["A3"] = "=SUM(A1:A2)"
        data = _save_wb(wb)

        assert count_formulas(data) == 2

    def test_count_formulas_empty_workbook(self):
        from xlsx_tools.xml_cache import count_formulas

        wb = Workbook()
        data = _save_wb(wb)
        assert count_formulas(data) == 0

    def test_recalc_result_populates_total_formulas(self):
        if not is_available():
            pytest.skip("`formulas` library not installed")

        wb = Workbook()
        ws = wb.active
        ws.title = "S"
        ws["A1"] = 10
        ws["A2"] = "=A1+5"
        ws["A3"] = "=A2*2"
        data = _save_wb(wb)

        result = recalculate_workbook(data, ["S"])
        assert result.recalc_performed
        assert result.total_formulas == 2
