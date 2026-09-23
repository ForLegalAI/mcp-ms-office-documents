"""NaN and Infinity are refused by the schema, with the field's path.

Regression from the PowerPoint review: pydantic accepts "NaN", "inf" and
1e999 for a float, and the build then failed deep inside — the chart workbook
("NAN/INF not supported in write_number()") or an int() of an EMU length — so
the caller lost the whole deck to a message naming no field. Chart values,
scatter points, table widths and blank-slide positions now reject them at
validation.
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from pptx_tools.schema import coerce_slides

BAD = ["NaN", "inf", "-Infinity", 1e999, float("nan")]


def _rejected(slide):
    with pytest.raises(ValueError) as exc:
        coerce_slides([slide])
    return str(exc.value)


@pytest.mark.parametrize("bad", BAD)
def test_chart_value(bad):
    message = _rejected({"type": "chart", "title": "C", "chart_type": "column",
                         "categories": ["a", "b"],
                         "series": [{"name": "s", "values": [1, bad]}]})
    assert "series.0.values.1" in message
    assert "finite" in message


@pytest.mark.parametrize("bad", BAD)
def test_scatter_point(bad):
    message = _rejected({"type": "scatter", "title": "S",
                         "series": [{"name": "s", "points": [[1, 2], [bad, 3]]}]})
    assert "series.0.points.1.0" in message


@pytest.mark.parametrize("bad", ["inf", 1e999])
def test_table_width(bad):
    message = _rejected({"type": "table", "rows": [["a", "b"], ["1", "2"]],
                         "widths": [bad, 1]})
    assert "widths.0" in message


@pytest.mark.parametrize("bad", [1e999, float("nan"), float("-inf")])
def test_blank_element_position(bad):
    message = _rejected({"type": "blank", "elements": [
        {"kind": "text", "text": "x", "x": bad, "y": 1, "w": 2, "h": 1}]})
    assert "elements.0" in message
    assert "finite" in message


def test_a_gap_is_still_null_and_ordinary_numbers_still_pass():
    slides = coerce_slides([{"type": "chart", "title": "C", "chart_type": "line",
                             "categories": ["a", "b", "c"],
                             "series": [{"name": "s", "values": [1, None, -2.5]}]}])
    assert slides[0].series[0].values == [1.0, None, -2.5]
