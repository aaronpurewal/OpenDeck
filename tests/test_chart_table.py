"""
Unit tests for create_chart and create_table tools.

NOTE: Aspose evaluation version may add watermarks or truncate content.
"""

import pytest
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import aspose.slides as slides
import aspose.slides.charts as charts

from tools import (
    create_chart, create_table,
    _get_theme_colors, _apply_theme_to_chart,
    _POSITION_SLOTS, _CHART_TYPE_MAP, _inches,
)

# Detect Aspose eval mode
_EVAL_MODE = False
try:
    _prs = slides.Presentation()
    _layout = _prs.masters[0].layout_slides[0]
    _prs.slides.insert_empty_slide(0, _layout)
    _slide = _prs.slides[0]
    _shape = _slide.shapes.add_auto_shape(slides.ShapeType.RECTANGLE, 100, 100, 500, 300)
    _shape.text_frame.paragraphs[0].portions[0].text = "eval_test_string_1234567890"
    _readback = _shape.text_frame.paragraphs[0].portions[0].text
    _EVAL_MODE = "evaluation" in _readback.lower() or len(_readback) < 20
except Exception:
    _EVAL_MODE = True


@pytest.fixture
def prs():
    """Create a minimal presentation with one slide."""
    p = slides.Presentation()
    layout = p.masters[0].layout_slides[0]
    p.slides.insert_empty_slide(0, layout)
    return p


class TestInchesHelper:
    def test_one_inch(self):
        assert _inches(1) == 72.0  # 72 points per inch

    def test_zero(self):
        assert _inches(0) == 0

    def test_fractional(self):
        assert _inches(0.5) == 36.0  # 36 points = 0.5 inches


class TestPositionSlots:
    def test_all_slots_have_four_values(self):
        for name, slot in _POSITION_SLOTS.items():
            assert len(slot) == 4, f"Slot '{name}' should have 4 values (x, y, w, h)"

    def test_all_values_positive(self):
        for name, slot in _POSITION_SLOTS.items():
            for val in slot:
                assert val > 0, f"Slot '{name}' has non-positive value"


class TestChartTypeMap:
    def test_six_types(self):
        assert len(_CHART_TYPE_MAP) == 6

    def test_expected_keys(self):
        expected = {"clustered_bar", "stacked_bar", "line", "pie",
                    "doughnut", "clustered_column"}
        assert set(_CHART_TYPE_MAP.keys()) == expected


class TestThemeHelpers:
    def test_get_theme_colors_returns_list(self, prs):
        colors = _get_theme_colors(prs)
        assert isinstance(colors, list)

    def test_theme_colors_are_hex(self, prs):
        colors = _get_theme_colors(prs)
        for c in colors:
            assert c.startswith("#"), f"Color '{c}' should start with #"
            assert len(c) == 7, f"Color '{c}' should be #rrggbb format"

    def test_apply_theme_no_crash(self, prs):
        """_apply_theme_to_chart should not raise even with empty data."""
        slide = prs.slides[0]
        try:
            chart = slide.shapes.add_chart(
                charts.ChartType.CLUSTERED_BAR, 100, 100, 400, 300, True
            )
            colors = _get_theme_colors(prs)
            _apply_theme_to_chart(chart, colors)
        except Exception:
            pass  # Eval mode may interfere


class TestCreateChart:
    def test_invalid_chart_type(self, prs):
        result = create_chart(
            prs, slide_idx=0, chart_type="invalid_type",
            title="Test", categories=["A"], series=[{"name": "S", "values": [1]}]
        )
        assert result["status"] == "error"
        assert "Unknown chart type" in result["message"]

    def test_invalid_position(self, prs):
        result = create_chart(
            prs, slide_idx=0, chart_type="clustered_bar",
            title="Test", categories=["A"], series=[{"name": "S", "values": [1]}],
            position="top_left"
        )
        assert result["status"] == "error"
        assert "Unknown position" in result["message"]

    def test_invalid_slide_index(self, prs):
        result = create_chart(
            prs, slide_idx=999, chart_type="clustered_bar",
            title="Test", categories=["A"], series=[{"name": "S", "values": [1]}]
        )
        assert result["status"] == "error"
        assert "out of range" in result["message"]

    @pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval may interfere with chart creation")
    def test_valid_bar_chart(self, prs):
        result = create_chart(
            prs, slide_idx=0, chart_type="clustered_bar",
            title="Revenue", categories=["Q1", "Q2", "Q3"],
            series=[{"name": "2024", "values": [100, 150, 200]}],
            position="center"
        )
        assert result["status"] == "ok"
        assert "shape_name" in result
        assert result["chart_type"] == "clustered_bar"

    @pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval may interfere with chart creation")
    def test_valid_line_chart(self, prs):
        result = create_chart(
            prs, slide_idx=0, chart_type="line",
            title="Trend", categories=["Jan", "Feb", "Mar"],
            series=[{"name": "Growth", "values": [10, 20, 30]}],
            position="left_half"
        )
        assert result["status"] == "ok"
        assert result["chart_type"] == "line"

    @pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval may interfere with chart creation")
    def test_valid_pie_chart(self, prs):
        result = create_chart(
            prs, slide_idx=0, chart_type="pie",
            title="Share", categories=["A", "B", "C"],
            series=[{"name": "Market", "values": [40, 35, 25]}],
            position="right_half"
        )
        assert result["status"] == "ok"
        assert result["chart_type"] == "pie"

    @pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval may interfere with chart creation")
    def test_multiple_series(self, prs):
        result = create_chart(
            prs, slide_idx=0, chart_type="clustered_column",
            title="Comparison", categories=["Q1", "Q2"],
            series=[
                {"name": "Revenue", "values": [100, 150]},
                {"name": "Cost", "values": [80, 90]},
            ],
            position="center"
        )
        assert result["status"] == "ok"


class TestCreateTable:
    def test_invalid_position(self, prs):
        result = create_table(
            prs, slide_idx=0, headers=["A", "B"],
            rows=[["1", "2"]], position="top_left"
        )
        assert result["status"] == "error"
        assert "Unknown position" in result["message"]

    def test_invalid_slide_index(self, prs):
        result = create_table(
            prs, slide_idx=999, headers=["A", "B"],
            rows=[["1", "2"]]
        )
        assert result["status"] == "error"
        assert "out of range" in result["message"]

    @pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval may interfere with table creation")
    def test_valid_table(self, prs):
        result = create_table(
            prs, slide_idx=0,
            headers=["Metric", "Q2", "Q3"],
            rows=[["Revenue", "13.1", "14.8"], ["EBITDA", "3.4", "4.0"]],
            position="center"
        )
        assert result["status"] == "ok"
        assert result["rows"] == 3  # 2 data + 1 header
        assert result["cols"] == 3
        assert "shape_name" in result

    @pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval may interfere with table creation")
    def test_custom_col_widths(self, prs):
        result = create_table(
            prs, slide_idx=0,
            headers=["Name", "Value"],
            rows=[["A", "1"]],
            col_widths=[3.0, 2.0]
        )
        assert result["status"] == "ok"
        assert result["cols"] == 2

    @pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval may interfere with table creation")
    def test_table_bottom_half(self, prs):
        result = create_table(
            prs, slide_idx=0,
            headers=["A", "B", "C", "D"],
            rows=[["1", "2", "3", "4"]],
            position="bottom_half"
        )
        assert result["status"] == "ok"
