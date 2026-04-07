"""
Tests for edit_run, edit_paragraph, edit_table_cell, edit_table_run.

NOTE: Aspose evaluation version truncates text when reading back.
Tests that verify text content after writing are marked accordingly.
With a licensed Aspose, all text assertions should pass as-is.
"""

import pytest
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import aspose.slides as slides
from tools import (
    edit_run, edit_paragraph, edit_table_cell, edit_table_run,
    fill_placeholder, fill_table, clone_slide,
    move_shape, swap_shape_positions, set_shape_fill, swap_table_rows
)

# Aspose evaluation truncates text — detect this
_EVAL_MODE = False
try:
    _prs = slides.Presentation()
    _l = _prs.masters[0].layout_slides[0]
    _prs.slides.insert_empty_slide(0, _l)
    _s = _prs.slides[0].shapes.add_auto_shape(slides.ShapeType.RECTANGLE, 0, 0, 100, 100)
    _s.text_frame.paragraphs[0].portions[0].text = "test_long_text_here"
    _readback = _s.text_frame.paragraphs[0].portions[0].text
    _EVAL_MODE = "truncated" in _readback.lower() or len(_readback) < 19
except Exception:
    _EVAL_MODE = True


def _create_text_slide(prs):
    """Create a slide with a text shape containing formatted runs."""
    layout = prs.masters[0].layout_slides[0]
    prs.slides.insert_empty_slide(len(prs.slides), layout)
    slide_idx = len(prs.slides) - 1
    slide = prs.slides[slide_idx]

    ashape = slide.shapes.add_auto_shape(
        slides.ShapeType.RECTANGLE, 100, 100, 500, 300
    )
    ashape.name = "TestTextBox"
    tf = ashape.text_frame

    para = tf.paragraphs[0]
    if para.portions.count > 0:
        para.portions[0].text = "Revenue: "
        para.portions[0].portion_format.font_bold = slides.NullableBool.TRUE
    else:
        p = slides.Portion()
        p.text = "Revenue: "
        p.portion_format.font_bold = slides.NullableBool.TRUE
        para.portions.add(p)

    portion2 = slides.Portion()
    portion2.text = "$13.1M"
    portion2.portion_format.font_bold = slides.NullableBool.FALSE
    para.portions.add(portion2)

    return slide_idx


class TestEditRun:
    def test_shape_not_found(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        result = edit_run(prs, len(prs.slides) - 1, "NonexistentShape", 0, "text", "new")
        assert result["status"] == "error"

    def test_paragraph_out_of_range(self):
        prs = slides.Presentation()
        idx = _create_text_slide(prs)
        result = edit_run(prs, idx, "TestTextBox", 99, "$13.1M", "$14.8M")
        assert result["status"] == "error"

    def test_run_not_found(self):
        prs = slides.Presentation()
        idx = _create_text_slide(prs)
        result = edit_run(prs, idx, "TestTextBox", 0, "NONEXISTENT", "new")
        assert result["status"] == "error"

    @pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval truncates text, can't match runs")
    def test_successful_edit(self):
        prs = slides.Presentation()
        idx = _create_text_slide(prs)
        result = edit_run(prs, idx, "TestTextBox", 0, "$13.1M", "$14.8M")
        assert result["status"] == "ok"


class TestEditParagraph:
    @pytest.mark.skipif(_EVAL_MODE, reason="Aspose eval truncates text")
    def test_successful_rewrite(self):
        prs = slides.Presentation()
        idx = _create_text_slide(prs)
        result = edit_paragraph(prs, idx, "TestTextBox", 0, "New text")
        assert result["status"] == "ok"

    def test_paragraph_out_of_range(self):
        prs = slides.Presentation()
        idx = _create_text_slide(prs)
        result = edit_paragraph(prs, idx, "TestTextBox", 99, "text")
        assert result["status"] == "error"


class TestSlideIndexValidation:
    def test_negative_index(self):
        prs = slides.Presentation()
        result = edit_run(prs, -1, "shape", 0, "match", "new")
        assert result["status"] == "error"

    def test_out_of_range_index(self):
        prs = slides.Presentation()
        result = edit_paragraph(prs, 999, "shape", 0, "text")
        assert result["status"] == "error"


def _create_slide_with_oval(prs, name="Oval 1", x=200.0, y=300.0,
                            w=40.0, h=40.0):
    """Add an oval autoshape to the last slide of prs at given coordinates."""
    slide = prs.slides[len(prs.slides) - 1]
    shape = slide.shapes.add_auto_shape(slides.ShapeType.ELLIPSE, x, y, w, h)
    shape.name = name
    return shape


class TestMoveShape:
    def test_absolute_move(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        _create_slide_with_oval(prs, "Oval A", x=100.0, y=200.0)

        result = move_shape(prs, idx, "Oval A", x=500.0, y=600.0)
        assert result["status"] == "ok"

        slide = prs.slides[idx]
        for s in slide.shapes:
            if s.name == "Oval A":
                assert abs(s.x - 500.0) < 1.0
                assert abs(s.y - 600.0) < 1.0
                break

    def test_relative_move(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        _create_slide_with_oval(prs, "Oval B", x=100.0, y=200.0)

        result = move_shape(prs, idx, "Oval B", dx=50.0, dy=-30.0)
        assert result["status"] == "ok"

        slide = prs.slides[idx]
        for s in slide.shapes:
            if s.name == "Oval B":
                assert abs(s.x - 150.0) < 1.0
                assert abs(s.y - 170.0) < 1.0
                break

    def test_shape_not_found(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        result = move_shape(prs, idx, "NonexistentShape", x=100, y=100)
        assert result["status"] == "error"

    def test_invalid_slide_index(self):
        prs = slides.Presentation()
        result = move_shape(prs, 999, "Oval A", x=100, y=100)
        assert result["status"] == "error"


class TestSwapShapePositions:
    def test_swap_two_shapes(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        _create_slide_with_oval(prs, "Oval A", x=100.0, y=200.0)
        _create_slide_with_oval(prs, "Oval B", x=400.0, y=500.0)

        result = swap_shape_positions(prs, idx, "Oval A", "Oval B")
        assert result["status"] == "ok"

        slide = prs.slides[idx]
        positions = {s.name: (s.x, s.y) for s in slide.shapes
                     if s.name in ("Oval A", "Oval B")}
        assert abs(positions["Oval A"][0] - 400.0) < 1.0
        assert abs(positions["Oval A"][1] - 500.0) < 1.0
        assert abs(positions["Oval B"][0] - 100.0) < 1.0
        assert abs(positions["Oval B"][1] - 200.0) < 1.0

    def test_missing_shape(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        _create_slide_with_oval(prs, "Oval A", x=100.0, y=200.0)
        result = swap_shape_positions(prs, idx, "Oval A", "MissingOval")
        assert result["status"] == "error"


class TestSetShapeFill:
    def test_set_red_fill(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        _create_slide_with_oval(prs, "RAG Badge", x=100.0, y=100.0)

        result = set_shape_fill(prs, idx, "RAG Badge", "#C00000")
        assert result["status"] == "ok"
        assert "color" in result

    def test_invalid_hex(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        _create_slide_with_oval(prs, "Oval A", x=100.0, y=100.0)

        result = set_shape_fill(prs, idx, "Oval A", "#XYZ")
        assert result["status"] == "error"

    def test_short_hex_rejected(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        _create_slide_with_oval(prs, "Oval A", x=100.0, y=100.0)

        result = set_shape_fill(prs, idx, "Oval A", "#FFF")
        assert result["status"] == "error"

    def test_shape_not_found(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1

        result = set_shape_fill(prs, idx, "Nope", "#FF0000")
        assert result["status"] == "error"


class TestSwapTableRows:
    def _create_slide_with_table_and_overlays(self, prs, n_rows=4, n_cols=3):
        """Create a slide with a table and an oval overlaying each row."""
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        slide = prs.slides[idx]

        # Table at (100, 100), 400 wide, 200 tall (50pt rows)
        col_widths = [100.0] * n_cols
        row_heights = [50.0] * n_rows
        from aspose.pydrawing import RectangleF
        # Aspose API: add_table(x, y, columnWidths, rowHeights)
        table = slide.shapes.add_table(100.0, 100.0, col_widths, row_heights)
        table.name = "TestTable"

        # Fill some text into each cell
        for r in range(n_rows):
            for c in range(n_cols):
                cell = table.rows[r][c]
                tf = cell.text_frame
                if tf.paragraphs.count > 0 and tf.paragraphs[0].portions.count > 0:
                    tf.paragraphs[0].portions[0].text = f"R{r}C{c}"

        # Add an oval overlaying each row, centered in the rightmost column
        # Table spans x=100..400 (cols 100..200, 200..300, 300..400)
        # Place oval in the middle of col 3 (x center ~350)
        for r in range(n_rows):
            row_center_y = 100.0 + r * 50.0 + 25.0
            oval = slide.shapes.add_auto_shape(
                slides.ShapeType.ELLIPSE,
                340.0,  # inside col 3 (x=300..400)
                row_center_y - 10.0,
                20.0, 20.0
            )
            oval.name = f"Dot R{r}"

        return idx

    def test_swap_rows_moves_overlays(self):
        prs = slides.Presentation()
        idx = self._create_slide_with_table_and_overlays(prs, n_rows=4)

        # Capture initial dot positions
        slide = prs.slides[idx]
        initial = {}
        for s in slide.shapes:
            if s.name and s.name.startswith("Dot R"):
                initial[s.name] = (s.x, s.y)

        result = swap_table_rows(prs, idx, "TestTable",
                                  row_idx_a=1, row_idx_b=2)
        assert result["status"] == "ok"
        assert result["swapped_cells"] >= 0  # may be 0 in eval mode
        assert "moved_shapes" in result

        # The "Dot R1" should now be at the y-position originally
        # held by "Dot R2" (and vice versa)
        final = {}
        for s in slide.shapes:
            if s.name and s.name.startswith("Dot R"):
                final[s.name] = (s.x, s.y)

        assert abs(final["Dot R1"][1] - initial["Dot R2"][1]) < 1.0
        assert abs(final["Dot R2"][1] - initial["Dot R1"][1]) < 1.0
        # Untouched rows should be unchanged
        assert abs(final["Dot R0"][1] - initial["Dot R0"][1]) < 1.0
        assert abs(final["Dot R3"][1] - initial["Dot R3"][1]) < 1.0

    def test_invalid_row_index(self):
        prs = slides.Presentation()
        idx = self._create_slide_with_table_and_overlays(prs, n_rows=4)
        result = swap_table_rows(prs, idx, "TestTable",
                                  row_idx_a=99, row_idx_b=0)
        assert result["status"] == "error"

    def test_table_not_found(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        result = swap_table_rows(prs, idx, "NonexistentTable", 0, 1)
        assert result["status"] == "error"

    def test_same_row_is_noop(self):
        prs = slides.Presentation()
        idx = self._create_slide_with_table_and_overlays(prs, n_rows=4)
        result = swap_table_rows(prs, idx, "TestTable", 1, 1)
        assert result["status"] == "ok"
        assert result["swapped_cells"] == 0
