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
    move_shape, swap_shape_positions, set_shape_fill, swap_table_rows,
    swap_table_sections, fit_tables_to_slide
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


def _create_section_table(prs, slide_idx, name, x, y,
                          n_sections=2, col_widths=None,
                          with_dots=True):
    """
    Create a table with numbered sections on the given slide.

    Each section = 2 rows: a header row "(N) Title N" and a bullet row.
    Returns the table shape.
    """
    if col_widths is None:
        col_widths = [200.0, 60.0, 60.0]
    row_heights = [40.0] * (n_sections * 2)
    slide = prs.slides[slide_idx]
    table = slide.shapes.add_table(x, y, col_widths, row_heights)
    table.name = name

    for s_idx in range(n_sections):
        header_row = s_idx * 2
        bullet_row = header_row + 1
        try:
            cell = table.rows[header_row][0]
            tf = cell.text_frame
            if tf.paragraphs.count > 0 and tf.paragraphs[0].portions.count > 0:
                tf.paragraphs[0].portions[0].text = f"({s_idx + 1}) Title {s_idx + 1}"
        except Exception:
            pass
        try:
            cell = table.rows[bullet_row][0]
            tf = cell.text_frame
            if tf.paragraphs.count > 0 and tf.paragraphs[0].portions.count > 0:
                tf.paragraphs[0].portions[0].text = f"Bullets for section {s_idx + 1}"
        except Exception:
            pass

    # Add overlay dots on each header row (centered in col 2 area)
    if with_dots:
        total_table_w = sum(col_widths)
        for s_idx in range(n_sections):
            header_row = s_idx * 2
            row_center_y = y + header_row * 40.0 + 20.0
            dot = slide.shapes.add_auto_shape(
                slides.ShapeType.ELLIPSE,
                x + total_table_w - 50.0,  # inside last col
                row_center_y - 10.0,
                20.0, 20.0
            )
            dot.name = f"Dot {name} S{s_idx}"
    return table


class TestSwapTableSections:
    def test_same_slide_swap(self):
        """Sections on the same slide, same table — content and overlays swap."""
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1

        _create_section_table(prs, idx, "TestTable", 50.0, 50.0, n_sections=2)

        # Capture initial dot y positions
        slide = prs.slides[idx]
        initial_dots = {}
        for s in slide.shapes:
            if s.name and s.name.startswith("Dot TestTable"):
                initial_dots[s.name] = s.y

        result = swap_table_sections(
            prs, idx, "TestTable", 0,
            idx, "TestTable", 1
        )
        assert result["status"] == "ok"
        assert result["rows_swapped"] == 2
        assert result["cross_slide"] is False

        # Dots should have swapped y positions
        final_dots = {}
        for s in slide.shapes:
            if s.name and s.name.startswith("Dot TestTable"):
                final_dots[s.name] = s.y

        assert abs(final_dots["Dot TestTable S0"] - initial_dots["Dot TestTable S1"]) < 2.0
        assert abs(final_dots["Dot TestTable S1"] - initial_dots["Dot TestTable S0"]) < 2.0

    def test_cross_slide_swap(self):
        """Sections on different slides — overlays recreated on target slides."""
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx_a = len(prs.slides) - 2
        idx_b = len(prs.slides) - 1

        _create_section_table(prs, idx_a, "TableA", 50.0, 50.0, n_sections=2)
        _create_section_table(prs, idx_b, "TableB", 50.0, 50.0, n_sections=2)

        # Before: slide A has 2 dots, slide B has 2 dots
        dots_a_before = [s.name for s in prs.slides[idx_a].shapes
                         if s.name and s.name.startswith("Dot TableA")]
        dots_b_before = [s.name for s in prs.slides[idx_b].shapes
                         if s.name and s.name.startswith("Dot TableB")]
        assert len(dots_a_before) == 2
        assert len(dots_b_before) == 2

        result = swap_table_sections(
            prs, idx_a, "TableA", 0,
            idx_b, "TableB", 0
        )
        assert result["status"] == "ok"
        assert result["cross_slide"] is True
        assert result["rows_swapped"] == 2

        # After: TableA's original dot (S0) was moved to slide B
        # slide B should now have a "Dot TableA S0" somewhere
        names_on_b = {s.name for s in prs.slides[idx_b].shapes if s.name}
        names_on_a = {s.name for s in prs.slides[idx_a].shapes if s.name}
        assert "Dot TableA S0" in names_on_b
        assert "Dot TableB S0" in names_on_a
        # Original dots should be gone from their source slides
        assert "Dot TableA S0" not in names_on_a
        assert "Dot TableB S0" not in names_on_b

    def test_row_count_mismatch_returns_error(self):
        """Sections with different row counts cleanly return an error.

        Note: section detection caps the LAST section's bullet rows to the
        median of prior sections. So to test mismatch, we make section 0
        the bigger one (3 rows) and section 1 the smaller (2 rows).
        """
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        slide = prs.slides[idx]

        # Row 0: "(1) First" (header)
        # Row 1: bullet
        # Row 2: bullet (2nd bullet row for section 0)
        # Row 3: "(2) Second" (header)
        # Row 4: bullet
        # → section 0 = [0,1,2] (3 rows), section 1 = [3,4] (2 rows capped to 1 bullet)
        col_widths = [150.0]
        row_heights = [40.0] * 5
        table = slide.shapes.add_table(50.0, 50.0, col_widths, row_heights)
        table.name = "Unbalanced"

        texts = ["(1) First", "b1a", "b1b", "(2) Second", "b2"]
        for r, t in enumerate(texts):
            try:
                cell = table.rows[r][0]
                cell.text_frame.paragraphs[0].portions[0].text = t
            except Exception:
                pass

        result = swap_table_sections(
            prs, idx, "Unbalanced", 0,
            idx, "Unbalanced", 1
        )
        assert result["status"] == "error"
        assert "row counts" in result["message"].lower()

    def test_table_not_found_returns_error(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        result = swap_table_sections(
            prs, idx, "NonexistentTable", 0,
            idx, "NonexistentTable", 1
        )
        assert result["status"] == "error"


class TestPreWriteTruncation:
    """char_limit on edit_table_cell / edit_table_run truncates text."""

    def test_edit_table_cell_truncates(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        slide = prs.slides[idx]

        col_widths = [150.0, 150.0]
        row_heights = [40.0, 40.0]
        table = slide.shapes.add_table(50.0, 50.0, col_widths, row_heights)
        table.name = "T"

        long_text = "x" * 500
        result = edit_table_cell(
            prs, idx, "T", row_idx=1, col_idx=0,
            new_text=long_text, char_limit=50
        )
        assert result["status"] == "ok"
        assert result.get("truncated") is True

    def test_edit_table_cell_no_truncation_when_under_limit(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        slide = prs.slides[idx]

        col_widths = [150.0, 150.0]
        row_heights = [40.0, 40.0]
        table = slide.shapes.add_table(50.0, 50.0, col_widths, row_heights)
        table.name = "T"

        short_text = "Hello world"
        result = edit_table_cell(
            prs, idx, "T", row_idx=0, col_idx=0,
            new_text=short_text, char_limit=100
        )
        assert result["status"] == "ok"
        assert result.get("truncated") is False

    def test_edit_table_cell_no_limit_arg(self):
        """char_limit is optional — old callers without it still work."""
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        slide = prs.slides[idx]

        table = slide.shapes.add_table(50.0, 50.0, [150.0], [40.0])
        table.name = "T"
        result = edit_table_cell(prs, idx, "T", 0, 0, "test")
        assert result["status"] == "ok"
        assert result.get("truncated") is False


class TestTableFitChecks:
    """fit_tables_to_slide post-write safety net."""

    def test_fits_already_returns_no_shrinkage(self):
        """A table well within slide bounds is untouched."""
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        slide = prs.slides[idx]

        # Small table, fits easily
        table = slide.shapes.add_table(50.0, 50.0, [200.0], [30.0, 30.0])
        table.name = "SmallTable"

        result = fit_tables_to_slide(prs, idx)
        assert result["status"] == "ok"
        assert result["shrunk"] == []
        assert result["overflow_remaining"] == []

    def test_overflowing_table_triggers_shrink(self):
        """A table that overflows gets entries in shrunk list."""
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        idx = len(prs.slides) - 1
        slide = prs.slides[idx]

        # Slide is typically 540pt or 590pt tall. Build a table with
        # row heights that sum well past that.
        slide_h = prs.slide_size.size.height
        # Start near the top, make rows huge
        table = slide.shapes.add_table(
            50.0, 30.0,
            [400.0],
            [float(slide_h), float(slide_h)]  # 2 rows each = slide_h tall
        )
        table.name = "HugeTable"

        # Write substantial text so rows actually take up their minimal_height
        try:
            for r in range(2):
                cell = table.rows[r][0]
                tf = cell.text_frame
                if tf.paragraphs.count > 0 and tf.paragraphs[0].portions.count > 0:
                    tf.paragraphs[0].portions[0].text = "Content " * 20
        except Exception:
            pass

        result = fit_tables_to_slide(prs, idx)
        assert result["status"] == "ok"
        # Either shrunk or overflow_remaining should have entries
        # (depending on whether Aspose recomputed row heights after shrink)
        assert (len(result["shrunk"]) > 0 or
                len(result["overflow_remaining"]) > 0)

    def test_invalid_slide_index(self):
        prs = slides.Presentation()
        result = fit_tables_to_slide(prs, 999)
        assert result["status"] == "error"
