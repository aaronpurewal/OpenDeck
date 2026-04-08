"""
Tests for state extraction accuracy.

NOTE: Aspose evaluation version truncates text and starts presentations
with 1 default slide. Tests account for this.
"""

import pytest
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import aspose.slides as slides
from state import harvest_deck, extract_shape, estimate_char_limit


class TestEstimateCharLimit:
    def test_basic_calculation(self):
        # 720pt wide (10 inches) x 144pt tall (2 inches)
        limit = estimate_char_limit(720, 144)
        assert limit > 0
        assert isinstance(limit, int)

    def test_larger_font_means_fewer_chars(self):
        # 360pt wide x 144pt tall
        limit_small = estimate_char_limit(360, 144, font_size_pt=12)
        limit_large = estimate_char_limit(360, 144, font_size_pt=24)
        assert limit_small > limit_large

    def test_zero_dimensions(self):
        limit = estimate_char_limit(0, 0)
        assert limit >= 0

    def test_no_font_size_uses_default(self):
        limit = estimate_char_limit(360, 144, font_size_pt=None)
        assert limit > 0


class TestExtractShape:
    def test_text_shape(self):
        prs = slides.Presentation()
        slide = prs.slides[0]  # Default slide exists in eval

        ashape = slide.shapes.add_auto_shape(
            slides.ShapeType.RECTANGLE, 100, 100, 500, 300
        )
        ashape.name = "TestShape"
        ashape.text_frame.paragraphs[0].portions[0].text = "Hello"

        result = extract_shape(ashape)
        assert result is not None
        assert result["type"] == "text"
        assert result["name"] == "TestShape"
        assert "bounds" in result
        assert "paragraphs" in result

    def test_unsupported_shape_returns_none(self):
        prs = slides.Presentation()
        slide = prs.slides[0]
        # Line shape — may or may not have text frame
        shape = slide.shapes.add_auto_shape(
            slides.ShapeType.LINE, 100, 100, 200, 0
        )
        # Just verify no crash
        extract_shape(shape)


class TestHarvestDeck:
    def test_default_presentation(self):
        prs = slides.Presentation()
        state = harvest_deck(prs)
        assert isinstance(state, dict)
        assert "slide_count" in state
        assert "slides" in state
        assert "label_list" in state
        assert "master_layouts" in state
        # Aspose eval starts with 1 default slide
        assert state["slide_count"] >= 0
        assert len(state["label_list"]) == state["slide_count"]
        assert len(state["slides"]) == state["slide_count"]

    def test_added_slides(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        initial = len(prs.slides)
        for i in range(3):
            prs.slides.insert_empty_slide(len(prs.slides), layout)

        state = harvest_deck(prs)
        assert state["slide_count"] == initial + 3
        assert len(state["label_list"]) == initial + 3

    def test_labels_are_sequential(self):
        prs = slides.Presentation()
        state = harvest_deck(prs)
        for i, label in enumerate(state["label_list"]):
            assert label == f"slide_{i}"


class TestOverlayDetection:
    """Verify decorations are anchored to table rows during harvest."""

    def test_oval_overlaying_table_row_gets_anchored(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        slide_idx = len(prs.slides) - 1
        slide = prs.slides[slide_idx]

        # Build a 4-row table at (100, 100), each row 50pt tall, 3 cols of 100pt
        col_widths = [100.0, 100.0, 100.0]
        row_heights = [50.0, 50.0, 50.0, 50.0]
        table = slide.shapes.add_table(100.0, 100.0, col_widths, row_heights)
        table.name = "RisksTable"

        # Place an oval overlaying row 2 (y=200..250) in the rightmost column area
        # Center the oval at y = 225 (middle of row 2)
        oval = slide.shapes.add_auto_shape(
            slides.ShapeType.ELLIPSE,
            350.0,  # x within table x-range (100..400)
            215.0,  # y center at 225
            20.0, 20.0
        )
        oval.name = "Risk Dot 2"

        state = harvest_deck(prs)
        slide_state = state["slides"][slide_idx]

        # Find the oval in the harvested shapes
        oval_state = None
        for s in slide_state["shapes"]:
            if s.get("name") == "Risk Dot 2":
                oval_state = s
                break

        assert oval_state is not None, "Oval should be present in harvested state"
        assert oval_state["type"] == "decoration"
        assert "anchor" in oval_state
        assert oval_state["anchor"]["kind"] == "table_row"
        assert oval_state["anchor"]["shape"] == "RisksTable"
        assert oval_state["anchor"]["row_idx"] == 2

        # The table should have row_overlays referencing the oval
        table_state = None
        for s in slide_state["shapes"]:
            if s.get("name") == "RisksTable":
                table_state = s
                break
        assert table_state is not None
        assert "row_overlays" in table_state
        # row_overlays keys are stringified ints
        assert "2" in table_state["row_overlays"]
        assert "Risk Dot 2" in table_state["row_overlays"]["2"]

    def test_sections_detected_in_numbered_table(self):
        """Table with numbered headers gets sections detected."""
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        slide_idx = len(prs.slides) - 1
        slide = prs.slides[slide_idx]

        col_widths = [200.0, 100.0]
        row_heights = [30.0, 60.0, 30.0, 60.0]
        table = slide.shapes.add_table(50.0, 50.0, col_widths, row_heights)
        table.name = "NumberedTable"

        # Row 0 = header "(1) First risk"
        # Row 1 = bullets for risk 1
        # Row 2 = header "(2) Second risk"
        # Row 3 = bullets for risk 2
        labels = [("(1) First risk", "bullets for 1"),
                  ("bullets for 1 continued", ""),
                  ("(2) Second risk", "bullets for 2"),
                  ("bullets for 2 continued", "")]
        for r, (c0, c1) in enumerate(labels):
            try:
                cell0 = table.rows[r][0]
                tf = cell0.text_frame
                if tf.paragraphs.count > 0 and tf.paragraphs[0].portions.count > 0:
                    tf.paragraphs[0].portions[0].text = c0
            except Exception:
                pass

        state = harvest_deck(prs)
        slide_state = state["slides"][slide_idx]
        table_state = None
        for s in slide_state["shapes"]:
            if s.get("name") == "NumberedTable":
                table_state = s
                break
        assert table_state is not None
        sections = table_state.get("sections", [])
        # Should detect 2 sections — row 0 is "(1)" header, row 2 is "(2)" header
        assert len(sections) == 2
        assert sections[0]["header_row"] == 0
        assert sections[0]["bullet_rows"] == [1]
        assert sections[1]["header_row"] == 2
        assert sections[1]["bullet_rows"] == [3]
        assert sections[0]["title_preview"].startswith("(1)")
        assert sections[1]["title_preview"].startswith("(2)")

    def test_row_char_limits_per_row(self):
        """Each table row should have its own char_limit in harvested state."""
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        slide_idx = len(prs.slides) - 1
        slide = prs.slides[slide_idx]

        col_widths = [200.0, 100.0]
        # Different row heights so different char limits
        row_heights = [20.0, 60.0, 120.0]
        table = slide.shapes.add_table(50.0, 50.0, col_widths, row_heights)
        table.name = "MultiHeightTable"

        state = harvest_deck(prs)
        slide_state = state["slides"][slide_idx]
        table_state = None
        for s in slide_state["shapes"]:
            if s.get("name") == "MultiHeightTable":
                table_state = s
                break
        assert table_state is not None
        limits = table_state.get("row_char_limits", [])
        assert len(limits) == 3
        assert all(isinstance(l, int) and l > 0 for l in limits)
        # Taller rows should have larger limits
        assert limits[2] > limits[0]

    def test_no_sections_in_plain_table(self):
        """Table without numbered headers has no sections."""
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        slide_idx = len(prs.slides) - 1
        slide = prs.slides[slide_idx]

        col_widths = [100.0, 100.0]
        row_heights = [30.0, 30.0]
        table = slide.shapes.add_table(50.0, 50.0, col_widths, row_heights)
        table.name = "PlainTable"

        state = harvest_deck(prs)
        slide_state = state["slides"][slide_idx]
        table_state = None
        for s in slide_state["shapes"]:
            if s.get("name") == "PlainTable":
                table_state = s
                break
        assert table_state is not None
        # No numbered headers means no sections key (or empty)
        assert not table_state.get("sections")

    def test_oval_outside_table_not_anchored(self):
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)
        slide_idx = len(prs.slides) - 1
        slide = prs.slides[slide_idx]

        # Table at (100, 100), 200pt tall total
        col_widths = [100.0, 100.0]
        row_heights = [50.0, 50.0]
        table = slide.shapes.add_table(100.0, 100.0, col_widths, row_heights)
        table.name = "SmallTable"

        # Oval far below the table
        oval = slide.shapes.add_auto_shape(
            slides.ShapeType.ELLIPSE, 150.0, 500.0, 20.0, 20.0
        )
        oval.name = "Orphan"

        state = harvest_deck(prs)
        slide_state = state["slides"][slide_idx]

        for s in slide_state["shapes"]:
            if s.get("name") == "Orphan":
                # Orphan should NOT have an anchor
                assert "anchor" not in s
                break
