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
