"""
Unit tests for structural tool functions.

NOTE: Aspose evaluation version starts presentations with 1 default slide.
"""

import pytest
import os
import tempfile
import aspose.slides as slides

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tools import (
    clone_slide, duplicate_slide, delete_slides, save_deck,
    list_layouts, get_slide_state, get_bounds
)
from state import harvest_deck

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
SAMPLE_DECK = os.path.join(FIXTURE_DIR, "sample_deck.pptx")


@pytest.fixture
def prs():
    """Load test presentation if available, otherwise create a minimal one."""
    if os.path.exists(SAMPLE_DECK):
        return slides.Presentation(SAMPLE_DECK)
    return slides.Presentation()


class TestHarvestDeck:
    def test_returns_dict(self, prs):
        state = harvest_deck(prs)
        assert isinstance(state, dict)
        assert "slide_count" in state
        assert "slides" in state
        assert "label_list" in state
        assert "master_layouts" in state

    def test_label_list_matches_slides(self, prs):
        state = harvest_deck(prs)
        assert len(state["label_list"]) == state["slide_count"]
        assert len(state["slides"]) == state["slide_count"]

    def test_labels_are_sequential(self, prs):
        state = harvest_deck(prs)
        for i, label in enumerate(state["label_list"]):
            assert label == f"slide_{i}"


class TestListLayouts:
    def test_returns_layouts(self, prs):
        result = list_layouts(prs)
        assert result["status"] == "ok"
        assert isinstance(result["layouts"], list)


class TestGetSlideState:
    def test_valid_index(self, prs):
        if len(prs.slides) > 0:
            result = get_slide_state(prs, 0)
            assert result["status"] == "ok"
            assert "slide" in result

    def test_invalid_index(self, prs):
        result = get_slide_state(prs, 999)
        assert result["status"] == "error"


class TestCloneSlide:
    def test_clone_existing_layout(self, prs):
        layouts = list_layouts(prs)
        if layouts["layouts"]:
            layout_name = layouts["layouts"][0]["name"]
            initial_count = len(prs.slides)
            result = clone_slide(prs, layout_name=layout_name)
            assert result["status"] == "ok"
            assert len(prs.slides) == initial_count + 1

    def test_clone_nonexistent_layout(self, prs):
        result = clone_slide(prs, layout_name="NONEXISTENT_LAYOUT_XYZ")
        assert result["status"] == "error"


class TestSaveDeck:
    def test_save_and_reopen(self, prs):
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            result = save_deck(prs, tmp.name)
            assert result["status"] == "ok"
            reopened = slides.Presentation(tmp.name)
            assert len(reopened.slides) == len(prs.slides)
            os.unlink(tmp.name)
