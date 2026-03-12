"""
End-to-end integration tests for the pipeline.

NOTE: Aspose evaluation version has limitations:
- Text is truncated when read back
- Default presentation starts with 1 slide
- Saved files include evaluation watermark
"""

import pytest
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import aspose.slides as slides
from state import harvest_deck
from validation import smoke_test

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
SAMPLE_DECK = os.path.join(FIXTURE_DIR, "sample_deck.pptx")


class TestSmokeTest:
    def test_valid_file(self):
        prs = slides.Presentation()
        # Eval starts with 1 slide, add 1 more
        layout = prs.masters[0].layout_slides[0]
        prs.slides.insert_empty_slide(len(prs.slides), layout)

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs.save(tmp.name, slides.export.SaveFormat.PPTX)
            result = smoke_test(tmp.name)
            assert result["status"] == "ok"
            assert result["slide_count"] >= 1
            os.unlink(tmp.name)

    def test_nonexistent_file(self):
        result = smoke_test("/nonexistent/path/file.pptx")
        assert result["status"] == "error"


class TestHarvestRoundTrip:
    def test_harvest_save_harvest(self):
        """Verify state is consistent after save and re-harvest."""
        prs = slides.Presentation()
        layout = prs.masters[0].layout_slides[0]
        for i in range(3):
            prs.slides.insert_empty_slide(len(prs.slides), layout)
            slide = prs.slides[len(prs.slides) - 1]
            ashape = slide.shapes.add_auto_shape(
                slides.ShapeType.RECTANGLE, 100, 100, 500, 300
            )
            ashape.name = f"TextBox_{i}"
            ashape.text_frame.paragraphs[0].portions[0].text = f"Content {i}"

        state1 = harvest_deck(prs)

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            prs.save(tmp.name, slides.export.SaveFormat.PPTX)
            prs2 = slides.Presentation(tmp.name)
            state2 = harvest_deck(prs2)

            assert state1["slide_count"] == state2["slide_count"]
            assert len(state1["label_list"]) == len(state2["label_list"])

            for s1, s2 in zip(state1["slides"], state2["slides"]):
                assert s1["label"] == s2["label"]

            os.unlink(tmp.name)


@pytest.mark.skipif(
    not os.path.exists(SAMPLE_DECK),
    reason="sample_deck.pptx not found in fixtures"
)
class TestWithFixture:
    def test_harvest_fixture_deck(self):
        prs = slides.Presentation(SAMPLE_DECK)
        state = harvest_deck(prs)
        assert state["slide_count"] > 0
        assert len(state["master_layouts"]) > 0
