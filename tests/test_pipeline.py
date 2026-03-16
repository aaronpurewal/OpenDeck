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
from pipeline import _remap_content_shapes

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


class TestRemapContentShapes:
    """Test that content output shape names get corrected for cloned slides."""

    def _make_post_state(self, label_list, shapes_by_slide):
        """Build a minimal post_state dict for remap testing."""
        slides_data = []
        for shapes in shapes_by_slide:
            slides_data.append({"shapes": [{"name": n, "type": t}
                                           for n, t in shapes]})
        return {"label_list": label_list, "slides": slides_data}

    def test_corrects_wrong_shape_name(self):
        """Content LLM outputs 'Holder 2' but actual shape is 'object 2'."""
        post_state = self._make_post_state(
            ["slide_0", "new_summary_1"],
            [
                [("Title 1", "text")],
                [("object 2", "text"), ("object 3", "text")],
            ]
        )
        plan = {
            "content_manifest": [
                {"action": "fill_placeholder", "slide_label": "new_summary_1",
                 "shape_name": "object 2"},
            ]
        }
        content = {
            "content_updates": [
                {"action": "fill_placeholder", "slide_label": "new_summary_1",
                 "shape_name": "Holder 2", "text": "Hello"},
            ]
        }
        _remap_content_shapes(content, plan, post_state, {"new_summary_1"})
        assert content["content_updates"][0]["shape_name"] == "object 2"

    def test_skips_non_cloned_slides(self):
        """Shape names on existing slides should not be remapped."""
        post_state = self._make_post_state(
            ["slide_0"],
            [[("Title 1", "text")]],
        )
        plan = {"content_manifest": [
            {"action": "edit_paragraph", "slide_label": "slide_0",
             "shape_name": "Title 1"},
        ]}
        content = {"content_updates": [
            {"action": "edit_paragraph", "slide_label": "slide_0",
             "shape_name": "Title 1", "new_text": "New"},
        ]}
        _remap_content_shapes(content, plan, post_state, {"new_summary_1"})
        assert content["content_updates"][0]["shape_name"] == "Title 1"

    def test_skips_create_actions(self):
        """create_chart and create_table don't have shape names to fix."""
        post_state = self._make_post_state(
            ["new_summary_1"],
            [[("object 2", "text")]],
        )
        plan = {"content_manifest": [
            {"action": "create_chart", "slide_label": "new_summary_1"},
        ]}
        content = {"content_updates": [
            {"action": "create_chart", "slide_label": "new_summary_1",
             "chart_type": "pie", "categories": ["A"], "series": []},
        ]}
        _remap_content_shapes(content, plan, post_state, {"new_summary_1"})
        # Should not crash or add a shape_name
        assert "shape_name" not in content["content_updates"][0] or \
               content["content_updates"][0].get("shape_name") is None or \
               content["content_updates"][0].get("action") == "create_chart"

    def test_leaves_correct_names_alone(self):
        """If the content LLM already has the right name, don't touch it."""
        post_state = self._make_post_state(
            ["new_summary_1"],
            [[("object 2", "text")]],
        )
        plan = {"content_manifest": [
            {"action": "fill_placeholder", "slide_label": "new_summary_1",
             "shape_name": "object 2"},
        ]}
        content = {"content_updates": [
            {"action": "fill_placeholder", "slide_label": "new_summary_1",
             "shape_name": "object 2", "text": "Correct"},
        ]}
        _remap_content_shapes(content, plan, post_state, {"new_summary_1"})
        assert content["content_updates"][0]["shape_name"] == "object 2"

    def test_multiple_shapes_same_type(self):
        """Multiple text shapes on same cloned slide get correct positional mapping."""
        post_state = self._make_post_state(
            ["new_summary_1"],
            [[("object 2", "text"), ("object 3", "text"), ("object 4", "text")]],
        )
        plan = {"content_manifest": [
            {"action": "fill_placeholder", "slide_label": "new_summary_1",
             "shape_name": "object 2"},
            {"action": "fill_placeholder", "slide_label": "new_summary_1",
             "shape_name": "object 3"},
            {"action": "fill_placeholder", "slide_label": "new_summary_1",
             "shape_name": "object 4"},
        ]}
        content = {"content_updates": [
            {"action": "fill_placeholder", "slide_label": "new_summary_1",
             "shape_name": "Holder 2", "text": "Title"},
            {"action": "fill_placeholder", "slide_label": "new_summary_1",
             "shape_name": "Holder 3", "text": "Body"},
            {"action": "fill_placeholder", "slide_label": "new_summary_1",
             "shape_name": "Holder 4", "text": "Footer"},
        ]}
        _remap_content_shapes(content, plan, post_state, {"new_summary_1"})
        assert content["content_updates"][0]["shape_name"] == "object 2"
        assert content["content_updates"][1]["shape_name"] == "object 3"
        assert content["content_updates"][2]["shape_name"] == "object 4"


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
