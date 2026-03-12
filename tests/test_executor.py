"""
Tests for the deterministic plan executor.

Tests label resolution, structural operations, and content dispatch
using mock plans against real Aspose presentations.
"""

import pytest
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import aspose.slides as slides
from executor import execute_plan
from tools import list_layouts
from state import harvest_deck


@pytest.fixture
def prs_with_slides():
    """Create a presentation with 3 slides."""
    prs = slides.Presentation()
    # Presentation starts with 0 slides, add 3
    layout = prs.masters[0].layout_slides[0]
    for i in range(3):
        prs.slides.insert_empty_slide(i, layout)
    return prs


class TestLabelResolution:
    def test_resolve_existing_label(self, prs_with_slides):
        label_list = ["slide_0", "slide_1", "slide_2"]
        plan = {"structural_changes": [], "content_updates": []}
        result = execute_plan(plan, prs_with_slides, label_list)
        assert result["status"] == "complete"

    def test_delete_updates_labels(self, prs_with_slides):
        label_list = ["slide_0", "slide_1", "slide_2"]
        plan = {
            "structural_changes": [
                {"action": "delete_slides", "args": {"labels": ["slide_1"]}}
            ],
            "content_updates": []
        }
        result = execute_plan(plan, prs_with_slides, label_list)
        assert result["status"] == "complete"
        assert "slide_1" not in label_list
        assert len(label_list) == 2
        assert label_list == ["slide_0", "slide_2"]

    def test_clone_inserts_label(self, prs_with_slides):
        label_list = ["slide_0", "slide_1", "slide_2"]
        layouts = list_layouts(prs_with_slides)
        layout_name = layouts["layouts"][0]["name"] if layouts["layouts"] else ""
        if not layout_name:
            pytest.skip("No layouts available")

        plan = {
            "structural_changes": [
                {"action": "clone_slide",
                 "args": {"layout_name": layout_name, "insert_at": 1},
                 "label": "new_summary"}
            ],
            "content_updates": []
        }
        result = execute_plan(plan, prs_with_slides, label_list)
        assert result["status"] == "complete"
        assert "new_summary" in label_list
        assert label_list.index("new_summary") == 1
        assert len(label_list) == 4


class TestContentDispatch:
    def test_unknown_action(self, prs_with_slides):
        label_list = ["slide_0", "slide_1", "slide_2"]
        plan = {
            "structural_changes": [],
            "content_updates": [
                {"action": "nonexistent_action",
                 "slide_label": "slide_0",
                 "shape_name": "test"}
            ]
        }
        result = execute_plan(plan, prs_with_slides, label_list)
        assert result["status"] == "complete"
        # Should log error for unknown action
        errors = [l for l in result["log"] if l["status"] == "error"]
        assert len(errors) == 1

    def test_unknown_slide_label(self, prs_with_slides):
        label_list = ["slide_0", "slide_1", "slide_2"]
        plan = {
            "structural_changes": [],
            "content_updates": [
                {"action": "edit_run",
                 "slide_label": "nonexistent_slide",
                 "shape_name": "test",
                 "para_idx": 0,
                 "run_match": "x",
                 "new_text": "y"}
            ]
        }
        result = execute_plan(plan, prs_with_slides, label_list)
        assert result["status"] == "complete"
        errors = [l for l in result["log"] if l["status"] == "error"]
        assert len(errors) == 1


class TestStructuralFailure:
    def test_clone_nonexistent_layout(self, prs_with_slides):
        label_list = ["slide_0", "slide_1", "slide_2"]
        plan = {
            "structural_changes": [
                {"action": "clone_slide",
                 "args": {"layout_name": "NONEXISTENT_XYZ"},
                 "label": "will_fail"}
            ],
            "content_updates": []
        }
        result = execute_plan(plan, prs_with_slides, label_list)
        assert result["status"] == "structural_failure"
