"""
Tests for the FastAPI server (server.py).

Mocks the pipeline layer to test the HTTP interface without
requiring Aspose or LLM calls.
"""

import io
import json
import os
import sys
import tempfile
from unittest.mock import MagicMock, patch

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from fastapi.testclient import TestClient
from server import app, _pending_jobs


@pytest.fixture
def client():
    return TestClient(app)


@pytest.fixture
def sample_pptx():
    return b"PK\x03\x04" + b"\x00" * 100


@pytest.fixture(autouse=True)
def clear_jobs():
    _pending_jobs.clear()
    yield
    _pending_jobs.clear()


class TestHealthEndpoint:
    def test_returns_ok(self, client):
        resp = client.get("/health")
        assert resp.status_code == 200
        body = resp.json()
        assert body["status"] == "ok"
        assert "provider" in body
        assert "version" in body


class TestEditEndpoint:
    def test_rejects_non_pptx(self, client):
        resp = client.post(
            "/edit",
            data={"instruction": "fix slide 3"},
            files={"file": ("deck.pdf", b"fake", "application/pdf")},
        )
        assert resp.status_code == 422

    def test_rejects_missing_instruction(self, client, sample_pptx):
        resp = client.post(
            "/edit",
            files={"file": ("deck.pptx", sample_pptx)},
        )
        assert resp.status_code == 422

    @patch("server.step1_harvest")
    @patch("server.step2_plan")
    @patch("server.step3_execute")
    def test_successful_edit(self, mock_exec, mock_plan, mock_harvest,
                             client, sample_pptx):
        mock_harvest.return_value = (MagicMock(), {"label_list": ["slide_0"]})
        mock_plan.return_value = {
            "structural_changes": [],
            "content_manifest": [
                {"action": "edit_paragraph", "slide_label": "slide_0",
                 "shape_name": "Title", "para_idx": 0, "new_text": "Hello"}
            ],
        }
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            f.write(b"fake pptx content")
            out_path = f.name
        mock_exec.return_value = {
            "status": "complete",
            "output_path": out_path,
            "log": [{"action": "edit_paragraph", "status": "ok"}],
            "data_warnings": [],
        }

        resp = client.post(
            "/edit",
            data={"instruction": "change title to Hello"},
            files={"file": ("deck.pptx", sample_pptx)},
        )
        assert resp.status_code == 200
        assert "X-Change-Summary" in resp.headers
        assert len(resp.content) > 0

    @patch("server.step1_harvest")
    @patch("server.step2_plan")
    def test_plan_failure_returns_502(self, mock_plan, mock_harvest,
                                     client, sample_pptx):
        mock_harvest.return_value = (MagicMock(), {"label_list": []})
        mock_plan.return_value = None

        resp = client.post(
            "/edit",
            data={"instruction": "do something"},
            files={"file": ("deck.pptx", sample_pptx)},
        )
        assert resp.status_code == 502


class TestPlanAndExecute:
    @patch("server.step1_harvest")
    @patch("server.step2_plan")
    def test_plan_returns_job_id(self, mock_plan, mock_harvest,
                                 client, sample_pptx):
        mock_harvest.return_value = (MagicMock(), {"label_list": ["slide_0"]})
        mock_plan.return_value = {
            "structural_changes": [],
            "content_manifest": [],
        }
        resp = client.post(
            "/plan",
            data={"instruction": "summarize"},
            files={"file": ("deck.pptx", sample_pptx)},
        )
        assert resp.status_code == 200
        body = resp.json()
        assert "job_id" in body
        assert "plan" in body
        assert body["job_id"] in _pending_jobs

    def test_execute_unknown_job_returns_404(self, client):
        resp = client.post(
            "/execute/nonexistent",
            data={},
        )
        assert resp.status_code == 404

    @patch("server.step1_harvest")
    @patch("server.step2_plan")
    @patch("server.step3_execute")
    def test_plan_then_execute(self, mock_exec, mock_plan, mock_harvest,
                               client, sample_pptx):
        mock_harvest.return_value = (MagicMock(), {"label_list": ["slide_0"]})
        mock_plan.return_value = {
            "structural_changes": [],
            "content_manifest": [],
        }

        plan_resp = client.post(
            "/plan",
            data={"instruction": "summarize"},
            files={"file": ("deck.pptx", sample_pptx)},
        )
        job_id = plan_resp.json()["job_id"]

        def fake_execute(plan, deck_state, prs, provider, output_path):
            with open(output_path, "wb") as f:
                f.write(b"output content")
            return {
                "status": "complete",
                "output_path": output_path,
                "log": [],
                "data_warnings": [],
            }

        mock_exec.side_effect = fake_execute

        exec_resp = client.post(f"/execute/{job_id}", data={})
        assert exec_resp.status_code == 200
        assert job_id not in _pending_jobs


class TestAuth:
    @patch("server.API_KEY", "test-secret-key")
    def test_rejects_missing_key(self, client):
        resp = client.get("/health")
        assert resp.status_code == 200

        resp = client.post(
            "/edit",
            data={"instruction": "test"},
            files={"file": ("deck.pptx", b"PK\x03\x04" + b"\x00" * 100)},
        )
        assert resp.status_code == 401

    @patch("server.API_KEY", "test-secret-key")
    def test_accepts_valid_key(self, client, sample_pptx):
        with patch("server.step1_harvest") as mh, \
             patch("server.step2_plan") as mp, \
             patch("server.step3_execute") as me:
            mh.return_value = (MagicMock(), {"label_list": []})
            mp.return_value = {"structural_changes": [], "content_manifest": []}
            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
                f.write(b"out")
                out_path = f.name
            me.return_value = {
                "status": "complete", "output_path": out_path,
                "log": [], "data_warnings": [],
            }
            resp = client.post(
                "/edit",
                data={"instruction": "test"},
                files={"file": ("deck.pptx", sample_pptx)},
                headers={"Authorization": "Bearer test-secret-key"},
            )
            assert resp.status_code == 200


class TestBuildChangeSummary:
    def test_basic_summary(self):
        from pipeline import build_change_summary
        plan = {
            "structural_changes": [
                {"action": "clone_slide", "label": "new_summary",
                 "layout_name": "Title Slide"},
            ],
            "content_manifest": [
                {"action": "fill_placeholder"},
                {"action": "fill_placeholder"},
                {"action": "edit_run"},
            ],
        }
        result = {
            "log": [
                {"action": "clone_slide", "status": "ok"},
                {"action": "fill_placeholder", "status": "ok"},
                {"action": "fill_placeholder", "status": "ok"},
                {"action": "edit_run", "status": "error"},
            ],
            "data_warnings": [],
        }
        summary = build_change_summary(plan, result)
        assert "new_summary" in summary
        assert "3 succeeded" in summary
        assert "1 failed" in summary

    def test_empty_plan(self):
        from pipeline import build_change_summary
        summary = build_change_summary(
            {"structural_changes": [], "content_manifest": []},
            {"log": [], "data_warnings": []},
        )
        assert "No changes" in summary

    def test_data_warnings_included(self):
        from pipeline import build_change_summary
        summary = build_change_summary(
            {"structural_changes": [], "content_manifest": []},
            {"log": [], "data_warnings": ["Revenue mismatch"]},
        )
        assert "Revenue mismatch" in summary
