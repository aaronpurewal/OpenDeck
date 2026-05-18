"""
FastAPI server exposing the OpenDeck pipeline as an HTTP API.

Endpoints:
  POST /edit    — upload PPTX + instruction, get modified PPTX back
  POST /plan    — upload PPTX + instruction, get plan JSON for review
  POST /execute — approve a plan, get modified PPTX back
  GET  /health  — liveness check
"""

import asyncio
import json
import os
import tempfile
import time
import urllib.parse
import uuid
from concurrent.futures import ThreadPoolExecutor

from fastapi import (
    FastAPI, File, Form, HTTPException, Request, UploadFile,
)
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask

from config import (
    API_KEY, LLM_PROVIDER, MAX_UPLOAD_MB, SERVER_HOST, SERVER_PORT, TEMP_DIR,
)
from pipeline import (
    build_change_summary, step1_harvest, step2_plan, step3_execute,
)


app = FastAPI(title="OpenDeck API", version="1.0.0")
_executor = ThreadPoolExecutor(max_workers=4)
_pending_jobs: dict[str, dict] = {}
_JOB_TTL = 3600


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def _check_auth(request: Request) -> None:
    if not API_KEY:
        return
    header = request.headers.get("Authorization", "")
    if header != f"Bearer {API_KEY}":
        raise HTTPException(status_code=401, detail="Invalid API key")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _cleanup(*paths: str) -> None:
    for p in paths:
        try:
            if p and os.path.exists(p):
                os.unlink(p)
        except OSError:
            pass


def _expire_jobs() -> None:
    now = time.time()
    expired = [k for k, v in _pending_jobs.items()
               if now - v["created_at"] > _JOB_TTL]
    for k in expired:
        job = _pending_jobs.pop(k, {})
        _cleanup(job.get("input_path", ""))


async def _save_upload(upload: UploadFile) -> str:
    if not upload.filename or not upload.filename.lower().endswith(".pptx"):
        raise HTTPException(422, "File must be a .pptx")
    content = await upload.read()
    if len(content) > MAX_UPLOAD_MB * 1024 * 1024:
        raise HTTPException(413, f"File exceeds {MAX_UPLOAD_MB}MB limit")
    os.makedirs(TEMP_DIR, exist_ok=True)
    path = os.path.join(TEMP_DIR, f"input_{uuid.uuid4().hex}.pptx")
    with open(path, "wb") as f:
        f.write(content)
    return path


def _run_full_pipeline(input_path: str, instruction: str,
                       provider: str) -> dict:
    os.makedirs(TEMP_DIR, exist_ok=True)
    output_path = os.path.join(TEMP_DIR, f"output_{uuid.uuid4().hex}.pptx")

    prs, deck_state = step1_harvest(input_path)

    plan = step2_plan(deck_state, instruction, provider)
    if plan is None:
        raise ValueError("Plan generation failed after retries")

    result = step3_execute(plan, deck_state, prs, provider, output_path)
    if result["status"] != "complete":
        msg = result.get("message", "Execution failed")
        raise ValueError(msg)

    summary = build_change_summary(plan, result)
    return {
        "output_path": result["output_path"],
        "summary": summary,
        "log": result.get("log", []),
        "data_warnings": result.get("data_warnings", []),
    }


def _run_plan_only(input_path: str, instruction: str,
                   provider: str) -> dict:
    prs, deck_state = step1_harvest(input_path)
    plan = step2_plan(deck_state, instruction, provider)
    if plan is None:
        raise ValueError("Plan generation failed after retries")
    return {
        "prs": prs,
        "deck_state": deck_state,
        "plan": plan,
        "input_path": input_path,
    }


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.get("/health")
async def health():
    return {"status": "ok", "provider": LLM_PROVIDER, "version": "1.0.0"}


@app.post("/edit")
async def edit_deck(
    request: Request,
    instruction: str = Form(...),
    file: UploadFile = File(...),
    provider: str = Form(None),
):
    _check_auth(request)
    input_path = await _save_upload(file)
    prov = provider or LLM_PROVIDER
    loop = asyncio.get_event_loop()
    try:
        result = await loop.run_in_executor(
            _executor, _run_full_pipeline, input_path, instruction, prov,
        )
    except ValueError as e:
        _cleanup(input_path)
        raise HTTPException(502, str(e))
    except Exception as e:
        _cleanup(input_path)
        raise HTTPException(500, f"Pipeline error: {e}")

    output_path = result["output_path"]
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument"
                    ".presentationml.presentation",
        filename="result.pptx",
        background=BackgroundTask(_cleanup, input_path, output_path),
        headers={
            "X-Change-Summary": urllib.parse.quote(result["summary"]),
            "X-Data-Warnings": json.dumps(result.get("data_warnings", [])),
        },
    )


@app.post("/plan")
async def plan_deck(
    request: Request,
    instruction: str = Form(...),
    file: UploadFile = File(...),
    provider: str = Form(None),
):
    _check_auth(request)
    _expire_jobs()
    input_path = await _save_upload(file)
    prov = provider or LLM_PROVIDER
    loop = asyncio.get_event_loop()
    try:
        data = await loop.run_in_executor(
            _executor, _run_plan_only, input_path, instruction, prov,
        )
    except ValueError as e:
        _cleanup(input_path)
        raise HTTPException(502, str(e))
    except Exception as e:
        _cleanup(input_path)
        raise HTTPException(500, f"Pipeline error: {e}")

    job_id = uuid.uuid4().hex[:12]
    _pending_jobs[job_id] = {
        "created_at": time.time(),
        "prs": data["prs"],
        "deck_state": data["deck_state"],
        "plan": data["plan"],
        "input_path": data["input_path"],
    }
    return {
        "job_id": job_id,
        "plan": data["plan"],
    }


@app.post("/execute/{job_id}")
async def execute_plan_endpoint(
    job_id: str,
    request: Request,
    provider: str = Form(None),
):
    _check_auth(request)
    job = _pending_jobs.pop(job_id, None)
    if job is None:
        raise HTTPException(404, "Job not found or expired")
    if time.time() - job["created_at"] > _JOB_TTL:
        _cleanup(job.get("input_path", ""))
        raise HTTPException(410, "Job expired")

    prov = provider or LLM_PROVIDER
    os.makedirs(TEMP_DIR, exist_ok=True)
    output_path = os.path.join(TEMP_DIR, f"output_{uuid.uuid4().hex}.pptx")
    input_path = job["input_path"]

    loop = asyncio.get_event_loop()
    try:
        result = await loop.run_in_executor(
            _executor,
            step3_execute,
            job["plan"], job["deck_state"], job["prs"], prov, output_path,
        )
    except Exception as e:
        _cleanup(input_path)
        raise HTTPException(500, f"Execution error: {e}")

    if result["status"] != "complete":
        _cleanup(input_path, output_path)
        raise HTTPException(
            500,
            json.dumps({"message": result.get("message", "Failed"),
                         "log": result.get("log", [])}),
        )

    summary = build_change_summary(job["plan"], result)
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument"
                    ".presentationml.presentation",
        filename="result.pptx",
        background=BackgroundTask(_cleanup, input_path, output_path),
        headers={
            "X-Change-Summary": urllib.parse.quote(summary),
            "X-Data-Warnings": json.dumps(result.get("data_warnings", [])),
        },
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host=SERVER_HOST, port=SERVER_PORT, reload=True)
