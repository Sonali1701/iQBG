import asyncio
import base64
import json
import os
import tempfile
from pathlib import Path
from typing import Dict, List, Optional
from uuid import uuid4

from fastapi import BackgroundTasks, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from document_exports import prepare_export_assets
from extractor import render_doc_question_preview
from pdf_sanitizer import remove_watermarks
from question_bank import build_question_bank, choose_questions, enrich_question_bank, load_bank, summarize_bank


BASE_DIR = Path(__file__).resolve().parent
APP_DIR = BASE_DIR.parent
FRONTEND_DIR = APP_DIR / "frontend-static"
WORKSPACE_DIR = Path(tempfile.gettempdir()) / "pdf_extraction_workspace"
WORKSPACE_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Question Bank Paper Builder")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/workspace", StaticFiles(directory=str(WORKSPACE_DIR)), name="workspace")
app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")


class PaperRequest(BaseModel):
    project_name: str = "Generated Paper"
    total_questions: int = Field(..., ge=1)
    difficulty_targets: Dict[str, int] = Field(default_factory=dict)
    type_targets: Dict[str, int] = Field(default_factory=dict)


class FinalizeRequest(BaseModel):
    project_name: str = "Generated Paper"
    selected_ids: List[str]


class PreviewSourcePdf(BaseModel):
    sourceDocId: str
    filename: str
    pdfBase64: str


class PreviewItemRequest(BaseModel):
    itemKey: str
    finalQNo: int = 0
    origQNo: int = Field(..., ge=1)
    sourceDocId: str
    questionFileId: str = ""
    chapter: str = ""
    questionType: str = ""
    previewKind: str = ""
    sourcePdf: PreviewSourcePdf


def _job_dir(job_id: str) -> Path:
    safe_job = "".join(ch for ch in job_id if ch.isalnum() or ch in "-_")
    if not safe_job:
        raise HTTPException(status_code=400, detail="Invalid job id.")
    path = WORKSPACE_DIR / safe_job
    if not path.exists():
        raise HTTPException(status_code=404, detail="Job not found.")
    return path


def _write_json(path: Path, data) -> None:
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def _read_json(path: Path, default):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def _status_path(job_dir: Path) -> Path:
    return job_dir / "analysis_status.json"


def _write_status(job_dir: Path, data) -> None:
    _write_json(_status_path(job_dir), data)


def _read_status(job_dir: Path):
    return _read_json(
        _status_path(job_dir),
        {"status": "unknown", "processed_questions": 0, "total_questions": 0},
    )


def _background_enrich(job_dir: Path) -> None:
    try:
        asyncio.run(enrich_question_bank(job_dir, concurrency=4, progress_path=_status_path(job_dir)))
    except Exception as exc:
        _write_status(
            job_dir,
            {
                "status": "failed",
                "processed_questions": 0,
                "total_questions": len(load_bank(job_dir).get("questions", [])),
                "error": str(exc),
            },
        )


def _check_preview_token(authorization: Optional[str]) -> None:
    expected = os.getenv("PREVIEW_API_TOKEN", "").strip()
    if not expected:
        return
    provided = (authorization or "").strip()
    if provided.lower().startswith("bearer "):
        provided = provided[7:].strip()
    if provided != expected:
        raise HTTPException(status_code=401, detail="Unauthorized preview request.")


@app.get("/api/health")
def health_check():
    return {"status": "ok"}


@app.post("/api/doc-preview-assets/item")
def doc_preview_asset_item(request: PreviewItemRequest, authorization: Optional[str] = Header(default=None)):
    _check_preview_token(authorization)

    source_pdf = request.sourcePdf or None
    if source_pdf is None or not source_pdf.pdfBase64:
        raise HTTPException(status_code=400, detail="Missing sourcePdf.pdfBase64.")

    try:
        pdf_bytes = base64.b64decode(source_pdf.pdfBase64)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid source PDF payload: {exc}") from exc

    try:
        rendered = render_doc_question_preview(pdf_bytes, request.origQNo, start_page=1, dpi=200)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Preview rendering failed: {exc}") from exc

    return {
        "itemKey": request.itemKey,
        "origQNo": request.origQNo,
        "previewImageBase64": base64.b64encode(rendered["png_bytes"]).decode("ascii"),
        "previewImageWidth": rendered["width"],
        "previewImageHeight": rendered["height"],
        "previewImageMode": "external",
        "previewKind": request.previewKind or "rich",
    }


@app.get("/")
def serve_frontend():
    return FileResponse(FRONTEND_DIR / "index.html")


@app.post("/api/analyze-pdfs")
async def analyze_pdfs(background_tasks: BackgroundTasks, files: List[UploadFile] = File(...), start_page: int = Form(1)):
    if not files:
        raise HTTPException(status_code=400, detail="Please upload at least one PDF.")

    job_id = uuid4().hex[:12]
    job_dir = WORKSPACE_DIR / job_id
    original_dir = job_dir / "uploads" / "original"
    cleaned_dir = job_dir / "uploads" / "cleaned"
    original_dir.mkdir(parents=True, exist_ok=True)
    cleaned_dir.mkdir(parents=True, exist_ok=True)

    pdf_paths: List[Path] = []
    preprocess_records = []
    for upload in files:
        if not upload.filename or not upload.filename.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail=f"Invalid file: {upload.filename}")
        original_path = original_dir / upload.filename
        cleaned_path = cleaned_dir / upload.filename
        original_path.write_bytes(await upload.read())

        sanitize_result = await asyncio.to_thread(remove_watermarks, original_path, cleaned_path)
        pdf_paths.append(cleaned_path)
        preprocess_records.append({
            "original_filename": upload.filename,
            "watermark_removed": sanitize_result["watermark_removed"],
            "removals": sanitize_result["removals"],
        })

    bank_data = await build_question_bank(job_id, job_dir, pdf_paths, start_page=start_page, classify=False)
    if preprocess_records:
        preprocess_map = {item["original_filename"]: item for item in preprocess_records}
        for document in bank_data.get("documents", []):
            extra = preprocess_map.get(document.get("original_filename"))
            if extra:
                document.update(extra)
    _write_json(job_dir / "selection.json", {"selected_ids": []})
    _write_status(
        job_dir,
        {
            "status": "queued",
            "processed_questions": 0,
            "total_questions": len(bank_data.get("questions", [])),
        },
    )
    background_tasks.add_task(_background_enrich, job_dir)

    return {
        **bank_data,
        "analysis_status": _read_status(job_dir),
    }


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str):
    job_dir = _job_dir(job_id)
    bank = load_bank(job_dir)
    selection = _read_json(job_dir / "selection.json", {"selected_ids": []})
    return {
        "job_id": job_id,
        "documents": bank.get("documents", []),
        "questions": bank.get("questions", []),
        "summary": bank.get("summary", summarize_bank(bank.get("questions", []))),
        "selected_ids": selection.get("selected_ids", []),
        "analysis_status": _read_status(job_dir),
    }


@app.get("/api/jobs/{job_id}/status")
def get_job_status(job_id: str):
    job_dir = _job_dir(job_id)
    bank = load_bank(job_dir)
    return {
        "job_id": job_id,
        "analysis_status": _read_status(job_dir),
        "summary": bank.get("summary", summarize_bank(bank.get("questions", []))),
        "documents": bank.get("documents", []),
    }


@app.post("/api/jobs/{job_id}/build-paper")
def build_paper(job_id: str, request: PaperRequest):
    job_dir = _job_dir(job_id)
    status = _read_status(job_dir)
    if status.get("status") not in {"completed"}:
        raise HTTPException(status_code=409, detail="Analysis is still running. Please wait until tagging finishes.")
    bank = load_bank(job_dir)
    questions = bank.get("questions", [])
    if not questions:
        raise HTTPException(status_code=400, detail="No analyzed questions found for this job.")

    suggested = choose_questions(
        questions,
        request.total_questions,
        request.difficulty_targets,
        request.type_targets,
    )
    selected_ids = [item["id"] for item in suggested["selected_questions"]]
    _write_json(job_dir / "selection.json", {"selected_ids": selected_ids})

    return {
        "job_id": job_id,
        "project_name": request.project_name,
        "selected_questions": suggested["selected_questions"],
        "summary": suggested["summary"],
        "shortages": suggested["shortages"],
        "all_questions": questions,
    }


@app.post("/api/jobs/{job_id}/finalize")
def finalize_paper(job_id: str, request: FinalizeRequest):
    job_dir = _job_dir(job_id)
    bank = load_bank(job_dir)
    question_map = {item["id"]: item for item in bank.get("questions", [])}

    if not request.selected_ids:
        raise HTTPException(status_code=400, detail="No selected questions received.")

    selected_questions = []
    for selected_id in request.selected_ids:
        item = question_map.get(selected_id)
        if item:
            selected_questions.append(item)

    if not selected_questions:
        raise HTTPException(status_code=400, detail="Selected questions are not available.")

    exports = prepare_export_assets(job_dir, selected_questions, request.project_name.strip() or "Generated Paper")
    _write_json(job_dir / "selection.json", {"selected_ids": request.selected_ids})

    return {
        "job_id": job_id,
        "status": "success",
        "downloads": {
            "zip": f"/workspace/{job_id}/{Path(exports['zip_path']).name}",
            "xlsx": f"/workspace/{job_id}/final_output/{Path(exports['xlsx_path']).name}",
            "word": f"/workspace/{job_id}/final_output/{Path(exports['word_path']).name}",
            "pdf": f"/workspace/{job_id}/final_output/{Path(exports['pdf_path']).name}",
            "ppt": [f"/workspace/{job_id}/final_output/{Path(path).name}" for path in exports["ppt_paths"]],
        },
    }


@app.get("/api/jobs/{job_id}/download/{artifact}")
def artifact_info(job_id: str, artifact: str):
    job_dir = _job_dir(job_id)
    export_dir = job_dir / "final_output"
    options = {
        "xlsx": next(export_dir.glob("*.xlsx"), None),
        "word": next(export_dir.glob("*.docx"), None),
        "pdf": next(export_dir.glob("*.pdf"), None),
        "zip": next(job_dir.glob("*_bundle.zip"), None),
    }
    path = options.get(artifact)
    if path is None or not path.exists():
        raise HTTPException(status_code=404, detail="Artifact not found.")
    return FileResponse(path)
