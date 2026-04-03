import asyncio
import base64
import json
import mimetypes
import os
import re
from pathlib import Path
from typing import Any, Dict, List

import requests
from dotenv import load_dotenv

from extractor import run_extraction

load_dotenv(Path(__file__).with_name(".env"))


QUESTION_TYPES = [
    "Single Choice",
    "Multiple Choice",
    "Matching",
    "Assertion & Reason",
    "Diagram Based",
    "Integer",
    "Subjective",
    "Unknown",
]

DIFFICULTIES = ["Easy", "Medium", "Hard"]
DEFAULT_ANALYSIS = {
    "question_type": "Unknown",
    "difficulty": "Medium",
    "subject": "",
    "chapter": "",
    "topic": "",
    "subtopic": "",
    "confidence": 0.0,
    "notes": "Pending Gemini analysis.",
}


def _safe_name(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned.strip("._") or "document"


def _read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json(path: Path, data: Any) -> None:
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def _gemini_payload(image_path: Path) -> Dict[str, Any]:
    prompt = """
You are analyzing one exam question screenshot.
Return JSON only. Do not add markdown fences.

Required JSON shape:
{
  "question_type": "Single Choice | Multiple Choice | Matching | Assertion & Reason | Diagram Based | Integer | Subjective | Unknown",
  "difficulty": "Easy | Medium | Hard",
  "subject": "",
  "chapter": "",
  "topic": "",
  "subtopic": "",
  "confidence": 0.0,
  "notes": ""
}

Rules:
- Use the visible question only.
- Keep subject/chapter/topic/subtopic short.
- If unsure, prefer "Unknown" for question_type and "Medium" for difficulty.
- confidence must be a number between 0 and 1.
""".strip()

    mime_type = mimetypes.guess_type(str(image_path))[0] or "image/png"
    encoded = base64.b64encode(image_path.read_bytes()).decode("utf-8")
    return {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"text": prompt},
                    {"inlineData": {"mimeType": mime_type, "data": encoded}},
                ],
            }
        ],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
            "maxOutputTokens": 2048,
        },
    }


def _normalize_analysis(data: Dict[str, Any]) -> Dict[str, Any]:
    q_type = str(data.get("question_type", "Unknown")).strip() or "Unknown"
    if q_type not in QUESTION_TYPES:
        q_type = "Unknown"

    difficulty = str(data.get("difficulty", "Medium")).strip().title() or "Medium"
    if difficulty not in DIFFICULTIES:
        difficulty = "Medium"

    try:
        confidence = float(data.get("confidence", 0.0))
    except (TypeError, ValueError):
        confidence = 0.0
    confidence = max(0.0, min(1.0, confidence))

    return {
        "question_type": q_type,
        "difficulty": difficulty,
        "subject": str(data.get("subject", "")).strip(),
        "chapter": str(data.get("chapter", "")).strip(),
        "topic": str(data.get("topic", "")).strip(),
        "subtopic": str(data.get("subtopic", "")).strip(),
        "confidence": confidence,
        "notes": str(data.get("notes", "")).strip(),
    }


def analyze_question_image(image_path: Path) -> Dict[str, Any]:
    api_key = os.environ.get("GEMINI_API_KEY", "").strip("'\" ")
    if not api_key:
        return {
            "question_type": "Unknown",
            "difficulty": "Medium",
            "subject": "",
            "chapter": "",
            "topic": "",
            "subtopic": "",
            "confidence": 0.0,
            "notes": "Gemini API key not configured.",
        }

    url = "https://aiplatform.googleapis.com/v1/publishers/google/models/gemini-2.5-pro:generateContent"
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
    }

    payload = _gemini_payload(image_path)
    response = requests.post(url, headers=headers, json=payload, timeout=180)
    response.raise_for_status()
    body = response.json()

    candidates = body.get("candidates", [])
    if not candidates:
        raise ValueError("Gemini returned no candidates.")

    text = candidates[0].get("content", {}).get("parts", [{}])[0].get("text", "")
    if not text:
        raise ValueError("Gemini returned empty content.")

    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()

    return _normalize_analysis(json.loads(cleaned))


def _build_bank_item(job_id: str, source_name: str, q_file: Dict[str, Any], analysis: Dict[str, Any], index: int) -> Dict[str, Any]:
    image_path = Path(q_file["filepath"])
    return {
        "id": f"{source_name}-{index}",
        "job_id": job_id,
        "source_pdf": source_name,
        "q_num": q_file["q_num"],
        "display_label": f"{source_name} / Q{q_file['q_num']}",
        "image_filename": image_path.name,
        "image_path": str(image_path),
        "image_url": "",
        "question_type": analysis["question_type"],
        "difficulty": analysis["difficulty"],
        "subject": analysis["subject"],
        "chapter": analysis["chapter"],
        "topic": analysis["topic"],
        "subtopic": analysis["subtopic"],
        "confidence": analysis["confidence"],
        "notes": analysis["notes"],
        "selected": False,
    }


async def build_question_bank(
    job_id: str,
    job_dir: Path,
    pdf_paths: List[Path],
    start_page: int = 1,
    classify: bool = False,
) -> Dict[str, Any]:
    bank_items: List[Dict[str, Any]] = []
    documents: List[Dict[str, Any]] = []

    for pdf_path in pdf_paths:
        source_name = _safe_name(pdf_path.stem)
        output_dir = job_dir / "bank" / source_name
        output_dir.mkdir(parents=True, exist_ok=True)

        extraction = await asyncio.to_thread(
            run_extraction,
            str(pdf_path),
            str(output_dir),
            start_page,
            False,
            "QUES",
            "",
            "English",
        )
        files = extraction.get("English", {}).get("files", [])
        documents.append({
            "name": source_name,
            "original_filename": pdf_path.name,
            "question_count": len(files),
        })

        for index, q_file in enumerate(files, start=1):
            if classify:
                try:
                    analysis = await asyncio.to_thread(analyze_question_image, Path(q_file["filepath"]))
                except Exception as exc:
                    analysis = dict(DEFAULT_ANALYSIS)
                    analysis["notes"] = f"Gemini analysis failed: {exc}"
            else:
                analysis = dict(DEFAULT_ANALYSIS)

            item = _build_bank_item(job_id, source_name, q_file, analysis, index)
            item["image_url"] = f"/workspace/{job_id}/bank/{source_name}/{item['image_filename']}"
            bank_items.append(item)

    summary = summarize_bank(bank_items)
    bank_data = {
        "job_id": job_id,
        "documents": documents,
        "questions": bank_items,
        "summary": summary,
    }
    _write_json(job_dir / "bank.json", bank_data)
    return bank_data


def load_bank(job_dir: Path) -> Dict[str, Any]:
    return _read_json(job_dir / "bank.json", {"documents": [], "questions": [], "summary": {}})


async def enrich_question_bank(job_dir: Path, concurrency: int = 4, progress_path: Path | None = None) -> Dict[str, Any]:
    bank = load_bank(job_dir)
    questions = bank.get("questions", [])
    total = len(questions)
    if total == 0:
        return bank

    sem = asyncio.Semaphore(max(1, concurrency))
    completed = 0

    def write_progress(status: str) -> None:
        if not progress_path:
            return
        progress_path.write_text(
            json.dumps(
                {
                    "status": status,
                    "processed_questions": completed,
                    "total_questions": total,
                },
                indent=2,
            ),
            encoding="utf-8",
        )

    async def classify_item(index: int, item: Dict[str, Any]) -> None:
        nonlocal completed
        async with sem:
            try:
                analysis = await asyncio.to_thread(analyze_question_image, Path(item["image_path"]))
            except Exception as exc:
                analysis = dict(DEFAULT_ANALYSIS)
                analysis["notes"] = f"Gemini analysis failed: {exc}"

            for key, value in analysis.items():
                item[key] = value
            questions[index] = item
            completed += 1
            bank["summary"] = summarize_bank(questions)
            _write_json(job_dir / "bank.json", bank)
            write_progress("running")

    write_progress("running")
    await asyncio.gather(*(classify_item(index, item) for index, item in enumerate(questions)))
    bank["summary"] = summarize_bank(questions)
    _write_json(job_dir / "bank.json", bank)
    write_progress("completed")
    return bank


def summarize_bank(questions: List[Dict[str, Any]]) -> Dict[str, Any]:
    difficulty_counts: Dict[str, int] = {name: 0 for name in DIFFICULTIES}
    type_counts: Dict[str, int] = {name: 0 for name in QUESTION_TYPES}

    for item in questions:
        difficulty_counts[item["difficulty"]] = difficulty_counts.get(item["difficulty"], 0) + 1
        type_counts[item["question_type"]] = type_counts.get(item["question_type"], 0) + 1

    return {
        "total_questions": len(questions),
        "difficulty_counts": difficulty_counts,
        "question_type_counts": type_counts,
    }


def choose_questions(
    questions: List[Dict[str, Any]],
    total_questions: int,
    difficulty_targets: Dict[str, int],
    type_targets: Dict[str, int],
) -> Dict[str, Any]:
    selected: List[Dict[str, Any]] = []
    used_ids = set()

    def pick_matching(items: List[Dict[str, Any]], count: int) -> None:
        for item in items:
            if len(selected) >= total_questions or count <= 0:
                break
            if item["id"] in used_ids:
                continue
            selected.append(item)
            used_ids.add(item["id"])
            count -= 1

    for difficulty, count in difficulty_targets.items():
        if count <= 0:
            continue
        pool = [q for q in questions if q["difficulty"] == difficulty]
        pick_matching(pool, count)

    for q_type, count in type_targets.items():
        if count <= 0:
            continue
        pool = [q for q in questions if q["question_type"] == q_type]
        pick_matching(pool, count)

    if len(selected) < total_questions:
        remaining = [q for q in questions if q["id"] not in used_ids]
        pick_matching(remaining, total_questions - len(selected))

    selected = selected[:total_questions]
    summary = summarize_bank(selected)

    shortages = {
        "total_shortfall": max(0, total_questions - len(selected)),
        "difficulty_shortfall": {},
        "type_shortfall": {},
    }
    for difficulty, count in difficulty_targets.items():
        actual = sum(1 for q in selected if q["difficulty"] == difficulty)
        if actual < count:
            shortages["difficulty_shortfall"][difficulty] = count - actual
    for q_type, count in type_targets.items():
        actual = sum(1 for q in selected if q["question_type"] == q_type)
        if actual < count:
            shortages["type_shortfall"][q_type] = count - actual

    return {
        "selected_questions": selected,
        "summary": summary,
        "shortages": shortages,
    }
