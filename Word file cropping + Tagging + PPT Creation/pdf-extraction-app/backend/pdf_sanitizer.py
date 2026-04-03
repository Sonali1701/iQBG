import re
import shutil
from collections import Counter
from pathlib import Path
from typing import Dict, List, Tuple

import fitz


WATERMARK_KEYWORDS = {
    "watermark",
    "sample",
    "demo",
    "draft",
    "confidential",
    "aakash",
    "allen",
    "unacademy",
    "byjus",
    "narayana",
    "pw",
}


def _normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().lower())


def _span_text(span: Dict) -> str:
    if "text" in span:
        return str(span.get("text", "")).strip()
    return "".join(ch.get("c", "") for ch in span.get("chars", [])).strip()


def _rgb_from_int(color_value: int) -> Tuple[int, int, int]:
    return ((color_value >> 16) & 255, (color_value >> 8) & 255, color_value & 255)


def _is_light(color_value: int) -> bool:
    r, g, b = _rgb_from_int(int(color_value or 0))
    return (r + g + b) / 3 >= 180


def _is_centered(page_rect, bbox) -> bool:
    x0, y0, x1, y1 = bbox
    cx = (x0 + x1) / 2
    cy = (y0 + y1) / 2
    return abs(cx - page_rect.width / 2) <= page_rect.width * 0.3 and abs(cy - page_rect.height / 2) <= page_rect.height * 0.3


def _is_rotated(line_dir) -> bool:
    if not line_dir:
        return False
    return abs(line_dir[1]) > 0.15


def _collect_candidates(doc: fitz.Document) -> List[Dict]:
    candidates: List[Dict] = []
    for page_index in range(len(doc)):
        page = doc[page_index]
        raw = page.get_text("rawdict")
        for block in raw.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                line_dir = line.get("dir", (1.0, 0.0))
                for span in line.get("spans", []):
                    text = _span_text(span)
                    norm = _normalize_text(text)
                    if len(norm) < 4:
                        continue

                    size = float(span.get("size", 0) or 0)
                    bbox = span.get("bbox")
                    if not bbox:
                        continue

                    keyword_hit = any(keyword in norm for keyword in WATERMARK_KEYWORDS)
                    centered = _is_centered(page.rect, bbox)
                    rotated = _is_rotated(line_dir)
                    light = _is_light(span.get("color", 0))
                    large = size >= 16

                    if keyword_hit or (centered and (rotated or light or large)):
                        candidates.append({
                            "page": page_index,
                            "text": text,
                            "norm": norm,
                            "bbox": bbox,
                        })
    return candidates


def remove_watermarks(input_pdf: Path, output_pdf: Path) -> Dict:
    doc = fitz.open(input_pdf)
    candidates = _collect_candidates(doc)
    counts = Counter(item["norm"] for item in candidates)
    removable_texts = {
        norm for norm, count in counts.items()
        if count >= 2 or any(keyword in norm for keyword in WATERMARK_KEYWORDS)
    }

    removals = 0
    for item in candidates:
        if item["norm"] not in removable_texts:
            continue
        page = doc[item["page"]]
        page.add_redact_annot(fitz.Rect(item["bbox"]), fill=(1, 1, 1))
        removals += 1

    if removals:
        for page in doc:
            page.apply_redactions()
        doc.save(output_pdf, garbage=4, deflate=True)
        doc.close()
        return {"processed_path": str(output_pdf), "watermark_removed": True, "removals": removals}

    doc.close()
    shutil.copy2(input_pdf, output_pdf)
    return {"processed_path": str(output_pdf), "watermark_removed": False, "removals": 0}
