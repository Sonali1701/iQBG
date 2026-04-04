import fitz
import re
import os

PAGE_WIDTH_MID = 297
COL_1_L, COL_1_R = 24, 291
COL_2_L, COL_2_R = 304, 568
FOOTER_Y = 808
BOTTOM_PAD = 6
QUESTION_X_COL1 = (20, 250)
QUESTION_X_COL2 = (280, 560)

def col_bounds(col):
    return (COL_1_L, COL_1_R) if col == 1 else (COL_2_L, COL_2_R)

def _question_match(text):
    text = text.strip()
    compact = re.sub(r"\s+", "", text)
    patterns = [
        re.compile(r"^(\d{1,3})\.(?:\s|\(|$)"),
        re.compile(r"^(\d{1,3})\)(?:\s|\(|$)"),
        re.compile(r"^\((\d{1,3})\)(?:\s|$)"),
        re.compile(r"^Q(?:ue)?\.?\s*(\d{1,3})(?:\s|\.|\)|$)", re.IGNORECASE),
    ]
    for pattern in patterns:
        match = pattern.match(text) or pattern.match(compact)
        if match:
            return int(match.group(1))
    return None

def _looks_like_question_start(x0, col):
    col_left, col_right = col_bounds(col)
    left_band = col_left + 18
    return col_left - 8 <= x0 <= left_band

def _collapse_bilingual_duplicates(questions):
    if not questions:
        return questions

    by_page_qnum = {}
    for item in questions:
        by_page_qnum.setdefault((item["page"], item["q_num"]), set()).add(item["col"])

    paired_qnums = {
        q_num
        for (_, q_num), cols in by_page_qnum.items()
        if 1 in cols and 2 in cols
    }
    unique_qnums = {item["q_num"] for item in questions}
    if len(paired_qnums) < max(8, len(unique_qnums) // 4):
        return questions

    chosen = []
    for q_num in sorted(unique_qnums):
        same_q = [item for item in questions if item["q_num"] == q_num]
        paired_left = [
            item for item in same_q
            if item["col"] == 1 and 2 in by_page_qnum.get((item["page"], q_num), set())
        ]
        if paired_left:
            best = sorted(paired_left, key=lambda item: (item["page"], item["bbox"][1]))[0]
        else:
            left_only = [item for item in same_q if item["col"] == 1]
            fallback_pool = left_only or same_q
            best = sorted(fallback_pool, key=lambda item: (item["page"], item["bbox"][1]))[0]
        chosen.append(best)

    chosen.sort(key=lambda q: (q["q_num"], q["page"], q["bbox"][1]))
    return chosen

def find_questions(doc, start_page=1, cols=None):
    if cols is None:
        cols = [1, 2]
    questions = []
    seen = set()
    for pn in range(start_page - 1, len(doc)):
        page = doc[pn]
        page_dict = page.get_text("dict")
        for b in page_dict["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                line_text = " ".join(s["text"].strip() for s in l["spans"] if s["text"].strip())
                for s in l["spans"]:
                    t = s["text"].strip()
                    q_num = _question_match(t)
                    x0, y0, x1, y1 = s["bbox"]
                    col = 1 if x0 < PAGE_WIDTH_MID else 2
                    if col not in cols:
                        continue
                    if q_num is None and s is l["spans"][0]:
                        q_num = _question_match(line_text)
                    if q_num is None:
                        continue
                    if not _looks_like_question_start(x0, col):
                        continue
                    key = (pn, col, q_num)
                    if key in seen:
                        continue
                    seen.add(key)
                    questions.append({
                        "q_num": q_num,
                        "page": pn,
                        "col": col,
                        "bbox": [x0, y0, x1, y1],
                    })
    questions.sort(key=lambda q: (q["page"], q["col"], q["bbox"][1]))
    if set(cols) == {1, 2}:
        questions = _collapse_bilingual_duplicates(questions)
    return questions

def content_bottom_in_col(page, col, from_y, to_y):
    max_y1 = from_y
    for b in page.get_text("dict")["blocks"]:
        bx0, by0, bx1, by1 = b["bbox"]
        cx = (bx0 + bx1) / 2
        if col == 1 and cx > PAGE_WIDTH_MID: continue
        if col == 2 and cx < PAGE_WIDTH_MID: continue
        if by1 <= from_y or by0 >= FOOTER_Y: continue
        if by0 >= to_y: continue
        if by1 > max_y1: max_y1 = by1
    from_x, to_x = col_bounds(col)
    for d in page.get_drawings():
        rect = d["rect"]
        cx = (rect.x0 + rect.x1) / 2
        if cx < from_x or cx > to_x: continue
        if rect.y1 <= from_y or rect.y0 >= FOOTER_Y: continue
        if rect.y0 >= to_y: continue
        if rect.y1 > max_y1: max_y1 = rect.y1
    return max_y1

def extract_col(doc, questions, output_dir, prefix="Q"):
    os.makedirs(output_dir, exist_ok=True)
    filepaths = []
    
    for i, q in enumerate(questions):
        page_num = q["page"]
        col = q["col"]
        
        # TOP PADDING: 12px breathing space above the detected question start
        start_y = max(0, q["bbox"][1] - 12)
        
        end_q = questions[i + 1] if i + 1 < len(questions) else None
        same_page_and_col = bool(end_q and end_q["page"] == page_num and end_q["col"] == col)
        
        # BOTTOM PADDING: Move 6px above the start of the next question.
        # For the last question in a column, stop exactly at the footer (no subtraction
        # needed since there is no following question to collide with).
        if same_page_and_col:
            end_y = end_q["bbox"][1] - 6
        else:
            end_y = FOOTER_Y
            
        # Safety clamp to ensure we never have a negative height
        if end_y <= start_y:
            end_y = start_y + 20

        page = doc[page_num]
        mx0, my0, mx1, my1 = q["bbox"]
        q_str = str(q["q_num"]) + "."
        
        # LEFT PADDING: Calculate where the text dot ends to clip perfectly after it
        approx_width = (len(q_str) * 6) + 2
        mask_x1 = min(mx0 + approx_width, mx1)
        clip_x0 = mask_x1
        
        # RIGHT PADDING: Extend clip to 574 for breathing space, masking the border line into pure white padding.
        if col == 1:
            clip_x1 = PAGE_WIDTH_MID - 2  # Safe bound before the mid split box
        else:
            clip_x1 = 574                  # Extend width to 574 to add generous breathing space
        
        # --- GLOBAL PAGE LINE ERASURE ---
        # White-out all structural layout lines on the page before snapping the crop.
        pw, ph = page.rect.width, page.rect.height
        page.draw_rect(fitz.Rect(0, 0, pw, 35), color=(1, 1, 1), fill=(1, 1, 1))              # Top header
        page.draw_rect(fitz.Rect(0, FOOTER_Y, pw, ph), color=(1, 1, 1), fill=(1, 1, 1))       # Bottom footer
        page.draw_rect(fitz.Rect(0, 0, 20, ph), color=(1, 1, 1), fill=(1, 1, 1))              # Left margin
        page.draw_rect(fitz.Rect(568, 0, pw, ph), color=(1, 1, 1), fill=(1, 1, 1))            # Right margin (starts at 568 to kill the line)
        page.draw_rect(fitz.Rect(296, 0, 298, ph), color=(1, 1, 1), fill=(1, 1, 1))           # Center divider line


        
        clip = fitz.Rect(clip_x0, start_y, clip_x1, end_y)
        if clip.height < 5 or clip.width < 5: continue
        
        pix = page.get_pixmap(clip=clip, dpi=300)
        
        base = f"{prefix}{q['q_num']}"
        fname = f"{base}.png"
        c = 2
        while os.path.exists(os.path.join(output_dir, fname)):
            fname = f"{base}_{c}.png"
            c += 1
            
        full_path = os.path.join(output_dir, fname)
        pix.save(full_path)
        
        filepaths.append({"q_num": q["q_num"], "filename": fname, "filepath": full_path})
        
    return filepaths

def run_extraction(pdf_path, output_base_dir, start_page=1, bilingual=False, doc_type="QUES", mid_code="", language_code=""):
    '''
    Language string mapping logic per user request.
    '''
    doc = fitz.open(pdf_path)
    extracted_data = {}
    
    def build_prefix(ctype, lang):
        parts = []
        if ctype:
            base = "QUES" if ctype.upper().startswith("Q") else "SOL"
            parts.append(base)
        if lang: parts.append(lang.upper())
        if mid_code: parts.append(mid_code.upper())
        suffix = "Q" if ctype.upper().startswith("Q") else "S"
        
        prefix_str = "_".join(parts)
        if prefix_str: return prefix_str + "_" + suffix
        return suffix
        
    if bilingual:
        eng_dir = os.path.join(output_base_dir, "English")
        hindi_dir = os.path.join(output_base_dir, "Hindi")
        eng_qs = find_questions(doc, start_page=start_page, cols=[1])
        hin_qs = find_questions(doc, start_page=start_page, cols=[2])
        
        eng_prefix = build_prefix(doc_type, "ENG")
        hin_prefix = build_prefix(doc_type, "HIN")
        
        eng_files = extract_col(doc, eng_qs, eng_dir, prefix=eng_prefix)
        hin_files = extract_col(doc, hin_qs, hindi_dir, prefix=hin_prefix)
        extracted_data["English"] = {"dir": eng_dir, "files": eng_files, "prefix": eng_prefix}
        extracted_data["Hindi"] = {"dir": hindi_dir, "files": hin_files, "prefix": hin_prefix}
    else:
        questions = find_questions(doc, start_page=start_page)
        lang_short = "ENG" if "eng" in str(language_code).lower() else "HIN" if "hin" in str(language_code).lower() else ""
        prefix = build_prefix(doc_type, lang_short)
        files = extract_col(doc, questions, output_base_dir, prefix=prefix)
        key = language_code if language_code else "Extracted"
        extracted_data[key] = {"dir": output_base_dir, "files": files, "prefix": prefix}
        
    doc.close()
    return extracted_data


def find_questions_doc_pdf(doc, start_page=1):
    questions = []
    seen_qnums = set()
    for pn in range(max(0, start_page - 1), len(doc)):
        page = doc[pn]
        page_dict = page.get_text("dict")
        for block in page_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                line_text = " ".join(s.get("text", "").strip() for s in spans if s.get("text", "").strip()).strip()
                if not line_text:
                    continue
                first_span = spans[0]
                q_num = _question_match(first_span.get("text", "").strip()) or _question_match(line_text)
                if q_num is None or q_num in seen_qnums:
                    continue
                x0, y0, x1, y1 = first_span["bbox"]
                questions.append({
                    "q_num": q_num,
                    "page": pn,
                    "bbox": [x0, y0, x1, y1],
                })
                seen_qnums.add(q_num)
    questions.sort(key=lambda q: (q["page"], q["bbox"][1], q["bbox"][0]))
    return questions


def _iter_page_content_boxes(page):
    page_dict = page.get_text("dict")
    for block in page_dict.get("blocks", []):
        bbox = block.get("bbox")
        if not bbox:
            continue
        yield fitz.Rect(bbox)

    for drawing in page.get_drawings():
        rect = drawing.get("rect")
        if rect:
            yield fitz.Rect(rect)


def render_doc_question_preview(pdf_bytes, orig_q_no, start_page=1, dpi=200):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        questions = find_questions_doc_pdf(doc, start_page=start_page)
        target = None
        target_index = -1
        for idx, item in enumerate(questions):
            if int(item["q_num"]) == int(orig_q_no):
                target = item
                target_index = idx
                break
        if target is None:
            raise ValueError(f"Question {orig_q_no} not found in PDF.")

        page = doc[target["page"]]
        page_h = float(page.rect.height)
        start_y = max(0, float(target["bbox"][1]) - 8)

        next_q = questions[target_index + 1] if target_index + 1 < len(questions) else None
        if next_q and next_q["page"] == target["page"]:
            limit_y = max(start_y + 20, float(next_q["bbox"][1]) - 8)
        else:
            limit_y = max(start_y + 20, page_h - 18)

        x0 = None
        y0 = start_y
        x1 = None
        y1 = start_y + 20

        for rect in _iter_page_content_boxes(page):
            if rect.y1 <= start_y or rect.y0 >= limit_y:
                continue
            if x0 is None:
                x0 = rect.x0
                x1 = rect.x1
                y0 = min(y0, rect.y0)
                y1 = max(y1, rect.y1)
            else:
                x0 = min(x0, rect.x0)
                x1 = max(x1, rect.x1)
                y0 = min(y0, rect.y0)
                y1 = max(y1, rect.y1)

        if x0 is None:
            x0 = max(0, float(target["bbox"][0]) - 8)
            x1 = min(float(page.rect.width), float(target["bbox"][2]) + 420)

        clip = fitz.Rect(
            max(0, x0 - 12),
            max(0, start_y - 4),
            min(float(page.rect.width), x1 + 12),
            min(page_h, max(y1 + 12, start_y + 40)),
        )
        if clip.width < 10 or clip.height < 10:
            raise ValueError(f"Invalid crop bounds for question {orig_q_no}.")

        pix = page.get_pixmap(clip=clip, dpi=dpi, alpha=False)
        return {
            "png_bytes": pix.tobytes("png"),
            "width": pix.width,
            "height": pix.height,
            "questions_found": len(questions),
        }
    finally:
        doc.close()
