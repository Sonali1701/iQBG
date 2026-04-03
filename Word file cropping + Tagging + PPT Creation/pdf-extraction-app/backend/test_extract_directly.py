import fitz
import re
import extractor
import os

pdf_q = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
doc = fitz.open(pdf_q)

questions = extractor.find_questions(doc, start_page=3, cols=[1])
print(f"Total questions found starting from page 3 col 1: {len(questions)}")

if len(questions) > 0:
    for i, q in enumerate(questions[:5]):
        page_num = q["page"]
        col = q["col"]
        start_y = q["bbox"][1]
        
        end_q = questions[i + 1] if i + 1 < len(questions) else None
        same_col = bool(end_q and end_q["page"] == page_num and end_q["col"] == col)
        
        if same_col:
            nominal = end_q["bbox"][1]
            real_bot = extractor.content_bottom_in_col(doc[page_num], col, start_y, nominal)
            end_y = min(real_bot + extractor.BOTTOM_PAD, nominal - 1)
            if end_y < start_y: end_y = nominal
        else:
            real_bot = extractor.content_bottom_in_col(doc[page_num], col, start_y, extractor.FOOTER_Y)
            end_y = min(real_bot + extractor.BOTTOM_PAD, extractor.FOOTER_Y)

        x0, x1 = extractor.col_bounds(col)
        clip = fitz.Rect(x0, start_y, x1, end_y)
        
        print(f"Q {q['q_num']} on page {page_num} col {col}: start_y={start_y}, end_y={end_y}, clipHeight={clip.height}")
        if clip.height < 5 or clip.width < 5:
            print(" -> SKIPPED (clip too small)")
