import fitz
import re

pdf_path = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
doc = fitz.open(pdf_path)
q_pat = re.compile(r'^([1-9]|[1-9][0-9]|1[0-7][0-9]|180)\.(?:\s|\(|$)')

found = 0
for b in doc[0].get_text("dict")["blocks"]:
    if b["type"] != 0: continue
    for l in b["lines"]:
        for s in l["spans"]:
            text = s["text"].strip()
            if text and ("1" in text or "Q" in text):
                print(f"Sample text on page 1: '{text}' at bbox {s['bbox']}")
            if q_pat.match(text):
                print(f"Match found: '{text}' at bbox {s['bbox']}")
                found += 1
print(f"Total matches on page 1: {found}")
