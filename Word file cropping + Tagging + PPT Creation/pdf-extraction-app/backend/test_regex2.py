import fitz
import re

pdf_path = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
doc = fitz.open(pdf_path)
q_pat = re.compile(r'^([1-9]|[1-9][0-9]|1[0-7][0-9]|180)\.(?:\s|\(|$)')

with open("output_regex.txt", "w", encoding="utf-8") as f:
    f.write("Matching elements:\n")
    found = 0
    for b in doc[0].get_text("dict")["blocks"]:
        if b["type"] != 0: continue
        for l in b["lines"]:
            for s in l["spans"]:
                text = s["text"].strip()
                if text:
                    f.write(f"Sample: '{text}' at bbox {s['bbox']}\n")
                    if q_pat.match(text):
                        f.write(f"MATCH! '{text}' at bbox {s['bbox']}\n")
                        found += 1
    f.write(f"Total matches on page 1: {found}\n")
