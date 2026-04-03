import fitz
import re
pdf_path = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
doc = fitz.open(pdf_path)
q_pat = re.compile(r'^([1-9]|[1-9][0-9]|1[0-7][0-9]|180)\.(?:\s|\(|$)')

total = 0
for pn in range(len(doc)):
    found = 0
    for b in doc[pn].get_text("dict")["blocks"]:
        if b["type"] != 0: continue
        for l in b["lines"]:
            for s in l["spans"]:
                text = s["text"].strip()
                m = q_pat.match(text)
                if m:
                    print(f"Page {pn+1}: MATCH '{text}' at {s['bbox']}")
                    found += 1
    if found > 0:
        print(f"Page {pn+1} summary: {found} questions")
    total += found
print(f"Total questions found: {total}")
