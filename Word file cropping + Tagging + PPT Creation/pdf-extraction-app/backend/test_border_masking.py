import fitz
import extractor
import os

pdf_q = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
doc = fitz.open(pdf_q)

qs1 = extractor.find_questions(doc, start_page=3, cols=[1, 2])
print(f"Docs found: {len(qs1)}")

out_dir = "test_out"
os.makedirs(out_dir, exist_ok=True)

# Delete existing images in out_dir
for f in os.listdir(out_dir):
    os.remove(os.path.join(out_dir, f))

# Extract the very first 2 items of Col 1 and first 2 of Col 2
col1_qs = [q for q in qs1 if q["col"] == 1][:2]
col2_qs = [q for q in qs1 if q["col"] == 2][:2]

extracted = extractor.extract_col(doc, col1_qs + col2_qs, out_dir, prefix="Q")
print(f"Extracted {len(extracted)} test images.")
for ext in extracted:
    print("->", ext["filepath"])
