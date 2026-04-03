import fitz
import extractor
import re

pdf_q = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"

doc = fitz.open(pdf_q)

qs1 = extractor.find_questions(doc, start_page=3, cols=[1])
qs2 = extractor.find_questions(doc, start_page=3, cols=[2])

print(f"Start Page 3 - Col 1 questions: {len(qs1)}")
print(f"Start Page 3 - Col 2 questions: {len(qs2)}")
