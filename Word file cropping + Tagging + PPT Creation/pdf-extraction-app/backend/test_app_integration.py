import requests
import json
import os

pdf_q = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
pdf_s = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-09_Set-1_Class-11th NEET (2025-26)_Date-15-03-2026_Solutions (1).pdf"

files = {
    "question_file": ("q.pdf", open(pdf_q, "rb"), "application/pdf"),
    "solution_file": ("s.pdf", open(pdf_s, "rb"), "application/pdf")
}
data = {
    "question_mode": "local",
    "solution_mode": "local",
    "mode": "bilingual",
    "language": "English",
    "question_start_page": 3,
    "solution_start_page": 2,
    "action": "images_xlsx",
    "destination_type": "download"
}

try:
    print("Sending request to port 8000...")
    res = requests.post("http://localhost:8000/process", files=files, data=data)
    print("STATUS:", res.status_code)
    
    j = res.json()
    msg = j.get("message", "No message")
    print("API DEBUG MESSAGE:", msg)
    
    imgs = j.get("preview_data", {}).get("images", [])
    print("OUTPUT IN PREVIEW IMAGES:", len(imgs))
except Exception as e:
    print("Error:", e)
