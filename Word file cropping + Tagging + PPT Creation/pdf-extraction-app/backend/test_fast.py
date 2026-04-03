import requests
import json

pdf_q = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
pdf_s = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-09_Set-1_Class-11th NEET (2025-26)_Date-15-03-2026_Solutions (1).pdf"

files = {
    "question_file": ("q.pdf", open(pdf_q, "rb"), "application/pdf"),
    "solution_file": ("s.pdf", open(pdf_s, "rb"), "application/pdf")
}
data = {"question_mode": "local", "solution_mode": "local", "mode": "bilingual", "language": "English", "question_start_page": 20, "solution_start_page": 20, "action": "images_xlsx", "destination_type": "download"}

try:
    res = requests.post("http://localhost:8001/process", files=files, data=data)
    print("STATUS:", res.status_code)
    j = res.json()
    print("API DEBUG MESSAGE:", j.get("message", "No message"))
    print("OUTPUT IN PREVIEW IMAGES:", len(j.get("preview_data", {}).get("images", [])))
except Exception as e:
    print("Error:", e)
