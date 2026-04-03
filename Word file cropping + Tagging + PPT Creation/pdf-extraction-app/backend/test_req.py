import requests
import json
import os

url = "http://localhost:8000/process"

q_pdf = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
s_pdf = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-09_Set-1_Class-11th NEET (2025-26)_Date-15-03-2026_Solutions (1).pdf"

with open(q_pdf, "rb") as qf, open(s_pdf, "rb") as sf:
    files = {
        "question_file": ("q.pdf", qf, "application/pdf"),
        "solution_file": ("s.pdf", sf, "application/pdf")
    }
    data = {
        "question_mode": "local",
        "solution_mode": "local",
        "question_start_page": 2,
        "solution_start_page": 3,
        "mode": "bilingual",
        "action": "images_xlsx",
        "destination_type": "download"
    }

    res = requests.post(url, files=files, data=data)
    print("Status:", res.status_code)
    try:
        print(json.dumps(res.json(), indent=2))
    except:
        print(res.text)
