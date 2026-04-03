from fastapi.testclient import TestClient
from main import app

client = TestClient(app)

q_pdf = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
s_pdf = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-09_Set-1_Class-11th NEET (2025-26)_Date-15-03-2026_Solutions (1).pdf"

with open(q_pdf, 'rb') as qf, open(s_pdf, 'rb') as sf:
    data = {
        "question_mode": "local",
        "solution_mode": "local",
        "question_start_page": "1",
        "solution_start_page": "1",
        "mode": "bilingual",
        "action": "images_xlsx",
        "destination_type": "download"
    }
    files = {
        "question_file": ("q.pdf", qf, "application/pdf"),
        "solution_file": ("s.pdf", sf, "application/pdf")
    }
    response = client.post("/process", data=data, files=files)
    print("Status:", response.status_code)
    try:
        data = response.json()
        print("Images count:", len(data.get("preview_data", {}).get("images", [])))
        print("XLSX files:", len(data.get("preview_data", {}).get("xlsx_files", [])))
    except Exception as e:
        print(e)
