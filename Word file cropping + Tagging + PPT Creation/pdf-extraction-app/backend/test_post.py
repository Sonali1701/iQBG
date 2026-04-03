import requests

url = "http://localhost:8001/process"
data = {
    "question_mode": "none",
    "solution_mode": "none",
    "include_tagging": "true",
    "tagging_config": '{"Physics":["Full Syllabus"]}',
    "action": "images_xlsx",
    "mode": "bilingual",
    "destination_type": "parent",
    "question_start_page": "1",
    "solution_start_page": "1"
}

print("SENDING POST...")
res = requests.post(url, data=data)
print(f"RESPONSE STATUS: {res.status_code}")
print(f"RESPONSE TEXT: {res.text}")
