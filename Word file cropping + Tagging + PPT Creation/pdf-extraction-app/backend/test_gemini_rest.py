import os
import requests
from dotenv import load_dotenv

load_dotenv()
api_key = os.environ.get("GEMINI_API_KEY")

if not api_key:
    print("NO API KEY")
    exit(1)
    
url = "https://aiplatform.googleapis.com/v1/publishers/google/models/gemini-2.5-pro:generateContent"

payload = {
    "contents": [
        {
            "role": "user",
            "parts": [
                {"text": "Who are you?"}
            ]
        }
    ]
}

headers = {
    "Content-Type": "application/json",
    "X-Goog-Api-Key": api_key
}

print("Calling REST URL with X-Goog-Api-Key header...")
try:
    response = requests.post(url, headers=headers, json=payload)
    print("STATUS:", response.status_code)
    print("BODY:", response.text)
except Exception as e:
    print("Error:", e)
