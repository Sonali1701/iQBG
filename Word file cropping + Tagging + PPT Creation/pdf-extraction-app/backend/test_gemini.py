import os
import traceback
from dotenv import load_dotenv
from google import genai

load_dotenv()
api_key = os.environ.get("GEMINI_API_KEY")

with open("out_gemini_py.txt", "w", encoding="utf-8") as f:
    f.write(f"API Key loaded? {bool(api_key)}\n")
    if not api_key:
        exit(1)
        
    client = genai.Client(api_key=api_key)
    try:
        f.write("Calling Gemini...\n")
        res = client.models.generate_content(
            model="gemini-2.5-flash", 
            contents="Hello, just testing authentication."
        )
        f.write(f"Success! Response: {res.text}\n")
    except Exception as e:
        f.write(f"Error during execution: {e}\n")
        f.write(traceback.format_exc())
