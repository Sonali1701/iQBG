import asyncio
import os
import sys
sys.path.insert(0, os.path.dirname(__file__))
from dotenv import load_dotenv
load_dotenv()

from tagger import run_gemini_tagging

PDF = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"

async def main():
    print("Testing asyncio.to_thread + wait_for pattern (exactly as main.py uses it)...")
    coro = asyncio.to_thread(run_gemini_tagging, PDF, "Physics", ["Full Syllabus"], "Test hint")
    try:
        res = await asyncio.wait_for(coro, timeout=300)
        print(f"Result: {len(res)} tags")
        print(f"Keys: {sorted(list(res.keys()))}")
        if res:
            key = list(res.keys())[0]
            print(f"Sample: {res[key]}")
    except Exception as e:
        import traceback
        print(f"FAILED: {e}")
        traceback.print_exc()

asyncio.run(main())
