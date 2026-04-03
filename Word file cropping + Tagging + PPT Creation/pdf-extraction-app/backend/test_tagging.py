import os
from dotenv import load_dotenv

import tagger

def main():
    load_dotenv()
    pdf_path = r"c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf"
    
    print("Testing manual gemini tagging...")
    try:
        res = tagger.run_gemini_tagging(pdf_path, "physics", ["Full Syllabus"], "")
        print(f"Result count: {len(res)}")
        if res:
            first_key = list(res.keys())[0]
            print(f"Sample row ({first_key}): {res[first_key]}")
        else:
            print("No tags returned.")
    except Exception as e:
        print(f"Script threw an error: {e}")

if __name__ == "__main__":
    main()
