import asyncio
from tagger import run_gemini_tagging
import os
import tempfile

async def main():
    workspaces = os.path.join(tempfile.gettempdir(), "pdf_extraction_workspace")
    jobs = [d for d in os.listdir(workspaces) if os.path.isdir(os.path.join(workspaces, d))]
    jobs.sort(key=lambda x: os.path.getmtime(os.path.join(workspaces, x)))
    latest_job = jobs[-1]
    
    pdf_path = os.path.join(workspaces, latest_job, "question.pdf")
    print(f"Testing on {pdf_path}")
    
    hint = "Identify and tag ONLY the Physics questions. IGNORE all other subjects."
    res = await asyncio.to_thread(run_gemini_tagging, pdf_path, "Physics", ["Full Syllabus"], hint)
    print("FINISHED")
    print(res)

if __name__ == "__main__":
    asyncio.run(main())
