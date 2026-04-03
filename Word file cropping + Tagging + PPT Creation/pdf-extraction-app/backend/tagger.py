import os
import re
import pandas as pd
from google import genai
from google.genai import types
from typing import List, Dict, Optional
from dotenv import load_dotenv

load_dotenv()

TAGS_DIR = os.path.join(os.path.dirname(__file__), 'tags')
PROMPT_FILE = os.path.join(TAGS_DIR, 'NEET_JEE Prompts.xlsx')

SUBJECT_FILES = {
    'physics': 'NEET_JEE Physics Tag.xlsx',
    'chemistry': 'NEET_JEE Chemistry Tag.xlsx',
    'botany': 'NEET Botany Tag.xlsx',
    'zoology': 'NEET Zoology Tag.xlsx',
}

_df_cache = {}
_prompt_cache = None

def get_df(subject: str) -> Optional[pd.DataFrame]:
    subject = subject.lower().strip()
    if subject in _df_cache:
        return _df_cache[subject]
    
    filename = SUBJECT_FILES.get(subject)
    if not filename:
        return None
    path = os.path.join(TAGS_DIR, filename)
    if not os.path.exists(path):
        return None
        
    df = pd.read_excel(path)
    for c in ["RowID", "Subject", "Chapter", "Topic", "Subtopic", "Chapter_name"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()
            
    _df_cache[subject] = df
    return df

def get_biology_df() -> Optional[pd.DataFrame]:
    bot = get_df('botany')
    zoo = get_df('zoology')
    if bot is not None and zoo is not None:
        return pd.concat([bot, zoo], ignore_index=True)
    return bot if bot is not None else zoo

def get_subject_df(subject: str) -> Optional[pd.DataFrame]:
    if subject.lower() == 'biology':
        return get_biology_df()
    return get_df(subject)

def get_chapters_for_subject(subject: str) -> List[str]:
    df = get_subject_df(subject)
    if df is None or 'Chapter_name' not in df.columns:
        return ["Full Syllabus"]
    chapters = df['Chapter_name'].unique().tolist()
    chapters = [c for c in chapters if isinstance(c, str) and c.strip()]
    return ["Full Syllabus"] + sorted(chapters)

def load_prompts() -> Dict[str, str]:
    global _prompt_cache
    if _prompt_cache is not None:
        return _prompt_cache
    
    if not os.path.exists(PROMPT_FILE):
        return {}
        
    df = pd.read_excel(PROMPT_FILE)
    prompts = {}
    for _, row in df.iterrows():
        s = str(row.get('Subject', '')).strip().lower()
        p = str(row.get('Prompt', '')).strip()
        if s and p:
            prompts[s] = p
            
    _prompt_cache = prompts
    return prompts

def build_tag_csv(subject: str, chapters: List[str]) -> str:
    _LOG = os.path.join(os.path.dirname(__file__), 'tagger_internal_log.txt')
    df = get_subject_df(subject)
    if df is None:
        with open(_LOG, 'a') as f: f.write(f"[build_tag_csv] df is None for subject='{subject}'\n")
        return ""
        
    if chapters and "Full Syllabus" not in chapters:
        df = df[df['Chapter_name'].isin(chapters)]
        
    cols = ['RowID','Subject','Subject_name','Chapter','Chapter_name','Topic','Topic_name','Subtopic','Subtopic_name']
    cols = [c for c in cols if c in df.columns]
    
    return df[cols].to_csv(index=False)

def parse_tagging_csv(text: str) -> Dict[int, Dict]:
    lines = text.strip().split('\n')
    out = {}
    for line in lines:
        line = line.replace('`', '').replace('"', '').strip()
        if not line or line.lower().startswith('question'): continue
        parts = line.split(',')
        if len(parts) >= 3:
            q_str = re.sub(r'\D', '', parts[0])
            if q_str:
                q_num = int(q_str)
                out[q_num] = {
                    "RowID": parts[1].strip(),
                    "Difficulty": parts[2].strip()
                }
    return out

def fetch_tag_meta(row_id: str, subject_hint: str) -> Dict:
    row_id = str(row_id).strip()
    df = get_subject_df(subject_hint)
    if df is not None and "RowID" in df.columns:
        match = df[df["RowID"] == row_id]
        if not match.empty:
            r = match.iloc[0]
            return {
                "Subject": r.get("Subject", "") if not pd.isna(r.get("Subject")) else "",
                "Chapter": r.get("Chapter", "") if not pd.isna(r.get("Chapter")) else "",
                "Topic": r.get("Topic", "") if not pd.isna(r.get("Topic")) else "",
                "Subtopic": r.get("Subtopic", "") if not pd.isna(r.get("Subtopic")) else ""
            }
    return {"Subject": "", "Chapter": "", "Topic": "", "Subtopic": ""}

def get_difficulty_name(diff_num) -> str:
    d = str(diff_num).strip()
    if d == '1': return 'Easy'
    if d == '2': return 'Medium'
    if d == '3': return 'Hard'
    return ''

def run_gemini_tagging(pdf_path: str, subject: str, chapters: List[str], full_test_hint: str = "") -> Dict[int, Dict]:
    _LOG = os.path.join(os.path.dirname(__file__), 'tagger_internal_log.txt')
    with open(_LOG, 'a') as f: f.write(f"[run_gemini_tagging] CALLED: subject={subject}, chapters={chapters}, pdf={pdf_path}\n")
    prompts = load_prompts()
    subj_norm = subject.lower().strip()
    base_prompt = prompts.get(subj_norm, prompts.get('default', ''))
    
    wrapper = "RUNTIME CONTEXT (do not override subject rules):\n"
    wrapper += f"• Input files:\n  1) Questions.pdf\n  2) TagList.csv\n• Subject scope: {subject}.\n"
    if full_test_hint:
        wrapper += f"• IMPORTANT: {full_test_hint}\n"
    wrapper += "• CRITICAL: You must extract and output a valid row for EVERY SINGLE question belonging to this subject in the PDF. DO NOT skip any questions. DO NOT stop early.\n"
    wrapper += "• For EVERY numbered question in Questions.pdf, choose exactly ONE RowID whose Subtopic/Subtopic_name best matches the core concept.\n"
    wrapper += "• Output Difficulty as a NUMBER only (1/2/3).\n"
    wrapper += "• Output EXACTLY in this CSV format: Question #,RowID,Difficulty\n"
    wrapper += "• CRITICAL: Your very last output line MUST be: ----- END OF TAGS -----\n\n"
    wrapper += "OUTPUT (STRICT):\nQuestion #,RowID,Difficulty\n"
    
    full_prompt = base_prompt + "\n\n" + wrapper
    tag_csv = build_tag_csv(subject, chapters)
    
    if not tag_csv:
        _LOG = os.path.join(os.path.dirname(__file__), 'tagger_internal_log.txt')
        with open(_LOG, 'a') as f: f.write(f"[run_gemini_tagging] EARLY EXIT: tag_csv is empty for subject='{subject}', chapters={chapters}\n")
        print(f"Warning: Tag CSV for {subject} is empty.")
        return {}
        
    csv_wrapper = f"----- BEGIN TagList.csv -----\n{tag_csv}\n----- END TagList.csv -----"
    
    api_key = os.environ.get("GEMINI_API_KEY", "").strip("'\" ")
    if not api_key:
        print("Warning: GEMINI_API_KEY environment variable not set.")
        return {}
        
    try:
        import base64
        with open(pdf_path, 'rb') as f:
            pdf_b64 = base64.b64encode(f.read()).decode("utf-8")
    except Exception as e:
        with open('tagger_internal_log.txt', 'a') as debugfile:
            debugfile.write(f"\n[TAGGER] FAIL READ: {e}\n")
        print(f"Failed to read PDF for base64 encoding: {e}")
        return {}

    url = "https://aiplatform.googleapis.com/v1/publishers/google/models/gemini-2.5-pro:generateContent"
    
    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"text": full_prompt},
                    {"text": csv_wrapper},
                    {"inlineData": {"mimeType": "application/pdf", "data": pdf_b64}}
                ]
            }
        ],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 65536}
    }
    
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key
    }
    
    import requests
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=1200)
        resp.raise_for_status()
        data = resp.json()
        
        candidates = data.get("candidates", [])
        if not candidates:
            raise ValueError(f"No candidates in response: {data}")
            
        cand = candidates[0]
        finish_reason = cand.get("finishReason", "UNKNOWN")
        text_resp = cand.get("content", {}).get("parts", [{}])[0].get("text", "")
        
        if not text_resp:
            return {}
            
        print(f"DEBUG: finishReason={finish_reason}, length={len(text_resp)}")
        parsed = parse_tagging_csv(text_resp)
        for _, d in parsed.items():
            meta = fetch_tag_meta(d["RowID"], subject)
            d.update(meta)
            d["Difficulty_name"] = get_difficulty_name(d["Difficulty"])
            
        with open('tagger_internal_log.txt', 'a') as debugfile:
            debugfile.write(f"\n[TAGGER] SUCCESS: {len(parsed)} tags parsed for {subject}.\n")
        return parsed
        
    except Exception as e:
        error_msg = str(e)
        if hasattr(e, 'response') and e.response is not None:
             error_msg += f" \nBody: {e.response.text}"
        print(f"Failed to fetch from Gemini Vertex REST API: {error_msg}")
        with open('tagger_internal_log.txt', 'a') as debugfile:
            debugfile.write(f"\n[TAGGER] EXCEPTION POST: {error_msg}\n")
        return {}
