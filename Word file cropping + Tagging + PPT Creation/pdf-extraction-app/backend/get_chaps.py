import pandas as pd
import json
import os

tags_dir = r"C:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\pdf-extraction-app\backend\tags"
files = {
    'Physics': 'NEET_JEE Physics Tag.xlsx',
    'Chemistry': 'NEET_JEE Chemistry Tag.xlsx',
    'Botany': 'NEET Botany Tag.xlsx',
    'Zoology': 'NEET Zoology Tag.xlsx',
}
out = {}
for subj, fname in files.items():
    p = os.path.join(tags_dir, fname)
    if os.path.exists(p):
        df = pd.read_excel(p)
        chaps = [str(c).strip() for c in df['Chapter_name'].dropna().unique().tolist() if str(c).strip() and str(c).strip() != 'nan']
        out[subj] = sorted(list(set(chaps)))
    else:
        out[subj] = []

out['Biology'] = sorted(list(set(out.get('Botany', []) + out.get('Zoology', []))))
out['Maths'] = []

print(json.dumps(out, indent=2))
