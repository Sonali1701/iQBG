import requests

url = 'http://localhost:8001/process'
files = {
    'question_file': ('AITS.pdf', open(r'c:\Users\SarvjeetSinghTomar\Desktop\Anti-Gravity\Word file cropping + Tagging + PPT Creation\AITS-07_Dropper_Set-01 NEET (2025-26)_Date-08-03-2026_Question paper (1).pdf', 'rb'), 'application/pdf')
}
data = {
    'question_mode': 'questions_only',
    'solution_mode': 'solutions_only',
    'mode': 'questions_only',
    'language': 'English',
    'include_tagging': 'true',
    'tagging_config': '{"Physics":["Full Syllabus"]}',
    'project_name': 'Test'
}
try:
    print('Sending POST /process...')
    r = requests.post(url, files=files, data=data, timeout=300)
    res = r.json()
    if 'preview_data' in res and 'xlsx_data' in res['preview_data']:
        xd = res['preview_data']['xlsx_data'].get('English', [])
        print(f'Got {len(xd)} excel rows.')
        if xd and len(xd) > 0:
            print(f'Row 1 sample (first 6 columns): {xd[0][:6]}')
    else:
        import json
        with open('out.json', 'w') as f:
            json.dump(res, f)
        print('Wrote missing xlsx_data response to out.json')
except Exception as e:
    print('Failed:', e)
