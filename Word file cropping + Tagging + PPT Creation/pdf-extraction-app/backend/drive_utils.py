from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import re
import time
from pathlib import Path

# Create a global drive instance lazily or upon request
drive = None

def get_drive():
    global drive
    if drive is None:
        gauth = GoogleAuth()
        # Expects client_secrets.json in the current working directory
        # Try to load saved client credentials
        gauth.LoadCredentialsFile("mycreds.txt")
        if gauth.credentials is None:
            # Authenticate if they're not there
            gauth.LocalWebserverAuth()
        elif gauth.access_token_expired:
            # Refresh them if expired
            gauth.Refresh()
        else:
            # Initialize the saved creds
            gauth.Authorize()
        gauth.SaveCredentialsFile("mycreds.txt")
        drive = GoogleDrive(gauth)
    return drive

def parse_drive_link(url: str):
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
    if m: return ("folder", m.group(1))

    m = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
    if m: return ("file", m.group(1))

    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    if m: return ("file", m.group(1))

    raise ValueError("Couldn't parse the Drive link.")

def download_file(file_id: str, local_path: Path):
    d = get_drive()
    f = d.CreateFile({'id': file_id})
    f.GetContentFile(str(local_path))
    return f

def list_pdfs_in_folder(folder_id: str):
    d = get_drive()
    q = f"'{folder_id}' in parents and trashed=false and mimeType='application/pdf'"
    return d.ListFile({'q': q}).GetList()

def get_parent_folder_id_of_file(file_id: str) -> str:
    d = get_drive()
    f = d.CreateFile({'id': file_id})
    f.FetchMetadata(fields='parents')
    parents = f.metadata.get('parents', [])
    return parents[0]['id'] if parents else 'root'

def ensure_folder(parent_id: str, folder_name: str) -> str:
    d = get_drive()
    q = (
        f"'{parent_id}' in parents and trashed=false and "
        f"mimeType='application/vnd.google-apps.folder' and title='{folder_name}'"
    )
    found = d.ListFile({'q': q}).GetList()
    if found: return found[0]['id']

    folder = d.CreateFile({
        'title': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [{'id': parent_id}]
    })
    folder.Upload()
    return folder['id']

def delete_existing_pngs(folder_id: str):
    d = get_drive()
    q = f"'{folder_id}' in parents and trashed=false"
    files = d.ListFile({'q': q}).GetList()
    for f in files:
        if f.get('title', '').lower().endswith('.png'):
            try: f.Delete()
            except: pass

def upload_file_unique(local_path: Path, folder_id: str, max_retries=5, initial_delay=1):
    d = get_drive()
    q = f"'{folder_id}' in parents and trashed=false and title='{local_path.name}'"
    existing = d.ListFile({'q': q}).GetList()
    for old in existing:
        try: old.Delete()
        except: pass

    retries = 0
    delay = initial_delay
    while retries < max_retries:
        try:
            f = d.CreateFile({
                'title': local_path.name,
                'parents': [{'id': folder_id}]
            })
            f.SetContentFile(str(local_path))
            f.Upload()
            return f
        except Exception as e:
            if "Transient failure" in str(e) or "503" in str(e):
                retries += 1
                time.sleep(delay)
                delay *= 2
            else:
                raise
    raise Exception(f"Failed to upload {local_path.name}")
