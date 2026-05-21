"""
gdrive.py - Αποθήκευση αρχείων εκθέσεων στο Google Drive
"""
import os
import io

import os as _os
# Ψάχνει το key σε πολλά μέρη
_KEY_PATHS = [
    _os.path.expanduser("~/Documents/experts360_credentials/service_account.json"),
    _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "service_account.json"),
    "/etc/experts360/service_account.json",
]
SERVICE_ACCOUNT_FILE = next((p for p in _KEY_PATHS if _os.path.exists(p)), None)
print(f"Service account file: {SERVICE_ACCOUNT_FILE}")
ROOT_FOLDER_ID = "1EWHhMHCmKPU2mHBJUBLBd_GuniExpSRh"
SCOPES = ['https://www.googleapis.com/auth/drive']

_service = None

def get_service():
    global _service
    if _service:
        return _service
    try:
        from googleapiclient.discovery import build
        from google.oauth2 import service_account
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        _service = build('drive', 'v3', credentials=creds)
        return _service
    except Exception as e:
        print(f"Google Drive error: {e}")
        return None

def get_or_create_folder(name: str, parent_id: str) -> str:
    """Βρίσκει ή δημιουργεί φάκελο."""
    svc = get_service()
    if not svc:
        return None
    # Ψάχνω αν υπάρχει
    q = f"name='{name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    res = svc.files().list(q=q, fields="files(id)").execute()
    files = res.get('files', [])
    if files:
        return files[0]['id']
    # Δημιουργώ νέο
    meta = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = svc.files().create(body=meta, fields='id').execute()
    return folder['id']

def upload_file(file_bytes: bytes, filename: str, folder_id: str,
                mimetype: str = 'application/octet-stream') -> str:
    """Ανεβάζει αρχείο στο Drive και επιστρέφει το file ID."""
    from googleapiclient.http import MediaIoBaseUpload
    svc = get_service()
    if not svc:
        return None
    meta = {'name': filename, 'parents': [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype)
    f = svc.files().create(body=meta, media_body=media, fields='id,webViewLink').execute()
    return f.get('webViewLink', f.get('id'))

def upload_ekthesi_files(ar_zimias: str, files: dict) -> dict:
    """
    Ανεβάζει αρχεία έκθεσης στο Drive.
    files = {'pdf': bytes, 'excel': bytes, 'photos': [(bytes, filename), ...]}
    Επιστρέφει dict με links.
    """
    try:
        svc = get_service()
        if not svc:
            return {}

        # Φάκελος για αυτή την έκθεση
        safe_name = ar_zimias.replace('/', '_').replace(' ', '_') if ar_zimias else 'ekthesi'
        folder_id = get_or_create_folder(safe_name, ROOT_FOLDER_ID)
        if not folder_id:
            return {}

        links = {}

        # PDF
        if files.get('pdf'):
            link = upload_file(files['pdf'], f"ekthesi_{safe_name}.pdf",
                             folder_id, 'application/pdf')
            if link: links['pdf'] = link

        # Excel
        if files.get('excel'):
            link = upload_file(files['excel'], f"ekthesi_{safe_name}.xlsx",
                             folder_id,
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if link: links['excel'] = link

        # Φωτογραφίες
        if files.get('photos'):
            photos_folder = get_or_create_folder('photos', folder_id)
            for i, (photo_bytes, photo_name) in enumerate(files['photos']):
                link = upload_file(photo_bytes, photo_name or f"photo_{i+1}.jpg",
                                 photos_folder, 'image/jpeg')
                if link:
                    links[f'photo_{i+1}'] = link

        return links
    except Exception as e:
        print(f"Drive upload error: {e}")
        return {}

def get_ekthesi_folder_link(ar_zimias: str) -> str:
    """Επιστρέφει το link του φακέλου έκθεσης."""
    try:
        svc = get_service()
        if not svc: return None
        safe_name = ar_zimias.replace('/', '_').replace(' ', '_')
        q = f"name='{safe_name}' and '{ROOT_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder'"
        res = svc.files().list(q=q, fields="files(id,webViewLink)").execute()
        files = res.get('files', [])
        if files:
            fid = files[0]['id']
            return f"https://drive.google.com/drive/folders/{fid}"
        return None
    except:
        return None
