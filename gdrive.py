"""
gdrive.py - Αποθήκευση αρχείων εκθέσεων στο Google Drive με OAuth
"""
import os
import io

TOKEN_FILE = os.path.expanduser("~/Documents/experts360_credentials/token.json")
CLIENT_SECRET_FILE = os.path.expanduser("~/Documents/experts360_credentials/client_secret.json")
SCOPES = ['https://www.googleapis.com/auth/drive.file']
ROOT_FOLDER_NAME = "Experts360"

_service = None

def get_service():
    global _service
    if _service:
        return _service
    try:
        from googleapiclient.discovery import build
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        import json

        creds = None

        # Πρώτα ψάχνω στα Streamlit secrets (cloud)
        token_json = None
        try:
            import streamlit as st
            token_json = st.secrets.get("GOOGLE_TOKEN", "")
        except:
            pass

        if token_json:
            creds = Credentials.from_authorized_user_info(
                json.loads(token_json), SCOPES)
        elif os.path.exists(TOKEN_FILE):
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                # Αποθήκευση ενημερωμένου token τοπικά
                if os.path.exists(os.path.dirname(TOKEN_FILE)):
                    with open(TOKEN_FILE, 'w') as f:
                        f.write(creds.to_json())
            else:
                return None
        _service = build('drive', 'v3', credentials=creds)
        return _service
    except Exception as e:
        print(f"Google Drive error: {e}")
        return None

def get_root_folder_id() -> str:
    """Βρίσκει ή δημιουργεί τον κεντρικό φάκελο Experts360."""
    svc = get_service()
    if not svc: return None
    q = f"name='{ROOT_FOLDER_NAME}' and 'root' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    res = svc.files().list(q=q, fields="files(id)").execute()
    files = res.get('files', [])
    if files:
        return files[0]['id']
    meta = {
        'name': ROOT_FOLDER_NAME,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    folder = svc.files().create(body=meta, fields='id').execute()
    return folder['id']

def get_or_create_folder(name: str, parent_id: str) -> str:
    """Βρίσκει ή δημιουργεί φάκελο."""
    svc = get_service()
    if not svc: return None
    q = f"name='{name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    res = svc.files().list(q=q, fields="files(id)").execute()
    files = res.get('files', [])
    if files:
        return files[0]['id']
    meta = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = svc.files().create(body=meta, fields='id').execute()
    return folder['id']

def upload_file(file_bytes: bytes, filename: str, folder_id: str,
                mimetype: str = 'application/octet-stream') -> str:
    """Ανεβάζει αρχείο στο Drive και επιστρέφει το webViewLink."""
    from googleapiclient.http import MediaIoBaseUpload
    svc = get_service()
    if not svc: return None
    meta = {'name': filename, 'parents': [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype)
    f = svc.files().create(body=meta, media_body=media, fields='id,webViewLink').execute()
    return f.get('webViewLink', f.get('id'))

def upload_ekthesi_files(ar_zimias: str, files: dict) -> dict:
    """
    Ανεβάζει αρχεία έκθεσης στο Drive.
    files = {'pdf': bytes, 'excel': bytes, 'photos': [(bytes, filename), ...]}
    """
    try:
        svc = get_service()
        if not svc:
            return {}

        safe_name = ar_zimias.replace('/', '_').replace(' ', '_') if ar_zimias else 'ekthesi'
        root_id = get_root_folder_id()
        if not root_id: return {}
        folder_id = get_or_create_folder(safe_name, root_id)
        if not folder_id: return {}

        links = {}

        if files.get('pdf'):
            link = upload_file(files['pdf'], f"ekthesi_{safe_name}.pdf",
                             folder_id, 'application/pdf')
            if link: links['pdf'] = link

        if files.get('excel'):
            link = upload_file(files['excel'], f"ekthesi_{safe_name}.xlsx",
                             folder_id,
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if link: links['excel'] = link

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
        root_id = get_root_folder_id()
        if not root_id: return None
        safe_name = ar_zimias.replace('/', '_').replace(' ', '_')
        q = f"name='{safe_name}' and '{root_id}' in parents and mimeType='application/vnd.google-apps.folder'"
        res = svc.files().list(q=q, fields="files(id)").execute()
        files = res.get('files', [])
        if files:
            return f"https://drive.google.com/drive/folders/{files[0]['id']}"
        return None
    except:
        return None
