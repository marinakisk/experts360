"""
gdrive_auth.py - Εκτέλεσε αυτό ΜΙΑ ΦΟΡΑ για να πάρεις το token.
"""
import os
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import json

SCOPES = ['https://www.googleapis.com/auth/drive.file']
TOKEN_FILE = os.path.expanduser("~/Documents/experts360_credentials/token.json")
CLIENT_SECRET_FILE = os.path.expanduser("~/Documents/experts360_credentials/client_secret.json")

def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as f:
            f.write(creds.to_json())
    return creds

if __name__ == "__main__":
    creds = get_credentials()
    print("✅ Authorization OK!")
    print(f"Token saved: {TOKEN_FILE}")
