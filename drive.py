"""
Google Drive integration for Export Docs Tool.
Handles OAuth2 flow and file upload to the correct folder structure:
Certificados Expo / {producto} / Editables
"""

import os
import json
from io import BytesIO
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

SCOPES = ['https://www.googleapis.com/auth/drive.file']
CLIENT_ID = os.environ.get('GOOGLE_CLIENT_ID', '')
CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET', '')
REDIRECT_URI = os.environ.get('GOOGLE_REDIRECT_URI', 'https://notco-export-docs.up.railway.app/oauth/callback')
ROOT_FOLDER_NAME = 'Certificados Expo'

def get_flow():
    client_config = {
        "web": {
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [REDIRECT_URI],
        }
    }
    flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=REDIRECT_URI)
    return flow

def get_auth_url(state=None):
    flow = get_flow()
    auth_url, _ = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        state=state or ''
    )
    return auth_url

def exchange_code(code):
    flow = get_flow()
    flow.fetch_token(code=code)
    creds = flow.credentials
    return {
        'token': creds.token,
        'refresh_token': creds.refresh_token,
        'token_uri': creds.token_uri,
        'client_id': creds.client_id,
        'client_secret': creds.client_secret,
        'scopes': list(creds.scopes) if creds.scopes else SCOPES,
    }

def get_service(creds_dict):
    creds = Credentials(
        token=creds_dict['token'],
        refresh_token=creds_dict.get('refresh_token'),
        token_uri=creds_dict.get('token_uri', 'https://oauth2.googleapis.com/token'),
        client_id=creds_dict.get('client_id', CLIENT_ID),
        client_secret=creds_dict.get('client_secret', CLIENT_SECRET),
        scopes=creds_dict.get('scopes', SCOPES),
    )
    return build('drive', 'v3', credentials=creds)

def find_or_create_folder(service, name, parent_id=None):
    """Find a folder by name (and optional parent), or create it."""
    query = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    results = service.files().list(q=query, fields='files(id, name)').execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    # Create it
    meta = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
    }
    if parent_id:
        meta['parents'] = [parent_id]
    folder = service.files().create(body=meta, fields='id').execute()
    return folder['id']

def upload_file(creds_dict, filename, file_bytes, producto):
    """
    Upload file to: Certificados Expo / {producto} / Editables
    Returns the file's web view URL.
    """
    service = get_service(creds_dict)

    # Build folder structure
    root_id = find_or_create_folder(service, ROOT_FOLDER_NAME)
    product_id = find_or_create_folder(service, producto, parent_id=root_id)
    editables_id = find_or_create_folder(service, 'Editables', parent_id=product_id)

    # Upload file
    file_meta = {'name': filename, 'parents': [editables_id]}
    media = MediaIoBaseUpload(
        BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        resumable=True
    )
    uploaded = service.files().create(
        body=file_meta,
        media_body=media,
        fields='id, webViewLink'
    ).execute()

    return uploaded.get('webViewLink', '')
