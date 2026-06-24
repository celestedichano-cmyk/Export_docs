"""
Google Drive integration for Export Docs Tool.
"""
import os
import json
import requests
from io import BytesIO

# El scope "drive.file" solo permite ver/crear archivos que la propia app
# creó — no alcanza para leer una carpeta compartida ya existente (como
# la carpeta de destino fija de abajo) que no fue creada por esta app.
# Por eso se amplía a "drive" (acceso completo a Drive), necesario para
# poder ubicar y escribir dentro de esa carpeta compartida.
SCOPES = ['https://www.googleapis.com/auth/drive']
CLIENT_ID = os.environ.get('GOOGLE_CLIENT_ID', '')
CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET', '')
REDIRECT_URI = os.environ.get('GOOGLE_REDIRECT_URI', 'https://notco-export-docs.up.railway.app/oauth/callback')

# Carpeta compartida de destino: https://drive.google.com/drive/folders/10NkxQ-2KV7Y6lMWyjTSenG1LuklqNN9n
# Dentro de ella se busca (o se crea) la carpeta del producto, y dentro de
# esa, la subcarpeta "Editables" donde se sube el archivo generado.
ROOT_FOLDER_ID = '10NkxQ-2KV7Y6lMWyjTSenG1LuklqNN9n'


def get_auth_url(state=None):
    from urllib.parse import urlencode
    params = {
        'client_id': CLIENT_ID,
        'redirect_uri': REDIRECT_URI,
        'response_type': 'code',
        'scope': ' '.join(SCOPES),
        'access_type': 'offline',
        'prompt': 'consent',
        'state': state or '',
    }
    return 'https://accounts.google.com/o/oauth2/v2/auth?' + urlencode(params)


def exchange_code(code):
    resp = requests.post('https://oauth2.googleapis.com/token', data={
        'code': code,
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'redirect_uri': REDIRECT_URI,
        'grant_type': 'authorization_code',
    })
    resp.raise_for_status()
    token = resp.json()
    return {
        'token': token.get('access_token'),
        'refresh_token': token.get('refresh_token'),
        'token_uri': 'https://oauth2.googleapis.com/token',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scopes': SCOPES,
    }


def get_service(creds_dict):
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    creds = Credentials(
        token=creds_dict['token'],
        refresh_token=creds_dict.get('refresh_token'),
        token_uri=creds_dict.get('token_uri', 'https://oauth2.googleapis.com/token'),
        client_id=creds_dict.get('client_id', CLIENT_ID),
        client_secret=creds_dict.get('client_secret', CLIENT_SECRET),
        scopes=creds_dict.get('scopes', SCOPES),
    )
    return build('drive', 'v3', credentials=creds)


def _escapar_nombre(name):
    """Escapa comillas simples para usarlas dentro de una query de Drive
    (ej. nombres de producto que traigan un apóstrofe)."""
    return str(name).replace("'", "\\'")


def find_or_create_folder(service, name, parent_id=None):
    """
    Busca una carpeta por nombre dentro de `parent_id` y, si no existe, la
    crea. Incluye supportsAllDrives/includeItemsFromAllDrives porque la
    carpeta raíz de destino es una unidad compartida (Shared Drive), no
    una carpeta común del Drive personal — sin estos parámetros, la API
    no encuentra ni puede crear contenido ahí.
    """
    nombre_escapado = _escapar_nombre(name)
    query = f"name='{nombre_escapado}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    results = service.files().list(
        q=query,
        fields='files(id, name)',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora='allDrives',
    ).execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    meta = {'name': name, 'mimeType': 'application/vnd.google-apps.folder'}
    if parent_id:
        meta['parents'] = [parent_id]
    folder = service.files().create(
        body=meta, fields='id', supportsAllDrives=True
    ).execute()
    return folder['id']


def upload_file(creds_dict, filename, file_bytes, producto):
    from googleapiclient.http import MediaIoBaseUpload
    service = get_service(creds_dict)
    product_id = find_or_create_folder(service, producto, parent_id=ROOT_FOLDER_ID)
    editables_id = find_or_create_folder(service, 'Editables', parent_id=product_id)
    file_meta = {'name': filename, 'parents': [editables_id]}
    media = MediaIoBaseUpload(
        BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        resumable=True
    )
    uploaded = service.files().create(
        body=file_meta, media_body=media, fields='id, webViewLink',
        supportsAllDrives=True,
    ).execute()
    return uploaded.get('webViewLink', '')
