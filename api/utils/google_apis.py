import os
import json
from google_auth_oauthlib.flow import InstalledAppFlow, Flow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# Scopes for Gmail
SCOPES = ['https://mail.google.com/']

def get_auth_flow(client_secret_file, redirect_uri='http://localhost:8000/auth/callback'):
    """Create a flow instance for multi-user OAuth."""
    return Flow.from_client_secrets_file(
        client_secret_file,
        scopes=SCOPES,
        redirect_uri=redirect_uri
    )

def create_service(client_secret_file, api_name, api_version, *scopes, prefix='', user_id=None):
    """
    Revised create_service that can handle multiple users or fall back to local token.
    """
    CLIENT_SECRET_FILE = client_secret_file
    API_SERVICE_NAME = api_name
    API_VERSION = api_version
    
    # If scopes are passed as a list in the first arg of *scopes
    if scopes and isinstance(scopes[0], (list, tuple)):
        final_scopes = list(scopes[0])
    else:
        final_scopes = list(scopes) if scopes else SCOPES

    creds = None
    token_dir = 'env'
    
    # Use user_id for token filename if provided, otherwise default
    token_filename = f"token_{API_SERVICE_NAME}_{API_VERSION}{prefix}.json"
    if user_id:
        token_filename = f"token_{user_id}_{API_SERVICE_NAME}.json"
        
    token_path = os.path.join(os.getcwd(), token_dir, token_filename)

    if not os.path.exists(os.path.dirname(token_path)):
        os.makedirs(os.path.dirname(token_path))

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, final_scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, final_scopes)
            creds = flow.run_local_server(port=0)

        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    
    return build(API_SERVICE_NAME, API_VERSION, credentials=creds, static_discovery=False)

def get_service_from_credentials(creds_json, api_name='gmail', api_version='v1'):
    """Create a service instance from credentials JSON (e.g., from DB)."""
    creds = Credentials.from_authorized_user_info(json.loads(creds_json))
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return build(api_name, api_version, credentials=creds, static_discovery=False)
