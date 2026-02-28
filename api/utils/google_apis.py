import os
import json
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from api.core.core import settings
from api.database.database import get_user_tokens, upsert_user_tokens

# Scopes for Gmail and User Profile
SCOPES = [
    'https://mail.google.com/',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
    'openid'
]

def get_client_config():
    """Build the Google Client config dictionary from settings"""
    return {
        "web": {
            "client_id": settings.GMAIL_CLIENT_ID,
            "project_id": settings.GMAIL_PROJECT_ID,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_secret": settings.GMAIL_CLIENT_SECRET,
            "redirect_uris": [
                "http://localhost:8000/auth/callback",
                "http://127.0.0.1:8000/auth/callback"
            ]
        }
    }

def get_auth_flow(redirect_uri='http://localhost:8000/auth/callback'):
    """Create a flow instance for multi-user OAuth."""
    return Flow.from_client_config(
        get_client_config(),
        scopes=SCOPES,
        redirect_uri=redirect_uri
    )

def create_service(api_name='gmail', api_version='v1', *scopes, user_id=None):
    """
    Revised create_service that fetches user token from Supabase based on user_id.
    """
    if not user_id:
        raise ValueError("User ID is required to create a Gmail service")

    # If scopes are passed as a list in the first arg of *scopes
    if scopes and isinstance(scopes[0], (list, tuple)):
        final_scopes = list(scopes[0])
    else:
        final_scopes = list(scopes) if scopes else SCOPES

    creds = None
    
    # Fetch from Supabase
    token_data = get_user_tokens(user_id)
    if token_data and token_data.get("token"):
        creds = Credentials(
            token=token_data.get("token"),
            refresh_token=token_data.get("refresh_token"),
            token_uri=token_data.get("token_uri"),
            client_id=token_data.get("client_id"),
            client_secret=token_data.get("client_secret"),
            scopes=token_data.get("scopes", final_scopes)
        )

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            # Save the refreshed credentials back to Supabase
            upsert_user_tokens(
                user_id=user_id,
                email=token_data.get("email", ""), # Keep the old email
                tokens={
                    "token": creds.token,
                    "refresh_token": creds.refresh_token,
                    "token_uri": creds.token_uri,
                    "client_id": creds.client_id,
                    "client_secret": creds.client_secret,
                    "scopes": creds.scopes
                }
            )
        else:
            raise ValueError(f"No valid credentials found for user_id: {user_id}. They need to re-authenticate.")

    return build(api_name, api_version, credentials=creds, static_discovery=False)

def get_service_from_credentials(creds_json, api_name='gmail', api_version='v1'):
    """Create a service instance from credentials JSON (e.g., from DB)."""
    creds = Credentials.from_authorized_user_info(json.loads(creds_json))
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return build(api_name, api_version, credentials=creds, static_discovery=False)
