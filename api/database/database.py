from supabase import create_client, Client
from api.core.core import settings
from typing import Optional, Dict

# Initialize Supabase client
supabase: Client = create_client(settings.SUPABASE_URL, settings.SUPABASE_ANON_KEY)

def upsert_user_tokens(user_id: str, email: str, tokens: Dict):
    """
    Store or update a user's Google OAuth tokens in the database.
    Requires table: user_tokens (id, email, access_token, refresh_token, token_uri, client_id, client_secret, scopes)
    """
    data = {
        "id": user_id,
        "email": email,
        "access_token": tokens.get("token"),
        "refresh_token": tokens.get("refresh_token"),
        "token_uri": tokens.get("token_uri"),
        "client_id": tokens.get("client_id"),
        "client_secret": tokens.get("client_secret"),
        "scopes": tokens.get("scopes", [])
    }
    
    # Supabase UPSERT
    response = supabase.table("user_tokens").upsert(data).execute()
    return response

def get_user_tokens(user_id: str) -> Optional[Dict]:
    """ Fetch the Google OAuth tokens for a user. """
    response = supabase.table("user_tokens").select("*").eq("id", user_id).execute()
    if response.data and len(response.data) > 0:
        row = response.data[0]
        return {
            "token": row.get("access_token"),
            "refresh_token": row.get("refresh_token"),
            "token_uri": row.get("token_uri"),
            "client_id": row.get("client_id"),
            "client_secret": row.get("client_secret"),
            "scopes": row.get("scopes", [])
        }
    return None