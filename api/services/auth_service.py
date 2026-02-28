from googleapiclient.discovery import build
from api.utils.google_apis import get_auth_flow
from api.database.database import upsert_user_tokens
from api.security import create_access_token

def generate_auth_url(redirect_uri: str) -> tuple[str, str]:
    """Generate the Google OAuth consent URL and state token."""
    flow = get_auth_flow(redirect_uri)
    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent'
    )
    return authorization_url, state

def process_oauth_callback(code: str, redirect_uri: str) -> dict:
    """Process the authorization code and return the JWT session data."""
    flow = get_auth_flow(redirect_uri)
    flow.fetch_token(code=code)
    credentials = flow.credentials
    
    # Get user info
    oauth2_service = build('oauth2', 'v2', credentials=credentials, static_discovery=False)
    user_info = oauth2_service.userinfo().get().execute()
    
    google_id = user_info.get("id")
    email = user_info.get("email")

    if not google_id or not email:
        raise ValueError("Could not retrieve user id or email from Google.")

    # Save tokens to Supabase
    upsert_user_tokens(
        user_id=google_id,
        email=email,
        tokens={
            "token": credentials.token,
            "refresh_token": credentials.refresh_token,
            "token_uri": credentials.token_uri,
            "scopes": credentials.scopes
        }
    )

    # Create JWT session
    jwt_token = create_access_token(data={"sub": google_id, "email": email})
    
    return {
        "message": "Successfully authenticated!",
        "access_token": jwt_token,
        "token_type": "bearer",
        "user": {
            "id": google_id,
            "email": email
        }
    }
