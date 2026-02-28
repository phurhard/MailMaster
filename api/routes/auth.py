from fastapi import APIRouter, Request, HTTPException
from fastapi.responses import RedirectResponse
from api.utils.google_apis import get_auth_flow
import os

router = APIRouter(prefix="/auth", tags=["Authentication"])

CLIENT_SECRET_FILE = 'client_secret.json'

@router.get("/login")
async def login():
    """
    Step 1: Redirect user to Google Authorization Page
    """
    flow = get_auth_flow(CLIENT_SECRET_FILE)
    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent'
    )
    # In a real app, you'd store the state in a session to verify it in callback
    return RedirectResponse(authorization_url)

@router.get("/callback")
async def callback(request: Request):
    """
    Step 2: Receive authorization code and exchange for tokens
    """
    code = request.query_params.get('code')
    if not code:
        raise HTTPException(status_code=400, detail="Authorization code not found")

    flow = get_auth_flow(CLIENT_SECRET_FILE)
    try:
        flow.fetch_token(code=code)
        credentials = flow.credentials
        
        # In a real app, we would store this in Supabase/Postgres
        # For now, let's return a success message
        return {
            "message": "Successfully authenticated!",
            "token": credentials.token,
            "refresh_token": credentials.refresh_token,
            "expiry": credentials.expiry
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
