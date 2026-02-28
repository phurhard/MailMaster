from fastapi import APIRouter, Request, HTTPException, Response, Cookie
from fastapi.responses import RedirectResponse
from api.services.auth_service import generate_auth_url, process_oauth_callback
from api.core.core import settings

from api.core.logger import setup_logger

logger = setup_logger(__name__)

router = APIRouter(prefix="/auth", tags=["Authentication"])

@router.get("/login")
async def login(request: Request):
    """
    Step 1: Redirect user to Google Authorization Page
    """
    redirect_uri = str(request.url_for("callback"))
    authorization_url, state = generate_auth_url(redirect_uri)
    
    # Store the state in a secure, HTTP-only cookie
    response = RedirectResponse(authorization_url)
    response.set_cookie(
        key="oauth_state",
        value=state,
        httponly=True,
        secure=settings.PRODUCTION,
        samesite="lax",
        max_age=600
    )
    return response

@router.get("/callback")
async def callback(request: Request, response: Response, oauth_state: str | None = Cookie(default=None)):
    """
    Step 2: Receive authorization code and exchange for tokens
    """
    code = request.query_params.get('code')
    returned_state = request.query_params.get('state')

    if not code:
        raise HTTPException(status_code=400, detail="Authorization code not found")
        
    if not oauth_state or returned_state != oauth_state:
        # Logging for the local console
        logger.warning(f"CSRF Trace - Cookie State: {oauth_state}, Query State: {returned_state}")
        raise HTTPException(status_code=400, detail="Invalid state parameter or state cookie missing. CSRF verification failed.")

    try:
        redirect_uri = str(request.url_for("callback"))
        result = process_oauth_callback(code, redirect_uri)
        
        # We can clear the state cookie now that it's been used
        response.delete_cookie('oauth_state')
        
        access_token = result.get("access_token")
        frontend_redirect_url = f"{settings.FRONTEND_URL}/login/success?token={access_token}"
        return RedirectResponse(frontend_redirect_url)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
