from fastapi import APIRouter, Request, HTTPException, Response, Cookie
from fastapi.responses import RedirectResponse
from api.services.auth_service import generate_auth_url, process_oauth_callback

router = APIRouter(prefix="/auth", tags=["Authentication"])

@router.get("/login")
async def login():
    """
    Step 1: Redirect user to Google Authorization Page
    """
    authorization_url, state = generate_auth_url()
    
    # Store the state in a secure, HTTP-only cookie
    response = RedirectResponse(authorization_url)
    response.set_cookie(
        key="oauth_state",
        value=state,
        httponly=True,
        secure=True,
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
        raise HTTPException(status_code=400, detail="Invalid state parameter or state cookie missing. CSRF verification failed.")

    try:
        result = process_oauth_callback(code)
        
        # We can clear the state cookie now that it's been used
        response.delete_cookie('oauth_state')
        
        return result
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
