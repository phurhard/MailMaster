import pytest
from fastapi.testclient import TestClient
from app import app
from api.core.core import settings
import urllib.parse

client = TestClient(app)

def test_settings_loaded():
    """Test that pydantic-settings loaded environment variables correctly"""
    assert settings.GMAIL_CLIENT_ID is not None
    assert settings.GMAIL_CLIENT_SECRET is not None
    assert settings.SUPABASE_URL is not None
    assert settings.SUPABASE_ANON_KEY is not None

def test_scalar_docs():
    """Test that the Scalar API docs are available"""
    response = client.get("/scalar")
    assert response.status_code == 200
    assert "Scalar" in response.text or "scalar" in response.text

def test_login_redirect():
    """Test the /auth/login endpoint redirects to Google OAuth Consent screen"""
    # Follow_redirects=False is important to catch the 302/303 redirect
    response = client.get("/auth/login", follow_redirects=False)
    
    # It should be a redirect
    assert response.status_code in (302, 303, 307, 308)
    
    # Check if we redirect to Google
    location = response.headers.get("Location", "")
    assert "accounts.google.com/o/oauth2/auth" in location
    
    # Check if client_id is in the URL
    parsed_url = urllib.parse.urlparse(location)
    query_params = urllib.parse.parse_qs(parsed_url.query)
    assert "client_id" in query_params
    assert query_params["client_id"][0] == settings.GMAIL_CLIENT_ID
    assert "redirect_uri" in query_params
    assert "auth/callback" in query_params["redirect_uri"][0]

def test_protected_emails_route_no_auth():
    """Test accessing a protected route without a JWT returns 401 Unauthorized"""
    response = client.get("/emails/search/test")
    assert response.status_code == 401

def test_oauth_callback_invalid_code():
    """
    Test the /auth/callback with an invalid code and missing state.
    Since we added state verification, it should return 400 immediately.
    """
    response = client.get("/auth/callback?code=invalid_code")
    assert response.status_code == 400
    assert "state" in response.text.lower()
