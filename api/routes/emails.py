from fastapi import APIRouter, Depends, HTTPException
from api.utils.google_apis import create_service
from api.security import get_current_user
from api.services.email_service import (
    search_and_format_emails,
    process_email_summary,
    process_email_category,
    get_formatted_cleanup_suggestions
)
from typing import List, Dict

router = APIRouter(prefix="/emails", tags=["Emails"])

def get_user_gmail_service(current_user: dict = Depends(get_current_user)):
    try:
        # Create Google API service dynamically for the logged-in user
        return create_service(user_id=current_user["id"])
    except ValueError as e:
        raise HTTPException(status_code=401, detail=str(e))

@router.get("/search/{keyword}")
async def search_emails_by_keyword(keyword: str, service = Depends(get_user_gmail_service)) -> List[Dict]:
    """
    Search emails containing the specified keyword
    """
    try:
        return search_and_format_emails(service, keyword)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/{email_id}/summarize")
async def summarize_email(email_id: str, service = Depends(get_user_gmail_service)):
    """
    Summarize a specific email using AI
    """
    try:
        return process_email_summary(service, email_id)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/{email_id}/categorize")
async def categorize_email(email_id: str, service = Depends(get_user_gmail_service)):
    """
    Categorize a specific email using AI
    """
    try:
        return process_email_category(service, email_id)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/cleanup-suggestions")
async def get_cleanup_suggestions(service = Depends(get_user_gmail_service)):
    """
    Get suggestions for emails to clean up (large attachments, etc.)
    """
    try:
        return get_formatted_cleanup_suggestions(service)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
