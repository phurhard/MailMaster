from fastapi import APIRouter, Depends, HTTPException
from api.services.gmail import search_emails, get_email_message_details, init_gmail_service
from typing import List, Dict
import os

router = APIRouter(prefix="/emails", tags=["Emails"])

CLIENT_SECRET_FILE = 'client_secret.json'

# Helper to get the gmail service (for now using the default local one)
# Later this will be dynamic based on the logged-in user session
def get_gmail_service():
    return init_gmail_service(CLIENT_SECRET_FILE)

@router.get("/search/{keyword}")
async def search_emails_by_keyword(keyword: str, service = Depends(get_gmail_service)) -> List[Dict]:
    """
    Search emails containing the specified keyword
    """
    try:
        email_messages = search_emails(service, keyword, max_results=10)
        results = []
        for email in email_messages:
            email_details = get_email_message_details(service, email['id'])
            results.append({
                'id': email['id'],
                'subject': email_details.get('subject', ''),
                'from': email_details.get('sender', ''), # Changed from 'from' to 'sender' to match gmail.py
                'date': email_details.get('date', ''),
                'snippet': email_details.get('snippet', ''),
                'labels': email_details.get('label', [])
            })
        return results
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

from api.services.ai_service import summarize_email_content, categorize_email_content

@router.post("/{email_id}/summarize")
async def summarize_email(email_id: str, service = Depends(get_gmail_service)):
    """
    Summarize a specific email using AI
    """
    try:
        details = get_email_message_details(service, email_id)
        summary = summarize_email_content(details['body'])
        return {"id": email_id, "summary": summary}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/{email_id}/categorize")
async def categorize_email(email_id: str, service = Depends(get_gmail_service)):
    """
    Categorize a specific email using AI
    """
    try:
        details = get_email_message_details(service, email_id)
        category = categorize_email_content(details['subject'], details['snippet'])
        return {"id": email_id, "category": category}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/cleanup-suggestions")
async def get_cleanup_suggestions(service = Depends(get_gmail_service)):
    """
    Get suggestions for emails to clean up (large attachments, etc.)
    """
    from api.services.gmail import list_large_emails
    try:
        # Find emails > 5MB
        large_emails = list_large_emails(service, min_size_mb=5, max_results=10)
        results = []
        for email in large_emails:
            details = get_email_message_details(service, email['id'])
            results.append({
                'id': email['id'],
                'subject': details.get('subject', ''),
                'sender': details.get('sender', ''),
                'size_mb': details.get('size_estimate', 0) / (1024 * 1024),
                'date': details.get('date', '')
            })
        return {
            "large_emails": results,
            "suggestion": "Review these large emails to free up space in your account."
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
