from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import Response
from api.utils.google_apis import create_service
from api.security import get_current_user
from api.services.gmail import get_attachment_data
from api.services.email_service import (
    search_and_format_emails,
    process_email_summary,
    process_email_category,
    get_formatted_cleanup_suggestions,
    batch_categorize_emails,
    apply_category_to_email,
    fetch_sent_emails
)
from pydantic import BaseModel
from typing import List, Dict
from api.utils.idempotency import idempotency_manager

class BatchRequest(BaseModel):
    email_ids: List[str]

class LabelRequest(BaseModel):
    label_name: str

router = APIRouter(prefix="/emails", tags=["Emails"])

def get_user_gmail_service(current_user: dict = Depends(get_current_user)):
    try:
        # Create Google API service dynamically for the logged-in user
        return create_service(user_id=current_user["id"])
    except ValueError as e:
        raise HTTPException(status_code=401, detail=str(e))

@router.get("/sent")
async def get_sent_emails(limit: int = 20, service = Depends(get_user_gmail_service)) -> List[Dict]:
    """Fetch sent emails."""
    try:
        return fetch_sent_emails(service, limit)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/search/{keyword}")
async def search_emails_by_keyword(keyword: str, limit: int = 20, service = Depends(get_user_gmail_service)) -> List[Dict]:
    """
    Search emails containing the specified keyword
    """
    try:
        return search_and_format_emails(service, keyword, limit)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/{email_id}/summarize")
async def summarize_email(email_id: str, current_user: dict = Depends(get_current_user), service = Depends(get_user_gmail_service)):
    """
    Summarize a specific email using AI
    """
    if not idempotency_manager.check_and_lock(current_user["id"], "summarize", email_id):
        raise HTTPException(status_code=429, detail="Summarization already in progress or requested recently.")
        
    try:
        return process_email_summary(service, email_id)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/{email_id}/categorize")
async def categorize_email(email_id: str, current_user: dict = Depends(get_current_user), service = Depends(get_user_gmail_service)):
    """
    Categorize a specific email using AI
    """
    if not idempotency_manager.check_and_lock(current_user["id"], "categorize", email_id):
        raise HTTPException(status_code=429, detail="Categorization already in progress or requested recently.")

    try:
        return process_email_category(service, email_id)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/cleanup-suggestions")
async def get_cleanup_suggestions(limit: int = 20, service = Depends(get_user_gmail_service)):
    """
    Get suggestions for emails to clean up (large attachments, etc.)
    """
    try:
        return get_formatted_cleanup_suggestions(service, limit)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/batch/categorize")
async def batch_categorize(request: BatchRequest, service = Depends(get_user_gmail_service)):
    """Categorize a batch of emails."""
    try:
        return batch_categorize_emails(service, request.email_ids)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/{email_id}/labels")
async def add_label(email_id: str, request: LabelRequest, service = Depends(get_user_gmail_service)):
    """Manually apply a label to an email."""
    try:
        from api.services.gmail import ensure_label_exists, modify_email_labels
        label_id = ensure_label_exists(service, request.label_name)
        modify_email_labels(service, 'me', email_id, add_labels=[label_id])
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.delete("/{email_id}/labels/{label_name}")
async def remove_label(email_id: str, label_name: str, service = Depends(get_user_gmail_service)):
    """Remove a label and return to inbox if it was a MailMaster label."""
    try:
        from api.services.gmail import modify_email_labels, list_labels
        labels = list_labels(service)
        label_id = next((l['id'] for l in labels if l['name'] == label_name), None)
        if not label_id:
            raise HTTPException(status_code=404, detail="Label not found")
            
        # Add back to inbox and remove label
        modify_email_labels(service, 'me', email_id, add_labels=['INBOX'], remove_labels=[label_id])
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/{email_id}/attachments/{attachment_id}")
async def download_email_attachment(email_id: str, attachment_id: str, filename: str, service = Depends(get_user_gmail_service)):
    """
    Download a specific attachment from an email
    """
    try:
        data = get_attachment_data(service, 'me', email_id, attachment_id)
        return Response(
            content=data,
            media_type="application/octet-stream",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/track")
async def track_email_open(tid: str, eid: str, service = Depends(get_user_gmail_service)):
    """
    Tracking pixel endpoint. Returns a 1x1 transparent GIF.
    Applies the 'MailMaster/Opened' label to the email.
    """
    import base64
    from api.services.gmail import ensure_label_exists, modify_email_labels
    try:
        label_id = ensure_label_exists(service, "MailMaster/Opened")
        modify_email_labels(service, 'me', eid, add_labels=[label_id])
    except Exception:
        pass
    
    # 1x1 Transparent GIF
    gif = base64.b64decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7")
    return Response(content=gif, media_type="image/gif")
