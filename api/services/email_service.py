from api.services.gmail import search_emails, get_email_message_details, get_batch_email_details, list_large_emails, modify_email_labels, ensure_label_exists
from api.services.ai_service import summarize_email_content, smart_categorize_email
from typing import List, Dict
import html

def search_and_format_emails(service, keyword: str, limit: int = 20) -> List[Dict]:
    email_messages = search_emails(service, query=keyword, max_results=limit)
    if not email_messages:
        return []

    msg_ids = [msg['id'] for msg in email_messages]
    email_details_list = get_batch_email_details(service, msg_ids)
    
    results = []
    for email_details in email_details_list:
        # Decode HTML entities like &#39; to '
        raw_snippet = email_details.get('snippet', '')
        clean_snippet = html.unescape(raw_snippet)
        
        # Check tracking label
        labels = email_details.get('label', '')
        is_opened = "MailMaster/Opened" in labels
        
        results.append({
            'id': email_details['id'],
            'subject': email_details.get('subject', ''),
            'from': email_details.get('sender', ''),
            'to': email_details.get('recipients', ''),
            'date': email_details.get('date', ''),
            'snippet': clean_snippet,
            'body': email_details.get('body', ''),
            'attachments': email_details.get('attachments', []),
            'labels': labels,
            'opened': is_opened
        })
    return results

def fetch_sent_emails(service, limit: int = 20) -> List[Dict]:
    """Fetch emails sent by the user."""
    return search_and_format_emails(service, "in:sent", limit)

def process_email_summary(service, email_id: str) -> dict:
    details = get_email_message_details(service, email_id)
    summary = summarize_email_content(details['body'])
    return {"id": email_id, "summary": summary}


def process_email_category(service, email_id: str) -> dict:
    details = get_email_message_details(service, email_id)
    result = smart_categorize_email(
        subject=details['subject'],
        snippet=details['snippet'],
        body=details.get('body', ''),
        headers=details.get('headers', {})
    )
    
    # Apply label if category is high confidence
    category = result["category"]
    apply_category_to_email(service, email_id, category)
    
    return {
        "id": email_id, 
        "category": category, 
        "tokens_used": result.get("tokens_used", 0),
        "reasoning": result.get("raw_response", "")
    }

def apply_category_to_email(service, email_id: str, category: str):
    """Ensure a label exists, apply it, and move out of inbox."""
    if not category or category == "Other":
        return
        
    label_name = f"MailMaster/{category}"
    label_id = ensure_label_exists(service, label_name)
    # Move to label and remove from INBOX (Archiving into category)
    modify_email_labels(service, 'me', email_id, add_labels=[label_id], remove_labels=['INBOX'])

def batch_categorize_emails(service, email_ids: List[str]) -> List[dict]:
    """Process a batch of emails for categorization."""
    results = []
    # Fetch details in batch first
    details_list = get_batch_email_details(service, email_ids)
    
    for details in details_list:
        email_id = details['id']
        res = smart_categorize_email(
            subject=details['subject'],
            snippet=details['snippet'],
            body=details.get('body', ''),
            headers=details.get('headers', {})
        )
        category = res["category"]
        apply_category_to_email(service, email_id, category)
        
        results.append({
            "id": email_id,
            "category": category,
            "tokens_used": res.get("tokens_used", 0)
        })
    return results


def get_formatted_cleanup_suggestions(service, limit: int = 20) -> dict:
    large_emails = list_large_emails(service, min_size_mb=5, max_results=limit)
    if not large_emails:
        return {"large_emails": [], "suggestion": "No large emails found for cleanup."}
        
    msg_ids = [msg['id'] for msg in large_emails]
    email_details_list = get_batch_email_details(service, msg_ids)

    results = []
    for details in email_details_list:
        results.append({
            'id': details['id'],
            'subject': details.get('subject', ''),
            'sender': details.get('sender', ''),
            'size_mb': details.get('size_estimate', 0) / (1024 * 1024),
            'date': details.get('date', '')
        })
    return {
        "large_emails": results,
        "suggestion": "Review these large emails to free up space in your account."
    }
