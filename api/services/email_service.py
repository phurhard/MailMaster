from api.services.gmail import search_emails, get_email_message_details, get_batch_email_details, list_large_emails
from api.services.ai_service import summarize_email_content, categorize_email_content
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
        
        results.append({
            'id': email_details['id'],
            'subject': email_details.get('subject', ''),
            'from': email_details.get('sender', ''),
            'date': email_details.get('date', ''),
            'snippet': clean_snippet,
            'labels': email_details.get('label', [])
        })
    return results

def process_email_summary(service, email_id: str) -> dict:
    details = get_email_message_details(service, email_id)
    summary = summarize_email_content(details['body'])
    return {"id": email_id, "summary": summary}


def process_email_category(service, email_id: str) -> dict:
    details = get_email_message_details(service, email_id)
    category = categorize_email_content(details['subject'], details['snippet'])
    return {"id": email_id, "category": category}


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
