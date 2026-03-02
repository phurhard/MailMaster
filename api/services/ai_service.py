import litellm
from litellm import completion
from api.core.core import settings
from api.core.logger import setup_logger
from api.utils.nlp_utils import clean_email_body, sandwich_truncate, get_rule_based_category

litellm._turn_on_debug()
logger = setup_logger(__name__)

def summarize_email_content(content: str) -> str:
    """
    Summarize email content using an LLM.
    Uses truncation to save tokens.
    """
    if not content or len(content.strip()) < 50:
        return content

    # Clean and truncate content
    cleaned_content = clean_email_body(content)
    truncated_content = sandwich_truncate(cleaned_content, 1500, 500)

    prompt = f"""
    Please provide a concise 1-2 sentence summary of the following email.
    Focus on the main action item or the core message.
    
    Email Content:
    {truncated_content}
    
    Summary:
    """

    try:
        response = completion(
            model=settings.OPENAI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            api_key=settings.OPENAI_API_KEY
        )
        # Note: In a real system, you'd log token usage here: response.usage.total_tokens
        logger.info(response.usage.total_tokens)
        return response.choices[0].message.content.strip()
    except Exception as e:
        logger.error(f"Error generating summary: {str(e)}")
        return f"Error generating summary: {str(e)}"

def smart_categorize_email(subject: str, snippet: str, body: str, headers: dict) -> dict:
    """
    Two-tier categorization:
    1. Rule-based pre-check (cheap)
    2. LLM validation (expensive but accurate)
    """
    # Tier 1: Rules
    rule_category, confidence = get_rule_based_category(subject, snippet, headers)
    
    # If we are very confident (e.g. List-Unsubscribe exists), we could skip LLM.
    # But for now, let's use LLM to verify and refine, especially to distinguish Ads for deletion.
    
    cleaned_body = clean_email_body(body)
    truncated_body = sandwich_truncate(cleaned_body, 1000, 500)
    
    header_info = f"List-Unsubscribe: {headers.get('list_unsubscribe')}, X-Mailer: {headers.get('x_mailer')}"

    prompt = f"""
    Categorize the following email for a personal inbox management tool.
    
    Categories:
    - Priority: Personal correspondence, important updates, direct requests.
    - Newsletter: Educational content, blogs, updates you signed up for.
    - Ad: Commercial advertisements, sales, promotional offers (SAFE TO DELETE).
    - Social: LinkedIn/Facebook/Twitter notifications.
    - Billing: Receipts, invoices, payment reminders.
    - Spam: Unsolicited junk.
    - Other: Anything else.

    Metadata:
    Subject: {subject}
    Snippet: {snippet}
    Headers: {header_info}
    Rule-Check Result: {rule_category if rule_category else "Unknown"}
    
    Cleaned Body Content:
    {truncated_body}
    
    Response format:
    Category: [Category Name]
    Reasoning: [1 sentence why]
    Confidence: [0.0 - 1.0]
    """

    try:
        response = completion(
            model=settings.OPENAI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            api_key=settings.OPENAI_API_KEY
        )
        content = response.choices[0].message.content.strip()
        
        # Simple parser for the structured response
        lines = content.split('\n')
        category = "Other"
        for line in lines:
            if line.startswith("Category:"):
                category = line.replace("Category:", "").strip()
                break
        
        return {
            "category": category,
            "raw_response": content,
            "tokens_used": response.get('usage', {}).get('total_tokens', 0)
        }
    except Exception as e:
        logger.error(f"Error smart-categorizing email: {str(e)}")
        return {"category": rule_category or "Other", "error": str(e), "tokens_used": 0}

def categorize_email_content(subject: str, snippet: str) -> str:
    """
    Legacy method for backward compatibility.
    """
    res = smart_categorize_email(subject, snippet, "", {})
    return res["category"]
