import re
from bs4 import BeautifulSoup

def clean_email_body(html_content: str) -> str:
    """
    Cleans email HTML body by stripping tags, styles, and scripts.
    Returns plain text.
    """
    if not html_content:
        return ""
    
    # Use BeautifulSoup to parse HTML and strip unwanted tags
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Remove script and style elements
    for script_or_style in soup(["script", "style"]):
        script_or_style.decompose()
    
    # Get text
    text = soup.get_text(separator=' ')
    
    # Clean whitespace
    lines = (line.strip() for line in text.splitlines())
    chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
    text = '\n'.join(chunk for chunk in chunks if chunk)
    
    return text

def sandwich_truncate(text: str, intro_len: int = 1000, footer_len: int = 500) -> str:
    """
    Truncates text by taking the first 'intro_len' characters and 
    the last 'footer_len' characters to preserve context and footers.
    """
    if len(text) <= (intro_len + footer_len + 50):
        return text
    
    intro = text[:intro_len]
    footer = text[-footer_len:]
    
    return f"{intro}\n\n[... TRUNCATED ...]\n\n{footer}"

def get_rule_based_category(subject: str, snippet: str, headers: dict) -> tuple:
    """
    Returns (category, confidence) based on simple rules.
    Confidence is 1.0 if it's a near-certain match.
    """
    subject_lower = subject.lower()
    snippet_lower = snippet.lower()
    
    # Rule 1: Newsletter/Promotional Headers
    if headers.get('list_unsubscribe'):
        return "Newsletter", 0.9
    
    # Rule 2: Ad-like keywords in subject
    ad_keywords = ['sale', 'offer', 'save', 'discount', 'deal', 'limited time', 'off', 'exclusive', 'promo']
    if any(keyword in subject_lower for keyword in ad_keywords):
        return "Newsletter", 0.7
    
    # Rule 3: Receipts/Billing
    billing_keywords = ['invoice', 'receipt', 'billing', 'subscription', 'payment', 'order confirmed', 'purchase']
    if any(keyword in subject_lower for keyword in billing_keywords):
        return "Billing", 0.8
        
    # Rule 4: Social
    social_keywords = ['linkedin', 'facebook', 'twitter', 'instagram', 'notification', 'new message from']
    if any(keyword in subject_lower for keyword in social_keywords):
        return "Social", 0.8
        
    return None, 0.0
