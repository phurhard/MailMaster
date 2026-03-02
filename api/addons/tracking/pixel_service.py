import uuid
import base64
from api.core.core import settings

def generate_tracking_pixel(email_id: str, recipient: str) -> str:
    """
    Generate a 1x1 transparent tracking pixel URL.
    This will be used for the upcoming Gmail Add-on.
    """
    tracking_id = str(uuid.uuid4())
    # In a real production environment, this would point to our API's /track endpoint
    # For now, we'll return a placeholder that we can integrate later.
    base_url = "http://localhost:8000/api/v1/emails/track"
    return f'<img src="{base_url}?tid={tracking_id}&eid={email_id}" width="1" height="1" style="display:none;" />'

def parse_tracking_labels(labels_string: str) -> dict:
    """
    Check if the Gmail labels indicate the email has been read/opened.
    Gmail doesn't natively add a 'Read' label for recipients, but some 
    third-party tools do. We will use this to track our custom 'MailMaster/Opened' label.
    """
    if not labels_string:
        return {"opened": False, "open_count": 0}
        
    labels = labels_string.split(', ')
    is_opened = any("Opened" in label for label in labels)
    return {
        "opened": is_opened,
        "open_count": 1 if is_opened else 0
    }
