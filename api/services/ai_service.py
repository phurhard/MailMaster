from litellm import completion
from api.core.core import settings

def summarize_email_content(content: str) -> str:
    """
    Summarize email content using an LLM.
    Uses truncation to save tokens.
    """
    if not content or len(content.strip()) < 50:
        return content

    # Truncate content to ~2000 characters to save tokens
    truncated_content = content[:2000]

    prompt = f"""
    Please provide a concise 1-2 sentence summary of the following email.
    Focus on the main action item or the core message.
    
    Email Content:
    {truncated_content}
    
    Summary:
    """

    try:
        response = completion(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            api_key=settings.OPENAI_API_KEY
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error generating summary: {str(e)}"

def categorize_email_content(subject: str, snippet: str) -> str:
    """
    Categorize email into: Priority, Newsletter, Social, Spam, Billing, or Other.
    """
    prompt = f"""
    Categorize the following email based on its subject and snippet.
    Categories: Priority, Newsletter, Social, Spam, Billing, Other.
    Return ONLY the category name.
    
    Subject: {subject}
    Snippet: {snippet}
    
    Category:
    """

    try:
        response = completion(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            api_key=settings.OPENAI_API_KEY
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return "Other"
