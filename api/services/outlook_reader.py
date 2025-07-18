import imaplib
import email
from email.header import decode_header
from textblob import TextBlob
import spacy
import tensorflow as tf
from textblob import TextBlob
import imaplib
import email
from email.header import decode_header


def connect_to_gmail(username: str, password: str):
    """Connect to Gmail using IMAP."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(username, password)
        return mail
    except Exception as e:
        print(f"Error connecting to Gmail: {e}")
        return None


def get_inbox(mail):
    """Get the inbox folder."""
    mail.select("inbox")
    return mail


def read_emails_gmail(mail, num_emails=10):
    """Read the specified number of emails from inbox."""
    try:
        status, messages = mail.search(None, "ALL")
        email_ids = messages[0].split()
        email_ids = email_ids[-num_emails:]  # Get the last num_emails

        for e_id in email_ids:
            status, msg_data = mail.fetch(e_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else 'utf-8')
            print(f"Subject: {subject}")
            print(f"From: {msg['From']}")
            print(f"Date: {msg['Date']}")
            body_preview = msg.get_payload(decode=True)
            if isinstance(body_preview, bytes):
                body_preview = body_preview.decode()  # Decode if it's in bytes
            print(f"Body Preview: {body_preview[:100]}...")
            if isinstance(body_preview, bytes):
                body_preview = body_preview.decode()  # Decode if it's in bytes

    except Exception as e:
        print(f"Error reading emails: {e}")


def main():
    # Connect to Outlook
    namespace = connect_to_outlook()
    if not namespace:
        return

    # Get inbox
    inbox = get_inbox(namespace)
    if not inbox:
        return

    # Read last 5 emails
    print("Reading the last 5 emails from your inbox...")
    read_emails(inbox, 5)

def analyze_sentiment(email_body: str) -> float:
    """Analyze the sentiment of the email body and return a sentiment score."""
    analysis = TextBlob(email_body)
    return analysis.sentiment.polarity

def categorize_email(sentiment_score: float) -> str:
    """Categorize the email based on sentiment score."""
    if sentiment_score > 0.5:
        return "High Priority"
    elif sentiment_score > 0:
        return "Medium Priority"
    else:
        return "Low Priority"

def summarize_email(email_body: str) -> str:
    """Summarize the email content."""
    # Simple summarization logic (can be improved)
    return email_body[:100] + "..."


def connect_to_gmail_imap(username: str, password: str):
    """Connect to Gmail using IMAP."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(username, password)
        return mail
    except Exception as e:
        print(f"Error connecting to Gmail: {e}")
        return None

def get_inbox_gmail(mail):
    """Get the inbox folder."""
    mail.select("inbox")
    return mail

def read_emails(mail, num_emails=10):
    """Read the specified number of emails from inbox."""
    try:
        status, messages = mail.search(None, "ALL")
        email_ids = messages[0].split()
        email_ids = email_ids[-num_emails:]  # Get the last num_emails

        for e_id in email_ids:
            status, msg_data = mail.fetch(e_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else 'utf-8')
            print(f"Subject: {subject}")
            print(f"From: {msg['From']}")
            print(f"Date: {msg['Date']}")
            print(f"Body Preview: {msg.get_payload(decode=True)[:100].decode()}...")

    except Exception as e:
        print(f"Error reading emails: {e}")


if __name__ == "__main__":
    main()
