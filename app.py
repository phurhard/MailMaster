from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from api.services.gmail import init_gmail_service, search_emails, get_email_message_details
from typing import List, Dict
import os

client_service_file = 'client_secret.json'

app = FastAPI(
    title="MailMaster API",
    description="API for email management and analysis",
    version="1.0.0"
)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize Gmail service at startup

@app.on_event("startup")
async def startup_event():
    global gmail_service
    gmail_service = init_gmail_service(client_service_file)

@app.get("/search/{keyword}")
async def search_emails_by_keyword(keyword: str) -> List[Dict]:
    """
    Search emails containing the specified keyword and return their details
    """
    try:
        # Search for emails containing the keyword
        email_messages = search_emails(gmail_service, keyword, max_results=10)
        
        # Get detailed information for each matching email
        results = []
        for email in email_messages:
            email_details = get_email_message_details(gmail_service, email['id'])
            results.append({
                'subject': email_details.get('subject', ''),
                'from': email_details.get('from', ''),
                'date': email_details.get('date', ''),
                'snippet': email_details.get('snippet', ''),
                'labels': email_details.get('labels', [])
            })
        return results
    except Exception as e:
        return {"error": str(e)}

@app.get("/")
async def root():
    return {"message": "Welcome to MailMaster API"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
