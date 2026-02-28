from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from api.routes import auth, emails

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

# Include Routers
app.include_router(auth.router)
app.include_router(emails.router)

@app.get("/")
async def root():
    return {
        "message": "Welcome to MailMaster API",
        "endpoints": {
            "auth": "/auth/login",
            "search": "/emails/search/{keyword}"
        }
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
