from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from scalar_fastapi import get_scalar_api_reference
from api.routes import auth, emails

app = FastAPI(
    title="MailMaster API",
    description="API for email management and analysis",
    version="1.0.0",
    docs_url=None, 
    redoc_url=None
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

@app.get("/scalar", include_in_schema=False)
async def scalar_html():
    return get_scalar_api_reference(
        openapi_url=app.openapi_url,
        title=app.title,
    )

@app.get("/")
async def root():
    return {
        "message": "Welcome to MailMaster API",
        "endpoints": {
            "auth": "/auth/login",
            "docs": "/scalar",
            "search": "/emails/search/{keyword}"
        }
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
