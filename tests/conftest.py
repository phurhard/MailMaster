import os

# Set dummy environment variables for tests BEFORE importing anything else
os.environ["SUPABASE_URL"] = "https://dummy.supabase.co"
os.environ["SUPABASE_ANON_KEY"] = "dummy_anon_key"
os.environ["GMAIL_CLIENT_ID"] = "dummy_client_id"
os.environ["GMAIL_CLIENT_SECRET"] = "dummy_client_secret"
os.environ["JWT_SECRET"] = "dummy_jwt_secret"
