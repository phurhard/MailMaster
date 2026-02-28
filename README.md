# MailMaster 🚀

MailMaster is a modern, AI-powered email categorization and management dashboard. Built to learn RAG concepts and showcase advanced automated email workflows, this application transforms how you interact with your inbox.

## Key Features ✨

*   **Smart Inbox:** Seamlessly view and search your latest emails with lightning-fast O(1) batched Google API integrations.
*   **AI Summarization & Categorization:** Powered by **LiteLLM** (defaulting to OpenAI's `gpt-4o-mini`), MailMaster instantly reads your email threads and provides concise 1-2 sentence executive summaries and smart labels (Priority, Spam, Newsletter, etc).
*   **Space Optimizer:** Automatically identifies the largest attachments hiding in your inbox so you can free up your Google Drive storage limits.
*   **Premium UI:** A stunning **React + Vite** frontend styled with **Tailwind CSS v4** featuring glassmorphism, dynamic gradients, and fluid animations.
*   **Secure Auth:** Full **Google OAuth 2.0** integration backed by **Supabase**. Client secrets are securely protected via environment variables, not exposed in the database.

## Architecture & Tech Stack 🛠️

### Frontend (`/web`)
*   **Framework:** React 18 + Vite
*   **Routing:** `wouter` for lightweight client-side routing.
*   **Styling:** Tailwind CSS v4 (with `@tailwindcss/vite` plugin).
*   **Data Fetching:** `@tanstack/react-query` for optimized API state management.
*   **Icons:** `lucide-react`

### Backend (`/`)
*   **Framework:** FastAPI (Python 3.12)
*   **Database:** Supabase (PostgreSQL) for storing OAuth tokens securely (`user_tokens` schema).
*   **AI SDK:** `litellm` for standardized LLM completions.
*   **Google Integration:** `google-api-python-client` and `google-auth` for interacting with Gmail.
*   **Configuration:** `pydantic-settings` mapping environment variables cleanly.
*   **Logging:** Centralized `RotatingFileHandler` customized to maintain logs at `<root>/logs/mailmaster.log`.

## Setup & Local Development 💻

### 1. Backend Setup
1. Clone the repository and navigate to the project root.
2. Create a virtual environment: `python3 -m venv venv`
3. Activate it: `source venv/bin/activate`
4. Install dependencies: `pip install -r requirements.txt`
5. Create a `.env` file in the root based on your credentials:
```env
GMAIL_CLIENT_ID="your-google-client-id"
GMAIL_CLIENT_SECRET="your-google-client-secret"
GMAIL_PROJECT_ID="your-google-project-id"
SUPABASE_URL="your-supabase-url"
SUPABASE_ANON_KEY="your-supabase-anon-key"
OPENAI_API_KEY="your-openai-api-key"
JWT_SECRET="a-very-secure-random-string-at-least-thirty-two-bytes-long"
FRONTEND_URL="http://localhost:5173"
```
6. Start the FastAPI server: `uvicorn app:app --port 8000 --reload`
7. View API documentation at `http://localhost:8000/scalar`

### 2. Frontend Setup
1. Navigate to the web directory: `cd web`
2. Install dependencies: `npm install`
3. Start the Vite dev server: `npm run dev`
4. Access the gorgeous UI at `http://localhost:5173`

## Testing 🧪
The backend is fully covered by an automated test suite.
Run the tests using pytest:
```bash
source venv/bin/activate
pytest tests/ -v
```
