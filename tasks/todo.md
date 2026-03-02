# MailMaster Development Roadmap

## Current Phase: Smart Categorization & UI Enhancement (COMPLETED ✅)
- [x] **Enhanced Data Extraction**: Extract `List-Unsubscribe`, `X-Mailer`, etc.
- [x] **Body Cleaning**: HTML-to-Text conversion for efficient AI processing.
- [x] **Two-Tier Validation**: NLP Rule-based pre-check + LLM fine-tuning.
- [x] **Idempotency**: Prevent duplicate AI requests.
- [x] **UI Polish**: 
    - Hover cursor hand pointer.
    - Disable buttons during AI processing.
    - "More/Less" toggle for long emails.
    - HTML rendering for email bodies.
    - Attachment metadata extraction and download.

## Next Phase: Gmail Add-on Development (UPCOMING 🚀)
**Goal**: Build a native Google Workspace add-on to bring MailMaster's intelligence directly into the Gmail interface.

### Features
- [ ] **Contextual Sidebar**: Show AI summary and category directly in Gmail when opening an email.
- [ ] **Batch Actions**: One-click "Cleanup Ads" directly within the Google interface.
- [ ] **Auth Sync**: Ensure seamless OAuth link between the web app and the Add-on.
- [ ] **Token Tracking Sync**: Reflect AI token consumption across all interfaces.

## Future Ideas
- [ ] **NLP-Powered Auto-Reply Suggestions**: Native draft generation in the add-on.
- [ ] **Attachment Analyzer**: Flag suspicious or oversized attachments inside Gmail.
- [ ] **Custom Labels Sync**: Automatically apply MailMaster categories as Gmail labels.
