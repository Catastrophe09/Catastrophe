# Outlook Email Tracker -> Microsoft To Do

This app connects to your Outlook mailbox via Microsoft Graph and automatically creates a Microsoft To Do task for every unread inbox email.

## What it does

- Authenticates with your Microsoft account using Device Code flow.
- Polls Inbox for unread messages.
- Creates one To Do task per email.
- Marks each processed email as read so it will not be duplicated.

## Setup

1. Create an Azure App Registration.
2. Add delegated API permissions:
   - `Mail.ReadWrite`
   - `Tasks.ReadWrite`
   - `User.Read`
3. Allow public client flows (for Device Code auth).
4. Copy `.env.example` to `.env` and fill values:
   - `MS_CLIENT_ID`: your app/client ID.
   - `MS_TENANT_ID`: `common` or your tenant GUID.
   - `MS_TASK_LIST_ID` (optional): force a specific To Do list.

## Run

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m src.email_tracker --once
# Or keep it running:
python -m src.email_tracker
```

On first run, you will receive a URL + code to sign in.

## Testing

```bash
pytest
```
