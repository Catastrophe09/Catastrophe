"""Outlook email-to-task tracker.

Polls the signed-in user's inbox and creates a Microsoft To Do task for each
unread email, then marks the email as read.
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from dataclasses import dataclass
from typing import Any, Dict, List


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.ReadWrite", "Tasks.ReadWrite", "offline_access", "User.Read"]


@dataclass
class Settings:
    client_id: str
    tenant_id: str
    poll_seconds: int = 30
    task_list_id: str | None = None


class GraphClient:
    def __init__(self, settings: Settings) -> None:
        import msal

        authority = f"https://login.microsoftonline.com/{settings.tenant_id}"
        self._app = msal.PublicClientApplication(
            client_id=settings.client_id,
            authority=authority,
        )
        self._settings = settings
        self._token: str | None = None

    def authenticate(self) -> None:
        accounts = self._app.get_accounts()
        result: Dict[str, Any] | None = None
        if accounts:
            result = self._app.acquire_token_silent(SCOPES, account=accounts[0])

        if not result:
            flow = self._app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise RuntimeError("Failed to start device flow.")
            print(flow["message"])
            result = self._app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError(f"Authentication failed: {result.get('error_description')}")

        self._token = result["access_token"]

    def _request(self, method: str, path: str, **kwargs: Any) -> Dict[str, Any]:
        if not self._token:
            raise RuntimeError("Client not authenticated.")

        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {self._token}"
        headers["Content-Type"] = "application/json"

        import requests

        response = requests.request(
            method=method,
            url=f"{GRAPH_BASE}{path}",
            headers=headers,
            timeout=30,
            **kwargs,
        )
        if response.status_code >= 400:
            raise RuntimeError(
                f"Graph API error [{response.status_code}] on {path}: {response.text}"
            )

        if response.text:
            return response.json()
        return {}

    def get_inbox_unread(self, top: int = 25) -> List[Dict[str, Any]]:
        query = (
            f"/me/mailFolders/inbox/messages?"
            f"$filter=isRead eq false&$top={top}&"
            "$select=id,subject,from,receivedDateTime,webLink"
        )
        data = self._request("GET", query)
        return data.get("value", [])

    def get_default_task_list(self) -> str:
        if self._settings.task_list_id:
            return self._settings.task_list_id
        data = self._request("GET", "/me/todo/lists")
        lists = data.get("value", [])
        if not lists:
            raise RuntimeError("No Microsoft To Do list available for the signed-in user.")
        return lists[0]["id"]

    def create_task(self, list_id: str, payload: Dict[str, Any]) -> None:
        self._request("POST", f"/me/todo/lists/{list_id}/tasks", json=payload)

    def mark_email_read(self, message_id: str) -> None:
        self._request("PATCH", f"/me/messages/{message_id}", json={"isRead": True})


def build_task_from_email(message: Dict[str, Any]) -> Dict[str, Any]:
    subject = message.get("subject") or "(No subject)"
    sender = message.get("from", {}).get("emailAddress", {}).get("address", "unknown")
    received = message.get("receivedDateTime", "")
    web_link = message.get("webLink", "")

    body_text = (
        f"From: {sender}\n"
        f"Received: {received}\n"
        f"Message Link: {web_link}\n"
        "\n"
        "Follow up on this email."
    )

    return {
        "title": subject,
        "body": {
            "content": body_text,
            "contentType": "text",
        },
    }


def load_settings() -> Settings:
    from dotenv import load_dotenv

    load_dotenv()

    client_id = os.getenv("MS_CLIENT_ID", "").strip()
    tenant_id = os.getenv("MS_TENANT_ID", "").strip()
    task_list_id = os.getenv("MS_TASK_LIST_ID", "").strip() or None
    poll_seconds = int(os.getenv("POLL_SECONDS", "30"))

    if not client_id or not tenant_id:
        raise RuntimeError(
            "Missing env vars. Set MS_CLIENT_ID and MS_TENANT_ID in your environment or .env"
        )

    return Settings(
        client_id=client_id,
        tenant_id=tenant_id,
        poll_seconds=poll_seconds,
        task_list_id=task_list_id,
    )


def run(once: bool = False) -> None:
    settings = load_settings()
    client = GraphClient(settings)
    client.authenticate()
    list_id = client.get_default_task_list()

    print("Email tracker started.")
    print(f"Using To Do list id: {list_id}")

    while True:
        unread = client.get_inbox_unread()
        if unread:
            print(f"Found {len(unread)} unread emails.")

        for message in unread:
            task_payload = build_task_from_email(message)
            client.create_task(list_id, task_payload)
            client.mark_email_read(message["id"])
            print(f"Created task for email: {task_payload['title']}")

        if once:
            break

        time.sleep(settings.poll_seconds)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Track Outlook emails into Microsoft To Do")
    parser.add_argument(
        "--once",
        action="store_true",
        help="Run one sync cycle and exit",
    )
    return parser.parse_args()


if __name__ == "__main__":
    try:
        args = parse_args()
        run(once=args.once)
    except KeyboardInterrupt:
        print("Stopped by user.")
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
