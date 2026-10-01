#!/usr/bin/env python3
"""Acquire Graph tokens via ROPC and read the user's latest received email."""
import argparse
import sys
import requests

CLIENT_ID = "07c030f6-5743-41b7-ba00-0a6e85f37c17"
TOKEN_ENDPOINT = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

DELEGATED_MAIL_SCOPES = [
    "offline_access",
    "Mail.Read",
    "Mail.ReadBasic",
    "Mail.ReadWrite",
    "Mail.Send",
    "Mail.Read.Shared",
    "Mail.ReadWrite.Shared",
    "Mail.Send.Shared",
    "MailboxSettings.Read",
    "MailboxSettings.ReadWrite",
]


def acquire_token(username: str, password: str) -> dict:
    data = {
        "client_id": CLIENT_ID,
        "grant_type": "password",
        "scope": " ".join(DELEGATED_MAIL_SCOPES),
        "username": username,
        "password": password,
    }
    resp = requests.post(TOKEN_ENDPOINT, data=data, timeout=30)
    if resp.status_code != 200:
        print(f"Token request failed: {resp.status_code}")
        print(resp.text)
        sys.exit(1)
    return resp.json()


def get_latest_received_mail(access_token: str) -> dict | None:
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {
        "$top": 1,
        "$orderby": "receivedDateTime desc",
        "$select": "subject,from,receivedDateTime,bodyPreview,hasAttachments,isRead",
    }
    url = f"{GRAPH_BASE}/me/mailFolders/Inbox/messages"
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    if resp.status_code != 200:
        print(f"Mail fetch failed: {resp.status_code}")
        print(resp.text)
        sys.exit(1)
    items = resp.json().get("value", [])
    return items[0] if items else None


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("username", help="UPN of the user (e.g. user@contoso.com)")
    parser.add_argument("password", help="User password")
    args = parser.parse_args()

    print("Requesting token via ROPC...")
    token = acquire_token(args.username, args.password)
    print(f"  token_type:  {token.get('token_type')}")
    print(f"  expires_in:  {token.get('expires_in')}s")
    print(f"  scope:       {token.get('scope')}")
    print(f"  has_refresh: {'refresh_token' in token}")

    print("\nFetching latest received email...")
    msg = get_latest_received_mail(token["access_token"])
    if msg is None:
        print("Inbox is empty.")
        return

    sender = msg.get("from", {}).get("emailAddress", {})
    print(f"  Subject:       {msg.get('subject')}")
    print(f"  From:          {sender.get('name')} <{sender.get('address')}>")
    print(f"  Received:      {msg.get('receivedDateTime')}")
    print(f"  Read:          {msg.get('isRead')}")
    print(f"  Attachments:   {msg.get('hasAttachments')}")
    print(f"  Preview:       {msg.get('bodyPreview')}")


if __name__ == "__main__":
    main()
