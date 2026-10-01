#!/usr/bin/env python3
"""ROPC for a custom API, then OBO exchange for Graph mail scopes."""
import argparse
import sys
import requests

CLIENT_ID = "08ad6f98-a4f8-4635-bb8d-f1a3044760f0"
RESOURCE_APP_ID = "07c030f6-5743-41b7-ba00-0a6e85f37c17"
API_SCOPE = f"api://{RESOURCE_APP_ID}/access_as_user"
TOKEN_ENDPOINT = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"

GRAPH_MAIL_SCOPES = [
    "offline_access",
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.ReadBasic",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/Mail.Read.Shared",
    "https://graph.microsoft.com/Mail.ReadWrite.Shared",
    "https://graph.microsoft.com/Mail.Send.Shared",
    "https://graph.microsoft.com/MailboxSettings.Read",
    "https://graph.microsoft.com/MailboxSettings.ReadWrite",
]


def acquire_user_token(username: str, password: str) -> dict:
    data = {
        "client_id": CLIENT_ID,
        "grant_type": "password",
        "scope": API_SCOPE,
        "username": username,
        "password": password,
    }
    resp = requests.post(TOKEN_ENDPOINT, data=data, timeout=30)
    print(f"[ROPC] HTTP {resp.status_code}")
    return resp.json()


def obo_exchange(user_assertion: str, resource_client_secret: str | None) -> dict:
    data = {
        "client_id": RESOURCE_APP_ID,
        "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
        "assertion": user_assertion,
        "scope": " ".join(GRAPH_MAIL_SCOPES),
        "requested_token_use": "on_behalf_of",
    }
    if resource_client_secret:
        data["client_secret"] = resource_client_secret
    resp = requests.post(TOKEN_ENDPOINT, data=data, timeout=30)
    print(f"[OBO]  HTTP {resp.status_code}")
    return resp.json()


def print_token(label: str, result: dict) -> None:
    if "access_token" in result:
        print(f"  {label}: acquired")
        print(f"    token_type:  {result.get('token_type')}")
        print(f"    expires_in:  {result.get('expires_in')}s")
        print(f"    scope:       {result.get('scope')}")
        print(f"    has_refresh: {'refresh_token' in result}")
        print(f"    access_token (first 40 chars): {result['access_token'][:40]}...")
    else:
        print(f"  {label}: NOT acquired")
        print(f"    error:             {result.get('error')}")
        print(f"    error_description: {result.get('error_description')}")
        print(f"    error_codes:       {result.get('error_codes')}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("username", help="UPN of the user")
    parser.add_argument("password", help="User password")
    parser.add_argument("--resource-client-secret", default=None,
                        help="Client secret for the resource (middle-tier) app, if confidential")
    args = parser.parse_args()

    print(f"Step 1: ROPC — client={CLIENT_ID}, scope={API_SCOPE}")
    user_token = acquire_user_token(args.username, args.password)
    print_token("API token", user_token)
    if "access_token" not in user_token:
        sys.exit(1)

    print(f"\nStep 2: OBO — middle-tier={RESOURCE_APP_ID}, target=Microsoft Graph (mail scopes)")
    graph_token = obo_exchange(user_token["access_token"], args.resource_client_secret)
    print_token("Graph token", graph_token)
    if "access_token" not in graph_token:
        sys.exit(1)


if __name__ == "__main__":
    main()
