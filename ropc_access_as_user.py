#!/usr/bin/env python3
"""Acquire a token for a custom API's access_as_user scope via ROPC."""
import argparse
import sys
import requests

CLIENT_ID = "08ad6f98-a4f8-4635-bb8d-f1a3044760f0"
RESOURCE_APP_ID = "07c030f6-5743-41b7-ba00-0a6e85f37c17"
SCOPE = f"api://{RESOURCE_APP_ID}/access_as_user"
TOKEN_ENDPOINT = "https://login.microsoftonline.com/organizations/oauth2/v2.0/token"


def acquire_token(username: str, password: str) -> dict:
    data = {
        "client_id": CLIENT_ID,
        "grant_type": "password",
        "scope": SCOPE,
        "username": username,
        "password": password,
    }
    resp = requests.post(TOKEN_ENDPOINT, data=data, timeout=30)
    print(f"HTTP {resp.status_code}")
    try:
        return resp.json()
    except ValueError:
        print(resp.text)
        sys.exit(1)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("username", help="UPN of the user")
    parser.add_argument("password", help="User password")
    args = parser.parse_args()

    print(f"Client:   {CLIENT_ID}")
    print(f"Resource: {RESOURCE_APP_ID}")
    print(f"Scope:    {SCOPE}\n")

    result = acquire_token(args.username, args.password)

    if "access_token" in result:
        print("Token acquired.")
        print(f"  token_type:  {result.get('token_type')}")
        print(f"  expires_in:  {result.get('expires_in')}s")
        print(f"  scope:       {result.get('scope')}")
        print(f"  has_refresh: {'refresh_token' in result}")
        print(f"  access_token (first 40 chars): {result['access_token'][:40]}...")
    else:
        print("Token NOT acquired.")
        print(f"  error:             {result.get('error')}")
        print(f"  error_description: {result.get('error_description')}")
        print(f"  error_codes:       {result.get('error_codes')}")


if __name__ == "__main__":
    main()
