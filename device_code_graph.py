#!/usr/bin/env python3
"""Device code flow against the Microsoft Graph PowerShell app for mail scopes."""
import sys
import time
import requests

CLIENT_ID = "fb78d390-0c51-40cd-8e17-fdbfab77341b"
AUTHORITY = "https://login.microsoftonline.com/common"
DEVICE_CODE_ENDPOINT = f"{AUTHORITY}/oauth2/v2.0/devicecode"
TOKEN_ENDPOINT = f"{AUTHORITY}/oauth2/v2.0/token"

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


def start_device_code() -> dict:
    data = {"client_id": CLIENT_ID, "scope": " ".join(DELEGATED_MAIL_SCOPES)}
    resp = requests.post(DEVICE_CODE_ENDPOINT, data=data, timeout=30)
    if resp.status_code != 200:
        print(f"Device code request failed: HTTP {resp.status_code}")
        print(resp.text)
        sys.exit(1)
    return resp.json()


def poll_for_token(device_code: str, interval: int, expires_in: int) -> dict:
    data = {
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "client_id": CLIENT_ID,
        "device_code": device_code,
    }
    deadline = time.time() + expires_in
    while time.time() < deadline:
        time.sleep(interval)
        resp = requests.post(TOKEN_ENDPOINT, data=data, timeout=30)
        body = resp.json()
        if resp.status_code == 200:
            return body
        err = body.get("error")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
            continue
        if err in ("expired_token", "authorization_declined", "bad_verification_code"):
            print(f"Device code auth failed: {err}: {body.get('error_description')}")
            sys.exit(1)
        print(f"Unexpected error: {err}: {body.get('error_description')}")
        sys.exit(1)
    print("Device code expired before authentication completed.")
    sys.exit(1)


def main() -> None:
    print(f"Starting device code flow against Microsoft Graph PowerShell ({CLIENT_ID})")
    print(f"Scopes: {' '.join(DELEGATED_MAIL_SCOPES)}\n")

    dc = start_device_code()
    print("=" * 70)
    print(dc.get("message") or
          f"Go to {dc['verification_uri']} and enter the code: {dc['user_code']}")
    print("=" * 70)
    print(f"User code:     {dc['user_code']}")
    print(f"Verify URL:    {dc['verification_uri']}")
    print(f"Expires in:    {dc['expires_in']}s")
    print(f"Poll interval: {dc['interval']}s")
    print("\nWaiting for sign-in to complete...\n")
    sys.stdout.flush()

    token = poll_for_token(dc["device_code"], dc["interval"], dc["expires_in"])

    print("Token acquired.")
    print(f"  token_type:    {token.get('token_type')}")
    print(f"  expires_in:    {token.get('expires_in')}s")
    print(f"  scope:         {token.get('scope')}")
    print(f"  has_refresh:   {'refresh_token' in token}")
    print(f"  has_id_token:  {'id_token' in token}")
    print(f"\naccess_token:\n{token.get('access_token')}")
    if "refresh_token" in token:
        print(f"\nrefresh_token:\n{token['refresh_token']}")
    if "id_token" in token:
        print(f"\nid_token:\n{token['id_token']}")


if __name__ == "__main__":
    main()
