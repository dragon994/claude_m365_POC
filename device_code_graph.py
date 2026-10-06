#!/usr/bin/env python3
"""Device code flow for Microsoft Graph delegated mail scopes."""
import argparse
import sys
import time
import requests

AUTHORITY = "https://login.microsoftonline.com/common"
DEVICE_CODE_ENDPOINT = f"{AUTHORITY}/oauth2/v2.0/devicecode"
TOKEN_ENDPOINT = f"{AUTHORITY}/oauth2/v2.0/token"

# Application (client) ID used to run the device code flow.
DEFAULT_CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48"

# Resource (client) ID the access token is requested for.
# 00000002-0000-0ff1-ce00-000000000000 = Office 365 Exchange Online
# (outlook.office.com). The scopes below are already fully qualified with
# this resource's URL prefix, so the resource is implied by the scope URIs.
RESOURCE_CLIENT_ID = "00000002-0000-0ff1-ce00-000000000000"

DELEGATED_MAIL_SCOPES = [
    "https://outlook.office.com/analytics.readwrite",
    "https://outlook.office.com/calendars.readwrite",
    "https://outlook.office.com/calendars.readwrite.all",
    "https://outlook.office.com/calendars.readwrite.shared",
    "https://outlook.office.com/calendars-internal.readwrite",
    "https://outlook.office.com/channel.create",
    "https://outlook.office.com/channel.readbasic.all",
    "https://outlook.office.com/channelmember.read.all",
    "https://outlook.office.com/channelmember.readwrite.all",
    "https://outlook.office.com/channelmessage.read.all",
    "https://outlook.office.com/chat.read",
    "https://outlook.office.com/chat.readwrite.all",
    "https://outlook.office.com/collab-internal.readwrite",
    "https://outlook.office.com/connectedaccount-internal.readwrite",
    "https://outlook.office.com/connectors.readwrite.shared",
    "https://outlook.office.com/contacts.readwrite",
    "https://outlook.office.com/contacts.readwrite.shared",
    "https://outlook.office.com/directory.read.global",
    "https://outlook.office.com/directory.read.local",
    "https://outlook.office.com/dwengine-internal.read",
    "https://outlook.office.com/eas.accessasuser.all",
    "https://outlook.office.com/files.readwrite.all",
    "https://outlook.office.com/files.readwrite.shared",
    "https://outlook.office.com/focusedinbox-internal.readwrite",
    "https://outlook.office.com/group.readwrite.all",
    "https://outlook.office.com/group.readwrite.all.sdp",
    "https://outlook.office.com/locations-internal.readwrite",
    "https://outlook.office.com/mail.readwrite",
    "https://outlook.office.com/mail.readwrite.all",
    "https://outlook.office.com/mail.readwrite.shared",
    "https://outlook.office.com/mail.send",
    "https://outlook.office.com/mail.send.shared",
    "https://outlook.office.com/mailboxsettings.readwrite",
    "https://outlook.office.com/mailboxsettings.readwrite.all",
    "https://outlook.office.com/notes.read",
    "https://outlook.office.com/notes.readwrite",
    "https://outlook.office.com/notes-internal.readwrite",
    "https://outlook.office.com/notifications-internal.readwrite",
    "https://outlook.office.com/onlinemeetings.readwrite",
    "https://outlook.office.com/outlookcopilot-internal.readwrite",
    "https://outlook.office.com/outlookcopilotlicense-internal.read.sdp",
    "https://outlook.office.com/outlookservice.accessasuser.all",
    "https://outlook.office.com/outlookservice.notificationschannel.all",
    "https://outlook.office.com/owa.accessasuser.all",
    "https://outlook.office.com/people.read",
    "https://outlook.office.com/people.readwrite",
    "https://outlook.office.com/peoplepredictions-internal.read",
    "https://outlook.office.com/peoplesettings.read.all",
    "https://outlook.office.com/place.read.all",
    "https://outlook.office.com/place.readwrite.all",
    "https://outlook.office.com/policy.read.all.sdp",
    "https://outlook.office.com/premium-internal.readwrite",
    "https://outlook.office.com/privilege.openassystem",
    "https://outlook.office.com/signal.readwrite",
    "https://outlook.office.com/signals.read",
    "https://outlook.office.com/signals.readwrite",
    "https://outlook.office.com/signals-internal.read.shared",
    "https://outlook.office.com/signals-internal.readwrite.shared",
    "https://outlook.office.com/substratesearch-internal.readwrite",
    "https://outlook.office.com/tags.readwrite",
    "https://outlook.office.com/tasks.readwrite",
    "https://outlook.office.com/tasks.readwrite.shared",
    "https://outlook.office.com/team.readbasic.all",
    "https://outlook.office.com/todo-internal.readwrite",
    "https://outlook.office.com/user.invite.all.sdp",
    "https://outlook.office.com/user.read.all",
    "https://outlook.office.com/user.read.sdp",
    "https://outlook.office.com/user.readbasic",
    "https://outlook.office.com/user.readbasic.all",
    "https://outlook.office.com/user.readbasic.shared",
    "https://outlook.office.com/user.readwrite",
    "https://outlook.office.com/user.readwrite.shared",
    "https://outlook.office.com/user-internal.readwrite",
    "https://outlook.office.com/.default",
]


def start_device_code(client_id: str) -> dict:
    data = {"client_id": client_id, "scope": " ".join(DELEGATED_MAIL_SCOPES)}
    resp = requests.post(DEVICE_CODE_ENDPOINT, data=data, timeout=30)
    if resp.status_code != 200:
        print(f"Device code request failed: HTTP {resp.status_code}")
        print(resp.text)
        sys.exit(1)
    return resp.json()


def poll_for_token(client_id: str, device_code: str, interval: int, expires_in: int) -> dict:
    data = {
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "client_id": client_id,
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
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "client_id",
        nargs="?",
        default=DEFAULT_CLIENT_ID,
        help="Entra app (client) ID to run device code flow against "
             f"(default: {DEFAULT_CLIENT_ID})",
    )
    args = parser.parse_args()

    print(f"Starting device code flow (client_id={args.client_id})")
    print(f"Resource: {RESOURCE_CLIENT_ID}")
    print(f"Scopes: {' '.join(DELEGATED_MAIL_SCOPES)}\n")

    dc = start_device_code(args.client_id)
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

    token = poll_for_token(args.client_id, dc["device_code"], dc["interval"], dc["expires_in"])

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
