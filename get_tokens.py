#!/usr/bin/env python3
"""Device code flow for a Microsoft Graph refresh token."""
import argparse
import sys
import time
import requests

AUTHORITY = "https://login.microsoftonline.com/organizations"
# Microsoft Teams first-party FOCI client; the resulting refresh token
# can be redeemed at any other FOCI family-1 client.
DEFAULT_CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"
SCOPE = "https://graph.microsoft.com/.default offline_access"


def start_device_code(client_id):
    r = requests.post(f"{AUTHORITY}/oauth2/v2.0/devicecode",
                      data={"client_id": client_id, "scope": SCOPE}, timeout=30)
    r.raise_for_status()
    return r.json()


def poll_for_token(client_id, device_code, interval, expires_in):
    data = {
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "client_id": client_id,
        "device_code": device_code,
    }
    deadline = time.time() + expires_in
    while time.time() < deadline:
        time.sleep(interval)
        r = requests.post(f"{AUTHORITY}/oauth2/v2.0/token", data=data, timeout=30)
        body = r.json()
        if r.status_code == 200:
            return body
        err = body.get("error")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
            continue
        sys.exit(f"Auth failed: {err}: {body.get('error_description')}")
    sys.exit("Device code expired before authentication completed.")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("client_id", nargs="?", default=DEFAULT_CLIENT_ID)
    args = ap.parse_args()

    dc = start_device_code(args.client_id)
    print(dc["message"])
    sys.stdout.flush()

    token = poll_for_token(args.client_id, dc["device_code"],
                           dc["interval"], dc["expires_in"])
    print(f"\naccess_token:\n{token['access_token']}")
    if "refresh_token" in token:
        print(f"\nrefresh_token:\n{token['refresh_token']}")


if __name__ == "__main__":
    main()
