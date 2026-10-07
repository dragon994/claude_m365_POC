#!/usr/bin/env python3
"""Send an email from a user to themselves via the Outlook REST API.

The access token must be issued for the Outlook resource
(aud = https://outlook.office.com) and carry a Mail.Send scope.

Usage: send_mail.py <upn> <access_token>
"""
import argparse
import json
import sys
import urllib.error
import urllib.request

OUTLOOK_API = "https://outlook.office.com/api/v2.0"


def send_mail(upn: str, token: str, subject: str, body: str) -> None:
    payload = {
        "Message": {
            "Subject": subject,
            "Body": {"ContentType": "Text", "Content": body},
            "ToRecipients": [{"EmailAddress": {"Address": upn}}],
        },
        "SaveToSentItems": True,
    }
    req = urllib.request.Request(
        f"{OUTLOOK_API}/users/{upn}/sendmail",
        data=json.dumps(payload).encode(),
        method="POST",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req) as resp:
        # Success is 202 Accepted with an empty body
        print(f"OK: HTTP {resp.status} - mail sent to {upn}")


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("upn", help="user principal name, e.g. user@contoso.com")
    p.add_argument("token", help="access token for https://outlook.office.com")
    p.add_argument("--subject", default="Test message")
    p.add_argument("--body", default="This is a test message sent to myself via the Outlook REST API.")
    args = p.parse_args()

    try:
        send_mail(args.upn, args.token, args.subject, args.body)
    except urllib.error.HTTPError as e:
        print(f"HTTP {e.code}: {e.read().decode(errors='replace')}", file=sys.stderr)
        return 1
    except urllib.error.URLError as e:
        print(f"Request failed: {e.reason}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
