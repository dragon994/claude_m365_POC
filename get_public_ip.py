#!/usr/bin/env python3
import urllib.request

def main():
    req = urllib.request.Request("https://ip.me", headers={"User-Agent": "curl/8.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        print(resp.read().decode().strip())

if __name__ == "__main__":
    main()
