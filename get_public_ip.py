#!/usr/bin/env python3
import urllib.request

def main():
    with urllib.request.urlopen("https://ip.me", timeout=10) as resp:
        print(resp.read().decode().strip())

if __name__ == "__main__":
    main()
