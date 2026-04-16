#!/usr/bin/env python3
"""
diagnose.py — IMAP connection diagnostics for backup.py
========================================================
Run this before backup.py to confirm your config.ini is correct.

Usage:
    python diagnose.py
    python diagnose.py --config path/to/config.ini
"""

import argparse
import configparser
import imaplib
import socket
import ssl
import sys


def load_config(path):
    cfg = configparser.ConfigParser()
    if not cfg.read(path, encoding="utf-8"):
        print(f"[ERROR] Config file not found: {path}")
        print("        Copy config.example.ini → config.ini and fill in your details.")
        sys.exit(1)
    return cfg


def probe(host, port, use_ssl, username, password, label):
    print(f"\n  Trying username: {label!r}")
    try:
        socket.setdefaulttimeout(15)
        if use_ssl:
            ctx = ssl.create_default_context()
            conn = imaplib.IMAP4_SSL(host, port, ssl_context=ctx)
        else:
            conn = imaplib.IMAP4(host, port)

        typ, data = conn.login(username, password)
        if typ == "OK":
            print(f"  [OK] Login succeeded with username: {label!r}")
            # List a few folders to confirm access
            typ2, folders = conn.list()
            if typ2 == "OK":
                print(f"  [OK] Folder listing works. Sample (first 5):")
                for f in (folders or [])[:5]:
                    print(f"       {f}")
            conn.logout()
            return True
        else:
            print(f"  [FAIL] Server said: {data}")
            return False
    except imaplib.IMAP4.error as e:
        print(f"  [FAIL] IMAP error: {e}")
        return False
    except ssl.SSLError as e:
        print(f"  [FAIL] SSL error: {e}")
        return False
    except (ConnectionRefusedError, socket.timeout, OSError) as e:
        print(f"  [FAIL] Connection error: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="IMAP connection diagnostics")
    parser.add_argument("--config", default="config.ini")
    args = parser.parse_args()

    cfg = load_config(args.config)

    host    = cfg.get("imap", "host")
    port    = cfg.getint("imap", "port", fallback=993)
    use_ssl = cfg.getboolean("imap", "ssl", fallback=True)
    username = cfg.get("imap", "username")
    password = cfg.get("imap", "password")

    print("=" * 60)
    print("  IMAP Diagnostics")
    print("=" * 60)
    print(f"  Host      : {host}")
    print(f"  Port      : {port}")
    print(f"  SSL       : {use_ssl}")
    print(f"  Username  : {username}")
    print(f"  Password  : {'*' * len(password)}")

    # ── Step 1: TCP reachability ──────────────────────────────────────
    print("\n[1] Testing TCP connection …")
    try:
        socket.setdefaulttimeout(10)
        s = socket.create_connection((host, port), timeout=10)
        s.close()
        print(f"  [OK] TCP connection to {host}:{port} succeeded.")
    except (ConnectionRefusedError, socket.timeout, OSError) as e:
        print(f"  [FAIL] Cannot reach {host}:{port} — {e}")
        print()
        print("  Possible fixes:")
        print("  • Check that 'host' and 'port' in config.ini are correct.")
        print("  • Verify IMAP is enabled on the server (ask your IT admin).")
        print("  • Check firewall / VPN — IMAP port 993 may be blocked.")
        sys.exit(1)

    # ── Step 2: IMAP greeting + capabilities ─────────────────────────
    print("\n[2] Reading server capabilities …")
    try:
        socket.setdefaulttimeout(15)
        if use_ssl:
            ctx = ssl.create_default_context()
            conn = imaplib.IMAP4_SSL(host, port, ssl_context=ctx)
        else:
            conn = imaplib.IMAP4(host, port)

        typ, caps = conn.capability()
        print(f"  [OK] Server capabilities: {caps}")
        conn.shutdown()
    except Exception as e:
        print(f"  [FAIL] Could not read capabilities: {e}")

    # ── Step 3: Try login with username as-is ─────────────────────────
    print("\n[3] Testing authentication …")
    if probe(host, port, use_ssl, username, password, username):
        print("\n[RESULT] Your config.ini is correct. Run backup.py normally.")
        sys.exit(0)

    # ── Step 4: Try with only the local part (before @) ───────────────
    local_part = username.split("@")[0] if "@" in username else None
    if local_part and local_part != username:
        if probe(host, port, use_ssl, local_part, password, local_part):
            print("\n[RESULT] SUCCESS with the short username.")
            print(f'         Change username in config.ini to:  username = {local_part}')
            sys.exit(0)

    # ── Step 5: Try alternate ports ───────────────────────────────────
    alt_configs = []
    if port == 993:
        alt_configs = [(143, False), (143, True)]
    elif port == 143:
        alt_configs = [(993, True)]

    if alt_configs:
        print("\n[4] Trying alternate port/SSL combinations …")
        for alt_port, alt_ssl in alt_configs:
            print(f"\n  Trying port {alt_port}, ssl={alt_ssl} …")
            try:
                socket.setdefaulttimeout(10)
                s = socket.create_connection((host, alt_port), timeout=10)
                s.close()
                print(f"  [OK] Port {alt_port} is open.")
                if probe(host, alt_port, alt_ssl, username, password, username):
                    print(f"\n[RESULT] SUCCESS on port {alt_port}, ssl={alt_ssl}.")
                    print(f"         Update config.ini:  port = {alt_port}  ssl = {alt_ssl}")
                    sys.exit(0)
                if local_part:
                    if probe(host, alt_port, alt_ssl, local_part, password, local_part):
                        print(f"\n[RESULT] SUCCESS with short username on port {alt_port}.")
                        print(f"         Update config.ini:  username = {local_part}")
                        print(f"                             port = {alt_port}  ssl = {alt_ssl}")
                        sys.exit(0)
            except (ConnectionRefusedError, socket.timeout, OSError):
                print(f"  Port {alt_port} not reachable.")

    # ── All attempts failed ───────────────────────────────────────────
    print("\n" + "=" * 60)
    print("[RESULT] Authentication failed for all tried combinations.")
    print()
    print("Things to check:")
    print("  1. PASSWORD: Verify you can log in to Horde webmail with")
    print("               the exact same password.")
    print("  2. IMAP OFF: Ask your mail admin whether IMAP access is")
    print("               enabled for your account.")
    print("  3. IP BLOCK: The server may only allow IMAP from inside")
    print("               the company network / VPN.")
    print("  4. 2FA / APP PASSWORD: If your account uses two-factor")
    print("               auth you need an 'App Password', not your")
    print("               regular password.")
    print("  5. DIFFERENT HOST: The IMAP server may be a different")
    print("               hostname from the Horde web URL. Ask IT.")
    print("=" * 60)
    sys.exit(1)


if __name__ == "__main__":
    main()
