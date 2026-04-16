#!/usr/bin/env python3
"""
Horde Email Backup Tool
=======================
Connects to an IMAP server and downloads a full (or incremental) backup of
all mailbox folders.  Emails are stored as raw .eml files; metadata and
full-text bodies are indexed into SQLite (with FTS5) for the offline explorer.

Usage:
    python backup.py                  # uses config.ini in the same directory
    python backup.py --config /path/to/config.ini
    python backup.py --full           # ignore incremental state, re-download all
    python backup.py --folder INBOX   # backup only one folder
"""

import argparse
import base64
import configparser
import email
import email.policy
import hashlib
import imaplib
import logging
import os
import re
import socket
import sqlite3
import ssl
import sys
import time
from datetime import datetime, timezone
from email.header import decode_header, make_header
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import Iterator, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("backup.log", encoding="utf-8"),
    ],
)
log = logging.getLogger("backup")


# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------

def load_config(path: str) -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    if not os.path.exists(path):
        log.error("Config file not found: %s", path)
        log.error("Copy config.example.ini → config.ini and fill in your credentials.")
        sys.exit(1)
    cfg.read(path, encoding="utf-8")
    return cfg


# ---------------------------------------------------------------------------
# IMAP connection
# ---------------------------------------------------------------------------

class IMAPClient:
    """Thin, robust wrapper around imaplib.IMAP4_SSL / IMAP4."""

    def __init__(self, host: str, port: int, use_ssl: bool, timeout: int = 60):
        self.host = host
        self.port = port
        self.use_ssl = use_ssl
        self.timeout = timeout
        self._conn: Optional[imaplib.IMAP4] = None

    def connect(self, username: str, password: str) -> None:
        log.info("Connecting to %s:%d (SSL=%s) …", self.host, self.port, self.use_ssl)
        socket.setdefaulttimeout(self.timeout)
        if self.use_ssl:
            ctx = ssl.create_default_context()
            self._conn = imaplib.IMAP4_SSL(self.host, self.port, ssl_context=ctx)
        else:
            self._conn = imaplib.IMAP4(self.host, self.port)
        typ, data = self._conn.login(username, password)
        if typ != "OK":
            raise RuntimeError(f"Login failed: {data}")
        log.info("Logged in as %s", username)

    def disconnect(self) -> None:
        if self._conn:
            try:
                self._conn.logout()
            except Exception:
                pass
            self._conn = None

    def list_folders(self) -> List[str]:
        """Return a list of decoded folder names."""
        typ, data = self._conn.list()
        if typ != "OK":
            raise RuntimeError("LIST command failed")
        folders = []
        for item in data:
            if item is None:
                continue
            if isinstance(item, bytes):
                item = item.decode("utf-8", errors="replace")
            # Parse: (\HasNoChildren) "/" "INBOX"
            m = re.match(r'\(.*?\)\s+"?([^"]+)"?\s+"?([^"]+)"?$', item.strip())
            if m:
                sep = m.group(1).strip('"')
                name = m.group(2).strip('"')
                folders.append(name)
        return folders

    def select_folder(self, folder: str) -> int:
        """Select a folder and return the message count."""
        typ, data = self._conn.select(f'"{folder}"', readonly=True)
        if typ != "OK":
            raise RuntimeError(f"SELECT failed for folder '{folder}': {data}")
        return int(data[0])

    def fetch_uids(self, folder: str, since_uid: int = 0) -> List[int]:
        """Return all UIDs in the folder, optionally only those > since_uid."""
        self.select_folder(folder)
        if since_uid > 0:
            typ, data = self._conn.uid("SEARCH", None, f"UID {since_uid + 1}:*")
        else:
            typ, data = self._conn.uid("SEARCH", None, "ALL")
        if typ != "OK":
            raise RuntimeError(f"UID SEARCH failed: {data}")
        raw = data[0]
        if not raw:
            return []
        return [int(u) for u in raw.split()]

    def fetch_email_batch(
        self, uids: List[int], batch_size: int = 25
    ) -> Iterator[Tuple[int, bytes]]:
        """Yield (uid, raw_rfc822_bytes) for each UID, in batches."""
        for i in range(0, len(uids), batch_size):
            batch = uids[i : i + batch_size]
            uid_set = ",".join(str(u) for u in batch)
            typ, data = self._conn.uid("FETCH", uid_set, "(RFC822)")
            if typ != "OK":
                log.warning("FETCH failed for UID batch %s", uid_set)
                continue
            # data is a list of alternating (header, body) tuples and separators
            for part in data:
                if not isinstance(part, tuple):
                    continue
                # part[0] looks like b'123 (UID 456 RFC822 {size}'
                header_str = part[0].decode("utf-8", errors="replace")
                uid_match = re.search(r"UID (\d+)", header_str)
                if not uid_match:
                    continue
                uid = int(uid_match.group(1))
                raw_bytes = part[1]
                yield uid, raw_bytes


# ---------------------------------------------------------------------------
# Local storage
# ---------------------------------------------------------------------------

class EmailStore:
    """
    Persists emails as .eml files in a folder hierarchy and maintains a
    SQLite database with FTS5 index for the offline explorer.

    Layout:
        data/
            emails/<folder_name>/<UID>.eml
            attachments/<folder_name>/<UID>/<filename>
            index.db
    """

    SCHEMA = """
    PRAGMA journal_mode=WAL;
    PRAGMA foreign_keys=ON;

    CREATE TABLE IF NOT EXISTS folders (
        id        INTEGER PRIMARY KEY,
        name      TEXT    UNIQUE NOT NULL,
        last_uid  INTEGER DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS emails (
        id            INTEGER PRIMARY KEY,
        folder_id     INTEGER NOT NULL REFERENCES folders(id),
        uid           INTEGER NOT NULL,
        message_id    TEXT,
        subject       TEXT,
        sender        TEXT,
        recipients    TEXT,
        date_sent     TEXT,
        date_received TEXT,
        has_attachments INTEGER DEFAULT 0,
        eml_path      TEXT NOT NULL,
        UNIQUE(folder_id, uid)
    );

    CREATE TABLE IF NOT EXISTS attachments (
        id           INTEGER PRIMARY KEY,
        email_id     INTEGER NOT NULL REFERENCES emails(id),
        filename     TEXT,
        content_type TEXT,
        size         INTEGER,
        file_path    TEXT NOT NULL
    );

    -- FTS5 virtual table for full-text search
    CREATE VIRTUAL TABLE IF NOT EXISTS emails_fts USING fts5(
        subject,
        sender,
        recipients,
        body_text,
        content='emails_body',
        content_rowid='rowid'
    );

    -- Stores the plain-text body separately (keeps emails table lean)
    CREATE TABLE IF NOT EXISTS emails_body (
        rowid    INTEGER PRIMARY KEY REFERENCES emails(id),
        body_text TEXT
    );

    -- Triggers to keep FTS in sync
    CREATE TRIGGER IF NOT EXISTS emails_fts_insert
        AFTER INSERT ON emails_body BEGIN
            INSERT INTO emails_fts(rowid, subject, sender, recipients, body_text)
            SELECT new.rowid,
                   e.subject, e.sender, e.recipients, new.body_text
            FROM emails e WHERE e.id = new.rowid;
        END;

    CREATE TRIGGER IF NOT EXISTS emails_fts_delete
        AFTER DELETE ON emails_body BEGIN
            INSERT INTO emails_fts(emails_fts, rowid, subject, sender, recipients, body_text)
            VALUES('delete', old.rowid,
                   (SELECT subject FROM emails WHERE id=old.rowid),
                   (SELECT sender  FROM emails WHERE id=old.rowid),
                   (SELECT recipients FROM emails WHERE id=old.rowid),
                   old.body_text);
        END;
    """

    def __init__(self, data_dir: str):
        self.data_dir = Path(data_dir)
        self.emails_dir = self.data_dir / "emails"
        self.attachments_dir = self.data_dir / "attachments"
        self.db_path = self.data_dir / "index.db"
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.emails_dir.mkdir(exist_ok=True)
        self.attachments_dir.mkdir(exist_ok=True)
        self._db = sqlite3.connect(str(self.db_path), check_same_thread=False)
        self._db.row_factory = sqlite3.Row
        self._db.executescript(self.SCHEMA)
        self._db.commit()

    def get_folder_id(self, folder_name: str) -> int:
        cur = self._db.execute(
            "INSERT OR IGNORE INTO folders(name) VALUES(?)", (folder_name,)
        )
        self._db.commit()
        row = self._db.execute(
            "SELECT id FROM folders WHERE name=?", (folder_name,)
        ).fetchone()
        return row["id"]

    def get_last_uid(self, folder_name: str) -> int:
        row = self._db.execute(
            "SELECT last_uid FROM folders WHERE name=?", (folder_name,)
        ).fetchone()
        return row["last_uid"] if row else 0

    def set_last_uid(self, folder_name: str, uid: int) -> None:
        self._db.execute(
            "UPDATE folders SET last_uid=? WHERE name=?", (uid, folder_name)
        )
        self._db.commit()

    def email_exists(self, folder_id: int, uid: int) -> bool:
        row = self._db.execute(
            "SELECT 1 FROM emails WHERE folder_id=? AND uid=?", (folder_id, uid)
        ).fetchone()
        return row is not None

    def save_email(
        self,
        folder_name: str,
        folder_id: int,
        uid: int,
        raw_bytes: bytes,
    ) -> None:
        """Parse and persist a raw email."""
        # Write .eml file
        folder_email_dir = self.emails_dir / _safe_path(folder_name)
        _makedirs(folder_email_dir)
        eml_path = folder_email_dir / f"{uid}.eml"
        with _open_for_write(eml_path) as f:
            f.write(raw_bytes)

        # Parse
        msg = email.message_from_bytes(raw_bytes, policy=email.policy.compat32)

        subject = _decode_header_value(msg.get("Subject", ""))
        sender = _decode_header_value(msg.get("From", ""))
        to_raw = msg.get("To", "")
        cc_raw = msg.get("Cc", "")
        recipients = "; ".join(
            filter(None, [_decode_header_value(to_raw), _decode_header_value(cc_raw)])
        )
        message_id = msg.get("Message-ID", "").strip()
        date_sent = _parse_date(msg.get("Date", ""))
        date_received = datetime.now(timezone.utc).isoformat()

        body_plain, body_html, attachments = _extract_parts(msg)
        has_attachments = 1 if attachments else 0
        body_text = body_plain or _html_to_text(body_html) or ""

        cur = self._db.execute(
            """
            INSERT OR IGNORE INTO emails
                (folder_id, uid, message_id, subject, sender, recipients,
                 date_sent, date_received, has_attachments, eml_path)
            VALUES (?,?,?,?,?,?,?,?,?,?)
            """,
            (
                folder_id, uid, message_id, subject, sender, recipients,
                date_sent, date_received, has_attachments,
                str(eml_path.relative_to(self.data_dir)),
            ),
        )
        email_id = cur.lastrowid
        if email_id is None or email_id == 0:
            # Already exists (IGNORE)
            return

        # Store body for FTS (triggers insert into emails_fts)
        self._db.execute(
            "INSERT OR IGNORE INTO emails_body(rowid, body_text) VALUES(?,?)",
            (email_id, body_text),
        )

        # Save attachments
        for filename, ctype, payload in attachments:
            att_dir = self.attachments_dir / _safe_path(folder_name) / str(uid)
            _makedirs(att_dir)
            safe_name = _safe_filename(filename)
            att_path = att_dir / safe_name
            # Avoid collisions
            counter = 1
            while att_path.exists():
                stem, suffix = os.path.splitext(safe_name)
                att_path = att_dir / f"{stem}_{counter}{suffix}"
                counter += 1
            with _open_for_write(att_path) as f:
                f.write(payload)
            self._db.execute(
                """INSERT INTO attachments(email_id, filename, content_type, size, file_path)
                   VALUES(?,?,?,?,?)""",
                (
                    email_id,
                    filename,
                    ctype,
                    len(payload),
                    str(att_path.relative_to(self.data_dir)),
                ),
            )

        self._db.commit()

    def close(self) -> None:
        self._db.close()


# ---------------------------------------------------------------------------
# Email parsing helpers
# ---------------------------------------------------------------------------

def _decode_header_value(raw: str) -> str:
    """Decode RFC-2047 encoded header values into a plain string."""
    if not raw:
        return ""
    try:
        return str(make_header(decode_header(raw)))
    except Exception:
        return raw


def _parse_date(raw: str) -> Optional[str]:
    if not raw:
        return None
    try:
        return parsedate_to_datetime(raw).isoformat()
    except Exception:
        return None


def _extract_parts(
    msg: email.message.Message,
) -> Tuple[str, str, List[Tuple[str, str, bytes]]]:
    """
    Walk the MIME tree and return:
        (body_plain, body_html, [(filename, content_type, bytes), ...])
    """
    body_plain_parts: List[str] = []
    body_html_parts: List[str] = []
    attachments: List[Tuple[str, str, bytes]] = []

    for part in msg.walk():
        ctype = part.get_content_type()
        disposition = part.get_content_disposition() or ""
        filename = part.get_filename()

        if filename or disposition.lower() == "attachment":
            payload = part.get_payload(decode=True)
            if payload is not None:
                name = _decode_header_value(filename or "attachment")
                attachments.append((name, ctype, payload))
            continue

        if ctype == "text/plain" and not filename:
            payload = part.get_payload(decode=True)
            if payload:
                charset = _normalize_charset(part.get_content_charset())
                body_plain_parts.append(payload.decode(charset, errors="replace"))

        elif ctype == "text/html" and not filename:
            payload = part.get_payload(decode=True)
            if payload:
                charset = _normalize_charset(part.get_content_charset())
                body_html_parts.append(payload.decode(charset, errors="replace"))

    return (
        "\n".join(body_plain_parts),
        "\n".join(body_html_parts),
        attachments,
    )


def _html_to_text(html: str) -> str:
    """Very simple HTML → plain text for FTS indexing (no external deps)."""
    if not html:
        return ""
    text = re.sub(r"<style[^>]*>.*?</style>", " ", html, flags=re.S | re.I)
    text = re.sub(r"<script[^>]*>.*?</script>", " ", text, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&gt;", ">", text)
    text = re.sub(r"&quot;", '"', text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _decode_imap_utf7(s: str) -> str:
    """Decode IMAP modified UTF-7 (RFC 3501) folder name to Unicode.

    IMAP encodes non-ASCII folder names as &<modified-base64>-.
    ',' is used instead of '/' in the base64 alphabet.
    '&-' is a literal '&'.
    """
    result = []
    i = 0
    while i < len(s):
        if s[i] == "&":
            j = s.find("-", i + 1)
            if j == -1:
                result.append(s[i:])
                break
            if j == i + 1:
                result.append("&")
            else:
                b64 = s[i + 1 : j].replace(",", "/")
                b64 += "=" * ((-len(b64)) % 4)
                try:
                    result.append(base64.b64decode(b64).decode("utf-16-be"))
                except Exception:
                    result.append(s[i : j + 1])
            i = j + 1
        else:
            result.append(s[i])
            i += 1
    return "".join(result)


# Maps non-standard charset labels (as seen in real-world email) to Python codec names.
_CHARSET_ALIASES: dict = {
    "cp-850": "cp850",
    "cp-852": "cp852",
    "cp-1250": "cp1250",
    "cp-1251": "cp1251",
    "cp-1252": "cp1252",
    "cp-1253": "cp1253",
    "cp-1254": "cp1254",
    "cp-1256": "cp1256",
    "x-mac-cyrillic": "mac_cyrillic",
    "x-mac-roman": "mac_roman",
    "x-mac-ce": "mac_latin2",
    "x-sjis": "shift_jis",
    "x-euc-jp": "euc_jp",
    "238": "cp1250",   # Windows Eastern European codepage number
    "204": "cp1251",   # Windows Cyrillic
    "161": "cp1253",   # Windows Greek
    "162": "cp1254",   # Windows Turkish
    "177": "cp1255",   # Windows Hebrew
    "178": "cp1256",   # Windows Arabic
    "850": "cp850",
    "437": "cp437",
    "1250": "cp1250",
    "1251": "cp1251",
    "1252": "cp1252",
    "1253": "cp1253",
}


def _normalize_charset(charset: Optional[str]) -> str:
    """Return a Python-recognised codec name for any charset label."""
    if not charset:
        return "utf-8"
    key = charset.lower().strip()
    return _CHARSET_ALIASES.get(key, key)


def _safe_path(folder_name: str) -> str:
    """Convert an IMAP folder name to a safe, short filesystem path segment.

    First decodes IMAP modified UTF-7 so Czech/Cyrillic etc. folders get
    their real Unicode names rather than the encoded &AOk-... form.
    Then strips characters illegal on Windows and caps length at 60 chars.
    """
    decoded = _decode_imap_utf7(folder_name)
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", decoded)
    return safe[:60]


def _safe_filename(name: str, max_len: int = 80) -> str:
    """Sanitize an attachment filename and cap its length.

    max_len=80 leaves ample room within Windows MAX_PATH even for deeply
    nested backup paths.  The file extension is always preserved.
    """
    name = os.path.basename(name) or "attachment"
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    if len(name) <= max_len:
        return name
    stem, ext = os.path.splitext(name)
    keep = max_len - len(ext)
    return stem[:keep] + ext


def _open_for_write(path: Path) -> "open":
    """Open a file for binary writing, using the \\\\?\\ prefix on Windows
    to bypass the 260-character MAX_PATH limit."""
    if sys.platform == "win32":
        path_str = "\\\\?\\" + str(path.resolve())
    else:
        path_str = str(path)
    return open(path_str, "wb")


def _makedirs(path: Path) -> None:
    """Create directories, using the \\\\?\\ prefix on Windows."""
    if sys.platform == "win32":
        path_str = "\\\\?\\" + str(path.resolve())
        os.makedirs(path_str, exist_ok=True)
    else:
        path.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Main backup logic
# ---------------------------------------------------------------------------

def run_backup(
    cfg: configparser.ConfigParser,
    full_backup: bool = False,
    only_folder: Optional[str] = None,
) -> None:
    imap_host = cfg.get("imap", "host")
    imap_port = cfg.getint("imap", "port", fallback=993)
    imap_ssl = cfg.getboolean("imap", "ssl", fallback=True)
    username = cfg.get("imap", "username")
    password = cfg.get("imap", "password")
    data_dir = cfg.get("backup", "data_dir", fallback="data")
    batch_size = cfg.getint("backup", "batch_size", fallback=25)
    timeout = cfg.getint("backup", "timeout", fallback=60)
    exclude_raw = cfg.get("backup", "exclude_folders", fallback="")
    exclude_folders = {f.strip() for f in exclude_raw.split(",") if f.strip()}

    store = EmailStore(data_dir)
    client = IMAPClient(imap_host, imap_port, imap_ssl, timeout)

    try:
        client.connect(username, password)

        folders = client.list_folders()
        log.info("Found %d folders: %s", len(folders), folders)

        if only_folder:
            folders = [f for f in folders if f == only_folder]
            if not folders:
                log.error("Folder '%s' not found on server.", only_folder)
                return

        for folder in folders:
            if folder in exclude_folders:
                log.info("Skipping excluded folder: %s", folder)
                continue

            log.info("--- Processing folder: %s ---", folder)
            folder_id = store.get_folder_id(folder)
            last_uid = 0 if full_backup else store.get_last_uid(folder)

            try:
                uids = client.fetch_uids(folder, since_uid=last_uid)
            except Exception as exc:
                log.warning("Could not fetch UIDs for '%s': %s", folder, exc)
                continue

            if not uids:
                log.info("  No new messages.")
                continue

            log.info("  Downloading %d messages (last known UID: %d) …", len(uids), last_uid)
            downloaded = 0
            max_uid = last_uid

            for uid, raw_bytes in client.fetch_email_batch(uids, batch_size):
                if store.email_exists(folder_id, uid):
                    max_uid = max(max_uid, uid)
                    continue
                try:
                    store.save_email(folder, folder_id, uid, raw_bytes)
                    downloaded += 1
                    max_uid = max(max_uid, uid)
                    if downloaded % 100 == 0:
                        log.info("    … %d saved so far", downloaded)
                except Exception as exc:
                    log.error("    Failed to save UID %d: %s", uid, exc)

            store.set_last_uid(folder, max_uid)
            log.info("  Done. Saved %d new emails. Max UID now: %d", downloaded, max_uid)

    except KeyboardInterrupt:
        log.warning("Interrupted by user.")
    except Exception as exc:
        log.error("Fatal error: %s", exc, exc_info=True)
        sys.exit(1)
    finally:
        client.disconnect()
        store.close()

    log.info("Backup complete.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Horde / IMAP Email Backup Tool")
    parser.add_argument(
        "--config", default="config.ini", help="Path to config file (default: config.ini)"
    )
    parser.add_argument(
        "--full",
        action="store_true",
        help="Force full re-download, ignoring incremental state",
    )
    parser.add_argument(
        "--folder",
        default=None,
        help="Only backup a specific folder (exact IMAP name)",
    )
    args = parser.parse_args()

    cfg = load_config(args.config)
    run_backup(cfg, full_backup=args.full, only_folder=args.folder)


if __name__ == "__main__":
    main()
