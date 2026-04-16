#!/usr/bin/env python3
"""
Horde Email Offline Explorer
=============================
A lightweight Flask web app that serves the email archive created by backup.py.
Runs on localhost; no internet connection required after backup.

Features:
  - Folder tree navigation
  - Full-text search (SQLite FTS5)
  - Filter by sender / recipient / date range / folder
  - HTML email rendering (sandboxed in iframe)
  - Plain-text fallback
  - Attachment download
  - Keyword highlighting
  - Email export (plain text / EML)
  - Dark mode toggle
  - Bookmarking / tagging emails

Entry points:
  Direct:      python app.py [--data ../data] [--port 5000]
  PyInstaller: python __main__.py  (same args)
"""

import argparse
import email as emaillib
import email.policy
import html as html_mod
import json
import mimetypes
import os
import re
import sqlite3
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from flask import (
    Flask,
    Response,
    abort,
    jsonify,
    render_template,
    request,
    send_file,
    send_from_directory,
)

# ---------------------------------------------------------------------------
# App factory
# ---------------------------------------------------------------------------

def create_app(data_dir: str) -> Flask:
    data_path = Path(data_dir).resolve()
    db_path = data_path / "index.db"

    if not db_path.exists():
        raise FileNotFoundError(
            f"Database not found at {db_path}.\n"
            "Run backup.py first to create the archive."
        )

    app = Flask(__name__, template_folder="templates", static_folder="static")
    app.config["DATA_DIR"] = str(data_path)
    app.config["DB_PATH"] = str(db_path)
    app.config["JSON_SORT_KEYS"] = False

    # ------------------------------------------------------------------
    # DB helper
    # ------------------------------------------------------------------

    def get_db() -> sqlite3.Connection:
        db = sqlite3.connect(app.config["DB_PATH"], check_same_thread=False)
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA journal_mode=WAL")
        db.execute("PRAGMA foreign_keys=ON")
        return db

    # Ensure bookmarks / tags tables exist before any request touches them.
    # _ensure_user_tables is idempotent (uses CREATE TABLE IF NOT EXISTS).
    with app.app_context():
        _db = get_db()
        _ensure_user_tables(_db)
        _db.close()

    @app.route("/")
    def index():
        return render_template("index.html")

    # ---------- Folders ----------

    @app.route("/api/folders")
    def api_folders():
        db = get_db()
        rows = db.execute(
            """
            SELECT f.id, f.name,
                   COUNT(e.id) AS total,
                   SUM(CASE WHEN e.has_attachments THEN 1 ELSE 0 END) AS with_attachments
            FROM folders f
            LEFT JOIN emails e ON e.folder_id = f.id
            GROUP BY f.id
            ORDER BY f.name
            """
        ).fetchall()
        db.close()
        return jsonify([dict(r) for r in rows])

    # ---------- Email list ----------

    @app.route("/api/emails")
    def api_emails():
        folder_id = request.args.get("folder_id", type=int)
        page = max(1, request.args.get("page", 1, type=int))
        per_page = min(200, request.args.get("per_page", 50, type=int))
        sort = request.args.get("sort", "date_sent")
        order = "ASC" if request.args.get("order", "desc").upper() == "ASC" else "DESC"
        sender_filter = request.args.get("sender", "").strip()
        recipient_filter = request.args.get("recipient", "").strip()
        subject_filter = request.args.get("subject", "").strip()
        date_from = request.args.get("date_from", "").strip()
        date_to = request.args.get("date_to", "").strip()
        has_attachments = request.args.get("has_attachments", "")

        allowed_sorts = {"date_sent", "date_received", "subject", "sender", "id"}
        if sort not in allowed_sorts:
            sort = "date_sent"

        db = get_db()
        conditions = []
        params: List[Any] = []

        if folder_id:
            conditions.append("e.folder_id = ?")
            params.append(folder_id)
        if sender_filter:
            conditions.append("e.sender LIKE ?")
            params.append(f"%{sender_filter}%")
        if recipient_filter:
            conditions.append("e.recipients LIKE ?")
            params.append(f"%{recipient_filter}%")
        if subject_filter:
            conditions.append("e.subject LIKE ?")
            params.append(f"%{subject_filter}%")
        if date_from:
            conditions.append("e.date_sent >= ?")
            params.append(date_from)
        if date_to:
            conditions.append("e.date_sent <= ?")
            params.append(date_to + "T23:59:59")
        if has_attachments == "1":
            conditions.append("e.has_attachments = 1")

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        count_row = db.execute(
            f"SELECT COUNT(*) AS n FROM emails e {where}", params
        ).fetchone()
        total = count_row["n"]

        offset = (page - 1) * per_page
        rows = db.execute(
            f"""
            SELECT e.id, e.uid, e.subject, e.sender, e.recipients,
                   e.date_sent, e.has_attachments,
                   f.name AS folder_name,
                   COALESCE(t.tags, '') AS tags,
                   COALESCE(bm.email_id IS NOT NULL, 0) AS bookmarked
            FROM emails e
            JOIN folders f ON f.id = e.folder_id
            LEFT JOIN (
                SELECT email_id, GROUP_CONCAT(tag, ',') AS tags
                FROM email_tags GROUP BY email_id
            ) t ON t.email_id = e.id
            LEFT JOIN bookmarks bm ON bm.email_id = e.id
            {where}
            ORDER BY e.{sort} {order}
            LIMIT ? OFFSET ?
            """,
            params + [per_page, offset],
        ).fetchall()
        db.close()

        return jsonify(
            {
                "total": total,
                "page": page,
                "per_page": per_page,
                "emails": [dict(r) for r in rows],
            }
        )

    # ---------- Full-text search ----------

    @app.route("/api/search")
    def api_search():
        q = request.args.get("q", "").strip()
        folder_id = request.args.get("folder_id", type=int)
        page = max(1, request.args.get("page", 1, type=int))
        per_page = min(200, request.args.get("per_page", 50, type=int))

        if not q:
            return jsonify({"total": 0, "page": page, "per_page": per_page, "emails": []})

        db = get_db()
        fts_query = _build_fts_query(q)

        # NOTE: snippet() crashes when the FTS content table (emails_body) has
        # fewer columns than the FTS virtual table.  We exclude snippet() from
        # the SQL and generate a plain-text excerpt in Python instead.
        folder_cond = "AND e.folder_id = ?" if folder_id else ""
        extra_params = [folder_id] if folder_id else []

        try:
            count_row = db.execute(
                f"""
                SELECT COUNT(*) AS n
                FROM emails_fts
                JOIN emails e ON e.id = emails_fts.rowid
                WHERE emails_fts MATCH ?
                {folder_cond}
                """,
                [fts_query] + extra_params,
            ).fetchone()
            total = count_row["n"]

            offset = (page - 1) * per_page
            rows = db.execute(
                f"""
                SELECT e.id, e.uid, e.subject, e.sender, e.recipients,
                       e.date_sent, e.has_attachments,
                       f.name AS folder_name
                FROM emails_fts
                JOIN emails e ON e.id = emails_fts.rowid
                JOIN folders f ON f.id = e.folder_id
                WHERE emails_fts MATCH ?
                {folder_cond}
                ORDER BY emails_fts.rank
                LIMIT ? OFFSET ?
                """,
                [fts_query] + extra_params + [per_page, offset],
            ).fetchall()
        except Exception as exc:
            db.close()
            return jsonify({"error": str(exc)}), 500

        # Build plain-text excerpts in Python (avoids the snippet() crash)
        keywords = [t.strip('"').rstrip('*') for t in fts_query.split(' AND ')]
        results = []
        for r in rows:
            d = dict(r)
            body_row = db.execute(
                "SELECT body_text FROM emails_body WHERE rowid = ?", (d["id"],)
            ).fetchone()
            d["snippet"] = _make_excerpt(
                body_row["body_text"] if body_row else "", keywords
            )
            results.append(d)

        db.close()
        return jsonify(
            {
                "total": total,
                "page": page,
                "per_page": per_page,
                "query": q,
                "emails": results,
            }
        )

    # ---------- Single email ----------

    @app.route("/api/email/<int:email_id>")
    def api_email(email_id: int):
        db = get_db()
        row = db.execute(
            """
            SELECT e.*, f.name AS folder_name,
                   eb.body_text,
                   COALESCE(bm.email_id IS NOT NULL, 0) AS bookmarked
            FROM emails e
            JOIN folders f ON f.id = e.folder_id
            LEFT JOIN emails_body eb ON eb.rowid = e.id
            LEFT JOIN bookmarks bm ON bm.email_id = e.id
            WHERE e.id = ?
            """,
            (email_id,),
        ).fetchone()
        if not row:
            abort(404)
        result = dict(row)

        # Attachments
        atts = db.execute(
            "SELECT id, filename, content_type, size, file_path FROM attachments WHERE email_id=?",
            (email_id,),
        ).fetchall()
        result["attachments"] = [dict(a) for a in atts]

        # Tags
        tags = db.execute(
            "SELECT tag FROM email_tags WHERE email_id=?", (email_id,)
        ).fetchall()
        result["tags"] = [t["tag"] for t in tags]

        db.close()
        return jsonify(result)

    # ---------- HTML body ----------

    @app.route("/api/email/<int:email_id>/html")
    def api_email_html(email_id: int):
        """Return sanitized HTML body for iframe rendering."""
        db = get_db()
        row = db.execute("SELECT eml_path FROM emails WHERE id=?", (email_id,)).fetchone()
        if not row:
            abort(404)
        db.close()

        eml_path = data_path / row["eml_path"]
        if not eml_path.exists():
            abort(404)

        raw = eml_path.read_bytes()
        msg = emaillib.message_from_bytes(raw, policy=emaillib.policy.compat32)

        html_body = None
        plain_body = None
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                continue
            ctype = part.get_content_type()
            if ctype == "text/html" and html_body is None:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    html_body = payload.decode(charset, errors="replace")
            elif ctype == "text/plain" and plain_body is None:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    plain_body = payload.decode(charset, errors="replace")

        if html_body:
            body = _sanitize_html(html_body)
        elif plain_body:
            body = f"<pre style='white-space:pre-wrap;font-family:inherit'>{html_mod.escape(plain_body)}</pre>"
        else:
            body = "<em>No readable body found.</em>"

        highlight = request.args.get("highlight", "")
        if highlight:
            body = _highlight_keywords(body, highlight)

        full_html = f"""<!DOCTYPE html><html><head>
<meta charset="utf-8">
<style>
  body {{font-family:Arial,sans-serif;font-size:14px;padding:12px;margin:0;
         color:#222;background:#fff;word-break:break-word;}}
  pre {{white-space:pre-wrap;}}
  mark {{background:#fff176;}}
  img {{max-width:100%;height:auto;}}
  a {{color:#0066cc;}}
</style>
</head><body>{body}</body></html>"""
        return Response(full_html, mimetype="text/html")

    # ---------- Download .eml ----------

    @app.route("/api/email/<int:email_id>/download")
    def api_email_download(email_id: int):
        db = get_db()
        row = db.execute(
            "SELECT eml_path, subject FROM emails WHERE id=?", (email_id,)
        ).fetchone()
        if not row:
            abort(404)
        db.close()
        eml_path = data_path / row["eml_path"]
        if not eml_path.exists():
            abort(404)
        safe_subject = re.sub(r"[^\w\s-]", "", row["subject"] or "email")[:60].strip()
        return send_file(
            str(eml_path),
            as_attachment=True,
            download_name=f"{safe_subject}.eml",
            mimetype="message/rfc822",
        )

    # ---------- Download attachment ----------

    @app.route("/api/attachment/<int:att_id>")
    def api_attachment(att_id: int):
        db = get_db()
        row = db.execute(
            "SELECT filename, content_type, file_path FROM attachments WHERE id=?",
            (att_id,),
        ).fetchone()
        if not row:
            abort(404)
        db.close()
        att_path = data_path / row["file_path"]
        if not att_path.exists():
            abort(404)
        mime = row["content_type"] or mimetypes.guess_type(row["filename"])[0] or "application/octet-stream"
        return send_file(
            str(att_path),
            as_attachment=True,
            download_name=row["filename"] or "attachment",
            mimetype=mime,
        )

    # ---------- Bookmarks ----------

    @app.route("/api/email/<int:email_id>/bookmark", methods=["POST", "DELETE"])
    def api_bookmark(email_id: int):
        db = get_db()
        _ensure_user_tables(db)
        if request.method == "POST":
            db.execute(
                "INSERT OR IGNORE INTO bookmarks(email_id) VALUES(?)", (email_id,)
            )
        else:
            db.execute("DELETE FROM bookmarks WHERE email_id=?", (email_id,))
        db.commit()
        db.close()
        return jsonify({"ok": True})

    @app.route("/api/bookmarks")
    def api_bookmarks():
        db = get_db()
        _ensure_user_tables(db)
        rows = db.execute(
            """
            SELECT e.id, e.subject, e.sender, e.date_sent, e.has_attachments,
                   f.name AS folder_name
            FROM bookmarks bm
            JOIN emails e ON e.id = bm.email_id
            JOIN folders f ON f.id = e.folder_id
            ORDER BY bm.created_at DESC
            """
        ).fetchall()
        db.close()
        return jsonify([dict(r) for r in rows])

    # ---------- Tags ----------

    @app.route("/api/email/<int:email_id>/tags", methods=["GET", "POST", "DELETE"])
    def api_tags(email_id: int):
        db = get_db()
        _ensure_user_tables(db)
        if request.method == "GET":
            rows = db.execute(
                "SELECT tag FROM email_tags WHERE email_id=?", (email_id,)
            ).fetchall()
            db.close()
            return jsonify([r["tag"] for r in rows])
        data = request.get_json(force=True) or {}
        tag = (data.get("tag") or "").strip()
        if not tag:
            abort(400)
        if request.method == "POST":
            db.execute(
                "INSERT OR IGNORE INTO email_tags(email_id, tag) VALUES(?,?)",
                (email_id, tag),
            )
        else:
            db.execute(
                "DELETE FROM email_tags WHERE email_id=? AND tag=?", (email_id, tag)
            )
        db.commit()
        db.close()
        return jsonify({"ok": True})

    @app.route("/api/tags")
    def api_all_tags():
        db = get_db()
        _ensure_user_tables(db)
        rows = db.execute(
            "SELECT tag, COUNT(*) AS count FROM email_tags GROUP BY tag ORDER BY count DESC"
        ).fetchall()
        db.close()
        return jsonify([dict(r) for r in rows])

    # ---------- Export ----------

    @app.route("/api/email/<int:email_id>/export/text")
    def api_export_text(email_id: int):
        db = get_db()
        row = db.execute(
            "SELECT e.*, f.name AS folder_name, eb.body_text "
            "FROM emails e JOIN folders f ON f.id=e.folder_id "
            "LEFT JOIN emails_body eb ON eb.rowid=e.id WHERE e.id=?",
            (email_id,),
        ).fetchone()
        if not row:
            abort(404)
        db.close()
        d = dict(row)
        content = (
            f"From: {d.get('sender','')}\n"
            f"To: {d.get('recipients','')}\n"
            f"Subject: {d.get('subject','')}\n"
            f"Date: {d.get('date_sent','')}\n"
            f"Folder: {d.get('folder_name','')}\n"
            f"{'='*60}\n\n"
            f"{d.get('body_text','')}"
        )
        safe_subject = re.sub(r"[^\w\s-]", "", d.get("subject", "email") or "email")[:60]
        return Response(
            content,
            mimetype="text/plain; charset=utf-8",
            headers={"Content-Disposition": f'attachment; filename="{safe_subject}.txt"'},
        )

    # ---------- Stats ----------

    @app.route("/api/stats")
    def api_stats():
        db = get_db()
        stats = {}
        stats["total_emails"] = db.execute("SELECT COUNT(*) FROM emails").fetchone()[0]
        stats["total_folders"] = db.execute("SELECT COUNT(*) FROM folders").fetchone()[0]
        stats["total_attachments"] = db.execute("SELECT COUNT(*) FROM attachments").fetchone()[0]
        stats["with_attachments"] = db.execute(
            "SELECT COUNT(*) FROM emails WHERE has_attachments=1"
        ).fetchone()[0]
        oldest = db.execute("SELECT MIN(date_sent) FROM emails").fetchone()[0]
        newest = db.execute("SELECT MAX(date_sent) FROM emails").fetchone()[0]
        stats["date_range"] = {"oldest": oldest, "newest": newest}
        db.close()
        return jsonify(stats)

    return app


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _ensure_user_tables(db: sqlite3.Connection) -> None:
    """Create bookmarks / tags tables on first use (not in main schema)."""
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS bookmarks (
            email_id   INTEGER PRIMARY KEY REFERENCES emails(id),
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS email_tags (
            email_id INTEGER NOT NULL REFERENCES emails(id),
            tag      TEXT    NOT NULL,
            PRIMARY KEY (email_id, tag)
        );
        """
    )
    db.commit()


def _build_fts_query(raw: str) -> str:
    """Convert a plain search string to an FTS5 MATCH expression."""
    cleaned = re.sub(r'[^\w\s@.\-]', ' ', raw)
    tokens = cleaned.split()
    if not tokens:
        return '""'
    return " AND ".join(f'"{t}"*' for t in tokens)


def _make_excerpt(text: str, keywords: List[str], radius: int = 120) -> str:
    """Return a short plain-text excerpt around the first keyword hit.

    Wraps matched terms in <mark>…</mark> for the frontend to display.
    Falls back to the first `radius*2` characters if no keyword is found.
    """
    if not text:
        return ""
    lower = text.lower()
    best = len(text)
    for kw in keywords:
        pos = lower.find(kw.lower())
        if 0 <= pos < best:
            best = pos
    start = max(0, best - radius // 2)
    end   = min(len(text), start + radius * 2)
    excerpt = ("…" if start else "") + text[start:end] + ("…" if end < len(text) else "")
    # Highlight each keyword (case-insensitive)
    for kw in keywords:
        if not kw:
            continue
        excerpt = re.sub(
            rf'(?i)({re.escape(kw)})',
            r'<mark>\1</mark>',
            excerpt,
        )
    return excerpt


_TAG_WHITELIST = {
    "a", "abbr", "acronym", "b", "blockquote", "br", "caption", "cite",
    "code", "col", "colgroup", "dd", "del", "dfn", "div", "dl", "dt", "em",
    "figcaption", "figure", "h1", "h2", "h3", "h4", "h5", "h6", "hr", "i",
    "img", "ins", "kbd", "li", "ol", "p", "pre", "q", "s", "samp", "small",
    "span", "strong", "sub", "sup", "table", "tbody", "td", "tfoot", "th",
    "thead", "time", "tr", "u", "ul", "var",
}

_ATTR_WHITELIST = {"href", "src", "alt", "title", "width", "height", "style", "class", "colspan", "rowspan"}

# Attributes that may contain javascript
_DANGEROUS_ATTRS = re.compile(r"^on\w+", re.I)

def _sanitize_html(html: str) -> str:
    """
    Basic HTML sanitizer: strips scripts, iframes, dangerous attrs.
    For a production deployment use bleach or nh3 instead.
    """
    # Remove dangerous block-level elements entirely
    for tag in ("script", "style", "iframe", "object", "embed", "form",
                "input", "button", "base", "link", "meta", "noscript"):
        html = re.sub(
            rf"<{tag}(\s[^>]*)?>.*?</{tag}>", "", html, flags=re.S | re.I
        )
        html = re.sub(rf"<{tag}(\s[^>]*)?/?>", "", html, flags=re.I)

    # Strip javascript: in href/src
    html = re.sub(r'(href|src)\s*=\s*["\']?\s*javascript:[^"\'>\s]*', r'\1="#"', html, flags=re.I)

    # Remove on* event attributes
    html = re.sub(r'\s+on\w+\s*=\s*(?:"[^"]*"|\'[^\']*\'|[^\s>]*)', '', html, flags=re.I)

    return html


def _highlight_keywords(html: str, keywords: str) -> str:
    """Wrap matched keywords in <mark> tags (case-insensitive)."""
    for word in keywords.split():
        if len(word) < 2:
            continue
        safe = re.escape(html_mod.escape(word))
        html = re.sub(
            rf"(?i)(?<![<\w])({safe})(?![>\w])",
            r"<mark>\1</mark>",
            html,
        )
    return html


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Horde Email Offline Explorer")
    parser.add_argument(
        "--data",
        default=os.path.join(os.path.dirname(__file__), "..", "data"),
        help="Path to backup data directory (default: ../data)",
    )
    parser.add_argument("--port", default=5000, type=int, help="Port to listen on (default: 5000)")
    parser.add_argument(
        "--no-browser", action="store_true", help="Do not open browser automatically"
    )
    args = parser.parse_args()

    try:
        app = create_app(args.data)
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    url = f"http://localhost:{args.port}"
    print(f"\n  Horde Email Explorer running at {url}\n  Press Ctrl+C to stop.\n")

    if not args.no_browser:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()

    app.run(host="127.0.0.1", port=args.port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
