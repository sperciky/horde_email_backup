# Horde Email Backup & Offline Explorer

A two-part toolkit for backing up a Horde webmail account via IMAP and browsing the archive offline on Windows.

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│ Part 1 — backup.py (runs anywhere: Linux / Mac / Windows)       │
│                                                                 │
│  IMAP Server ──► IMAPClient ──► EmailStore                      │
│  (read-only)      (imaplib)     ├─ data/emails/<folder>/<uid>.eml│
│                                 ├─ data/attachments/…           │
│                                 └─ data/index.db  (SQLite+FTS5) │
└─────────────────────────────────────────────────────────────────┘
          │  (data/ directory)
          ▼
┌─────────────────────────────────────────────────────────────────┐
│ Part 2 — explorer/ (Flask local web app, runs on Windows)       │
│                                                                 │
│  Browser ◄──► Flask app ──► SQLite FTS5 index                   │
│  localhost        │         + .eml files                        │
│  :5000            └──► HordeExplorer.exe  (PyInstaller bundle)  │
└─────────────────────────────────────────────────────────────────┘
```

---

## Technology Choices

| Component | Choice | Reason |
|---|---|---|
| IMAP client | `imaplib` (stdlib) | Zero dependencies; full RFC 3501 support |
| MIME parsing | `email` (stdlib) | Handles all MIME structures correctly |
| Storage — raw | `.eml` files | Universal format; any email client can open them |
| Storage — index | SQLite + FTS5 | Single file, no server, fast full-text search |
| Explorer server | Flask | Minimal, easy to bundle with PyInstaller |
| Frontend | Vanilla JS + CSS | No npm/node required; works fully offline |
| Windows packaging | PyInstaller | Creates a self-contained `.exe` folder |

---

## Project Structure

```
horde_email_backup/
├── backup.py               # Part 1: IMAP backup script
├── config.example.ini      # Config template (copy to config.ini)
├── requirements.txt        # Python deps (only Flask for explorer)
├── run_backup.bat          # Windows: run backup
├── run_explorer.bat        # Windows: start explorer in browser
├── build_windows.bat       # Windows: build .exe with PyInstaller
├── .gitignore
│
├── explorer/
│   ├── app.py              # Flask web application
│   ├── __main__.py         # PyInstaller entry-point
│   ├── templates/
│   │   └── index.html      # Single-page UI
│   └── static/
│       ├── style.css       # Responsive, dark-mode styles
│       └── app.js          # Vanilla JS (no framework)
│
└── data/                   # Created by backup.py (git-ignored)
    ├── index.db            # SQLite database + FTS5 index
    ├── emails/
    │   ├── INBOX/
    │   │   ├── 1.eml
    │   │   └── 2.eml
    │   └── Sent/
    │       └── 1.eml
    └── attachments/
        └── INBOX/
            └── 2/
                └── document.pdf
```

---

## Part 1 — Email Backup

### SQLite Schema

```sql
folders        id, name, last_uid
emails         id, folder_id, uid, message_id, subject, sender,
               recipients, date_sent, date_received,
               has_attachments, eml_path
emails_body    rowid→emails.id, body_text     ← plain-text for FTS
emails_fts     FTS5 virtual table (subject, sender, recipients, body_text)
attachments    id, email_id, filename, content_type, size, file_path
bookmarks      email_id, created_at           ← created by explorer
email_tags     email_id, tag                  ← created by explorer
```

### Configuring IMAP Access

1. Copy `config.example.ini` → `config.ini`
2. Fill in your server details:

```ini
[imap]
host     = mail.example.com   # Your IMAP host
port     = 993                # 993=SSL (recommended), 143=plain
ssl      = true
username = you@example.com
password = your_password

[backup]
data_dir        = data        # Where to store the archive
batch_size      = 25          # Emails per FETCH request
exclude_folders = [Gmail]/All Mail
```

**Common IMAP hosts:**

| Provider | Host | Port | SSL |
|---|---|---|---|
| Gmail | `imap.gmail.com` | 993 | true |
| Outlook/Hotmail | `outlook.office365.com` | 993 | true |
| Yahoo | `imap.mail.yahoo.com` | 993 | true |
| Self-hosted Horde | Ask your admin | 993 | true |

> **Gmail users:** Enable IMAP in Gmail Settings → Forwarding and POP/IMAP, and use an [App Password](https://myaccount.google.com/apppasswords) instead of your main password.

### Running the Backup

**Linux / Mac:**
```bash
pip install -r requirements.txt
python backup.py                # incremental (only new emails)
python backup.py --full         # full re-download
python backup.py --folder INBOX # one folder only
```

**Windows (double-click or cmd):**
```
run_backup.bat
run_backup.bat --full
run_backup.bat --folder INBOX
```

The script is **strictly read-only** — it never moves, flags, or deletes any email on the server.

### Incremental Backups

After the first run, each subsequent run only downloads emails with UIDs greater than the last seen UID per folder. This is stored in the `folders.last_uid` column. Run it daily via Windows Task Scheduler or a cron job for automatic updates.

---

## Part 2 — Offline Explorer

### Running (development mode)

```bash
# From the project root:
python explorer/app.py --data data --port 5000
# Then open http://localhost:5000 in your browser
```

Or on Windows, double-click `run_explorer.bat`.

### Explorer Features

| Feature | How to use |
|---|---|
| **Folder navigation** | Click any folder in the left sidebar |
| **Full-text search** | Type in the search box and press Enter or click 🔍 |
| **Filter** | Use the filter bar (sender, recipient, subject, date range, has attachments) |
| **Sort** | Sort dropdown (date/sender/subject, newest/oldest) |
| **Pagination** | 50 emails per page; Prev / Next buttons |
| **HTML rendering** | Email body shown in sandboxed iframe |
| **Plain text** | Click "Plain" button to switch |
| **Attachments** | Shown below the header; click filename to download |
| **Download .eml** | "↋ EML" button downloads the raw email file |
| **Export text** | "↋ TXT" button downloads plain-text version |
| **Bookmark** | Star button (★) — bookmarked emails appear in the Bookmarks view |
| **Tags** | Type a tag and press Enter or + to add; click × to remove |
| **Dark mode** | Moon/sun button in the top-left |
| **Keyword highlight** | Search results automatically highlight matched terms |
| **Keyboard shortcut** | Press `/` or `s` to focus the search box |

### API Endpoints (for scripting/integration)

| Endpoint | Description |
|---|---|
| `GET /api/folders` | List all folders with counts |
| `GET /api/emails?folder_id=&page=&sort=&sender=…` | Paginated email list |
| `GET /api/search?q=&folder_id=&page=` | FTS5 full-text search |
| `GET /api/email/<id>` | Email metadata + attachment list |
| `GET /api/email/<id>/html` | Sanitized HTML body (for iframe) |
| `GET /api/email/<id>/download` | Download raw .eml |
| `GET /api/email/<id>/export/text` | Export as plain text |
| `GET /api/attachment/<id>` | Download attachment |
| `POST/DELETE /api/email/<id>/bookmark` | Add/remove bookmark |
| `GET/POST/DELETE /api/email/<id>/tags` | Manage tags |
| `GET /api/stats` | Overall statistics |

---

## Building the Windows Executable

### Prerequisites (one-time)

```
pip install pyinstaller flask
```

### Build

```
build_windows.bat
```

This produces `dist\HordeExplorer\` — a self-contained folder containing:
- `HordeExplorer.exe` (double-click to launch)
- All Python runtime files bundled inside

### Distributing

1. Copy the entire `dist\HordeExplorer\` folder to the target Windows machine.
2. Place your `data\` folder (created by backup.py) next to `HordeExplorer.exe`.
3. Double-click `HordeExplorer.exe`.
4. The browser opens automatically at `http://localhost:5000`.

**Final layout on the target machine:**
```
HordeExplorer\
├── HordeExplorer.exe
├── _internal\          (Python runtime — don't delete)
└── data\
    ├── index.db
    ├── emails\
    └── attachments\
```

> To change the port: run from the command line:
> `HordeExplorer.exe --port 8080`

---

## Known Limitations & Future Improvements

### Current limitations

| Limitation | Workaround |
|---|---|
| HTML sanitization is regex-based | Install `bleach` / `nh3` for production-grade sanitization |
| No STARTTLS support (only SSL or plain) | Set `ssl=false` for plain; most servers use SSL anyway |
| No OAuth 2.0 / XOAUTH2 | Required for some corporate accounts; add via `imaplib` AUTH extension |
| Attachments are stored as flat files | Works well up to ~100 GB; consider compression for larger archives |
| FTS index is not updated if .eml files are edited externally | Re-run backup or trigger a manual `INSERT INTO emails_fts(emails_fts) VALUES('rebuild')` |
| Explorer is single-user only (no auth) | Acceptable for localhost use; add Flask-Login if exposing on a network |

### Future improvements

- [ ] STARTTLS support (`ssl=false` + `STARTTLS` negotiation)
- [ ] OAuth 2.0 / XOAUTH2 for Google / Microsoft accounts
- [ ] Export selected emails to PDF (via WeasyPrint or wkhtmltopdf)
- [ ] Export to Markdown
- [ ] Multi-account support (multiple config sections)
- [ ] Email threading / conversation view
- [ ] Configurable per-account exclusions in the explorer
- [ ] Drag-and-drop import of additional .eml / .mbox files
- [ ] Windows system tray icon for the explorer server
- [ ] Automatic scheduled backup (via Windows Task Scheduler integration)
