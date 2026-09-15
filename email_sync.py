"""
DMB-Dashboard — SharePoint Mailbox Sync to data/ folder
======================================================
Monitors mailbox for updated SharePoint Excel workbooks:
  1. Functional DMB Review Sheets-.xlsx
  2. Masterfile_DMB_Dashboard.xlsx

Key Features:
- Preserves all historical emails in the mailbox while always pulling the newest data.
- Writes incoming workbooks directly into DMB-Dashboard/data/ directory.
- Automatically commits updated Excel workbooks to GitHub (kavyabhardwajj30/DMB-Dashboard).
- Allows live reload of data_loader.py without server restarts.
"""

from __future__ import annotations

import email
import email.utils
import imaplib
import logging
import os
from email.header import decode_header, make_header
from pathlib import Path
from typing import Callable, Iterator

logger = logging.getLogger("dmb.email_sync")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

XLSX_MAGIC = b"PK\x03\x04"


def _env(name: str, default: str = "") -> str:
    return (os.getenv(name) or default).strip()


def _decode(value: str | None) -> str:
    if not value:
        return ""
    try:
        import urllib.parse
        decoded = str(make_header(decode_header(value)))
        return urllib.parse.unquote(decoded)
    except Exception:
        return value


def _excel_attachments(msg: email.message.Message) -> Iterator[tuple[str, bytes]]:
    """Yield (filename, bytes) for every real .xlsx/.xlsm/.xls attachment on the message."""
    subject = _decode(msg.get("Subject", "")).lower()
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = _decode(part.get_filename())
        if not filename:
            cd = part.get("Content-Disposition", "")
            if "filename=" in cd:
                for chunk in cd.split(";"):
                    if "filename=" in chunk:
                        filename = _decode(chunk.split("=", 1)[1].strip(' "\''))
        if not filename:
            continue

        try:
            payload = part.get_payload(decode=True)
        except Exception:
            continue

        if payload and (payload.startswith(XLSX_MAGIC) or filename.lower().endswith((".xlsx", ".xlsm", ".xls"))):
            fname_lower = filename.lower()
            # Standardize known DMB filenames
            if "functional" in fname_lower or "review" in fname_lower or "sheet" in fname_lower:
                clean_name = "Functional DMB Review Sheets-.xlsx"
            elif "master" in fname_lower or "dmb" in fname_lower:
                clean_name = "Masterfile_DMB_Dashboard.xlsx"
            else:
                clean_name = Path(filename).name
            yield clean_name, payload


def push_to_github_if_configured(filename: str, content: bytes) -> bool:
    """
    If GITHUB_TOKEN (or GH_TOKEN) is configured, automatically commits
    the updated Excel file directly to the GitHub repository main branch.
    """
    token = _env("GITHUB_TOKEN") or _env("GH_TOKEN")
    repo = _env("GITHUB_REPOSITORY", "kavyabhardwajj30/DMB-Dashboard")
    if not token:
        logger.debug("GITHUB_TOKEN not configured; skipping automatic GitHub commit.")
        return False
    try:
        import base64
        import json
        import urllib.request

        path = f"data/{filename}"
        api_url = f"https://api.github.com/repos/{repo}/contents/{path}"

        # Get existing file SHA if it exists
        sha = None
        req = urllib.request.Request(
            api_url,
            headers={
                "Authorization": f"token {token}",
                "Accept": "application/vnd.github.v3+json",
                "User-Agent": "DMB-Dashboard-Sync",
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode())
                sha = data.get("sha")
        except Exception:
            pass

        payload = {
            "message": f"Auto-sync {filename} from SharePoint via Mailbox [skip ci]",
            "content": base64.b64encode(content).decode("ascii"),
            "branch": "main",
        }
        if sha:
            payload["sha"] = sha

        put_req = urllib.request.Request(
            api_url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"token {token}",
                "Accept": "application/vnd.github.v3+json",
                "Content-Type": "application/json",
                "User-Agent": "DMB-Dashboard-Sync",
            },
            method="PUT",
        )
        with urllib.request.urlopen(put_req, timeout=15) as resp:
            if resp.status in (200, 201):
                logger.info("Successfully pushed data/%s to GitHub repo (%s)", filename, repo)
                return True
    except Exception as e:
        logger.warning("Could not auto-commit to GitHub: %s", e)
    return False


def sync_once(resolve_target: Callable[[str], Path] | None = None) -> int:
    """
    Polls the mailbox over IMAP and writes any workbooks found.
    Searches for emails with subject 'DMB-Sync', 'DMB', or general subjects.
    Scans from NEWEST to OLDEST so newest updates take precedence while retaining
    all historical emails in the mailbox archive.

    Returns the number of files written/updated.
    """
    host = _env("IMAP_HOST")
    user = _env("IMAP_USER")
    password = _env("IMAP_PASSWORD")
    if not (host and user and password):
        logger.debug("IMAP not configured; skipping mailbox sync.")
        return 0

    allowed_raw = _env("SYNC_ALLOWED_SENDER")
    allowed = {a.strip().lower() for a in allowed_raw.split(",") if a.strip()}
    if not allowed:
        allowed = {"*"}

    marker = _env("SYNC_SUBJECT_MARKER", "DMB-Sync")
    written = 0
    found_targets = set()
    needed_workbooks = {
        "Functional DMB Review Sheets-.xlsx",
        "Masterfile_DMB_Dashboard.xlsx",
    }

    try:
        with imaplib.IMAP4_SSL(host, int(_env("IMAP_PORT", "993")), timeout=15) as imap:
            clean_password = password.replace(" ", "") if "gmail.com" in host.lower() else password
            imap.login(user, clean_password)
            imap.select(_env("IMAP_FOLDER", "INBOX"))

            search_queries = [
                f'SUBJECT "{marker}"',
                'SUBJECT "DMB-Sync"',
                'SUBJECT "DMB-SYNC"',
                'SUBJECT "DMB"',
                'SUBJECT "Functional"',
                'SUBJECT "Masterfile"',
                'ALL',
            ]

            seen_ids = set()
            ordered_ids = []

            for query in search_queries:
                try:
                    status, data = imap.search(None, query)
                    if status == "OK" and data and data[0]:
                        for msg_id in data[0].split():
                            if msg_id not in seen_ids:
                                seen_ids.add(msg_id)
                                ordered_ids.append(msg_id)
                except Exception:
                    pass

            if not ordered_ids:
                return 0

            # Sort message IDs numerically (higher ID = newer email)
            ordered_ids.sort(key=lambda x: int(x) if x.isdigit() else 0)

            # Process messages from NEWEST to OLDEST (up to 200 messages)
            recent_ids = list(reversed(ordered_ids))[:200]

            for num in recent_ids:
                status, raw = imap.fetch(num, "(RFC822)")
                if status != "OK" or not raw or not raw[0]:
                    continue

                msg = email.message_from_bytes(raw[0][1])
                sender = email.utils.parseaddr(msg.get("From", ""))[1].lower()

                # Sender validation
                is_allowed = False
                if "*" in allowed:
                    is_allowed = True
                else:
                    for allow_item in allowed:
                        if allow_item.startswith("@") and sender.endswith(allow_item):
                            is_allowed = True
                            break
                        elif sender == allow_item:
                            is_allowed = True
                            break

                if not is_allowed:
                    continue

                file_found_in_msg = False
                for filename, payload in _excel_attachments(msg):
                    if resolve_target:
                        target = resolve_target(filename)
                    else:
                        target = DATA_DIR / filename

                    target_name = target.name
                    target.parent.mkdir(parents=True, exist_ok=True)

                    # Only update if this file has not yet been updated in this pass
                    if target_name not in found_targets:
                        found_targets.add(target_name)
                        is_new = True
                        if target.exists():
                            try:
                                if target.read_bytes() == payload:
                                    is_new = False
                            except Exception:
                                pass
                        if is_new:
                            target.write_bytes(payload)
                            written += 1
                            logger.info(
                                "Updated data/%s from mail (%s, %d bytes)", target_name, filename, len(payload)
                            )
                            push_to_github_if_configured(target_name, payload)
                        file_found_in_msg = True

                if file_found_in_msg:
                    imap.store(num, "+FLAGS", "\\Seen")

                # If both DMB core workbooks are found, we can finish early
                if needed_workbooks.issubset(found_targets):
                    logger.info("All DMB workbooks successfully retrieved from mailbox.")
                    break

    except Exception:
        logger.exception("DMB Mailbox sync failed")
        return 0

    return written
