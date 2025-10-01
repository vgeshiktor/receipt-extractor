#!/usr/bin/env python3
import argparse
import base64
import hashlib
import mimetypes
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow

# =========================
# Config & Constants
# =========================

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.file",
]

DEFAULT_DOWNLOAD_DIR = Path("downloaded_receipts")
DEFAULT_LINK_LOG = Path("external_links.txt")

# =========================
# Auth
# =========================

def ensure_creds(
    token_path: Path = Path("token.json"),
    client_secret_path: Path = Path("client_secret.json"),
    scopes: List[str] = SCOPES,
):
    creds: Optional[Credentials] = None

    def run_flow() -> Credentials:
        if not client_secret_path.exists():
            raise FileNotFoundError(
                f"Missing {client_secret_path}. Download OAuth client secrets from Google Cloud Console."
            )
        flow = InstalledAppFlow.from_client_secrets_file(str(client_secret_path), scopes)
        # Force refresh_token issuance on first consent
        creds_local = flow.run_local_server(
            host="localhost",
            port=8080,
            access_type="offline",
            prompt="consent",
            include_granted_scopes="true",
        )
        token_path.write_text(creds_local.to_json(), encoding="utf-8")
        return creds_local

    # Try loading existing token.json
    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), scopes)
        except ValueError:
            # token.json malformed or missing fields -> redo flow
            creds = run_flow()

    # Refresh or run flow if needed
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                token_path.write_text(creds.to_json(), encoding="utf-8")
            except Exception:
                creds = run_flow()
        else:
            creds = run_flow()

    return creds

# =========================
# Gmail helpers
# =========================

def gmail_build(creds: Credentials):
    return build("gmail", "v1", credentials=creds, cache_discovery=False)

def drive_build(creds: Credentials):
    return build("drive", "v3", credentials=creds, cache_discovery=False)

def gmail_query_string(after: Optional[str], before: Optional[str], extra: Optional[str]) -> str:
    """
    Compose a Gmail search query. Dates must be YYYY/MM/DD.
    """
    parts = ["has:attachment"]
    if after:
        parts.append(f"after:{after}")
    if before:
        parts.append(f"before:{before}")
    if extra:
        parts.append(extra)
    return " ".join(parts)

def list_message_ids(gmail, user_id: str, query: str) -> List[str]:
    """
    Paginate through Gmail results and return list of message IDs.
    """
    ids: List[str] = []
    request = gmail.users().messages().list(userId=user_id, q=query)
    while request is not None:
        resp = request.execute()
        for m in resp.get("messages", []):
            ids.append(m["id"])
        request = gmail.users().messages().list_next(request, resp)
    return ids

def traverse_parts(payload: Dict) -> Iterable[Dict]:
    """
    Depth-first traversal of MIME parts to yield all parts.
    """
    if not payload:
        return
    stack = [payload]
    while stack:
        part = stack.pop()
        yield part
        for child in part.get("parts", []) or []:
            stack.append(child)

URL_REGEX = re.compile(r"https?://[^\s<>\"]+", re.IGNORECASE)

def extract_urls_from_message(gmail, user_id: str, msg_id: str) -> List[str]:
    """
    Extract URLs from text/plain or text/html parts of an email.
    """
    message = gmail.users().messages().get(userId=user_id, id=msg_id, format="full").execute()
    urls: List[str] = []
    for part in traverse_parts(message.get("payload", {})):
        mime = part.get("mimeType", "")
        body = part.get("body", {})
        data = body.get("data")
        if data and (mime.startswith("text/plain") or mime.startswith("text/html")):
            try:
                text = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
                urls.extend(URL_REGEX.findall(text))
            except Exception:
                # Non-fatal: skip body decoding errors
                pass
    # Deduplicate
    return sorted(set(urls))

def get_attachment_parts(gmail, user_id: str, msg_id: str) -> List[Tuple[str, str]]:
    """
    Return list of (filename, attachment_id) for all non-empty filename parts.
    """
    message = gmail.users().messages().get(userId=user_id, id=msg_id, format="full").execute()
    results: List[Tuple[str, str]] = []
    for part in traverse_parts(message.get("payload", {})):
        fname = part.get("filename")
        body = part.get("body", {})
        att_id = body.get("attachmentId")
        if fname and att_id:
            results.append((fname, att_id))
    return results

def download_attachment_bytes(gmail, user_id: str, msg_id: str, attachment_id: str) -> bytes:
    att = gmail.users().messages().attachments().get(
        userId=user_id, messageId=msg_id, id=attachment_id
    ).execute()
    data = att.get("data")
    if not data:
        raise RuntimeError("Attachment has no data")
    return base64.urlsafe_b64decode(data)

# =========================
# Drive helpers
# =========================

def safe_filename(name: str) -> str:
    name = name.strip().replace("/", "_").replace("\\", "_")
    if not name:
        name = "attachment"
    return name

def file_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def guess_mimetype(path: Path) -> str:
    mtype, _ = mimetypes.guess_type(str(path))
    return mtype or "application/octet-stream"

def upload_to_drive(drive, folder_id: str, local_path: Path, name_override: Optional[str] = None) -> str:
    file_name = name_override or local_path.name
    file_metadata = {"name": file_name, "parents": [folder_id]}
    media = MediaFileUpload(str(local_path), mimetype=guess_mimetype(local_path))
    res = drive.files().create(body=file_metadata, media_body=media, fields="id").execute()
    return res.get("id")

# =========================
# Orchestration
# =========================

def parse_args():
    parser = argparse.ArgumentParser(description="Extract receipt attachments from Gmail and upload to Drive.")
    date = parser.add_mutually_exclusive_group(required=False)
    date.add_argument("--months-back", type=int, help="Number of months back from today (e.g., 2).")
    parser.add_argument("--after", type=str, help="Start date (YYYY/MM/DD).")
    parser.add_argument("--before", type=str, help="End date (YYYY/MM/DD).")

    parser.add_argument("--query-extra", type=str, default=None,
                        help="Extra Gmail query filters (e.g., 'subject:(invoice OR receipt) OR category:finance').")

    parser.add_argument("--drive-folder-id", type=str, required=True, help="Target Google Drive folder ID.")
    parser.add_argument("--download-dir", type=Path, default=DEFAULT_DOWNLOAD_DIR, help="Local temp download directory.")
    parser.add_argument("--log-links", type=Path, default=DEFAULT_LINK_LOG, help="Where to log external URLs.")
    parser.add_argument("--dedupe", action="store_true", help="Avoid uploading duplicates in a single run by file hash.")
    parser.add_argument("--client-secret", type=Path, default=Path("client_secret.json"), help="OAuth client secret path.")
    parser.add_argument("--token", type=Path, default=Path("token.json"), help="OAuth token path.")
    return parser.parse_args()

def compute_dates(after: Optional[str], before: Optional[str], months_back: Optional[int]) -> Tuple[Optional[str], Optional[str]]:
    if months_back and not (after or before):
        today = datetime.now().date()
        start_month = today.replace(day=1)
        # months_back full months; approximate by subtracting months
        m = months_back
        year = start_month.year
        month = start_month.month - m
        while month <= 0:
            month += 12
            year -= 1
        start = datetime(year, month, 1).date()
        # Gmail before is exclusive. Use first day of next month.
        if start_month.month == 12:
            next_month_first = datetime(start_month.year + 1, 1, 1).date()
        else:
            next_month_first = datetime(start_month.year, start_month.month + 1, 1).date()
        return (start.strftime("%Y/%m/%d"), next_month_first.strftime("%Y/%m/%d"))
    # Otherwise pass-through (user-specified)
    return after, before

def main():
    args = parse_args()

    # Date handling
    after, before = compute_dates(args.after, args.before, args.months_back)
    query = gmail_query_string(after, before, args.query_extra)

    # Auth + clients
    creds = ensure_creds(token_path=args.token, client_secret_path=args.client_secret, scopes=SCOPES)
    gmail = gmail_build(creds)
    drive = drive_build(creds)

    # Ensure local dirs
    args.download_dir.mkdir(parents=True, exist_ok=True)

    # Search
    print(f"[INFO] Gmail query: {query}")
    msg_ids = list_message_ids(gmail, "me", query)
    print(f"[INFO] Found {len(msg_ids)} emails.")

    # Dedupe set for a single run
    seen_hashes: set = set()

    total_files = 0
    total_links = 0
    errors: List[str] = []

    # Process emails
    for i, msg_id in enumerate(msg_ids, start=1):
        try:
            # External links (manual step)
            urls = extract_urls_from_message(gmail, "me", msg_id)
            if urls:
                with args.log_links.open("a", encoding="utf-8") as f:
                    for u in urls:
                        f.write(u + "\n")
                total_links += len(urls)

            # Attachments
            parts = get_attachment_parts(gmail, "me", msg_id)
            if not parts:
                continue
            for (filename, att_id) in parts:
                try:
                    data = download_attachment_bytes(gmail, "me", msg_id, att_id)
                    if args.dedupe:
                        h = file_sha256(data)
                        if h in seen_hashes:
                            print(f"[INFO] Skip duplicate by hash: {filename}")
                            continue
                        seen_hashes.add(h)

                    safe_name = safe_filename(filename)
                    local_path = args.download_dir / safe_name
                    with open(local_path, "wb") as f:
                        f.write(data)

                    # Upload to Drive
                    file_id = upload_to_drive(drive, args.drive_folder_id, local_path)
                    total_files += 1
                    print(f"[OK] Uploaded: {safe_name} → Drive ID {file_id}")
                except Exception as e:
                    errors.append(f"Attachment error (msg {msg_id}, {filename}): {e}")
        except Exception as e:
            errors.append(f"Message error (msg {msg_id}): {e}")

    # Summary
    print("\n===== SUMMARY =====")
    print(f"Processed emails: {len(msg_ids)}")
    print(f"Uploaded files:   {total_files}")
    print(f"Logged links:     {total_links} → {args.log_links}")
    if errors:
        print(f"Errors: {len(errors)}")
        for e in errors:
            print(f"  - {e}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[INFO] Interrupted by user.", file=sys.stderr)
        sys.exit(130)
