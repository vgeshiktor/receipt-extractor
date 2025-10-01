#!/usr/bin/env python3
import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Tuple

from auth import ensure_creds
from config import (
    SCOPES,
    DEFAULT_CLIENT_SECRET_PATH,
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_LINK_LOG,
    DEFAULT_TOKEN_PATH,
)
from gmail_client import gmail_build, gmail_query_string, list_message_ids
from drive_client import drive_build
from receipts import process_receipts

def parse_args():
    parser = argparse.ArgumentParser(description="Extract receipt attachments from Gmail and upload to Drive.")
    date = parser.add_mutually_exclusive_group(required=False)
    date.add_argument("--months-back", type=int, help="Number of months back from today (e.g., 2).")
    parser.add_argument("--after", type=str, help="Start date (YYYY/MM/DD).")
    parser.add_argument("--before", type=str, help="End date (YYYY/MM/DD).")

    parser.add_argument("--query-extra", type=str, default=None,
                        help="Extra Gmail filters (e.g., 'subject:(invoice OR receipt) OR category:finance').")

    parser.add_argument("--drive-folder-id", type=str, required=True, help="Target Google Drive folder ID.")
    parser.add_argument("--download-dir", type=Path, default=DEFAULT_DOWNLOAD_DIR, help="Local temp download directory.")
    parser.add_argument("--log-links", type=Path, default=DEFAULT_LINK_LOG, help="Where to log external URLs.")
    parser.add_argument("--dedupe", action="store_true", help="Avoid uploading duplicates in a single run by file hash.")
    parser.add_argument("--client-secret", type=Path, default=DEFAULT_CLIENT_SECRET_PATH, help="OAuth client secret path.")
    parser.add_argument("--token", type=Path, default=DEFAULT_TOKEN_PATH, help="OAuth token path.")
    return parser.parse_args()

def compute_dates(after: str | None, before: str | None, months_back: int | None) -> Tuple[str | None, str | None]:
    if months_back and not (after or before):
        today = datetime.now().date()
        start_month = today.replace(day=1)
        m = months_back
        year = start_month.year
        month = start_month.month - m
        while month <= 0:
            month += 12
            year -= 1
        start = datetime(year, month, 1).date()
        # Gmail 'before' is exclusive → first day of current month
        if start_month.month == 12:
            next_month_first = datetime(start_month.year + 1, 1, 1).date()
        else:
            next_month_first = datetime(start_month.year, start_month.month + 1, 1).date()
        return (start.strftime("%Y/%m/%d"), next_month_first.strftime("%Y/%m/%d"))
    return after, before

def main():
    args = parse_args()

    after, before = compute_dates(args.after, args.before, args.months_back)
    query = gmail_query_string(after, before, args.query_extra)

    creds = ensure_creds(token_path=args.token, client_secret_path=args.client_secret, scopes=SCOPES)
    gmail = gmail_build(creds)
    drive = drive_build(creds)

    print(f"[INFO] Gmail query: {query}")
    msg_ids = list_message_ids(gmail, "me", query)
    print(f"[INFO] Found {len(msg_ids)} emails.")

    result = process_receipts(
        gmail=gmail,
        drive=drive,
        msg_ids=msg_ids,
        drive_folder_id=args.drive_folder_id,
        download_dir=args.download_dir,
        links_log_path=args.log_links,
        dedupe=args.dedupe,
    )

    print("\n===== SUMMARY =====")
    print(f"Processed emails: {result.processed_emails}")
    print(f"Uploaded files:   {result.uploaded_files}")
    print(f"Logged links:     {result.logged_links} → {args.log_links}")
    if result.errors:
        print(f"Errors: {len(result.errors)}")
        for e in result.errors:
            print(f"  - {e}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[INFO] Interrupted by user.", file=sys.stderr)
        sys.exit(130)
