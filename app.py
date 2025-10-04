#!/usr/bin/env python3
import argparse
import json
import sys
from datetime import datetime, timezone
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
from gmail_client import gmail_build, gmail_query_string, list_message_ids as gmail_list
from drive_client import drive_build
from receipts import process_receipts

# Outlook imports
import os
from outlook_client import (
    msal_acquire_token_device_flow,
    outlook_build,
    outlook_query_filter,
    list_message_ids as outlook_list,
    extract_urls_from_message as outlook_extract_urls,
    get_attachment_parts as outlook_get_parts,
    download_attachment_bytes as outlook_download_bytes,
)

def parse_args():
    parser = argparse.ArgumentParser(description="Extract receipt attachments from email and upload to Drive.")
    parser.add_argument("--config", type=Path, help="Path to accounts config JSON. If set, overrides single-provider flags.")

    # single-provider mode (backward compatible)
    parser.add_argument("--provider", choices=["gmail", "outlook"], default="gmail", help="Email provider.")
    date = parser.add_mutually_exclusive_group(required=False)
    date.add_argument("--months-back", type=int, help="Number of months back from today (e.g., 2).")
    parser.add_argument("--after", type=str, help="Start date (YYYY/MM/DD).")
    parser.add_argument("--before", type=str, help="End date (YYYY/MM/DD).")
    parser.add_argument("--query-extra", type=str, default=None,
                        help="Gmail extra filters OR Outlook OData extra filter (advanced).")
    parser.add_argument("--drive-folder-id", type=str, required=False, help="Target Google Drive folder ID.")
    parser.add_argument("--download-dir", type=Path, default=DEFAULT_DOWNLOAD_DIR, help="Local temp download directory.")
    parser.add_argument("--log-links", type=Path, default=DEFAULT_LINK_LOG, help="Where to log external URLs.")
    parser.add_argument("--dedupe", action="store_true", help="Avoid uploading duplicates in a single run by file hash.")
    parser.add_argument("--client-secret", type=Path, default=DEFAULT_CLIENT_SECRET_PATH, help="OAuth client secret path.")
    parser.add_argument("--token", type=Path, default=DEFAULT_TOKEN_PATH, help="OAuth token path.")
    # Outlook-specific
    parser.add_argument("--ms-client-id", type=str, default=os.environ.get("AZURE_CLIENT_ID"))
    parser.add_argument("--ms-tenant", type=str, default=os.environ.get("AZURE_TENANT", "common"))
    return parser.parse_args()

def compute_dates(after: str | None, before: str | None, months_back: int | None) -> Tuple[str | None, str | None]:
    if not months_back or (after or before):
        return after, before
    today = datetime.now().date()
    start_month = today.replace(day=1)
    m = months_back
    year = start_month.year
    month = start_month.month - m
    while month <= 0:
        month += 12
        year -= 1
    start = datetime(year, month, 1).date()
    # Gmail 'before' is exclusive: first day of current month
    if start_month.month == 12:
        next_month_first = datetime(start_month.year + 1, 1, 1).date()
    else:
        next_month_first = datetime(start_month.year, start_month.month + 1, 1).date()
    return (start.strftime("%Y/%m/%d"), next_month_first.strftime("%Y/%m/%d"))

def _to_iso_utc(date_yyyy_mm_dd: str | None) -> str | None:
    if not date_yyyy_mm_dd:
        return None
    dt = datetime.strptime(date_yyyy_mm_dd, "%Y/%m/%d")
    return dt.replace(tzinfo=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

def run_gmail(creds, drive, drive_folder_id: str, download_dir: Path, log_links: Path,
              dedupe: bool, after: str | None, before: str | None, query_extra: str | None):
    gmail = gmail_build(creds)
    query = gmail_query_string(after, before, query_extra)
    print(f"[INFO] Gmail query: {query}")
    msg_ids = gmail_list(gmail, "me", query)
    print(f"[INFO] Gmail found {len(msg_ids)} emails.")

    return process_receipts(
        gmail=gmail,
        drive=drive,
        msg_ids=msg_ids,
        drive_folder_id=drive_folder_id,
        download_dir=download_dir,
        links_log_path=log_links,
        dedupe=dedupe,
    )

def run_outlook(ms_client_id: str, ms_tenant: str, drive, drive_folder_id: str, download_dir: Path, log_links: Path,
                dedupe: bool, after: str | None, before: str | None, query_extra: str | None):
    authority = f"https://login.microsoftonline.com/{ms_tenant}"
    print(f"[DEBUG] MSAL authority={authority}, client_id={ms_client_id[:6]}... (len={len(ms_client_id)})")
    token = msal_acquire_token_device_flow(client_id=ms_client_id, authority=authority)
    outlook = outlook_build(token)

    after_iso = _to_iso_utc(after)
    before_iso = _to_iso_utc(before)
    filter_expr = outlook_query_filter(after_iso, before_iso, query_extra)
    print(f"[INFO] Outlook $filter: {filter_expr or '(none)'}")
    msg_ids = outlook_list(outlook, filter_expr=filter_expr)
    print(f"[INFO] Outlook found {len(msg_ids)} emails.")

    from receipts import ReceiptRunResult
    from utils import safe_filename, file_sha256
    download_dir.mkdir(parents=True, exist_ok=True)
    seen_hashes = set()
    result = ReceiptRunResult()
    result.processed_emails = len(msg_ids)

    def log_message_urls(outlook, msg_id, log_links, result):
        if urls := outlook_extract_urls(outlook, msg_id):
            with log_links.open("a", encoding="utf-8") as f:
                for u in urls:
                    f.write(u + "\n")
            result.logged_links += len(urls)

    def process_attachment(outlook, msg_id, filename, att_id, dedupe, seen_hashes, safe_filename, download_dir, drive, drive_folder_id, result, file_sha256):
        data = outlook_download_bytes(outlook, msg_id, att_id)
        if dedupe:
            h = file_sha256(data)
            if h in seen_hashes:
                print(f"[INFO] Skip duplicate by hash: {filename}")
                return False
            seen_hashes.add(h)

        safe_name = safe_filename(filename)
        local_path = download_dir / safe_name
        local_path.write_bytes(data)

        file_id = drive.files().create(
            body={"name": safe_name, "parents": [drive_folder_id]},
            media_body=str(local_path),
            fields="id",
        ).execute()["id"]
        result.uploaded_files += 1
        print(f"[OK] Uploaded: {safe_name} → Drive ID {file_id}")
        return True

    for msg_id in msg_ids:
        try:
            log_message_urls(outlook, msg_id, log_links, result)
            parts = outlook_get_parts(outlook, msg_id)
            for (filename, att_id) in parts:
                try:
                    process_attachment(
                        outlook, msg_id, filename, att_id, dedupe, seen_hashes,
                        safe_filename, download_dir, drive, drive_folder_id, result, file_sha256
                    )
                except Exception as e:
                    result.errors.append(f"Attachment error (msg {msg_id}, {filename}): {e}")
        except Exception as e:
            result.errors.append(f"Message error (msg {msg_id}): {e}")

    return result

def main():
    args = parse_args()

    creds = ensure_creds(token_path=args.token, client_secret_path=args.client_secret, scopes=SCOPES)
    drive = drive_build(creds)

    if args.config and args.config.exists():
        return process_config_mode(args, drive)

    return process_single_provider_mode(args, drive)

def process_config_mode(args, drive):
    cfg = json.loads(args.config.read_text(encoding="utf-8"))
    mb = cfg.get("date", {}).get("months_back")
    after = cfg.get("date", {}).get("after")
    before = cfg.get("date", {}).get("before")
    after, before = compute_dates(after, before, mb)

    defaults = cfg.get("defaults", {})
    default_folder = defaults.get("drive_folder_id")
    default_download_dir = Path(defaults.get("download_dir", str(DEFAULT_DOWNLOAD_DIR)))
    default_log = Path(defaults.get("log_links", str(DEFAULT_LINK_LOG)))
    default_dedupe = bool(defaults.get("dedupe", True))

    grand_total_processed = 0
    grand_total_uploaded = 0
    grand_total_links = 0
    all_errors = []

    for acc in cfg.get("accounts", []):
        res = process_account(acc, after, before, default_folder, default_download_dir, default_log, default_dedupe, drive)
        grand_total_processed += res.processed_emails
        grand_total_uploaded += res.uploaded_files
        grand_total_links += res.logged_links
        all_errors.extend(res.errors)

    print("\n===== GRAND SUMMARY =====")
    print(f"Processed emails: {grand_total_processed}")
    print(f"Uploaded files:   {grand_total_uploaded}")
    print(f"Logged links:     {grand_total_links}")
    if all_errors:
        print(f"Errors: {len(all_errors)}")
        for e in all_errors:
            print(f"  - {e}")

def process_account(acc, after, before, default_folder, default_download_dir, default_log, default_dedupe, drive):
    provider = acc["provider"]
    name = acc.get("name", provider)
    print(f"\n===== PROVIDER: {provider} ({name}) =====")

    drive_folder_id = acc.get("drive_folder_id", default_folder)
    if not drive_folder_id:
        raise SystemExit(f"Missing drive_folder_id for account {name} and no default provided.")
    download_dir = Path(acc.get("download_dir", str(default_download_dir)))
    log_links = Path(acc.get("log_links", str(default_log)))
    dedupe = bool(acc.get("dedupe", default_dedupe))
    query_extra = acc.get("query_extra")

    if provider == "gmail":
        client_secret = Path(acc.get("client_secret", str(DEFAULT_CLIENT_SECRET_PATH)))
        token_path = Path(acc.get("token", str(DEFAULT_TOKEN_PATH)))
        gmail_creds = ensure_creds(token_path=token_path, client_secret_path=client_secret, scopes=SCOPES)
        gmail_drive = drive_build(gmail_creds)
        return run_gmail(
            creds=gmail_creds,
            drive=gmail_drive,
            drive_folder_id=drive_folder_id,
            download_dir=download_dir,
            log_links=log_links,
            dedupe=dedupe,
            after=after,
            before=before,
            query_extra=query_extra,
        )
    elif provider == "outlook":
        ms_client_id = acc.get("ms_client_id") or os.environ.get("AZURE_CLIENT_ID")
        ms_tenant = acc.get("ms_tenant", os.environ.get("AZURE_TENANT", "common"))
        if not ms_client_id:
            raise SystemExit(f"Outlook account {name} missing ms_client_id (or AZURE_CLIENT_ID).")
        return run_outlook(
            ms_client_id=ms_client_id,
            ms_tenant=ms_tenant,
            drive=drive,
            drive_folder_id=drive_folder_id,
            download_dir=download_dir,
            log_links=log_links,
            dedupe=dedupe,
            after=after,
            before=before,
            query_extra=query_extra,
        )
    else:
        raise SystemExit(f"Unknown provider: {provider}")

def process_single_provider_mode(args, drive):
    if not args.drive_folder_id:
        raise SystemExit("--drive-folder-id is required when --config is not provided.")

    after, before = compute_dates(args.after, args.before, args.months_back)
    if args.provider == "gmail":
        res = run_gmail(
            creds=ensure_creds(token_path=args.token, client_secret_path=args.client_secret, scopes=SCOPES),
            drive=drive,
            drive_folder_id=args.drive_folder_id,
            download_dir=args.download_dir,
            log_links=args.log_links,
            dedupe=args.dedupe,
            after=after,
            before=before,
            query_extra=args.query_extra,
        )
    elif ms_client_id := args.ms_client_id:
        res = run_outlook(
            ms_client_id=ms_client_id,
            ms_tenant=args.ms_tenant,
            drive=drive,
            drive_folder_id=args.drive_folder_id,
            download_dir=args.download_dir,
            log_links=args.log_links,
            dedupe=args.dedupe,
            after=after,
            before=before,
            query_extra=args.query_extra,
        )
    else:
        raise SystemExit("Missing --ms-client-id (or AZURE_CLIENT_ID).")
    print("\n===== SUMMARY =====")
    print(f"Processed emails: {res.processed_emails}")
    print(f"Uploaded files:   {res.uploaded_files}")
    print(f"Logged links:     {res.logged_links} → {args.log_links}")
    if res.errors:
        print(f"Errors: {len(res.errors)}")
        for e in res.errors:
            print(f"  - {e}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[INFO] Interrupted by user.", file=sys.stderr)
        sys.exit(130)
