from pathlib import Path
from typing import List

from gmail_client import (
    extract_urls_from_message,
    get_attachment_parts,
    download_attachment_bytes,
)
from drive_client import upload_to_drive
from utils import safe_filename, file_sha256

class ReceiptRunResult:
    def __init__(self):
        self.processed_emails = 0
        self.uploaded_files = 0
        self.logged_links = 0
        self.errors: List[str] = []

def process_receipts(
    gmail,
    drive,
    msg_ids: list[str],
    drive_folder_id: str,
    download_dir: Path,
    links_log_path: Path,
    dedupe: bool,
) -> ReceiptRunResult:
    download_dir.mkdir(parents=True, exist_ok=True)
    seen_hashes: set[str] = set()
    result = ReceiptRunResult()
    result.processed_emails = len(msg_ids)

    for msg_id in msg_ids:
        try:
            urls = extract_urls_from_message(gmail, "me", msg_id)
            if urls:
                with links_log_path.open("a", encoding="utf-8") as f:
                    for u in urls:
                        f.write(u + "\n")
                result.logged_links += len(urls)

            parts = get_attachment_parts(gmail, "me", msg_id)
            for (filename, att_id) in parts:
                try:
                    data = download_attachment_bytes(gmail, "me", msg_id, att_id)
                    if dedupe:
                        h = file_sha256(data)
                        if h in seen_hashes:
                            print(f"[INFO] Skip duplicate by hash: {filename}")
                            continue
                        seen_hashes.add(h)

                    safe_name = safe_filename(filename)
                    local_path = download_dir / safe_name
                    local_path.write_bytes(data)

                    file_id = upload_to_drive(drive, drive_folder_id, local_path)
                    result.uploaded_files += 1
                    print(f"[OK] Uploaded: {safe_name} → Drive ID {file_id}")
                except Exception as e:
                    result.errors.append(f"Attachment error (msg {msg_id}, {filename}): {e}")
        except Exception as e:
            result.errors.append(f"Message error (msg {msg_id}): {e}")

    return result
