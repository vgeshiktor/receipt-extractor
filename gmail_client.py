import base64
import re
from typing import Dict, Iterable, List, Tuple

from googleapiclient.discovery import build

URL_REGEX = re.compile(r"https?://[^\s<>\"]+", re.IGNORECASE)

def gmail_build(creds):
    return build("gmail", "v1", credentials=creds, cache_discovery=False)

def gmail_query_string(after: str | None, before: str | None, extra: str | None) -> str:
    parts = ["has:attachment"]
    if after:
        parts.append(f"after:{after}")
    if before:
        parts.append(f"before:{before}")
    if extra:
        parts.append(extra)
    return " ".join(parts)

def list_message_ids(gmail, user_id: str, query: str) -> List[str]:
    ids: List[str] = []
    request = gmail.users().messages().list(userId=user_id, q=query)
    while request is not None:
        resp = request.execute()
        for m in resp.get("messages", []):
            ids.append(m["id"])
        request = gmail.users().messages().list_next(request, resp)
    return ids

def traverse_parts(payload: Dict) -> Iterable[Dict]:
    if not payload:
        return
    stack = [payload]
    while stack:
        part = stack.pop()
        yield part
        for child in part.get("parts", []) or []:
            stack.append(child)

def extract_urls_from_message(gmail, user_id: str, msg_id: str) -> List[str]:
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
                pass
    return sorted(set(urls))

def get_attachment_parts(gmail, user_id: str, msg_id: str) -> List[Tuple[str, str]]:
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
