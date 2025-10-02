import base64
import os
import re
from typing import Dict, List, Tuple

import requests
import msal

URL_REGEX = re.compile(r"https?://[^\s<>\"]+", re.IGNORECASE)
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Scopes מפורשים (Delegated) ל-Device Code Flow
GRAPH_SCOPES = ["User.Read", "Mail.ReadWrite", "Mail.Send", "email"]

def _token_cache_path():
    return os.environ.get("MSAL_TOKEN_CACHE", "token_ms.json")

def _load_msal_cache():
    cache = msal.SerializableTokenCache()
    path = _token_cache_path()
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())
    return cache

def _save_msal_cache(cache: msal.SerializableTokenCache):
    path = _token_cache_path()
    with open(path, "w", encoding="utf-8") as f:
        f.write(cache.serialize())

def msal_acquire_token_device_flow(client_id: str, authority: str) -> str:
    """
    מחזיר access token ל-Microsoft Graph באמצעות Device Code Flow.
    """
    cache = _load_msal_cache()
    app = msal.PublicClientApplication(client_id=client_id, authority=authority, token_cache=cache)

    # נסיון שקט
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_msal_cache(cache)
            return result["access_token"]

    # Device Code Flow
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to create device flow for MSAL. Details: {flow}")
    print(f"[Outlook] Go to {flow['verification_uri']} and enter code: {flow['user_code']}")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Device code flow failed: {result}")
    _save_msal_cache(cache)
    return result["access_token"]

def outlook_build(access_token: str):
    """
    אובייקט "לקוח" מינימלי עבור Outlook (שומר את הטוקן).
    """
    return {"access_token": access_token}

def _headers(token: str):
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}

def outlook_query_filter(after_iso: str | None, before_iso: str | None, extra_odata: str | None = None) -> str:
    """
    מחזיר ביטוי $filter ל-Graph, למשל:
    receivedDateTime ge 2025-07-01T00:00:00Z and receivedDateTime lt 2025-08-01T00:00:00Z
    """
    filters = []
    if after_iso:
        filters.append(f"receivedDateTime ge {after_iso}")
    if before_iso:
        filters.append(f"receivedDateTime lt {before_iso}")
    base = " and ".join(filters) if filters else ""
    if extra_odata:
        base = f"{base} and ({extra_odata})" if base else extra_odata
    return base

def list_message_ids(outlook, user: str = "me", filter_expr: str = "", token: str | None = None) -> List[str]:
    """
    מחזיר רשימת message IDs הכוללים קבצים מצורפים (hasAttachments=true).
    """
    token = token or outlook["access_token"]
    ids: List[str] = []
    url = f"{GRAPH_BASE}/me/messages"
    params = {
        "$select": "id,hasAttachments",
        "$orderby": "receivedDateTime DESC",
        "$top": "50",
    }
    if filter_expr:
        params["$filter"] = filter_expr

    while True:
        resp = requests.get(url, headers=_headers(token), params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        for item in data.get("value", []):
            if item.get("hasAttachments"):
                ids.append(item["id"])
        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        url = next_link
        params = None  # nextLink כבר מכיל את הפרמטרים
    return ids

def _get_message(token: str, msg_id: str) -> Dict:
    url = f"{GRAPH_BASE}/me/messages/{msg_id}"
    resp = requests.get(url, headers=_headers(token), params={"$select": "id,body,bodyPreview"}, timeout=30)
    resp.raise_for_status()
    return resp.json()

def extract_urls_from_message(outlook, msg_id: str) -> List[str]:
    """
    חילוץ קישורים מה-HTML של ההודעה.
    """
    token = outlook["access_token"]
    msg = _get_message(token, msg_id)
    body = msg.get("body", {})
    content = body.get("content") or ""
    urls = URL_REGEX.findall(content)
    return sorted(set(urls))

def get_attachment_parts(outlook, msg_id: str) -> List[Tuple[str, str]]:
    """
    מחזיר [(filename, attachment_id)] עבור קובצי FileAttachment (מדלג על ItemAttachment במ״פ).
    """
    token = outlook["access_token"]
    url = f"{GRAPH_BASE}/me/messages/{msg_id}/attachments"
    parts: List[Tuple[str, str]] = []
    while True:
        resp = requests.get(url, headers=_headers(token), timeout=30)
        resp.raise_for_status()
        data = resp.json()
        for att in data.get("value", []):
            if "@odata.type" in att and att["@odata.type"].endswith("FileAttachment"):
                name = att.get("name") or att.get("fileName") or "attachment"
                parts.append((name, att["id"]))
        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        url = next_link
    return parts

def download_attachment_bytes(outlook, msg_id: str, attachment_id: str) -> bytes:
    """
    מוריד תוכן של קובץ מצורף (contentBytes) ומחזיר bytes.
    """
    token = outlook["access_token"]
    url = f"{GRAPH_BASE}/me/messages/{msg_id}/attachments/{attachment_id}"
    resp = requests.get(url, headers=_headers(token), timeout=30)
    resp.raise_for_status()
    att = resp.json()
    content = att.get("contentBytes")
    if not content:
        raise RuntimeError("Attachment has no contentBytes.")
    return base64.b64decode(content)
