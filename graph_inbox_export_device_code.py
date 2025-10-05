#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Graph Inbox Export (Device Code, PublicClientApplication)
--------------------------------------------------------
Reads your Outlook inbox via Microsoft Graph using **MSAL Device Code** flow
(no client secret required) and exports messages to JSON and/or CSV.

Prereqs:
  pip install msal httpx

Usage:
  export MS_CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
  python3 graph_inbox_export_device_code.py --out-json inbox.json --out-csv inbox.csv --max 500

Notes:
- Default authority targets **personal Microsoft accounts** ("consumers").
  Use --authority common if you registered for any-org + personal accounts.
- The first run will prompt you to visit https://microsoft.com/devicelogin
  and enter a code. Subsequent runs will be silent (token cache persisted).

Examples:
  # Last 200 messages from Inbox with a custom field selection
  python3 graph_inbox_export_device_code.py --max 200 --select "id,subject,from,receivedDateTime,bodyPreview,hasAttachments,isRead,webLink"

  # Filter only messages from a sender since a date
  python3 graph_inbox_export_device_code.py --filter "from/emailAddress/address eq 'someone@example.com' and receivedDateTime ge 2025-09-01T00:00:00Z"

  # Target a different well-known folder (e.g., 'SentItems')
  python3 graph_inbox_export_device_code.py --folder sentitems --max 100
"""

import os
import sys
import csv
import json
import time
import argparse
from pathlib import Path
from typing import Any, Dict, List, Optional

import httpx
import msal

DEFAULT_AUTHORITY = "https://login.microsoftonline.com/consumers"  # personal Microsoft accounts
TOKEN_CACHE_PATH = Path.home() / ".msal_dc_token_cache.bin"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read"]  # delegated

# Default fields to export
DEFAULT_SELECT = ",".join([
    "id",
    "subject",
    "from",
    "sender",
    "receivedDateTime",
    "sentDateTime",
    "bodyPreview",
    "hasAttachments",
    "isRead",
    "importance",
    "internetMessageId",
    "conversationId",
    "webLink",
    "toRecipients",
    "ccRecipients",
    "bccRecipients",
    "replyTo",
    "categories",
])


# ---------------------------
# Auth (Device Code) helpers
# ---------------------------
def build_public_app(client_id: str, authority: str, cache_path: Path) -> msal.PublicClientApplication:
    cache = msal.SerializableTokenCache()
    if cache_path.exists():
        cache.deserialize(cache_path.read_text(encoding="utf-8"))
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=cache,
    )
    app._cache_path = cache_path  # attach for persistence later
    app._cache_obj = cache
    return app


def persist_cache(app: msal.PublicClientApplication) -> None:
    cache: msal.SerializableTokenCache = getattr(app, "_cache_obj", None)
    path: Path = getattr(app, "_cache_path", None)
    if cache and path and cache.has_state_changed:
        path.write_text(cache.serialize(), encoding="utf-8")


def acquire_token_device_code(app: msal.PublicClientApplication) -> str:
    if accounts := app.get_accounts():
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            persist_cache(app)
            return result["access_token"]

    # 2) Device code flow (interactive once)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError("Failed to create device code flow.")
    print(flow["message"])  # e.g., "To sign in, use a web browser to open ... and enter the code ..."
    result = app.acquire_token_by_device_flow(flow)  # will block until done or timeout

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {json.dumps(result, indent=2)}")

    persist_cache(app)
    return result["access_token"]


# ---------------------------
# Graph client with retries
# ---------------------------
class GraphClient:
    def __init__(self, token_provider, timeout: float = 30.0, max_retries: int = 4):
        self.token_provider = token_provider
        self.access_token = token_provider()
        self.client = httpx.Client(timeout=timeout)
        self.max_retries = max_retries

    def _headers(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self.access_token}"}

    def request(self, method: str, url: str, **kwargs) -> httpx.Response:
        backoff = 0.5
        last_exc = None
        for attempt in range(self.max_retries):
            try:
                merged_headers = self._headers()
                if extra_headers := kwargs.get("headers"):
                    merged_headers.update(extra_headers)
                request_kwargs = {k: v for k, v in kwargs.items() if k != "headers"}
                resp = self.client.request(method, url, headers=merged_headers, **request_kwargs)
            except httpx.HTTPError as e:
                last_exc = e
                time.sleep(backoff)
                backoff = min(backoff * 2, 8.0)
                continue

            # Refresh on 401 once
            if resp.status_code == 401 and attempt == 0:
                self.access_token = self.token_provider(force=True)
                continue

            # Retry on 429/5xx with backoff/Retry-After
            if resp.status_code in (429, 500, 502, 503, 504):
                ra = resp.headers.get("Retry-After")
                sleep_for = float(ra) if ra and ra.isdigit() else backoff
                time.sleep(sleep_for)
                backoff = min(backoff * 2, 8.0)
                continue

            return resp

        if last_exc:
            raise last_exc
        return resp

    def get(self, url: str, params: Optional[Dict[str, Any]] = None, headers: Optional[Dict[str, str]] = None) -> httpx.Response:
        return self.request("GET", url, params=params, headers=headers or {})


# ---------------------------
# Message retrieval + export
# ---------------------------
def flatten_addresses(recipients: Optional[List[Dict[str, Any]]]) -> str:
    if not recipients:
        return ""
    emails = []
    for r in recipients:
        # Validate recipient is a dict and has "emailAddress" as a dict
        if not isinstance(r, dict):
            continue
        addr = r.get("emailAddress")
        if not isinstance(addr, dict):
            continue
        name = addr.get("name") or ""
        email = addr.get("address") or ""
        if not email:
            continue  # skip if no email address
        emails.append(f"{name} <{email}>" if name else email)
    return "; ".join(emails)


def message_to_row(m: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "id": m.get("id"),
        "subject": m.get("subject"),
        "from": (m.get("from") or {}).get("emailAddress", {}).get("address"),
        "from_name": (m.get("from") or {}).get("emailAddress", {}).get("name"),
        "sender": (m.get("sender") or {})
        .get("emailAddress", {})
        .get("address"),
        "receivedDateTime": m.get("receivedDateTime"),
        "sentDateTime": m.get("sentDateTime"),
        "bodyPreview": m.get("bodyPreview"),
        "hasAttachments": m.get("hasAttachments"),
        "isRead": m.get("isRead"),
        "importance": m.get("importance"),
        "internetMessageId": m.get("internetMessageId"),
        "conversationId": m.get("conversationId"),
        "webLink": m.get("webLink"),
        "toRecipients": flatten_addresses(m.get("toRecipients")),
        "ccRecipients": flatten_addresses(m.get("ccRecipients")),
        "bccRecipients": flatten_addresses(m.get("bccRecipients")),
        "replyTo": flatten_addresses(m.get("replyTo")),
        "categories": "; ".join(m.get("categories") or []),
    }


def list_messages(graph: GraphClient, folder: str, select: str, odata_filter: Optional[str], top: int, max_results: int) -> List[Dict[str, Any]]:
    base = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
    params: Dict[str, Any] = {"$select": select, "$top": min(top, max_results)}
    if odata_filter:
        params["$filter"] = odata_filter

    messages: List[Dict[str, Any]] = []
    url = base
    while url and len(messages) < max_results:
        resp = graph.get(url, params=params)
        if resp.status_code != 200:
            raise RuntimeError(f"Graph error {resp.status_code}: {resp.text}")
        payload = resp.json()
        batch = payload.get("value", [])
        messages.extend(batch)
        url = payload.get("@odata.nextLink")
        params = None
        if url and len(messages) + top > max_results:
            # reduce next page size, update or add $top in nextLink
            from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

            remaining = max_results - len(messages)
            parsed = urlparse(url)
            query = parse_qs(parsed.query)
            query["$top"] = [str(remaining)]
            new_query = urlencode(query, doseq=True)
            url = urlunparse(parsed._replace(query=new_query))
    return messages[:max_results]


def save_json(messages: List[Dict[str, Any]], path: Path) -> None:
    path.write_text(json.dumps(messages, indent=2, ensure_ascii=False), encoding="utf-8")


def save_csv(messages: List[Dict[str, Any]], path: Path) -> None:
    rows = [message_to_row(m) for m in messages]
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    fieldnames = list(rows[0].keys())
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)


# ---------------------------
# CLI
# ---------------------------
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Export Outlook Inbox to JSON/CSV using MSAL Device Code (PublicClientApplication)")
    p.add_argument("--client-id", default=os.getenv("MS_CLIENT_ID") or "", help="Application (client) ID. Defaults to env MS_CLIENT_ID.")
    p.add_argument("--authority", default=DEFAULT_AUTHORITY, help="Authority URL (default: consumers). Use https://login.microsoftonline.com/common if needed.")
    p.add_argument("--folder", default="inbox", help="Mail folder to read (well-known name or folder ID). Default: inbox")
    p.add_argument("--select", default=DEFAULT_SELECT, help=f"Comma-separated fields to select. Default: {DEFAULT_SELECT}")
    p.add_argument("--filter", dest="odata_filter", help="OData $filter string, e.g. \"receivedDateTime ge 2025-09-01T00:00:00Z and isRead eq false\"")
    p.add_argument("--top", type=int, default=50, help="Page size per request (Graph $top). Default: 50")
    p.add_argument("--max", dest="max_results", type=int, default=200, help="Max messages to fetch. Default: 200")
    p.add_argument("--out-json", type=Path, help="Path to write JSON export")
    p.add_argument("--out-csv", type=Path, help="Path to write CSV export")
    p.add_argument("--timeout", type=float, default=30.0, help="HTTP timeout seconds (default: 30)")
    p.add_argument("--retries", type=int, default=4, help="Max retries for 429/5xx (default: 4)")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    if not args.client_id:
        print("ERROR: Please supply --client-id or set MS_CLIENT_ID.", file=sys.stderr)
        sys.exit(1)
    if not (args.out_json or args.out_csv):
        print("ERROR: Please provide at least one of --out-json or --out-csv.", file=sys.stderr)
        sys.exit(1)

    app = build_public_app(args.client_id, args.authority, TOKEN_CACHE_PATH)

    # Token provider closure with optional forced refresh (used on 401)
    def token_provider(force: bool = False) -> str:
        if not force:
            # try silent first
            accounts = app.get_accounts()
            if accounts:
                res = app.acquire_token_silent(SCOPES, account=accounts[0])
                if res and "access_token" in res:
                    persist_cache(app)
                    return res["access_token"]
        # device code (may be skipped if still valid refresh tokens exist)
        return acquire_token_device_code(app)

    graph = GraphClient(token_provider, timeout=args.timeout, max_retries=args.retries)

    print(f"Fetching up to {args.max_results} messages from folder '{args.folder}' ...")
    msgs = list_messages(
        graph=graph,
        folder=args.folder,
        select=args.select,
        odata_filter=args.odata_filter,
        top=args.top,
        max_results=args.max_results,
    )
    print(f"Retrieved {len(msgs)} message(s).")

    if args.out_json:
        save_json(msgs, args.out_json)
        print(f"Wrote JSON: {args.out_json}")

    if args.out_csv:
        save_csv(msgs, args.out_csv)
        print(f"Wrote CSV:  {args.out_csv}")


if __name__ == "__main__":
    main()
