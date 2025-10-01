from pathlib import Path
from typing import List, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

def ensure_creds(
    token_path: Path,
    client_secret_path: Path,
    scopes: List[str],
) -> Credentials:
    """
    Return valid Google OAuth2 credentials (authorized-user JSON).
    Resilient to malformed/old token; forces refresh_token on first consent.
    """
    creds: Optional[Credentials] = None

    def run_flow() -> Credentials:
        if not client_secret_path.exists():
            raise FileNotFoundError(
                f"Missing {client_secret_path}. Download OAuth client secrets from Google Cloud Console."
            )
        flow = InstalledAppFlow.from_client_secrets_file(str(client_secret_path), scopes)
        creds_local = flow.run_local_server(
            host="localhost",
            port=8080,
            access_type="offline",
            prompt="consent",
            include_granted_scopes="true",
        )
        token_path.write_text(creds_local.to_json(), encoding="utf-8")
        return creds_local

    # Try loading existing token
    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), scopes)
        except ValueError:
            # token.json malformed/missing fields -> redo flow
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
