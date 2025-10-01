from pathlib import Path

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.file",
]

DEFAULT_DOWNLOAD_DIR = Path("downloaded_receipts")
DEFAULT_LINK_LOG = Path("external_links.txt")

DEFAULT_TOKEN_PATH = Path("token.json")
DEFAULT_CLIENT_SECRET_PATH = Path("client_secret.json")
