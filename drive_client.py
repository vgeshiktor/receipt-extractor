import mimetypes
from pathlib import Path
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

def drive_build(creds):
    return build("drive", "v3", credentials=creds, cache_discovery=False)

def guess_mimetype(path: Path) -> str:
    mtype, _ = mimetypes.guess_type(str(path))
    return mtype or "application/octet-stream"

def upload_to_drive(drive, folder_id: str, local_path: Path, name_override: str | None = None) -> str:
    file_name = name_override or local_path.name
    file_metadata = {"name": file_name, "parents": [folder_id]}
    media = MediaFileUpload(str(local_path), mimetype=guess_mimetype(local_path))
    res = drive.files().create(body=file_metadata, media_body=media, fields="id").execute()
    return res.get("id")
