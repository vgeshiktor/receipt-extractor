import hashlib

def safe_filename(name: str) -> str:
    name = name.strip().replace("/", "_").replace("\\", "_")
    return name or "attachment"

def file_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()
