from __future__ import annotations

from pathlib import Path
from dotenv import load_dotenv
import os

BASE_DIR = Path(__file__).resolve().parents[1]
load_dotenv(BASE_DIR / ".env")

DATABASE_URL = os.getenv(
    "DATABASE_URL", f"sqlite:///{(BASE_DIR / 'packing.db').as_posix()}"
)
MSAL_CLIENT_ID = os.getenv("MSAL_CLIENT_ID")
MSAL_TENANT_ID = os.getenv("MSAL_TENANT_ID")
MSAL_AUTHORITY = os.getenv(
    "MSAL_AUTHORITY", f"https://login.microsoftonline.com/{MSAL_TENANT_ID or 'common'}"
)
SHAREPOINT_URL = os.getenv("SHAREPOINT_URL")
SHAREPOINT_LOCAL_DIR = Path(
    os.getenv("SHAREPOINT_LOCAL_DIR", str(BASE_DIR / "build" / "sharepoint_cache"))
)
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "build" / "outputs"
MAPPINGS_FILE = BASE_DIR / "config" / "mappings.json"
APP_NAME = "Beta Packing List"


def ensure_runtime_dirs() -> None:
    """Create directories used by the desktop app if they do not exist."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SHAREPOINT_LOCAL_DIR.mkdir(parents=True, exist_ok=True)


def resolve_path(value: str | None, default: Path) -> Path:
    if not value:
        return default
    path = Path(value)
    return path if path.is_absolute() else BASE_DIR / path
