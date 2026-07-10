from __future__ import annotations

import sys
from pathlib import Path
from dotenv import load_dotenv
import os


def _frozen() -> bool:
    return getattr(sys, "frozen", False)


def app_dir() -> Path:
    """Directorio donde está el .exe (o el script). Para .env, logs, etc."""
    if _frozen():
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


def bundle_dir() -> Path:
    """Directorio de datos empaquetados (sys._MEIPASS). Para templates/, .ico, etc."""
    if _frozen():
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parents[1]


BASE_DIR = bundle_dir()
APP_DIR = app_dir()

# ─── Cargar .env ──────────────────────────────────────────────────────────────
if _frozen():
    # Compilado: .env está empaquetado dentro del .exe (sys._MEIPASS)
    env_path = BASE_DIR / ".env"
else:
    # Desarrollo: .env junto al script
    env_path = APP_DIR / ".env"

if env_path.exists():
    load_dotenv(env_path)

# ─── Supabase ─────────────────────────────────────────────────────────────────
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

# ─── SharePoint (override opcional — si no está, se busca automáticamente) ───
_sharepoint_env = os.getenv("SHAREPOINT_LOCAL_DIR", "")
SHAREPOINT_LOCAL_DIR: Path | None = Path(_sharepoint_env) if _sharepoint_env else None

# ─── Rutas locales ────────────────────────────────────────────────────────────
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = Path.home() / "Desktop" / "Beta Packing - Generados"
LOGS_DIR = APP_DIR / "logs"
APP_NAME = "Beta Packing List"


def ensure_runtime_dirs() -> None:
    """Crea los directorios necesarios para la app si no existen."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
