"""Lokasi data aplikasi yang aman untuk aplikasi Windows lokal."""
import os
from pathlib import Path


APP_NAME = "MappingKeuangan"
_LOCAL_APP_DATA = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
APP_DATA_DIR = Path(_LOCAL_APP_DATA) / APP_NAME if _LOCAL_APP_DATA else Path.home() / f".{APP_NAME}"
UPLOADS_DIR = APP_DATA_DIR / "uploads"
EXPORTS_DIR = APP_DATA_DIR / "exports"
CORRECTED_PDF_DIR = APP_DATA_DIR / "corrected-pdf"
DB_PATH = APP_DATA_DIR / "database.sqlite3"


def ensure_app_dirs() -> Path:
    """Buat folder data aplikasi dan kembalikan root-nya."""
    for path in (APP_DATA_DIR, UPLOADS_DIR, EXPORTS_DIR, CORRECTED_PDF_DIR):
        path.mkdir(parents=True, exist_ok=True)
    return APP_DATA_DIR
