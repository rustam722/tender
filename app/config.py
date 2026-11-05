"""Пути программы. Больше здесь ничего нет."""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
DB_PATH = DATA_DIR / "tender.db"
