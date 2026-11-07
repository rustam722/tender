"""Пути программы. Больше здесь ничего нет.

Каждая папка названа одним именем на всю программу и определена ровно здесь.
Остальные модули читают их через модуль (`from app import config`,
`config.CELLMAPS_DIR` в момент вызова), а не копируют имя к себе: тогда
подмена в тесте (`monkeypatch.setattr(config, ...)`) видна всем разом.
"""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "tender.db"
TEMPLATES_DOCS_DIR = BASE_DIR / "templates_docs"      # шаблоны документов пакета
CELLMAPS_DIR = BASE_DIR / "cellmaps"                  # карты ячеек книг Excel
LAYOUTS_DIR = BASE_DIR / "layouts"                    # раскладки колонок таблиц
