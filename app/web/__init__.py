"""Локальный веб-сервер программы: справочники, тендеры, расчёт, сборка пакета."""
from pathlib import Path

from flask import Flask

from app import db
from app.web import routes


def create_app(db_path: Path) -> Flask:
    """Приложение для базы по пути db_path. База создаётся и доводится до
    последней схемы здесь, один раз; запросы открывают собственные соединения
    (см. routes.get_repo). Тесты создают приложение на временном файле."""
    app = Flask(__name__)
    app.config["DB_PATH"] = Path(db_path)
    conn = db.connect(app.config["DB_PATH"])
    try:
        db.migrate(conn)
    finally:
        conn.close()
    routes.register(app)
    return app
