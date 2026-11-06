"""Запуск: `python -m app.web`. Сервер слушает только этот компьютер.

Адрес 127.0.0.1, а не 0.0.0.0: данные -- коммерческая информация фирмы, и
программа однопользовательская. С другого компьютера сети сюда не достучаться.
Порт занят или база не открывается -- понятная фраза в окне программы вместо
трассировки (окно видит пользователь, а не только разработчик).
"""
import logging
import sqlite3
import sys

from werkzeug.serving import ThreadedWSGIServer

from app.config import DB_PATH
from app.web import create_app

HOST = "127.0.0.1"
PORT = 8765


class BindError(Exception):
    """Не удалось занять адрес и порт (чаще всего порт уже занят)."""


class _LocalServer(ThreadedWSGIServer):
    """Поток на запрос; соединения с базой -- свои у каждого запроса.

    allow_reuse_address у Werkzeug включён, а на Windows SO_REUSEADDR позволяет
    второму серверу занять уже занятый порт: двойной щелчок по ярлыку поднял бы
    вторую копию, и запросы делились бы между ними без единой ошибки. Здесь
    он выключен, и занятый порт -- отказ.

    Отказ при привязке Werkzeug сам печатает по-английски и завершает процесс
    (sys.exit), минуя нас. Поэтому OSError превращается в BindError: исключение,
    не являющееся OSError, Werkzeug пропускает дальше."""
    allow_reuse_address = False

    def server_bind(self):
        try:
            super().server_bind()
        except OSError as exc:
            raise BindError(str(exc)) from exc


def build_server(app, port: int = PORT):
    return _LocalServer(HOST, port, app)


def main() -> int:
    try:
        app = create_app(DB_PATH)
    except (OSError, sqlite3.Error):
        logging.getLogger(__name__).exception("Не удалось открыть базу данных %s", DB_PATH)
        print(f"Не удалось открыть базу данных {DB_PATH}: проверьте, что файл не занят "
              "другой программой, не повреждён и что в папке можно записывать.")
        return 1
    try:
        server = build_server(app)
    except BindError:
        print(f"Не удалось запустить программу: порт {PORT} занят. Скорее всего, программа "
              f"уже запущена в другом окне -- откройте http://{HOST}:{PORT} в браузере. "
              "Если нет, закройте программу, которая занимает порт, и запустите снова.")
        return 1
    logging.getLogger("werkzeug").setLevel(logging.WARNING)     # без строки на каждый запрос
    print(f"Программа запущена: http://{HOST}:{PORT}")
    print("Не закрывайте это окно, пока работаете. Закройте его, чтобы остановить программу.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
