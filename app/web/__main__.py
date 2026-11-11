"""Запуск: `python -m app.web`. Сервер слушает только этот компьютер.

Адрес 127.0.0.1, а не 0.0.0.0: данные -- коммерческая информация фирмы, и
программа однопользовательская. С другого компьютера сети сюда не достучаться.
Порт занят или база не открывается -- понятная фраза в окне программы вместо
трассировки (окно видит пользователь, а не только разработчик).

На пустой базе справочники наполняются сами: первый запуск на новом
компьютере не должен показывать пустоту (см. fill_references).
"""
import logging
import sqlite3
import sys
from pathlib import Path

from werkzeug.serving import ThreadedWSGIServer

from app import config, db
from app.repo import Repo
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


def fill_references(db_path) -> None:
    """Первый запуск на пустой базе: справочники наполняются сами.

    Иначе фирма на новом компьютере открывает «Справочники» и видит пустоту --
    организации, сотрудников, 23 позиции МТР, каталог услуг и три профиля
    заказчиков пришлось бы набрать заново, хотя они уже написаны. Докстринг
    скрипта наполнения путём к этому знанию не является.

    Наполнение живёт в scripts/, а не в пакете app, поэтому импорт локальный,
    а корень проекта добавляется в путь: ярлык может запустить программу из
    любой папки. Сбой наполнения программу не останавливает -- справочники
    можно завести руками, и фраза говорит, что случилось.
    """
    if str(config.BASE_DIR) not in sys.path:
        sys.path.insert(0, str(config.BASE_DIR))
    try:
        from scripts.seed import SeedError, seed_if_empty
    except ImportError:
        logging.getLogger(__name__).exception("Нет модуля наполнения справочников")
        print("Справочники не заполнены: рядом с программой нет папки scripts. "
              "Справочники можно завести вручную на экране «Справочники».")
        return
    conn = db.connect(db_path)
    try:
        if seed_if_empty(Repo(conn)):
            print("Справочники заполнены данными «Техносферы»: "
                  "это первый запуск программы на этой базе.")
    except (SeedError, OSError, sqlite3.Error) as exc:
        logging.getLogger(__name__).exception("Не удалось заполнить справочники")
        print(f"Справочники не заполнены: {exc if isinstance(exc, SeedError) else ''}"
              " Программа работает, справочники можно завести вручную "
              "на экране «Справочники».")
    finally:
        conn.close()


def save_backup(db_path: Path) -> None:
    """Копия базы перед запуском -- до того, как `create_app` доведёт схему.

    Именно до миграций: если миграция окажется неверной, копия останется
    снятой с базы ПРЕЖНЕЙ схемы, и терять будет нечего. После миграций копия
    была бы уже испорченной тем же, от чего защищает.

    Сбой копирования программу не останавливает: диск полон, папка защищена,
    файл занят -- всё это повод сказать человеку, а не отказать ему в работе.
    Молча сбой не проходит: без копии он должен знать, что копии нет.
    """
    try:
        copy = db.backup(db_path)
    except (OSError, sqlite3.Error) as exc:
        logging.getLogger(__name__).exception("Не удалось сохранить копию базы")
        print(f"Копия базы НЕ сохранена: {exc}. Программа работает, но данные "
              "сейчас без резервной копии -- проверьте, есть ли место на диске "
              f"и права на запись в папку {Path(db_path).parent}.")
        return
    if copy is not None:
        print(f"Копия базы сохранена: {copy}")


def main(db_path=None) -> int:
    """Запуск программы. `db_path` -- база, с которой работать; None -- рабочая
    база программы (`config.DB_PATH`).

    Путь читается из `config` в момент вызова, а не копируется именем при
    загрузке модуля: так подмена пути в тесте видна и здесь, и всем остальным.
    Параметр нужен тому, кто зовёт `main` не из ярлыка: тест передаёт
    временный путь и ничего не подменяет.
    """
    db_path = Path(config.DB_PATH if db_path is None else db_path)
    save_backup(db_path)
    try:
        app = create_app(db_path)
    except (OSError, sqlite3.Error):
        logging.getLogger(__name__).exception("Не удалось открыть базу данных %s", db_path)
        print(f"Не удалось открыть базу данных {db_path}: проверьте, что файл не занят "
              "другой программой, не повреждён и что в папке можно записывать.")
        return 1
    fill_references(db_path)
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
