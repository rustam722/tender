import sqlite3
from decimal import Decimal

import pytest

from app import db


@pytest.fixture
def conn(tmp_path):
    c = db.connect(tmp_path / "t.db")
    yield c
    c.close()


def test_migrate_creates_tables(conn):
    db.migrate(conn)
    names = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"organization", "employee", "asset", "experience",
            "service", "org_settings", "customer_profile",
            "tender", "lot_item"} <= names


def test_migrate_is_idempotent(conn):
    db.migrate(conn)
    db.migrate(conn)  # второй раз не падает
    assert conn.execute("PRAGMA user_version").fetchone()[0] == db.SCHEMA_VERSION


def test_fresh_database_has_docs_and_no_old_profile_columns(conn):
    # описание документа -- одна колонка docs; layouts и file_prefixes сняты
    db.migrate(conn)
    cols = {r[1] for r in conn.execute("PRAGMA table_info(customer_profile)")}
    assert "docs" in cols
    assert not {"layouts", "file_prefixes"} & cols
    assert db.SCHEMA_VERSION == 6


def test_foreign_keys_are_enforced(conn):
    db.migrate(conn)
    with pytest.raises(sqlite3.IntegrityError):
        conn.execute("INSERT INTO lot_item (tender_id, name) VALUES (999, 'x')")


def test_decimal_round_trips(conn):
    db.migrate(conn)
    conn.execute(
        "INSERT INTO experience (customer_name, amount) VALUES (?, ?)",
        ("Коряжма", Decimal("5985000.00")))
    (amount,) = conn.execute("SELECT amount FROM experience").fetchone()
    assert isinstance(amount, Decimal)
    assert amount == Decimal("5985000.00")
    assert str(amount) == "5985000.00"  # копейки не теряются


def test_decimal_keeps_precision_float_would_lose(conn):
    db.migrate(conn)
    conn.execute(
        "INSERT INTO employee (full_name, hourly_rate) VALUES (?, ?)",
        ("Иванов И. И.", Decimal("270.10")))
    (rate,) = conn.execute("SELECT hourly_rate FROM employee").fetchone()
    assert rate == Decimal("270.10")
    assert not isinstance(rate, float)


def test_decimal_with_many_digits_is_not_rounded(conn):
    db.migrate(conn)
    big = Decimal("12345678901234567890.12")  # float держит ~15 значащих цифр
    conn.execute("INSERT INTO experience (amount) VALUES (?)", (big,))
    (amount,) = conn.execute("SELECT amount FROM experience").fetchone()
    assert amount == big


def test_null_decimal_stays_none(conn):
    db.migrate(conn)
    conn.execute("INSERT INTO organization (inn) VALUES ('5904188041')")
    (vat,) = conn.execute("SELECT vat_rate FROM organization").fetchone()
    assert vat is None


# ---- резервная копия базы -----------------------------------------------------------

# Справочники фирмы -- кадры, оборудование, МТР, опыт, ставки, профили -- и все
# тендеры лежат в одном файле. Копии не было никакой: сломался файл, и работы нет.

def test_a_backup_is_a_readable_database_with_the_same_rows(tmp_path):
    path = tmp_path / "рабочая.db"
    conn = db.connect(path)
    try:
        db.migrate(conn)
        conn.execute("INSERT INTO organization (inn, short_name) VALUES (?, ?)",
                     ("5904188041", "ООО «Техносфера»"))
        conn.commit()
    finally:
        conn.close()

    copy = db.backup(path)
    assert copy.is_file() and copy != path
    back = db.connect(copy)
    try:
        (name,) = back.execute("SELECT short_name FROM organization").fetchone()
        assert name == "ООО «Техносфера»"
    finally:
        back.close()


def test_a_backup_can_be_taken_while_the_program_holds_the_database_open(tmp_path):
    """Копия снимается, пока программа держит базу открытой, и содержит всё
    записанное: занятый файл -- не повод отказать.

    Имя теста утверждает ровно это и ни слова больше. Прежнее имя было
    «снимается с открытой базы, а shutil.copy отдал бы файл в середине записи»,
    и оно лгало: мутант, заменивший `Connection.backup` на `shutil.copy`, этот
    тест не уронил. С нынешним режимом журнала (не WAL) разницы и нет --
    почему способ всё равно выбран другой, сказано в `db.backup`."""
    path = tmp_path / "рабочая.db"
    conn = db.connect(path)
    try:
        db.migrate(conn)
        conn.execute("INSERT INTO organization (inn) VALUES ('5904188041')")
        conn.commit()
        copy = db.backup(path)
    finally:
        conn.close()
    back = db.connect(copy)
    try:
        assert back.execute("SELECT COUNT(*) FROM organization").fetchone()[0] == 1
    finally:
        back.close()


def test_the_oldest_backups_are_dropped_and_the_newest_are_kept(tmp_path):
    path = tmp_path / "рабочая.db"
    conn = db.connect(path)
    try:
        db.migrate(conn)
    finally:
        conn.close()
    made = [db.backup(path, keep=3) for _ in range(5)]
    folder = path.parent / db.BACKUPS_DIR_NAME
    left = sorted(p.name for p in folder.glob("*.db"))
    assert len(left) == 3, left
    assert left == sorted(p.name for p in made[-3:])      # остались последние три


def test_two_backups_in_the_same_second_do_not_overwrite_each_other(tmp_path):
    path = tmp_path / "рабочая.db"
    conn = db.connect(path)
    try:
        db.migrate(conn)
    finally:
        conn.close()
    first, second = db.backup(path), db.backup(path)
    assert first != second
    assert first.is_file() and second.is_file()


def test_the_name_of_a_backup_says_when_it_was_made(tmp_path):
    path = tmp_path / "рабочая.db"
    conn = db.connect(path)
    try:
        db.migrate(conn)
    finally:
        conn.close()
    from datetime import date
    name = db.backup(path).name
    assert name.startswith("tender-") and name.endswith(".db")
    assert date.today().isoformat() in name        # человек ищет копию по дате


def test_there_is_nothing_to_copy_on_the_first_launch(tmp_path):
    """Базы ещё нет -- копировать нечего, и папка копий не создаётся: пустая
    папка «копии» на новом компьютере заставила бы человека искать, что
    пропало."""
    path = tmp_path / "ещё-нет.db"
    assert db.backup(path) is None
    assert not (path.parent / db.BACKUPS_DIR_NAME).exists()


def test_the_backups_live_beside_the_database_they_copy(tmp_path):
    """Копии лежат рядом со своей базой, а не в одном месте на всю программу:
    иначе тест на временной базе писал бы в папку фирмы."""
    path = tmp_path / "вложено" / "рабочая.db"
    conn = db.connect(path)
    try:
        db.migrate(conn)
    finally:
        conn.close()
    copy = db.backup(path)
    assert copy.parent == path.parent / db.BACKUPS_DIR_NAME
