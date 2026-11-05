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
