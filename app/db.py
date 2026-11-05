"""Соединение с SQLite, типы Decimal и миграции схемы."""
import sqlite3
from decimal import Decimal
from pathlib import Path

# Встроенного адаптера Decimal в sqlite3 нет. Храним строкой, чтобы не терять
# копейки. Конвертер выбирается по первому слову объявленного типа колонки
# (DECIMAL), а аффинность SQLite определяется по всему объявлению. Голый DECIMAL
# имеет аффинность NUMERIC: строку '5985000.00' база молча превратила бы в
# целое 5985000, а '270.10' -- в REAL. Поэтому колонки объявлены «DECIMAL TEXT»:
# конвертер срабатывает, аффинность TEXT, строка хранится как есть.
# Внимание: SUM() и другие выражения теряют тип колонки и вернут не Decimal --
# суммировать нужно в Python.
sqlite3.register_adapter(Decimal, str)
sqlite3.register_converter("DECIMAL", lambda raw: Decimal(raw.decode("ascii")))

# Миграция N переводит схему с версии N-1 на N. Версия живёт в PRAGMA user_version.
_MIGRATIONS = [
    """
    CREATE TABLE IF NOT EXISTS organization (
        id INTEGER PRIMARY KEY,
        role TEXT NOT NULL DEFAULT 'participant'
            CHECK (role IN ('participant', 'customer')),
        inn TEXT NOT NULL DEFAULT '',
        kpp TEXT NOT NULL DEFAULT '',
        ogrn TEXT NOT NULL DEFAULT '',
        short_name TEXT NOT NULL DEFAULT '',
        full_name TEXT NOT NULL DEFAULT '',
        legal_address TEXT NOT NULL DEFAULT '',
        postal_address TEXT NOT NULL DEFAULT '',
        phones TEXT NOT NULL DEFAULT '',
        email TEXT NOT NULL DEFAULT '',
        website TEXT NOT NULL DEFAULT '',
        bank_name TEXT NOT NULL DEFAULT '',
        bank_bik TEXT NOT NULL DEFAULT '',
        bank_account TEXT NOT NULL DEFAULT '',
        bank_corr_account TEXT NOT NULL DEFAULT '',
        head_name TEXT NOT NULL DEFAULT '',
        head_position TEXT NOT NULL DEFAULT '',
        head_basis TEXT NOT NULL DEFAULT '',
        okpo TEXT NOT NULL DEFAULT '',
        oktmo TEXT NOT NULL DEFAULT '',
        okved TEXT NOT NULL DEFAULT '',
        vat_rate DECIMAL TEXT,
        msp_status TEXT NOT NULL DEFAULT ''
    );
    CREATE INDEX IF NOT EXISTS idx_organization_inn ON organization (inn);

    CREATE TABLE IF NOT EXISTS employee (
        id INTEGER PRIMARY KEY,
        full_name TEXT NOT NULL DEFAULT '',
        position TEXT NOT NULL DEFAULT '',
        education TEXT NOT NULL DEFAULT '',
        graduation_year INTEGER,
        specialty TEXT NOT NULL DEFAULT '',
        qualification TEXT NOT NULL DEFAULT '',
        experience_years INTEGER,
        hourly_rate DECIMAL TEXT NOT NULL DEFAULT '0'
    );

    CREATE TABLE IF NOT EXISTS asset (
        id INTEGER PRIMARY KEY,
        name TEXT NOT NULL DEFAULT '',
        kind TEXT NOT NULL DEFAULT '',
        maker_model TEXT NOT NULL DEFAULT '',
        quantity INTEGER NOT NULL DEFAULT 1,
        year INTEGER,
        location TEXT NOT NULL DEFAULT '',
        ownership_basis TEXT NOT NULL DEFAULT '',
        ownership_doc TEXT NOT NULL DEFAULT '',
        purpose TEXT NOT NULL DEFAULT '',
        condition TEXT NOT NULL DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS service (
        id INTEGER PRIMARY KEY,
        code TEXT NOT NULL DEFAULT '',
        name TEXT NOT NULL DEFAULT '',
        stage_wording TEXT NOT NULL DEFAULT ''
    );

    -- нормы трудозатрат: часов должности на услугу
    CREATE TABLE IF NOT EXISTS service_norm (
        service_id INTEGER NOT NULL REFERENCES service (id) ON DELETE CASCADE,
        position TEXT NOT NULL,
        hours DECIMAL TEXT NOT NULL DEFAULT '0',
        PRIMARY KEY (service_id, position)
    );

    CREATE TABLE IF NOT EXISTS experience (
        id INTEGER PRIMARY KEY,
        customer_name TEXT NOT NULL DEFAULT '',
        contract_number TEXT NOT NULL DEFAULT '',
        contract_date TEXT NOT NULL DEFAULT '',
        subject TEXT NOT NULL DEFAULT '',
        amount DECIMAL TEXT NOT NULL DEFAULT '0',
        term TEXT NOT NULL DEFAULT '',
        status TEXT NOT NULL DEFAULT '',
        service_id INTEGER REFERENCES service (id) ON DELETE SET NULL,
        contract_scan TEXT NOT NULL DEFAULT '',
        act_scan TEXT NOT NULL DEFAULT ''
    );

    -- одна строка настроек; значения по умолчанию задаёт репозиторий
    CREATE TABLE IF NOT EXISTS org_settings (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        insurance_rate DECIMAL TEXT NOT NULL,
        injury_rate DECIMAL TEXT NOT NULL,
        overhead_rate DECIMAL TEXT NOT NULL,
        profit_rate DECIMAL TEXT NOT NULL,
        vat_rate DECIMAL TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS customer_profile (
        id INTEGER PRIMARY KEY,
        customer_id INTEGER REFERENCES organization (id) ON DELETE CASCADE,
        documents TEXT NOT NULL DEFAULT '[]',      -- JSON: состав пакета
        file_prefixes TEXT NOT NULL DEFAULT '{}',  -- JSON: документ -> префикс
        declaration_text TEXT NOT NULL DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS tender (
        id INTEGER PRIMARY KEY,
        customer_id INTEGER REFERENCES organization (id),
        participant_id INTEGER REFERENCES organization (id),
        purchase_number TEXT NOT NULL DEFAULT '',
        purchase_date TEXT NOT NULL DEFAULT '',
        method TEXT NOT NULL DEFAULT '',
        subject TEXT NOT NULL DEFAULT '',
        object_name TEXT NOT NULL DEFAULT '',
        object_address TEXT NOT NULL DEFAULT '',
        price_wo_vat DECIMAL TEXT NOT NULL DEFAULT '0',
        vat_rate DECIMAL TEXT,
        deadline TEXT NOT NULL DEFAULT '',
        payment_terms TEXT NOT NULL DEFAULT '',
        doc_date TEXT NOT NULL DEFAULT '',
        platform TEXT NOT NULL DEFAULT '',
        output_dir TEXT NOT NULL DEFAULT '',
        deal_answers TEXT NOT NULL DEFAULT '{}'    -- JSON: условие -> ответ
    );

    CREATE TABLE IF NOT EXISTS lot_item (
        id INTEGER PRIMARY KEY,
        tender_id INTEGER NOT NULL REFERENCES tender (id) ON DELETE CASCADE,
        position INTEGER NOT NULL DEFAULT 0,
        name TEXT NOT NULL DEFAULT '',
        price DECIMAL TEXT NOT NULL DEFAULT '0'
    );
    CREATE INDEX IF NOT EXISTS idx_lot_item_tender ON lot_item (tender_id, position, id);
    """,
    # 2: трудозатраты тендера. Хранятся, а не выводятся из норм каталога:
    # численность в КП -- плановая величина под тендер, а не состав штата.
    """
    CREATE TABLE IF NOT EXISTS tender_labor (
        id INTEGER PRIMARY KEY,
        tender_id INTEGER NOT NULL REFERENCES tender (id) ON DELETE CASCADE,
        position INTEGER NOT NULL DEFAULT 0,
        post TEXT NOT NULL DEFAULT '',
        headcount INTEGER NOT NULL DEFAULT 1,
        hours DECIMAL TEXT NOT NULL DEFAULT '0',
        rate DECIMAL TEXT NOT NULL DEFAULT '0'
    );
    CREATE INDEX IF NOT EXISTS idx_tender_labor_tender
        ON tender_labor (tender_id, position, id);
    """,
    # 3: раскладки в профиле заказчика, отбор опыта при тендере, материалы и
    # субподряд. ALTER TABLE не имеет IF NOT EXISTS, поэтому скрипт -- одна
    # транзакция вместе с номером версии: обрыв не оставит половину колонок,
    # на которой следующий запуск упал бы с «duplicate column».
    """
    BEGIN;
    ALTER TABLE customer_profile
        ADD COLUMN layouts TEXT NOT NULL DEFAULT '{}';   -- JSON: документ -> файл
    ALTER TABLE tender
        ADD COLUMN experience_ids TEXT NOT NULL DEFAULT '[]';  -- JSON: id опыта
    CREATE TABLE IF NOT EXISTS tender_cost_item (
        id INTEGER PRIMARY KEY,
        tender_id INTEGER NOT NULL REFERENCES tender (id) ON DELETE CASCADE,
        kind TEXT NOT NULL DEFAULT 'materials'
            CHECK (kind IN ('materials', 'subcontract')),
        position INTEGER NOT NULL DEFAULT 0,
        name TEXT NOT NULL DEFAULT '',
        unit TEXT NOT NULL DEFAULT '',
        quantity DECIMAL TEXT NOT NULL DEFAULT '1',
        price DECIMAL TEXT NOT NULL DEFAULT '0',
        note TEXT NOT NULL DEFAULT ''
    );
    CREATE INDEX IF NOT EXISTS idx_tender_cost_item_tender
        ON tender_cost_item (tender_id, kind, position, id);
    PRAGMA user_version = 3;  -- номер этой миграции: при копировании сменить
    COMMIT;
    """,
    # 4: до какого знака округляется итог без НДС (0 -- целые рубли, как
    # ROUND(..., +D88) в книге КП.xlsx). Уже сохранённые настройки получают 0.
    """
    BEGIN;
    ALTER TABLE org_settings
        ADD COLUMN total_rounding_digits INTEGER NOT NULL DEFAULT 0;
    PRAGMA user_version = 4;  -- номер этой миграции: при копировании сменить
    COMMIT;
    """,
    # 5: описание документа у заказчика -- один JSON-объект на документ вместо
    # двух словарей с одними ключами (layouts, file_prefixes). Данные переносятся:
    # ключ документа берётся из обоих старых словарей, недостающее поле пусто;
    # шаблон и карта ячеек у перенесённого описания тоже пусты (умолчания).
    # Всё в одной транзакции вместе с номером версии, как миграции 3 и 4.
    # DROP COLUMN требует SQLite 3.35+.
    """
    BEGIN;
    ALTER TABLE customer_profile
        ADD COLUMN docs TEXT NOT NULL DEFAULT '{}';  -- JSON: документ -> описание
    UPDATE customer_profile SET docs = COALESCE((
        SELECT json_group_object(k.name, json_object(
                   'template', '',
                   'cellmap', '',
                   'layout', COALESCE(l.value, ''),
                   'prefix', COALESCE(p.value, '')))
        FROM (SELECT key AS name FROM json_each(customer_profile.layouts)
              UNION
              SELECT key FROM json_each(customer_profile.file_prefixes)) AS k
        LEFT JOIN json_each(customer_profile.layouts) AS l ON l.key = k.name
        LEFT JOIN json_each(customer_profile.file_prefixes) AS p ON p.key = k.name
    ), '{}');
    ALTER TABLE customer_profile DROP COLUMN layouts;
    ALTER TABLE customer_profile DROP COLUMN file_prefixes;
    PRAGMA user_version = 5;  -- номер этой миграции: при копировании сменить
    COMMIT;
    """,
]

SCHEMA_VERSION = len(_MIGRATIONS)


def connect(path: Path) -> sqlite3.Connection:
    """Открывает базу: типы Decimal, внешние ключи, строки как sqlite3.Row."""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path, detect_types=sqlite3.PARSE_DECLTYPES)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def migrate(conn: sqlite3.Connection) -> None:
    """Доводит схему до последней версии. Повторный вызов ничего не меняет."""
    version = conn.execute("PRAGMA user_version").fetchone()[0]
    for number in range(version + 1, SCHEMA_VERSION + 1):
        # executescript закрывает открытую транзакцию и сам не откатывается,
        # поэтому каждая миграция идемпотентна (IF NOT EXISTS).
        conn.executescript(_MIGRATIONS[number - 1])
        conn.execute(f"PRAGMA user_version = {number}")
