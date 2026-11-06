"""Веб-ручки: справочники, тендеры, расчёт, сборка пакета.

Запросы идут через `app.test_client()` к настоящей базе во временной папке.
Главное, что здесь проверяется, -- не «ручка отвечает 200», а то, что пользователь
видит при отказе: русский текст без трассировок и без английских исключений,
и что сборка не пишет вне выбранной папки.
"""
import json
import logging
import re
import sqlite3
import threading
from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal
from pathlib import Path

import docx
import pytest

from app import db
from app.config import BASE_DIR
from app.inn.fns import CaptchaRequired, InnNotFound, SourceUnavailable
from app.models import ROLE_CUSTOMER, Organization
from app.render import package as package_module
from app.repo import Repo
from app.web import create_app
from app.web import routes as routes_module
from app.web.errors import ApiError, translate_integrity
from app.web.schema import ORGANIZATIONS

CYRILLIC = re.compile("[А-Яа-яЁё]")
TRACEBACK_MARKERS = ("Traceback", 'File "', "line ", "Error:", "Exception")
ENGLISH_SQLITE = ("constraint", "FOREIGN KEY", "UNIQUE", "NOT NULL", "CHECK")


# ---- помощники ----------------------------------------------------------------

@pytest.fixture
def app(tmp_path):
    return create_app(tmp_path / "web.db")


@pytest.fixture
def client(app):
    return app.test_client()


@pytest.fixture
def dirs(tmp_path, monkeypatch):
    """Папки шаблонов и карт -- в папке теста, а не в проекте."""
    templates = tmp_path / "templates_docs"
    cellmaps = tmp_path / "cellmaps"
    templates.mkdir()
    cellmaps.mkdir()
    monkeypatch.setattr(package_module, "TEMPLATES_DOCS_DIR", templates)
    monkeypatch.setattr(package_module, "CELLMAPS_DIR", cellmaps)
    return templates, cellmaps


def make_template(path: Path, *lines: str) -> Path:
    doc = docx.Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)
    return path


def text_of(path: Path) -> list[str]:
    return [p.text for p in docx.Document(path).paragraphs]


def ok(response, status=200):
    assert response.status_code == status, response.get_data(as_text=True)
    return response.get_json()


def refused(response, status=400):
    """Отказ: статус, тело {"error": ...} с русским текстом и ничего лишнего
    из английских исключений."""
    assert response.status_code == status, response.get_data(as_text=True)
    body = response.get_json()
    assert isinstance(body["error"], str)
    assert CYRILLIC.search(body["error"]), body["error"]
    raw = response.get_data(as_text=True)
    for marker in TRACEBACK_MARKERS + ENGLISH_SQLITE:
        assert marker not in raw, f"{marker!r} в ответе: {raw}"
    return body


def create(client, resource, /, **fields):
    return ok(client.post(f"/api/{resource}", json=fields), 201)


def make_tender(client, **fields):
    me = create(client, "organizations", role="participant", short_name="ООО «Техносфера»",
                inn="5904188041")
    customer = create(client, "organizations", role="customer", short_name="МУП «Полигон»",
                      inn="2904016130")
    return create(client, "tenders", customer_id=customer["id"], participant_id=me["id"],
                  subject="Разработка КЭР", **fields)


def add_labor(client, tender_id, post="Инженер", headcount=1, hours="100", rate="1000"):
    return create(client, f"tenders/{tender_id}/labor", post=post, headcount=headcount,
                  hours=hours, rate=rate)


# ---- брифовые тесты -----------------------------------------------------------

def test_crud_round_trip_for_organizations(client):
    created = create(client, "organizations", role="customer",
                     short_name="МУП «Полигон»", inn="2904016130", vat_rate="0.05")
    assert created["id"] > 0
    assert created["vat_rate"] == "0.05"           # деньги и ставки -- строкой, не float

    one = ok(client.get(f"/api/organizations/{created['id']}"))
    assert one["short_name"] == "МУП «Полигон»"

    assert [o["id"] for o in ok(client.get("/api/organizations"))] == [created["id"]]
    assert ok(client.get("/api/organizations?role=participant")) == []

    updated = ok(client.post("/api/organizations",
                             json={"id": created["id"], "short_name": "МУП «Полигон-2»"}))
    assert updated["short_name"] == "МУП «Полигон-2»"
    assert updated["inn"] == "2904016130"          # не присланное поле не стирается

    assert ok(client.delete(f"/api/organizations/{created['id']}")) == {"deleted": created["id"]}
    refused(client.get(f"/api/organizations/{created['id']}"), 404)
    assert ok(client.get("/api/organizations")) == []


def test_inn_endpoint_returns_draft_and_warnings(client, monkeypatch):
    seen = {}

    def fake(inn, *, rusprofile_url=None, session=None):
        seen["args"] = (inn, rusprofile_url)
        return (Organization(role=ROLE_CUSTOMER, inn=inn, short_name="МУП «Полигон»"),
                ["Банковские реквизиты не найдены."])

    monkeypatch.setattr(routes_module, "build_draft", fake)
    body = ok(client.get("/api/inn/2904016130?rusprofile_url=https://www.rusprofile.ru/id/1"))
    assert body["organization"]["short_name"] == "МУП «Полигон»"
    assert body["organization"]["id"] is None       # черновик не сохраняется
    assert body["warnings"] == ["Банковские реквизиты не найдены."]
    assert seen["args"] == ("2904016130", "https://www.rusprofile.ru/id/1")
    assert ok(client.get("/api/organizations")) == []


def test_inn_not_found_gives_400_with_russian_message(client, monkeypatch):
    def fake(inn, **kw):
        raise InnNotFound("Организация с таким ИНН не найдена")
    monkeypatch.setattr(routes_module, "build_draft", fake)
    body = refused(client.get("/api/inn/2904016130"))
    assert body["error"] == "Организация с таким ИНН не найдена"


def test_source_unavailable_gives_502(client, monkeypatch):
    def down(inn, **kw):
        raise SourceUnavailable("Сервис ФНС не отвечает. Попробуйте позже.")

    def captcha(inn, **kw):
        raise CaptchaRequired("Сервис ФНС требует ввести капчу.")

    monkeypatch.setattr(routes_module, "build_draft", down)
    assert refused(client.get("/api/inn/2904016130"), 502)["error"].startswith("Сервис ФНС")
    monkeypatch.setattr(routes_module, "build_draft", captcha)
    assert "капчу" in refused(client.get("/api/inn/2904016130"), 502)["error"]


def test_solve_below_target_gives_400_not_500(client):
    tender = make_tender(client)
    add_labor(client, tender["id"])
    create(client, f"tenders/{tender['id']}/costs", kind="materials", name="Бумага",
           unit="шт.", quantity="1", price="500000")
    body = refused(client.post(f"/api/tenders/{tender['id']}/solve",
                               json={"target": "100000"}))
    assert "не больше неизбежных расходов" in body["error"]


def test_build_returns_written_and_skipped(client, tmp_path, dirs):
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "Предмет: {{ tender.subject }}")
    tender = make_tender(client)
    out = tmp_path / "out" / "tender"
    body = ok(client.post(f"/api/tenders/{tender['id']}/build", json={
        "documents": ["Декларация", "Нет такого документа"], "output_dir": str(out)}))
    assert [w["name"] for w in body["written"]] == ["Декларация.docx"]
    assert Path(body["written"][0]["path"]) == out / "Декларация.docx"
    assert text_of(out / "Декларация.docx") == ["Предмет: Разработка КЭР"]
    assert [s["document"] for s in body["skipped"]] == ["Нет такого документа"]
    assert "не найден" in body["skipped"][0]["reason"]


def test_index_serves_html(client):
    response = client.get("/")
    assert response.status_code == 200
    assert response.mimetype == "text/html"
    assert "<html" in response.get_data(as_text=True)


def test_no_traceback_in_any_error_body(client, monkeypatch, caplog):
    """Каждая дорога отказа: тело -- русская фраза, ни трассировки, ни имён
    исключений, ни английских слов SQLite."""
    tender = make_tender(client)
    tid = tender["id"]
    responses = [
        client.get("/api/nothing"),                                           # 404
        client.get("/api/organizations/999"),                                 # 404 записи
        client.put("/api/organizations", json={}),                            # 405
        client.post("/api/organizations", data="не json",
                    content_type="application/json"),                         # битый JSON
        client.post("/api/organizations", data="x", content_type="text/plain"),  # 415
        client.post("/api/organizations", json={"nope": 1}),                  # неизвестное поле
        client.post("/api/tenders", json={"customer_id": 999}),               # внешний ключ
        client.delete(f"/api/organizations/{tender['customer_id']}"),         # внешний ключ
        client.post(f"/api/tenders/{tid}/solve", json={}),                    # нет цели
        client.post(f"/api/tenders/{tid}/solve", json={"target": "1"}),  # нет работ
        client.post(f"/api/tenders/{tid}/build", json={}),                    # нет документов
        client.get("/api/inn/abc"),                                           # плохой ИНН
        client.get("/api/organizations", headers={"Host": "evil.example"}),   # чужой Host
    ]
    for response in responses:
        assert response.status_code >= 400
        refused(response, response.status_code)

    # непредвиденное: деталь уходит в журнал, а не в ответ
    def boom(self):
        raise RuntimeError("секретная деталь: C:\\Users\\x\\password.txt")

    monkeypatch.setattr(Repo, "list_employees", boom)
    with caplog.at_level(logging.ERROR):
        response = client.get("/api/employees")
    body = refused(response, 500)
    assert "секретная" not in response.get_data(as_text=True)
    assert "RuntimeError" not in response.get_data(as_text=True)
    assert "секретная деталь" in caplog.text           # а в журнале -- есть
    assert set(body) == {"error"}


# ---- справочники --------------------------------------------------------------

def test_every_directory_round_trips(client):
    emp = create(client, "employees", full_name="Рабичко Е.М.", position="Директор",
                 graduation_year=1999, hourly_rate="1 250,50")
    assert emp["hourly_rate"] == "1250.50"          # запятая и пробелы разобраны
    assert emp["graduation_year"] == 1999

    asset = create(client, "assets", name="Компьютер", kind="equipment", quantity=9, year=2021)
    assert asset["quantity"] == 9

    svc = create(client, "services", code="КЭР", name="Комплексное экологическое разрешение",
                 norms={"Инженер-эколог": "120", "Руководитель": "30,5"})
    assert svc["norms"] == {"Инженер-эколог": "120", "Руководитель": "30.5"}

    exp = create(client, "experience", customer_name="МУП «Полигон»", amount="5985000",
                 service_id=svc["id"])
    assert ok(client.get(f"/api/experience?service_id={svc['id']}"))[0]["id"] == exp["id"]
    assert ok(client.get(f"/api/experience?service_id={svc['id'] + 1}")) == []

    ok(client.delete(f"/api/services/{svc['id']}"))
    assert ok(client.get(f"/api/experience/{exp['id']}"))["service_id"] is None  # SET NULL


def test_update_of_missing_record_is_404_and_does_not_create_one(client):
    refused(client.post("/api/employees", json={"id": 77, "full_name": "Призрак"}), 404)
    assert ok(client.get("/api/employees")) == []


def test_delete_of_missing_record_is_404(client):
    refused(client.delete("/api/assets/5"), 404)


def test_settings_default_update_and_ranges(client):
    first = ok(client.get("/api/settings"))
    assert first["insurance_rate"] == "0.30" and first["vat_rate"] == "0.05"
    assert first["total_rounding_digits"] == 0

    new = ok(client.post("/api/settings", json={"vat_rate": "0,20", "total_rounding_digits": 2}))
    assert new["vat_rate"] == "0.20" and new["total_rounding_digits"] == 2
    assert new["profit_rate"] == "0.10"             # остальное осталось
    assert ok(client.get("/api/settings")) == new

    body = refused(client.post("/api/settings", json={"vat_rate": "5"}))
    assert body["field"] == "vat_rate" and "от 0 до 1" in body["error"]
    refused(client.post("/api/settings", json={"total_rounding_digits": 50}))
    assert ok(client.get("/api/settings")) == new   # отказ ничего не изменил


def test_profile_nested_docs_round_trip_and_errors_point_at_the_field(client):
    customer = create(client, "organizations", role="customer", short_name="Заказчик")
    profile = create(client, "profiles", customer_id=customer["id"],
                     documents=["Справка о МТР"],
                     docs={"Справка о МТР": {"layout": "assets_rnpk.json", "prefix": "8.1_"}},
                     declaration_text="Текст")
    assert profile["docs"] == {"Справка о МТР": {
        "template": "", "cellmap": "", "layout": "assets_rnpk.json", "prefix": "8.1_"}}

    body = refused(client.post("/api/profiles", json={
        "id": profile["id"], "docs": {"Справка о МТР": {"layout": 5}}}))
    assert body["path"] == ["docs", "Справка о МТР", "layout"]
    assert body["field"] == "docs"
    assert "Раскладка" in body["error"] and "Справка о МТР" in body["error"]

    body = refused(client.post("/api/profiles", json={
        "id": profile["id"], "docs": {"Анкета": {"layuot": "x"}}}))
    assert body["path"] == ["docs", "Анкета", "layuot"]

    body = refused(client.post("/api/profiles", json={
        "id": profile["id"], "docs": {"ТКП": {"template": "форма.xlsx"}}}))   # книга без карты
    assert body["path"] == ["docs", "ТКП"] and "карта ячеек" in body["error"]

    assert ok(client.get(f"/api/profiles/{profile['id']}"))["docs"] == profile["docs"]
    assert ok(client.get(f"/api/profiles?customer_id={customer['id']}"))[0]["id"] == profile["id"]


def test_second_profile_for_one_customer_is_refused(client):
    customer = create(client, "organizations", role="customer", short_name="Заказчик")
    create(client, "profiles", customer_id=customer["id"])
    body = refused(client.post("/api/profiles", json={"customer_id": customer["id"]}))
    assert "уже есть профиль" in body["error"]


def test_tender_children_round_trip(client):
    tender = make_tender(client, purchase_date="2025-10-24", price_wo_vat="5700000",
                         deal_answers={"крупная сделка": "нет"}, experience_ids=[3, 1])
    tid = tender["id"]
    assert tender["deal_answers"] == {"крупная сделка": "нет"}
    assert tender["experience_ids"] == [3, 1]

    item = create(client, f"tenders/{tid}/items", position=1, name="Разработка КЭР",
                  price="5 700 000,00")
    assert item["tender_id"] == tid and item["price"] == "5700000.00"
    ok(client.post(f"/api/tenders/{tid}/items", json={"id": item["id"], "name": "КЭР"}))
    assert [i["name"] for i in ok(client.get(f"/api/tenders/{tid}/items"))["rows"]] == ["КЭР"]

    labor = add_labor(client, tid)
    assert labor["hours"] == "100"
    cost = create(client, f"tenders/{tid}/costs", kind="subcontract", name="Замеры",
                  unit="усл.ед.", quantity="1", price="20000")
    create(client, f"tenders/{tid}/costs", kind="materials", name="Бумага", price="10")
    assert len(ok(client.get(f"/api/tenders/{tid}/costs"))["rows"]) == 2
    assert [c["id"] for c in ok(client.get(f"/api/tenders/{tid}/costs?kind=subcontract"))["rows"]] \
        == [cost["id"]]
    refused(client.get(f"/api/tenders/{tid}/costs?kind=drova"))

    ok(client.delete(f"/api/tenders/{tid}/items/{item['id']}"))
    ok(client.delete(f"/api/tenders/{tid}/labor/{labor['id']}"))
    assert ok(client.get(f"/api/tenders/{tid}/items"))["rows"] == []
    assert ok(client.get(f"/api/tenders/{tid}/labor"))["rows"] == []

    # удаление тендера уносит дочерние строки; на чужой тендер ручки отвечают 404
    ok(client.delete(f"/api/tenders/{tid}"))
    refused(client.get(f"/api/tenders/{tid}/items"), 404)
    refused(client.post(f"/api/tenders/{tid}/labor", json={"post": "x"}), 404)


def test_child_row_of_another_tender_is_not_reachable(client):
    a, b = make_tender(client), make_tender(client)
    row = add_labor(client, a["id"])
    refused(client.delete(f"/api/tenders/{b['id']}/labor/{row['id']}"), 404)
    refused(client.post(f"/api/tenders/{b['id']}/labor", json={"id": row["id"], "hours": "1"}), 404)
    assert ok(client.get(f"/api/tenders/{a['id']}/labor"))["rows"][0]["hours"] == "100"
    body = refused(client.post(f"/api/tenders/{a['id']}/labor",
                               json={"tender_id": b["id"], "post": "х"}))
    assert body["field"] == "tender_id"


# ---- проверка ввода -----------------------------------------------------------

def test_money_is_exact_not_float(client):
    # JSON-число читается как Decimal, минуя float: 0.1 + 0.2 здесь не появится
    raw = '{"price_wo_vat": 5985000.10, "subject": "x"}'
    response = client.post("/api/tenders", data=raw, content_type="application/json")
    assert ok(response, 201)["price_wo_vat"] == "5985000.10"


@pytest.mark.parametrize("value", ["abc", "", "NaN", "Infinity", "-1", "1e400", "99999999999999",
                                   True, [], {}])
def test_bad_decimal_is_refused_in_russian_with_field(client, value):
    body = refused(client.post("/api/employees", json={"hourly_rate": value}))
    assert body["field"] == "hourly_rate" and body["path"] == ["hourly_rate"]
    assert "Ставка" in body["error"]


@pytest.mark.parametrize("raw", ['{"hourly_rate": NaN}', '{"hourly_rate": Infinity}',
                                 '{"hourly_rate": 1e999999999}', "[1]", '"текст"', "{"])
def test_bad_json_values_are_refused(client, raw):
    refused(client.post("/api/employees", data=raw, content_type="application/json"))


@pytest.mark.parametrize("field,value", [
    ("graduation_year", "год"), ("graduation_year", 1999.5), ("graduation_year", -1),
    ("graduation_year", True), ("graduation_year", 2 ** 70), ("full_name", 5),
    ("full_name", ["a"]), ("position", None)])
def test_bad_scalar_is_refused(client, field, value):
    body = refused(client.post("/api/employees", json={field: value}))
    assert body["field"] == field


def test_blank_optional_number_means_not_set(client):
    """Пустое поле формы приходит как "": для необязательного числа или ссылки это
    «не задано». Для обязательного числа пустая строка -- по-прежнему отказ."""
    org = create(client, "organizations", vat_rate="", short_name="x")
    assert org["vat_rate"] is None
    emp = create(client, "employees", graduation_year=" ", experience_years="")
    assert emp["graduation_year"] is None and emp["experience_years"] is None
    assert ok(client.post("/api/tenders", json={"customer_id": "", "vat_rate": ""}), 201)["customer_id"] is None
    refused(client.post("/api/employees", json={"hourly_rate": ""}))
    assert ok(client.post("/api/organizations", json={"id": org["id"], "vat_rate": "0.20"}))["vat_rate"] == "0.20"
    assert ok(client.post("/api/organizations", json={"id": org["id"], "vat_rate": ""}))["vat_rate"] is None


def test_unknown_field_is_refused_not_dropped(client):
    body = refused(client.post("/api/employees", json={"full_nme": "опечатка"}))
    assert "full_nme" in body["error"]


def test_choices_dates_and_rates_are_checked(client):
    assert refused(client.post("/api/organizations", json={"role": "boss"}))["field"] == "role"
    assert refused(client.post("/api/organizations",
                               json={"msp_status": "huge"}))["field"] == "msp_status"
    assert refused(client.post("/api/tenders", json={"purchase_date": "24.10.2025"}))["field"] \
        == "purchase_date"
    assert refused(client.post("/api/tenders", json={"doc_date": "2025-13-40"}))["field"] \
        == "doc_date"
    refused(client.post("/api/tenders", json={"vat_rate": "5"}))
    tender = make_tender(client)
    refused(client.post(f"/api/tenders/{tender['id']}/costs", json={"kind": "drova"}))
    ok(client.post("/api/tenders", json={"vat_rate": "0", "doc_date": "2025-10-24"}), 201)


def test_nested_values_report_their_path(client):
    body = refused(client.post("/api/tenders", json={"deal_answers": {"условие": 5}}))
    assert body["path"] == ["deal_answers", "условие"]
    body = refused(client.post("/api/tenders", json={"experience_ids": [1, "x"]}))
    assert body["path"] == ["experience_ids", 1] and "№2" in body["error"]
    body = refused(client.post("/api/services", json={"norms": {"Инженер": "много"}}))
    assert body["path"] == ["norms", "Инженер"]
    body = refused(client.post("/api/services", json={"norms": ["a"]}))
    assert body["path"] == ["norms"]


# ---- ограничения базы по-русски -----------------------------------------------

def test_foreign_key_on_save_names_the_field(client):
    body = refused(client.post("/api/tenders", json={"customer_id": 999}))
    assert body["field"] == "customer_id" and "Заказчик" in body["error"]
    me = create(client, "organizations", role="participant", short_name="Я")
    body = refused(client.post("/api/tenders", json={"customer_id": me["id"],
                                                     "participant_id": 888}))
    assert body["field"] == "participant_id" and "Участник" in body["error"]
    assert "Заказчик" not in body["error"]          # существующую ссылку не обвиняет
    assert ok(client.get("/api/tenders")) == []


def test_foreign_key_on_delete_says_what_uses_the_record(client):
    tender = make_tender(client)
    body = refused(client.delete(f"/api/organizations/{tender['customer_id']}"))
    assert "тендер" in body["error"]
    assert ok(client.get(f"/api/organizations/{tender['customer_id']}"))["id"]


def test_translate_integrity_covers_every_constraint_kind(app):
    """Настоящие нарушения из SQLite, не выдуманные строки."""
    conn = db.connect(app.config["DB_PATH"])
    kinds = {}
    try:
        for name, sql in {
            "unique": "INSERT INTO org_settings (id, insurance_rate, injury_rate, overhead_rate, "
                      "profit_rate, vat_rate) VALUES (1, '0','0','0','0','0')",
            "check": "INSERT INTO organization (role) VALUES ('boss')",
            "notnull": "INSERT INTO lot_item (tender_id) VALUES (NULL)",
            "fk": "INSERT INTO lot_item (tender_id) VALUES (999)",
        }.items():
            # у настроек уже может не быть строки: сначала создаём её
            if name == "unique":
                conn.execute("INSERT OR IGNORE INTO org_settings (id, insurance_rate, injury_rate,"
                             " overhead_rate, profit_rate, vat_rate) VALUES (1,'0','0','0','0','0')")
            with pytest.raises(sqlite3.IntegrityError) as caught:
                conn.execute(sql)
            kinds[name] = caught.value
    finally:
        conn.close()
    # с описанием справочника ограничение называет поле по-русски
    assert "«Роль»" in translate_integrity(kinds["check"], ORGANIZATIONS).message
    assert translate_integrity(kinds["check"], ORGANIZATIONS).path == ["role"]
    for name, exc in kinds.items():
        error = translate_integrity(exc)
        assert isinstance(error, ApiError) and error.status == 400
        assert CYRILLIC.search(error.message), (name, error.message)
        assert not any(w in error.message for w in ENGLISH_SQLITE), error.message
        assert "org_settings" not in error.message and "lot_item" not in error.message


def test_unique_violation_from_repo_is_translated_with_field_name(client, monkeypatch):
    def duplicate(self, org):
        raise sqlite3.IntegrityError("UNIQUE constraint failed: organization.inn")

    monkeypatch.setattr(Repo, "save_organization", duplicate)
    body = refused(client.post("/api/organizations", json={"inn": "2904016130"}))
    assert "ИНН" in body["error"] and "уже есть" in body["error"]


# ---- отказ без трассировки ----------------------------------------------------

def test_lookup_error_subclasses_are_not_mistaken_for_missing_record(client, monkeypatch):
    """KeyError -- потомок LookupError. Баг в коде не должен выглядеть как «не найдено»
    и показывать английский текст исключения."""
    def broken(self):
        raise KeyError("english_secret")

    monkeypatch.setattr(Repo, "list_assets", broken)
    response = client.get("/api/assets")
    refused(response, 500)
    assert "english_secret" not in response.get_data(as_text=True)


def test_unexpected_error_during_write_is_500_in_russian(client, monkeypatch):
    def broken(self, employee):
        raise ZeroDivisionError("division by zero")

    monkeypatch.setattr(Repo, "save_employee", broken)
    response = client.post("/api/employees", json={"full_name": "x"})
    body = refused(response, 500)
    assert "division" not in body["error"]


def test_database_busy_is_503_in_russian(client, monkeypatch):
    def locked(self):
        raise sqlite3.OperationalError("database is locked")

    monkeypatch.setattr(Repo, "list_employees", locked)
    body = refused(client.get("/api/employees"), 503)
    assert "занята" in body["error"]


def test_oversized_body_is_refused_in_russian(client):
    response = client.post("/api/employees", data=b"x" * (3 * 1024 * 1024),
                           content_type="application/json")
    refused(response, 413)


def test_huge_record_id_is_404_not_overflow(client):
    refused(client.get(f"/api/organizations/{2 ** 70}"), 404)
    refused(client.delete(f"/api/organizations/{2 ** 70}"), 404)


# ---- защита локального сервера ------------------------------------------------

def test_foreign_host_header_is_refused(client):
    refused(client.get("/api/organizations", headers={"Host": "evil.example:8765"}), 403)
    refused(client.get("/", headers={"Host": "192.168.1.5:8765"}), 403)
    for host in ("127.0.0.1:8765", "localhost:8765", "[::1]:8765"):
        assert client.get("/api/organizations", headers={"Host": host}).status_code == 200


def test_foreign_origin_cannot_write(client):
    response = client.post("/api/employees", json={"full_name": "x"},
                           headers={"Origin": "http://evil.example"})
    refused(response, 403)
    same = client.post("/api/employees", json={"full_name": "x"},
                       headers={"Origin": "http://localhost"})
    assert same.status_code == 201
    refused(client.delete("/api/employees/1", headers={"Origin": "null"}), 403)
    assert len(ok(client.get("/api/employees"))) == 1


def test_post_must_be_json(client):
    refused(client.post("/api/employees", data='{"full_name": "x"}', content_type="text/plain"), 415)
    refused(client.post("/api/employees", data='{"full_name": "x"}'), 415)
    assert ok(client.get("/api/employees")) == []


def test_responses_are_not_cached_or_framed(client):
    response = client.get("/api/organizations")
    assert response.headers["Cache-Control"] == "no-store"
    assert response.headers["X-Frame-Options"] == "DENY"
    assert response.headers["X-Content-Type-Options"] == "nosniff"


# ---- потоки -------------------------------------------------------------------

def test_requests_on_different_threads_work(app):
    """Настоящий сервер Werkzeug с потоком на запрос. Барьер держит оба запроса
    внутри ручки одновременно, поэтому они заведомо на разных потоках и каждый
    выполняет запрос к базе через своё соединение."""
    from werkzeug.serving import make_server

    barrier = threading.Barrier(2, timeout=10)
    idents = []

    @app.get("/api/_probe")
    def probe():
        idents.append(threading.get_ident())
        barrier.wait()
        repo = routes_module.get_repo()
        repo.save_organization(Organization(short_name="из потока"))
        return {"count": len(repo.list_organizations()), "thread": threading.get_ident()}

    server = make_server("127.0.0.1", 0, app, threaded=True)
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        import urllib.request

        def fetch(_):
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/_probe", timeout=15) as r:
                return json.loads(r.read())

        with ThreadPoolExecutor(2) as pool:
            results = list(pool.map(fetch, range(2)))
        assert len(set(idents)) == 2                 # два разных потока обработчика
        assert results[0]["thread"] != results[1]["thread"]
        # соединение не «прилипло» к потоку: третий запрос с третьего потока
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/organizations", timeout=15) as r:
            assert len(json.loads(r.read())) == 2
    finally:
        server.shutdown()
        server.server_close()
        thread.join(5)


def test_many_parallel_writes_do_not_fail(app):
    """Параллельные записи с разных потоков: ни ProgrammingError, ни «database is locked»."""
    def work(n):
        client = app.test_client()
        return client.post("/api/employees", json={"full_name": f"Сотрудник {n}"}).status_code

    with ThreadPoolExecutor(8) as pool:
        codes = list(pool.map(work, range(40)))
    assert codes == [201] * 40
    assert len(app.test_client().get("/api/employees").get_json()) == 40


def test_create_app_is_repeatable_on_the_same_file(tmp_path):
    first = create_app(tmp_path / "x.db").test_client()
    first.post("/api/employees", json={"full_name": "Остаётся"})
    second = create_app(tmp_path / "x.db").test_client()
    assert second.get("/api/employees").get_json()[0]["full_name"] == "Остаётся"


# ---- расчёт -------------------------------------------------------------------

def test_estimate_matches_the_calculator(client):
    tender = make_tender(client)
    tid = tender["id"]
    add_labor(client, tid)                                   # 1 чел. * 100 ч * 1000 = 100 000
    body = ok(client.post(f"/api/tenders/{tid}/estimate", json={}))
    assert body["estimate"] == {
        "materials": "0.00", "labor": "100000.00", "insurance": "30200.00",
        "subcontract": "0.00", "overhead": "6510.00", "profit": "13671.00",
        "total_wo_vat": "150381.00", "vat": "7519.05", "total_with_vat": "157900.05"}
    assert body["vat_rate"] == "0.05" and body["vat_percent"] == "5"
    assert body["lines"]["labor"] == ["100000.00"]


def test_estimate_uses_materials_subcontract_and_tender_vat(client):
    tender = make_tender(client, vat_rate="0.20")
    tid = tender["id"]
    add_labor(client, tid, headcount=2, hours="50")
    create(client, f"tenders/{tid}/costs", kind="materials", name="Бумага", quantity="2",
           price="5000")
    create(client, f"tenders/{tid}/costs", kind="subcontract", name="Замеры", quantity="1",
           price="20000")
    body = ok(client.post(f"/api/tenders/{tid}/estimate", json={}))
    assert body["estimate"]["total_wo_vat"] == "182381.00"
    assert body["estimate"]["vat"] == "36476.20"             # 20 % из тендера, не 5 % из настроек
    assert body["vat_percent"] == "20"
    assert body["lines"]["materials"] == ["10000.00"]


def test_estimate_for_unknown_tender_is_404(client):
    refused(client.post("/api/tenders/404/estimate", json={}), 404)


def test_numbers_that_would_overflow_the_calculator_are_refused_on_save(client):
    """Строку, чьё произведение не округлить до копеек, нельзя было бы ни
    показать, ни удалить через список: отказ -- при сохранении, поле названо."""
    tender = make_tender(client)
    url = f"/api/tenders/{tender['id']}/labor"
    for field, value in (("headcount", 10 ** 7), ("hours", "1000001"), ("rate", "1000001")):
        body = refused(client.post(url, json={field: value}))
        assert body["field"] == field and "слишком велико" in body["error"]
    refused(client.post(f"/api/tenders/{tender['id']}/costs", json={"quantity": "1000000001"}))
    # предельные значения проходят, и расчёт с ними не падает
    add_labor(client, tender["id"], headcount=10 ** 6, hours="1000000", rate="1000000")
    ok(client.get(url))
    ok(client.post(f"/api/tenders/{tender['id']}/estimate", json={}))


def test_decimal_overflow_in_the_calculator_is_a_russian_400(client, monkeypatch):
    import decimal
    tender = make_tender(client)

    def broken(*args, **kwargs):
        raise decimal.InvalidOperation("[<class 'decimal.InvalidOperation'>]")

    monkeypatch.setattr(routes_module, "build_context", broken)
    body = refused(client.post(f"/api/tenders/{tender['id']}/estimate", json={}))
    assert "слишком велики" in body["error"]


def test_solve_returns_proposal_and_does_not_save_it(client):
    tender = make_tender(client)
    tid = tender["id"]
    row = add_labor(client, tid)
    body = ok(client.post(f"/api/tenders/{tid}/solve", json={"target": "300762"}))
    assert body["labor"][0]["id"] == row["id"]
    assert body["labor"][0]["hours"] == "200"                # вдвое больше -- вдвое дороже
    assert body["estimate"]["total_wo_vat"] == "300762.00"
    assert body["target"] == "300762" and body["target_includes_vat"] is False
    assert ok(client.get(f"/api/tenders/{tid}/labor"))["rows"][0]["hours"] == "100"   # не сохранено


def test_solve_refusals(client):
    tender = make_tender(client)
    tid = tender["id"]
    assert "нет трудозатрат" in refused(client.post(
        f"/api/tenders/{tid}/solve", json={"target": "1000"}))["error"]
    body = refused(client.post(f"/api/tenders/{tid}/solve", json={}))
    assert body["field"] == "target"
    refused(client.post(f"/api/tenders/{tid}/solve", json={"target": "-5"}))
    refused(client.post(f"/api/tenders/{tid}/solve", json={"target": "много"}))
    refused(client.post(f"/api/tenders/{tid}/solve", json={"target": "0"}))
    refused(client.post(f"/api/tenders/{tid}/solve", json={"target_wo_vat": "1"}))      # прежнего имени поля больше нет


# ---- сборка -------------------------------------------------------------------

def test_build_existing_file_is_skipped_then_replaced_on_overwrite(client, tmp_path, dirs):
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "{{ tender.subject }}")
    tender = make_tender(client)
    out = tmp_path / "out"
    url = f"/api/tenders/{tender['id']}/build"
    request = {"documents": ["Декларация"], "output_dir": str(out)}

    first = ok(client.post(url, json=request))
    assert len(first["written"]) == 1
    client.post("/api/tenders", json={"id": tender["id"], "subject": "Новый предмет"})

    again = ok(client.post(url, json=request))
    assert again["written"] == [] and "уже есть" in again["skipped"][0]["reason"]
    assert text_of(out / "Декларация.docx") == ["Разработка КЭР"]

    replaced = ok(client.post(url, json={**request, "overwrite": True}))
    assert len(replaced["written"]) == 1
    assert text_of(out / "Декларация.docx") == ["Новый предмет"]
    refused(client.post(url, json={**request, "overwrite": "да"}))


def test_plan_lists_files_and_which_exist_and_writes_nothing(client, tmp_path, dirs):
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "x")
    make_template(templates / "Справка о МТР.docx", "x")
    customer = create(client, "organizations", role="customer", short_name="Заказчик")
    create(client, "profiles", customer_id=customer["id"],
           docs={"Справка о МТР": {"prefix": "8.1_"}})
    tender = create(client, "tenders", customer_id=customer["id"])
    out = tmp_path / "plan_out"
    out.mkdir()
    (out / "Декларация.docx").write_bytes("старый".encode())
    before = sorted(p.name for p in out.iterdir())
    body = ok(client.post(f"/api/tenders/{tender['id']}/plan", json={
        "documents": ["Декларация", "Справка о МТР", "Нет шаблона"], "output_dir": str(out)}))
    assert [(f["document"], f["name"], f["exists"]) for f in body["files"]] == [
        ("Декларация", "Декларация.docx", True),
        ("Справка о МТР", "8.1_Справка о МТР.docx", False)]
    assert sorted(p.name for p in out.iterdir()) == before
    assert (out / "Декларация.docx").read_bytes() == "старый".encode()


def test_plan_and_build_agree_on_paths(client, tmp_path, dirs):
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "x")
    tender = make_tender(client)
    request = {"documents": ["Декларация"], "output_dir": str(tmp_path / "same")}
    planned = ok(client.post(f"/api/tenders/{tender['id']}/plan", json=request))["files"]
    built = ok(client.post(f"/api/tenders/{tender['id']}/build", json=request))["written"]
    assert [f["path"] for f in planned] == [w["path"] for w in built]


def test_build_documents_default_to_the_customer_profile(client, tmp_path, dirs):
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "x")
    customer = create(client, "organizations", role="customer", short_name="Заказчик")
    create(client, "profiles", customer_id=customer["id"], documents=["Декларация"])
    tender = create(client, "tenders", customer_id=customer["id"],
                    output_dir=str(tmp_path / "remembered"))
    body = ok(client.post(f"/api/tenders/{tender['id']}/build", json={}))   # ни состава, ни папки
    assert [w["name"] for w in body["written"]] == ["Декларация.docx"]
    assert (tmp_path / "remembered" / "Декларация.docx").is_file()


def test_build_warnings_reach_the_response(client, tmp_path, dirs):
    templates, _ = dirs
    make_template(templates / "Справка об опыте.docx", "x")
    customer = create(client, "organizations", role="customer", short_name="Заказчик")
    create(client, "profiles", customer_id=customer["id"], documents=["Справка об опыте"])
    tender = create(client, "tenders", customer_id=customer["id"])
    body = ok(client.post(f"/api/tenders/{tender['id']}/build", json={
        "documents": ["Справка об опыте"], "output_dir": str(tmp_path / "w"),
        "experience_ids": [4242]}))
    texts = [w["text"] for w in body["warnings"]]
    assert any("не назначена раскладка" in t for t in texts)          # колонки по умолчанию
    assert any("id=4242" in t and "не найдена" in t for t in texts)   # исчезнувший опыт
    assert {w["document"] for w in body["warnings"]} == {"Справка об опыте"}


def test_build_passes_experience_ids_to_the_package(client, tmp_path, dirs, monkeypatch):
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "x")
    tender = make_tender(client)
    seen = {}
    real = routes_module.build_package

    def spy(*args, **kwargs):
        seen.update(kwargs)
        return real(*args, **kwargs)

    monkeypatch.setattr(routes_module, "build_package", spy)
    ok(client.post(f"/api/tenders/{tender['id']}/build", json={
        "documents": ["Декларация"], "output_dir": str(tmp_path / "e"),
        "experience_ids": [3, 1], "overwrite": True}))
    assert seen["experience_ids"] == [3, 1] and seen["overwrite"] is True
    seen.clear()
    ok(client.post(f"/api/tenders/{tender['id']}/build", json={
        "documents": ["Декларация"], "output_dir": str(tmp_path / "e")}))
    assert seen["experience_ids"] is None and seen["overwrite"] is False   # сохранённый отбор


def test_build_refusals(client, tmp_path, dirs):
    templates, _ = dirs
    tender = make_tender(client)
    url = f"/api/tenders/{tender['id']}/build"
    refused(client.post("/api/tenders/999/build", json={"documents": ["x"],
                                                        "output_dir": str(tmp_path)}), 404)
    assert "документ" in refused(client.post(url, json={"output_dir": str(tmp_path)}))["error"]
    body = refused(client.post(url, json={"documents": ["x"]}))
    assert body["field"] == "output_dir"
    refused(client.post(url, json={"documents": "Декларация", "output_dir": str(tmp_path)}))
    refused(client.post(url, json={"documents": ["x"], "output_dir": str(tmp_path), "extra": 1}))
    refused(client.post(url, json={"documents": ["x"], "output_dir": str(tmp_path),
                                   "experience_ids": ["a"]}))


def test_unusable_output_folder_is_a_russian_400(client, tmp_path, dirs):
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "x")
    tender = make_tender(client)
    occupied = tmp_path / "файл.txt"
    occupied.write_text("это файл, а не папка")
    body = refused(client.post(f"/api/tenders/{tender['id']}/build", json={
        "documents": ["Декларация"], "output_dir": str(occupied / "внутри")}))
    assert "Не удалось создать папку" in body["error"]
    assert occupied.read_text() == "это файл, а не папка"


def test_build_without_templates_folder_is_a_russian_400(client, tmp_path, monkeypatch):
    monkeypatch.setattr(package_module, "TEMPLATES_DOCS_DIR", tmp_path / "нет такой")
    tender = make_tender(client)
    body = refused(client.post(f"/api/tenders/{tender['id']}/build", json={
        "documents": ["Декларация"], "output_dir": str(tmp_path / "o")}))
    assert "Папка шаблонов" in body["error"]
    assert not (tmp_path / "o").exists()


# ---- куда сборка может писать -------------------------------------------------

@pytest.mark.parametrize("bad", [
    "", "   ", "relative\\folder", "..\\..\\out", "out", "C:folder", "\\\\server\\share\\out",
    "//server/share/out", "\\\\?\\C:\\out", "\\\\.\\C:\\out", "C:\\out:stream",
    "C:\\out\x00evil", "C:\\out\nevil", "\\rooted", 5, None, ["C:\\out"]])
def test_output_dir_must_be_an_absolute_local_path(client, bad, tmp_path, dirs):
    tender = make_tender(client)
    for url in (f"/api/tenders/{tender['id']}/build", f"/api/tenders/{tender['id']}/plan"):
        body = refused(client.post(url, json={"documents": ["Декларация"], "output_dir": bad}))
        assert body["field"] == "output_dir"
    if bad not in ("", "   "):                 # пустая папка в тендере -- «ещё не выбрана»
        refused(client.post("/api/tenders", json={"output_dir": bad}))
    assert list(tmp_path.glob("out*")) == []


def test_output_dir_inside_the_program_folder_is_refused(client, dirs):
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "шаблон")
    tender = make_tender(client)
    for inside in (BASE_DIR, BASE_DIR / "templates_docs", BASE_DIR / "app" / "web",
                   BASE_DIR / "data" / "x", BASE_DIR / "layouts" / ".." / "app"):
        body = refused(client.post(f"/api/tenders/{tender['id']}/build", json={
            "documents": ["Декларация"], "output_dir": str(inside), "overwrite": True}))
        assert "программ" in body["error"]
    assert not (BASE_DIR / "data" / "x").exists()


def test_dotdot_in_output_dir_is_normalized_not_followed_blindly(client, tmp_path, dirs):
    """«C:\\a\\..\\b» -- это папка b: пользователь выбрал именно её, и файлы
    лежат в ней, а не в a."""
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "x")
    tender = make_tender(client)
    chosen = tmp_path / "a" / ".." / "b"
    (tmp_path / "a").mkdir()
    body = ok(client.post(f"/api/tenders/{tender['id']}/build", json={
        "documents": ["Декларация"], "output_dir": str(chosen)}))
    assert Path(body["written"][0]["path"]) == tmp_path / "b" / "Декларация.docx"
    assert list((tmp_path / "a").iterdir()) == []


def test_hostile_document_names_and_prefixes_cannot_leave_the_folder(client, tmp_path, dirs):
    templates, _ = dirs
    make_template(templates / "Декларация.docx", "x")
    out = tmp_path / "sandbox" / "out"
    outside_before = sorted(p.relative_to(tmp_path).as_posix() for p in tmp_path.rglob("*"))
    customer = create(client, "organizations", role="customer", short_name="Заказчик")
    create(client, "profiles", customer_id=customer["id"], documents=["Декларация"],
           docs={"Декларация": {"prefix": "..\\..\\..\\evil_"}})
    tender = create(client, "tenders", customer_id=customer["id"])
    url = f"/api/tenders/{tender['id']}/build"

    ok(client.post(url, json={"documents": ["Декларация"], "output_dir": str(out)}))
    for name in ("..\\..\\x", "../../x", "a/b", "C:\\x", "x\x00y", "", "  "):
        body = refused(client.post(url, json={"documents": [name], "output_dir": str(out)}))
        assert body["path"][0] == "documents"

    for path in tmp_path.rglob("*"):
        rel = path.relative_to(tmp_path).as_posix()
        if rel in outside_before:
            continue
        assert path == out or out in path.parents or path in out.parents \
            or path.parent == templates.parent, f"вне папки: {rel}"
    written = [p.name for p in out.iterdir()]
    assert len(written) == 1 and written[0].endswith("evil_Декларация.docx")
    assert "\\" not in written[0] and "/" not in written[0]   # разделители стали дефисами


def test_startup_failures_are_russian_not_tracebacks(app, monkeypatch, capsys):
    from app.web import __main__ as entry

    def broken(path):
        raise sqlite3.OperationalError("unable to open database file")

    monkeypatch.setattr(entry, "create_app", broken)
    assert entry.main() == 1
    out = capsys.readouterr().out
    assert "базу данных" in out and "Traceback" not in out

    monkeypatch.setattr(entry, "create_app", lambda path: app)

    def busy(app, port=entry.PORT):
        raise entry.BindError("[WinError 10048] Only one usage of each socket address")

    monkeypatch.setattr(entry, "build_server", busy)
    assert entry.main() == 1
    out = capsys.readouterr().out
    assert "порт 8765 занят" in out and "Traceback" not in out


def test_second_server_cannot_take_an_occupied_port(app):
    """Двойной щелчок по ярлыку: вторая копия должна получить отказ, а не молча
    занять тот же порт (так делает SO_REUSEADDR на Windows)."""
    from app.web.__main__ import BindError, build_server
    first = build_server(app, port=0)
    try:
        port = first.server_address[1]
        with pytest.raises(BindError):        # а не SystemExit с английским текстом
            build_server(app, port=port)
    finally:
        first.server_close()


def test_server_binds_to_loopback_only(app):
    from app.web.__main__ import HOST, PORT, build_server
    assert HOST == "127.0.0.1" and PORT == 8765
    server = build_server(app, port=0)
    try:
        assert server.server_address[0] == "127.0.0.1"
    finally:
        server.server_close()


# ---- деньги не считаются на экране: цель «с НДС», расхождение, готовые строки ---

NBSP = " "


def test_estimate_also_carries_ready_to_show_money_strings(client):
    tender = make_tender(client)
    add_labor(client, tender["id"])
    body = ok(client.post(f"/api/tenders/{tender['id']}/estimate", json={}))
    assert body["estimate_text"]["total_wo_vat"] == f"150{NBSP}381,00"
    assert body["estimate_text"]["total_with_vat"] == f"157{NBSP}900,05"
    assert set(body["estimate_text"]) == set(body["estimate"])
    assert body["lines_text"]["labor"] == [f"100{NBSP}000,00"]
    assert "difference" not in body                      # цели нет -- сравнивать не с чем


def test_estimate_with_a_target_returns_the_gap_from_the_calculator(client):
    tender = make_tender(client)
    add_labor(client, tender["id"])                      # итог без НДС 150 381
    body = ok(client.post(f"/api/tenders/{tender['id']}/estimate",
                          json={"target": "150 000,50"}))
    assert body["difference"] == "380.50"                # итог выше цели
    assert body["difference_text"] == "380,50"
    assert body["target_text"] == f"150{NBSP}000,50"
    assert body["target_includes_vat"] is False

    with_vat = ok(client.post(f"/api/tenders/{tender['id']}/estimate",
                              json={"target": "158000", "target_includes_vat": True}))
    assert with_vat["difference"] == "-99.95"            # 157 900,05 - 158 000
    assert with_vat["difference_text"] == "-99,95"
    assert with_vat["target_includes_vat"] is True

    refused(client.post(f"/api/tenders/{tender['id']}/estimate", json={"target": "0"}))
    refused(client.post(f"/api/tenders/{tender['id']}/estimate",
                        json={"target": "1", "target_includes_vat": "да"}))
    refused(client.post(f"/api/tenders/{tender['id']}/estimate", json={"target_includes_vat": True}))


def test_solve_target_with_vat_and_gap(client):
    tender = make_tender(client)
    add_labor(client, tender["id"])
    url = f"/api/tenders/{tender['id']}/solve"
    # 200 часов -> итог без НДС 300 762, с НДС 315 800,10
    exact = ok(client.post(url, json={"target": "315800.10", "target_includes_vat": True}))
    assert exact["labor"][0]["hours"] == "200"
    assert exact["estimate"]["total_with_vat"] == "315800.10"
    assert exact["difference"] == "0.00" and exact["difference_text"] == "0,00"
    assert exact["target_includes_vat"] is True
    assert exact["estimate_text"]["total_with_vat"] == f"315{NBSP}800,10"
    assert exact["labor"][0]["rate_text"] == f"1{NBSP}000,00"

    plain = ok(client.post(url, json={"target": "300762"}))             # без НДС (по умолчанию)
    assert plain["labor"][0]["hours"] == "200"
    assert plain["difference"] == "0.00" and plain["target_includes_vat"] is False

    # расхождение приходит из расчёта, а не из сложения статей: итог округлён до рублей
    off = ok(client.post(url, json={"target": "300761.83"}))
    est = off["estimate"]
    from decimal import Decimal as D
    assert D(off["difference"]) == D(est["total_wo_vat"]) - D("300761.83")
    assert off["difference"] == "0.17"
    assert off["difference_text"] == "0,17"


def test_solve_takes_target_and_flag_not_the_old_target_wo_vat_field(client):
    tender = make_tender(client)
    add_labor(client, tender["id"])
    body = refused(client.post(f"/api/tenders/{tender['id']}/solve",
                               json={"target_wo_vat": "300762"}))
    assert body["field"] == "target_wo_vat"
    refused(client.post(f"/api/tenders/{tender['id']}/solve",
                        json={"target": "1", "target_includes_vat": 1}))


def test_solve_below_target_with_vat_is_400_in_terms_of_vat(client):
    tender = make_tender(client)
    add_labor(client, tender["id"])
    create(client, f"tenders/{tender['id']}/costs", kind="materials", name="Бумага",
           unit="шт.", quantity="1", price="500000")
    body = refused(client.post(f"/api/tenders/{tender['id']}/solve",
                               json={"target": "100000", "target_includes_vat": True}))
    assert "с НДС" in body["error"]


# ---- замена позиций лота ------------------------------------------------------

def test_put_items_replaces_the_whole_lot_in_one_go(client):
    tender = make_tender(client)
    tid = tender["id"]
    create(client, f"tenders/{tid}/items", position=1, name="старая", price="1")
    pasted = [{"position": n, "name": f"Позиция {n}", "price": f"{n}00,50"} for n in range(1, 12)]
    saved = ok(client.put(f"/api/tenders/{tid}/items", json={"items": pasted}))["rows"]
    assert [i["name"] for i in saved] == [f"Позиция {n}" for n in range(1, 12)]
    assert saved[0]["price"] == "100.50" and saved[0]["tender_id"] == tid
    assert "старая" not in [i["name"] for i in saved]                       # прежняя заменена
    assert ok(client.get(f"/api/tenders/{tid}/items"))["rows"] == saved

    # отправленные обратно как есть (с id и tender_id) позиции -- тоже годятся
    again = ok(client.put(f"/api/tenders/{tid}/items", json={"items": saved}))["rows"]
    assert [i["name"] for i in again] == [i["name"] for i in saved]

    assert ok(client.put(f"/api/tenders/{tid}/items", json={"items": []}))["rows"] == []
    assert ok(client.get(f"/api/tenders/{tid}/items"))["rows"] == []


def test_put_items_failing_on_the_seventh_changes_nothing(client):
    tender = make_tender(client)
    tid = tender["id"]
    create(client, f"tenders/{tid}/items", position=1, name="прежняя", price="5")
    pasted = [{"position": n, "name": f"П{n}", "price": "10"} for n in range(1, 12)]
    pasted[6]["price"] = "много"
    body = refused(client.put(f"/api/tenders/{tid}/items", json={"items": pasted}))
    assert body["path"] == ["items", 6, "price"] and body["field"] == "items"
    assert "Позиция №7" in body["error"] and "Цена" in body["error"]
    assert [i["name"] for i in ok(client.get(f"/api/tenders/{tid}/items"))["rows"]] == ["прежняя"]


def test_put_items_refusals(client):
    tender = make_tender(client)
    other = make_tender(client)
    url = f"/api/tenders/{tender['id']}/items"
    refused(client.put("/api/tenders/999/items", json={"items": []}), 404)
    assert refused(client.put(url, json={}))["field"] == "items"
    refused(client.put(url, json={"items": "список"}))
    refused(client.put(url, json={"items": ["строка"]}))
    refused(client.put(url, json={"items": [{"name": "x"}], "extra": 1}))
    body = refused(client.put(url, json={"items": [{"nme": "x"}]}))
    assert body["path"] == ["items", 0, "nme"]
    body = refused(client.put(url, json={"items": [{"name": "x", "tender_id": other["id"]}]}))
    assert body["path"] == ["items", 0, "tender_id"]
    refused(client.put(url, json={"items": [{"name": "x"}] * 2001}))
    refused(client.put(url, data='{"items": []}', content_type="text/plain"), 415)
    assert ok(client.get(url))["rows"] == []


def test_put_is_guarded_like_post(client):
    tender = make_tender(client)
    url = f"/api/tenders/{tender['id']}/items"
    refused(client.put(url, json={"items": []}, headers={"Origin": "http://evil.example"}), 403)
    refused(client.put(url, json={"items": []}, headers={"Host": "evil.example"}), 403)


# ---- у каждого числа на экране -- готовая строка от сервера ---------------------
# Экран не считает и не форматирует ничего: в JavaScript нет Decimal. Рядом с
# сырым значением (его экран отправляет обратно) лежит <поле>_text (через money),
# у ставки -- <поле>_percent, у списка под итогом -- итог.

def decimal_field_names() -> set[str]:
    """Имена всех Decimal-полей моделей: по ним проверяется, что у каждого есть пара."""
    import dataclasses
    import typing
    from app import models

    names = set()
    for obj in vars(models).values():
        if dataclasses.is_dataclass(obj):
            for name, hint in typing.get_type_hints(obj).items():
                if "Decimal" in str(hint):
                    names.add(name)
    return names


DECIMAL_FIELDS = decimal_field_names() | {"difference", "target", "total", "amount", "vat",
                                                      "vat_total"}
NUMERIC = re.compile(r"-?\d+(\.\d+)?")
TWINNED_DICTS = {"estimate", "lines", "totals", "norms"}


def bare_numbers(value, path=""):
    """Пути денежных/дробных значений без готовой пары рядом: то, что экрану
    пришлось бы форматировать самому."""
    found = []
    if isinstance(value, list):
        for i, item in enumerate(value):
            found += bare_numbers(item, f"{path}[{i}]")
    elif isinstance(value, dict):
        for key, item in value.items():
            if key.endswith(("_text", "_percent")):
                continue
            where = f"{path}.{key}"
            if isinstance(item, str) and key in DECIMAL_FIELDS and NUMERIC.fullmatch(item):
                if not ({key + "_text", key + "_percent"} & set(value)
                        or (key == "vat_rate" and "vat_percent" in value)):
                    found.append(where)
            elif key in TWINNED_DICTS and isinstance(item, dict):
                twin = value.get(key + "_text")
                if not isinstance(twin, dict) or set(twin) != set(item):
                    found.append(where + " (нет пары " + key + "_text)")
            elif key in TWINNED_DICTS and isinstance(item, list):
                twin = value.get(key + "_text")
                if not isinstance(twin, list) or len(twin) != len(item):
                    found.append(where + " (нет пары " + key + "_text)")
            else:
                found += bare_numbers(item, where)
    return found


def test_scalar_money_and_percent_fields_have_ready_strings(client):
    emp = create(client, "employees", full_name="Рабичко", hourly_rate="1250.5")
    assert emp["hourly_rate"] == "1250.5"                       # сырое -- для отправки обратно
    assert emp["hourly_rate_text"] == f"1{NBSP}250,50"

    exp = create(client, "experience", customer_name="МУП", amount="5985000")
    assert exp["amount_text"] == f"5{NBSP}985{NBSP}000,00"

    org = create(client, "organizations", role="customer", vat_rate="0.055")
    assert org["vat_rate_percent"] == "5,5"
    assert create(client, "organizations", role="customer", vat_rate="0.05")["vat_rate_percent"] == "5"
    assert create(client, "organizations", role="customer")["vat_rate_percent"] is None

    tender = create(client, "tenders", price_wo_vat="5700000", vat_rate="0.2")
    assert tender["price_wo_vat_text"] == f"5{NBSP}700{NBSP}000,00"
    assert tender["vat_rate_percent"] == "20"
    assert create(client, "tenders")["vat_rate_percent"] is None
    assert create(client, "tenders")["price_wo_vat_text"] == "0,00"

    service = create(client, "services", code="КЭР", norms={"Инженер": "12.50", "Директор": "30"})
    assert service["norms_text"] == {"Инженер": "12,5", "Директор": "30"}

    # то же в списках и в одной записи
    assert ok(client.get("/api/employees"))[0]["hourly_rate_text"] == f"1{NBSP}250,50"
    assert ok(client.get(f"/api/experience/{exp['id']}"))["amount_text"] == exp["amount_text"]


def test_settings_have_percent_strings(client):
    settings = ok(client.get("/api/settings"))
    assert settings["insurance_rate_percent"] == "30"
    assert settings["injury_rate_percent"] == "0,2"
    assert settings["overhead_rate_percent"] == "5"
    assert settings["profit_rate_percent"] == "10"
    assert settings["vat_rate_percent"] == "5"
    assert settings["total_rounding_digits"] == 0
    changed = ok(client.post("/api/settings", json={"vat_rate": "0.205"}))
    assert changed["vat_rate_percent"] == "20,5"


def test_lot_items_come_with_texts_and_a_server_side_total(client):
    tender = make_tender(client)
    tid = tender["id"]
    # одиннадцать позиций, как у Роснефти; копейки, чтобы сложение было видно
    prices = [f"{n}00{n}.{n % 10}5" for n in range(1, 12)]
    pasted = [{"position": n, "name": f"П{n}", "price": p} for n, p in enumerate(prices, 1)]
    body = ok(client.put(f"/api/tenders/{tid}/items", json={"items": pasted}))
    from decimal import Decimal as D
    expected = sum((D(p) for p in prices), D("0.00"))
    assert body["total"] == format(expected, "f")
    assert body["total_text"] == f"{expected:,.2f}".translate({ord(","): NBSP, ord("."): ","})
    assert [r["price_text"] for r in body["rows"]][0] == f"1{NBSP}001,15"
    assert [r["price"] for r in body["rows"]][0] == "1001.15"

    again = ok(client.get(f"/api/tenders/{tid}/items"))
    assert again == body
    assert ok(client.get(f"/api/tenders/{create(client, 'tenders')['id']}/items")) == {
        "rows": [], "total": "0.00", "total_text": "0,00", "vat_total": "0.00",
        "vat_total_text": "0,00", "vat_percent": "5"}


def test_lot_total_is_the_sum_of_printed_prices_not_of_raw_ones(client):
    # три цены по 0,005: напечатаны будут по 0,01, итог напечатанных -- 0,03.
    # Сумма сырых (0,015 -> «0,02») разошлась бы со строками над ней.
    tender = make_tender(client)
    body = ok(client.put(f"/api/tenders/{tender['id']}/items", json={
        "items": [{"name": str(n), "price": "0.005"} for n in range(3)]}))
    assert [r["price_text"] for r in body["rows"]] == ["0,01"] * 3
    assert body["total"] == "0.03" and body["total_text"] == "0,03"


def test_single_lot_item_save_has_its_text(client):
    tender = make_tender(client)
    item = create(client, f"tenders/{tender['id']}/items", name="x", price="1234.5")
    assert item["price_text"] == f"1{NBSP}234,50" and item["price"] == "1234.5"


def test_labor_rows_have_texts_and_a_server_side_fot(client):
    tender = make_tender(client)
    tid = tender["id"]
    first = add_labor(client, tid, headcount=2, hours="12.5", rate="270")       # 6 750,00
    assert first["rate_text"] == f"270,00" and first["hours_text"] == "12,5"
    assert first["amount"] == "6750.00" and first["amount_text"] == f"6{NBSP}750,00"
    add_labor(client, tid, post="Начальник", headcount=1, hours="360", rate="420")  # 151 200,00
    body = ok(client.get(f"/api/tenders/{tid}/labor"))
    assert [r["amount"] for r in body["rows"]] == ["6750.00", "151200.00"]
    assert body["total"] == "157950.00" and body["total_text"] == f"157{NBSP}950,00"
    # ФОТ тот же, что в расчёте: экран не складывает строки сам
    assert ok(client.post(f"/api/tenders/{tid}/estimate", json={}))["estimate"]["labor"] \
        == body["total"]


def test_cost_rows_have_texts_and_totals_per_kind(client):
    tender = make_tender(client)
    tid = tender["id"]
    mat = create(client, f"tenders/{tid}/costs", kind="materials", name="Бумага",
                 unit="пач.", quantity="2.5", price="199.99")          # 499,975 -> 499,98
    assert mat["quantity_text"] == "2,5" and mat["price_text"] == "199,99"
    assert mat["amount"] == "499.98" and mat["amount_text"] == "499,98"
    create(client, f"tenders/{tid}/costs", kind="materials", name="Тонер", quantity="1",
           price="1000")
    create(client, f"tenders/{tid}/costs", kind="subcontract", name="Замеры", quantity="1",
           price="20000")
    body = ok(client.get(f"/api/tenders/{tid}/costs"))
    assert [r["amount_text"] for r in body["rows"]] == [
        "499,98", f"1{NBSP}000,00", f"20{NBSP}000,00"]
    assert body["totals"] == {"materials": "1499.98", "subcontract": "20000.00"}
    assert body["totals_text"] == {"materials": f"1{NBSP}499,98",
                                   "subcontract": f"20{NBSP}000,00"}
    only = ok(client.get(f"/api/tenders/{tid}/costs?kind=subcontract"))
    assert [r["name"] for r in only["rows"]] == ["Замеры"]
    assert only["totals"] == body["totals"]
    estimate = ok(client.post(f"/api/tenders/{tid}/estimate", json={}))["estimate"]
    assert estimate["materials"] == body["totals"]["materials"]           # те же числа
    assert estimate["subcontract"] == body["totals"]["subcontract"]


def test_solved_hours_have_texts(client):
    tender = make_tender(client)
    add_labor(client, tender["id"])
    body = ok(client.post(f"/api/tenders/{tender['id']}/solve", json={"target": "300762"}))
    assert body["labor"][0]["hours_text"] == "200"
    assert body["labor"][0]["rate_text"] == f"1{NBSP}000,00"


def test_no_screen_ever_has_to_format_a_number(client, monkeypatch):
    """Приёмка координатора: число на экране, которого нет в ответе готовой
    строкой, значит посчитал и отформатировал браузер. Обход всех ответов,
    несущих деньги, ставки и часы."""
    from decimal import Decimal

    from app.models import Organization as Org
    tender = make_tender(client, price_wo_vat="5700000", vat_rate="0.05")
    tid = tender["id"]
    create(client, "employees", full_name="А", hourly_rate="100")
    create(client, "experience", customer_name="Б", amount="10")
    create(client, "services", code="КЭР", norms={"И": "1"})
    create(client, f"tenders/{tid}/items", name="п", price="10")
    add_labor(client, tid)
    create(client, f"tenders/{tid}/costs", kind="materials", name="м", quantity="1", price="5")
    monkeypatch.setattr(routes_module, "build_draft",
                        lambda inn, **kw: (Org(vat_rate=Decimal("0.05")), []))
    responses = [
        client.get("/api/employees"), client.get("/api/experience"),
        client.get("/api/services"), client.get("/api/organizations"),
        client.get("/api/tenders"), client.get(f"/api/tenders/{tid}"),
        client.get("/api/settings"),
        client.get(f"/api/tenders/{tid}/items"), client.get(f"/api/tenders/{tid}/labor"),
        client.get(f"/api/tenders/{tid}/costs"),
        client.post(f"/api/tenders/{tid}/estimate", json={"target": "100000"}),
        client.post(f"/api/tenders/{tid}/solve", json={"target": "300762"}),
        client.put(f"/api/tenders/{tid}/items", json={"items": [{"name": "x", "price": "1"}]}),
        client.get("/api/inn/2904016130"),
        client.post("/api/settings", json={"vat_rate": "0.06"}),
    ]
    for response in responses:
        bare = bare_numbers(ok(response))
        assert bare == [], (response.request.path, bare)


def test_a_record_can_be_sent_back_exactly_as_received(client):
    """Экран отправляет записи обратно: готовые строки и суммы -- только для показа,
    сервер их игнорирует, а неизвестные ключи по-прежнему отказ."""
    tender = make_tender(client, price_wo_vat="5700000", vat_rate="0.05")
    tid = tender["id"]
    labor = add_labor(client, tid)
    cost = create(client, f"tenders/{tid}/costs", kind="materials", name="м", quantity="2",
                  price="5")
    service = create(client, "services", code="КЭР", norms={"И": "1.5"})
    employee = create(client, "employees", full_name="А", hourly_rate="100")
    for path, record in ((f"tenders/{tid}/labor", labor), (f"tenders/{tid}/costs", cost),
                         ("services", service), ("employees", employee),
                         ("tenders", ok(client.get(f"/api/tenders/{tid}"))),
                         ("settings", ok(client.get("/api/settings")))):
        back = ok(client.post(f"/api/{path}", json=record))
        assert back == record, path
    refused(client.post("/api/employees", json={**employee, "hourly_rate_texts": "1"}))


# ---- арифметики в слое веба нет: итоги и НДС позиций, суммы строк подбора -------

def test_lot_items_carry_vat_per_row_and_a_vat_total(client):
    tender = make_tender(client)                       # НДС из настроек, 5 %
    tid = tender["id"]
    body = ok(client.put(f"/api/tenders/{tid}/items", json={"items": [
        {"position": 1, "name": "а", "price": "100"},
        {"position": 2, "name": "б", "price": "0.50"}]}))
    assert [r["vat"] for r in body["rows"]] == ["5.00", "0.03"]       # 0,025 -> 0,03 на строке
    assert [r["vat_text"] for r in body["rows"]] == ["5,00", "0,03"]
    assert body["vat_total"] == "5.03" and body["vat_total_text"] == "5,03"
    assert body["vat_percent"] == "5"
    assert body["total"] == "100.50"


def test_lot_vat_total_is_the_sum_of_printed_row_vat_not_vat_of_the_total(client):
    # три позиции по 0,50: НДС строки 0,025 печатается «0,03»; итог напечатанных --
    # 0,09, а НДС от итога 1,50 был бы 0,075 -> «0,08». Форма Роснефти складывает
    # колонку НДС её же формулами, поэтому сходиться надо со строками.
    tender = make_tender(client)
    body = ok(client.put(f"/api/tenders/{tender['id']}/items", json={
        "items": [{"name": str(n), "price": "0.50"} for n in range(3)]}))
    assert body["vat_total"] == "0.09"


def test_lot_vat_uses_the_tender_rate(client):
    tender = make_tender(client, vat_rate="0.20")
    body = ok(client.put(f"/api/tenders/{tender['id']}/items", json={
        "items": [{"name": "а", "price": "10.05"}]}))
    assert body["rows"][0]["vat"] == "2.01" and body["vat_percent"] == "20"
    one = create(client, f"tenders/{tender['id']}/items", name="б", price="100")
    assert one["vat"] == "20.00" and one["vat_text"] == "20,00"      # одна запись -- тоже с НДС


def test_empty_lot_has_zero_vat_total(client):
    tender = make_tender(client)
    body = ok(client.get(f"/api/tenders/{tender['id']}/items"))
    assert body["vat_total"] == "0.00" and body["vat_total_text"] == "0,00"
    assert body["total"] == "0.00" and body["rows"] == []


def test_solved_rows_carry_amounts_like_stored_labor_rows(client):
    tender = make_tender(client)
    tid = tender["id"]
    add_labor(client, tid)                                              # 1 чел., 100 ч, 1000
    add_labor(client, tid, post="Помощник", headcount=2, hours="50", rate="100")   # 10 000
    body = ok(client.post(f"/api/tenders/{tid}/solve", json={"target": "300000"}))
    rows = body["labor"]
    assert [set(r) for r in rows] == [set(ok(client.get(f"/api/tenders/{tid}/labor"))["rows"][0])] * 2
    for row in rows:
        hours, rate = Decimal(row["hours"]), Decimal(row["rate"])
        assert Decimal(row["amount"]) == row["headcount"] * hours * rate
        assert row["amount_text"] == f"{Decimal(row['amount']):,.2f}".translate(
            {ord(","): NBSP, ord("."): ","})
    # решение не сохранено, но строки подбора -- в той же форме, что у сохранённых
    assert ok(client.get(f"/api/tenders/{tid}/labor"))["rows"][0]["hours"] == "100"
    assert body["estimate"]["labor"] == format(sum(Decimal(r["amount"]) for r in rows), "f")
