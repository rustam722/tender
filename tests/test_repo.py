import sqlite3
from decimal import Decimal

import pytest

from app import db
from app.models import (
    ROLE_CUSTOMER, ROLE_PARTICIPANT, Asset, CustomerProfile, Employee,
    ExperienceRecord, LotItem, Organization, OrgSettings, Service, Tender,
    TenderCostItem, TenderLaborRow,
)
from app.repo import Repo


@pytest.fixture
def repo(tmp_path):
    conn = db.connect(tmp_path / "t.db")
    db.migrate(conn)
    yield Repo(conn)
    conn.close()


def make_org(**kw) -> Organization:
    kw.setdefault("short_name", "ООО «Техносфера»")
    return Organization(**kw)


def make_tender(repo, **kw) -> int:
    customer = repo.save_organization(make_org(role=ROLE_CUSTOMER, inn="111"))
    kw.setdefault("customer_id", customer)
    return repo.save_tender(Tender(**kw))


def test_save_then_get_organization(repo):
    oid = repo.save_organization(make_org(inn="5904188041"))
    assert repo.get_organization(oid).inn == "5904188041"


def test_save_with_id_updates_in_place(repo):
    oid = repo.save_organization(make_org(inn="1"))
    org = repo.get_organization(oid)
    org.inn = "2"
    assert repo.save_organization(org) == oid
    assert len(repo.list_organizations()) == 1
    assert repo.get_organization(oid).inn == "2"


def test_list_organizations_filters_by_role(repo):
    mine = repo.save_organization(make_org(role=ROLE_PARTICIPANT, inn="1"))
    theirs = repo.save_organization(make_org(role=ROLE_CUSTOMER, inn="2"))
    assert [o.id for o in repo.list_organizations(ROLE_PARTICIPANT)] == [mine]
    assert [o.id for o in repo.list_organizations(ROLE_CUSTOMER)] == [theirs]
    assert [o.id for o in repo.list_organizations()] == [mine, theirs]


def test_find_by_inn_returns_none_when_absent(repo):
    assert repo.find_organization_by_inn("5904188041") is None


def test_find_by_inn_returns_organization(repo):
    oid = repo.save_organization(make_org(inn="5904188041"))
    assert repo.find_organization_by_inn("5904188041").id == oid


def test_get_missing_raises_russian_lookup_error(repo):
    with pytest.raises(LookupError, match="не найден"):
        repo.get_organization(999)


def test_nullable_decimal_none_stays_none(repo):
    oid = repo.save_organization(make_org(vat_rate=None))
    assert repo.get_organization(oid).vat_rate is None
    tid = make_tender(repo, vat_rate=None)
    assert repo.get_tender(tid).vat_rate is None


def test_decimal_scale_survives_round_trip(repo):
    oid = repo.save_organization(make_org(vat_rate=Decimal("0.050")))
    assert str(repo.get_organization(oid).vat_rate) == "0.050"
    tid = make_tender(repo, price_wo_vat=Decimal("5985000.00"))
    price = repo.get_tender(tid).price_wo_vat
    assert isinstance(price, Decimal) and str(price) == "5985000.00"


def test_lot_items_keep_insertion_order(repo):
    tid = make_tender(repo)
    names = [f"Позиция {n}" for n in range(1, 12)]
    for n, name in enumerate(names):
        repo.save_lot_item(LotItem(tender_id=tid, position=n, name=name,
                                   price=Decimal("270.10")))
    got = repo.list_lot_items(tid)
    assert [i.name for i in got] == names
    assert all(i.price == Decimal("270.10") for i in got)


def test_lot_items_sorted_by_position_not_insertion(repo):
    tid = make_tender(repo)
    repo.save_lot_item(LotItem(tender_id=tid, position=2, name="вторая"))
    repo.save_lot_item(LotItem(tender_id=tid, position=1, name="первая"))
    assert [i.name for i in repo.list_lot_items(tid)] == ["первая", "вторая"]


def test_lot_items_are_scoped_to_tender(repo):
    first, second = make_tender(repo), make_tender(repo)
    repo.save_lot_item(LotItem(tender_id=first, name="a"))
    repo.save_lot_item(LotItem(tender_id=second, name="b"))
    assert [i.name for i in repo.list_lot_items(first)] == ["a"]


def test_settings_have_defaults_on_empty_db(repo):
    s = repo.get_settings()
    assert s.insurance_rate == Decimal("0.30")
    assert s.injury_rate == Decimal("0.002")
    assert s.overhead_rate == Decimal("0.05")
    assert s.profit_rate == Decimal("0.10")
    assert s.vat_rate == Decimal("0.05")


def test_settings_save_and_update(repo):
    repo.save_settings(OrgSettings(Decimal("0.302"), Decimal("0.004"),
                                   Decimal("0.07"), Decimal("0.15"),
                                   Decimal("0.20")))
    repo.save_settings(OrgSettings(Decimal("0.30"), Decimal("0.002"),
                                   Decimal("0.05"), Decimal("0.10"),
                                   Decimal("0.22")))
    s = repo.get_settings()
    assert s.vat_rate == Decimal("0.22")
    assert str(s.insurance_rate) == "0.30"
    assert repo.conn.execute("SELECT COUNT(*) FROM org_settings").fetchone()[0] == 1


def test_service_norms_round_trip_and_replace(repo):
    sid = repo.save_service(Service(
        code="СЗЗ", name="Проект СЗЗ",
        norms={"Инженер-эколог": Decimal("120.50"),
               "Руководитель": Decimal("8.00")}))
    got = repo.get_service(sid)
    assert got.norms == {"Инженер-эколог": Decimal("120.50"),
                         "Руководитель": Decimal("8.00")}
    assert str(got.norms["Руководитель"]) == "8.00"
    got.norms = {"Руководитель": Decimal("10")}
    repo.save_service(got)
    assert repo.get_service(sid).norms == {"Руководитель": Decimal("10")}
    assert [s.norms for s in repo.list_services()] == [{"Руководитель": Decimal("10")}]


def test_find_service_by_code(repo):
    sid = repo.save_service(Service(code="НДВ", name="Проект НДВ"))
    assert repo.find_service_by_code("НДВ").id == sid
    assert repo.find_service_by_code("КЭР") is None


def test_customer_profile_json_round_trip(repo):
    cid = repo.save_organization(make_org(role=ROLE_CUSTOMER))
    pid = repo.save_customer_profile(CustomerProfile(
        customer_id=cid,
        documents=["Заявка", "Анкета"],
        file_prefixes={"Заявка": "01"},
        declaration_text="Декларация"))
    got = repo.get_customer_profile(pid)
    assert got.documents == ["Заявка", "Анкета"]
    assert got.file_prefixes == {"Заявка": "01"}
    assert repo.find_customer_profile(cid).id == pid
    assert repo.find_customer_profile(cid + 1) is None


def test_json_is_stored_without_ascii_escaping(repo):
    cid = repo.save_organization(make_org(role=ROLE_CUSTOMER))
    repo.save_customer_profile(CustomerProfile(customer_id=cid, documents=["Заявка"]))
    raw = repo.conn.execute("SELECT documents FROM customer_profile").fetchone()[0]
    assert "Заявка" in raw


def test_tender_deal_answers_round_trip(repo):
    tid = make_tender(repo, deal_answers={"крупная сделка": "нет"},
                      purchase_date="2026-11-05")
    got = repo.get_tender(tid)
    assert got.deal_answers == {"крупная сделка": "нет"}
    assert got.purchase_date == "2026-11-05"
    assert [t.id for t in repo.list_tenders()] == [tid]


def test_employee_and_asset_round_trip(repo):
    eid = repo.save_employee(Employee(full_name="Иванов И. И.",
                                      hourly_rate=Decimal("270.10"),
                                      graduation_year=None))
    emp = repo.get_employee(eid)
    assert emp.hourly_rate == Decimal("270.10") and emp.graduation_year is None
    aid = repo.save_asset(Asset(name="Ноутбук", quantity=3, year=2022))
    asset = repo.get_asset(aid)
    assert (asset.quantity, asset.year) == (3, 2022)
    assert [e.id for e in repo.list_employees()] == [eid]
    assert [a.id for a in repo.list_assets()] == [aid]


def test_experience_round_trip_and_service_filter(repo):
    sid = repo.save_service(Service(code="СЗЗ", name="СЗЗ"))
    a = repo.save_experience(ExperienceRecord(
        customer_name="А", amount=Decimal("5700000.00"), service_id=sid))
    repo.save_experience(ExperienceRecord(customer_name="Б"))
    assert str(repo.get_experience(a).amount) == "5700000.00"
    assert len(repo.list_experience()) == 2
    assert [e.id for e in repo.list_experience(service_id=sid)] == [a]


def test_delete_removes_record(repo):
    oid = repo.save_organization(make_org())
    repo.delete_organization(oid)
    assert repo.list_organizations() == []


def test_delete_tender_cascades_to_lot_items(repo):
    tid = make_tender(repo)
    repo.save_lot_item(LotItem(tender_id=tid, name="a"))
    repo.delete_tender(tid)
    assert repo.conn.execute("SELECT COUNT(*) FROM lot_item").fetchone()[0] == 0


def test_delete_service_cascades_norms_and_nulls_experience(repo):
    sid = repo.save_service(Service(code="СЗЗ", norms={"Эколог": Decimal("1")}))
    eid = repo.save_experience(ExperienceRecord(service_id=sid))
    repo.delete_service(sid)
    assert repo.conn.execute("SELECT COUNT(*) FROM service_norm").fetchone()[0] == 0
    assert repo.get_experience(eid).service_id is None


def test_delete_lot_item_and_customer_profile(repo):
    tid = make_tender(repo)
    lid = repo.save_lot_item(LotItem(tender_id=tid, name="a"))
    repo.delete_lot_item(lid)
    assert repo.list_lot_items(tid) == []
    cid = repo.save_organization(make_org(role=ROLE_CUSTOMER))
    pid = repo.save_customer_profile(CustomerProfile(customer_id=cid))
    repo.delete_customer_profile(pid)
    assert repo.find_customer_profile(cid) is None


def test_saves_are_committed(tmp_path):
    path = tmp_path / "c.db"
    conn = db.connect(path)
    db.migrate(conn)
    Repo(conn).save_organization(make_org(inn="42"))
    conn.close()
    conn2 = db.connect(path)
    assert Repo(conn2).find_organization_by_inn("42") is not None
    conn2.close()


# ---- строки ФОТ тендера (миграция 2) ------------------------------------------

def test_tender_labor_round_trips(repo):
    tid = make_tender(repo)
    rid = repo.save_tender_labor(TenderLaborRow(
        tender_id=tid, position=1, post="Эколог-проектировщик", headcount=4,
        hours=Decimal("1200.00"), rate=Decimal("270.50")))
    (got,) = repo.list_tender_labor(tid)
    assert got.id == rid
    assert (got.tender_id, got.position, got.post, got.headcount) == (
        tid, 1, "Эколог-проектировщик", 4)
    assert isinstance(got.hours, Decimal) and isinstance(got.rate, Decimal)
    # масштаб Decimal не теряется: «1200.00» не становится 1200, «270.50» -- 270.5
    assert str(got.hours) == "1200.00"
    assert str(got.rate) == "270.50"


def test_tender_labor_save_with_id_updates_in_place(repo):
    tid = make_tender(repo)
    rid = repo.save_tender_labor(TenderLaborRow(tender_id=tid, post="a"))
    row = repo.list_tender_labor(tid)[0]
    row.headcount = 3
    assert repo.save_tender_labor(row) == rid
    assert [r.headcount for r in repo.list_tender_labor(tid)] == [3]


def test_tender_labor_cascades_with_tender(repo):
    tid = make_tender(repo)
    repo.save_tender_labor(TenderLaborRow(tender_id=tid, post="a"))
    repo.delete_tender(tid)
    assert repo.conn.execute("SELECT COUNT(*) FROM tender_labor").fetchone()[0] == 0


def test_tender_labor_keeps_position_order(repo):
    tid = make_tender(repo)
    # вставляем не по порядку: порядок задаёт position, а не вставка
    repo.save_tender_labor(TenderLaborRow(tender_id=tid, position=2, post="вторая"))
    repo.save_tender_labor(TenderLaborRow(tender_id=tid, position=1, post="первая"))
    repo.save_tender_labor(TenderLaborRow(tender_id=tid, position=3, post="третья"))
    assert [r.post for r in repo.list_tender_labor(tid)] == [
        "первая", "вторая", "третья"]


def test_tender_labor_is_scoped_to_tender(repo):
    first, second = make_tender(repo), make_tender(repo)
    repo.save_tender_labor(TenderLaborRow(tender_id=first, post="a"))
    repo.save_tender_labor(TenderLaborRow(tender_id=second, post="b"))
    assert [r.post for r in repo.list_tender_labor(first)] == ["a"]


def test_delete_tender_labor_removes_only_that_row(repo):
    tid = make_tender(repo)
    gone = repo.save_tender_labor(TenderLaborRow(tender_id=tid, position=1, post="a"))
    repo.save_tender_labor(TenderLaborRow(tender_id=tid, position=2, post="b"))
    repo.delete_tender_labor(gone)
    assert [r.post for r in repo.list_tender_labor(tid)] == ["b"]


def test_tender_labor_money_columns_are_declared_decimal_text(repo):
    # голый DECIMAL даёт NUMERIC-аффинность и молча теряет масштаб
    cols = {r[1]: r[2] for r in repo.conn.execute("PRAGMA table_info(tender_labor)")}
    assert cols["hours"] == "DECIMAL TEXT"
    assert cols["rate"] == "DECIMAL TEXT"


def test_migrate_twice_keeps_version_and_data(tmp_path):
    conn = db.connect(tmp_path / "t.db")
    db.migrate(conn)
    repo = Repo(conn)
    tid = make_tender(repo, experience_ids=[4, 7])
    repo.save_tender_labor(TenderLaborRow(tender_id=tid, post="a"))
    repo.save_tender_cost_item(TenderCostItem(tender_id=tid, kind="materials"))
    db.migrate(conn)  # повторный вызов ничего не меняет и не стирает
    assert conn.execute("PRAGMA user_version").fetchone()[0] == db.SCHEMA_VERSION == 3
    assert len(repo.list_tender_labor(tid)) == 1
    assert len(repo.list_tender_cost_items(tid)) == 1
    assert repo.get_tender(tid).experience_ids == [4, 7]
    conn.close()


def old_db(path, version):
    """База, созданная до последних миграций: применены только первые version."""
    conn = db.connect(path)
    for number in range(version):
        conn.executescript(db._MIGRATIONS[number])
    conn.execute(f"PRAGMA user_version = {version}")
    conn.execute("INSERT INTO organization (id, role, inn) VALUES (1, 'customer', '1')")
    conn.execute(
        "INSERT INTO tender (id, customer_id, purchase_number, price_wo_vat) "
        "VALUES (1, 1, '42', '5700000.00')")
    conn.execute(
        "INSERT INTO customer_profile (customer_id, documents, declaration_text) "
        "VALUES (1, '[\"Заявка\"]', 'Декларация')")
    conn.commit()
    return conn


def test_version_1_database_upgrades_to_latest_keeping_data(tmp_path):
    conn = old_db(tmp_path / "old1.db", 1)
    db.migrate(conn)
    repo = Repo(conn)
    assert conn.execute("PRAGMA user_version").fetchone()[0] == db.SCHEMA_VERSION
    assert repo.get_tender(1).purchase_number == "42"
    assert repo.list_tender_labor(1) == []
    repo.save_tender_labor(TenderLaborRow(tender_id=1, post="a"))
    assert len(repo.list_tender_labor(1)) == 1
    conn.close()


def test_version_2_database_upgrades_to_3_keeping_data(tmp_path):
    conn = old_db(tmp_path / "old2.db", 2)
    conn.execute("INSERT INTO tender_labor (tender_id, post, hours, rate) "
                 "VALUES (1, 'Эколог', '1200.00', '270.50')")
    conn.commit()
    cols = {r[1] for r in conn.execute("PRAGMA table_info(tender)")}
    assert "experience_ids" not in cols

    db.migrate(conn)
    db.migrate(conn)                      # и повторный вызов безопасен

    repo = Repo(conn)
    assert conn.execute("PRAGMA user_version").fetchone()[0] == 3
    tender = repo.get_tender(1)
    assert tender.purchase_number == "42"
    assert str(tender.price_wo_vat) == "5700000.00"
    assert tender.experience_ids == []            # значение по умолчанию
    profile = repo.find_customer_profile(1)
    assert profile.declaration_text == "Декларация"
    assert profile.documents == ["Заявка"]
    assert profile.layouts == {}
    (row,) = repo.list_tender_labor(1)
    assert str(row.hours) == "1200.00" and str(row.rate) == "270.50"
    repo.save_tender_cost_item(TenderCostItem(tender_id=1, kind="subcontract"))
    assert len(repo.list_tender_cost_items(1)) == 1
    conn.close()


def test_migration_3_is_one_transaction(tmp_path):
    # обрыв посередине не должен оставить половину ALTER: следующий запуск
    # упал бы на «duplicate column». Скрипт сам открывает и закрывает транзакцию.
    script = db._MIGRATIONS[2].strip().upper()
    assert script.startswith("BEGIN") and script.rstrip().endswith("COMMIT;")


# ---- профиль: раскладки, тендер: отбор опыта ----------------------------------

def test_customer_profile_layouts_round_trip(repo):
    cid = repo.save_organization(make_org(role=ROLE_CUSTOMER))
    pid = repo.save_customer_profile(CustomerProfile(
        customer_id=cid,
        layouts={"Справка о МТР": "assets_rnpk.json",
                 "Справка о кадрах": "employees_koryazhma.json"}))
    got = repo.get_customer_profile(pid)
    assert got.layouts == {"Справка о МТР": "assets_rnpk.json",
                           "Справка о кадрах": "employees_koryazhma.json"}
    assert list(got.layouts) == ["Справка о МТР", "Справка о кадрах"]   # порядок
    assert CustomerProfile().layouts == {}


def test_tender_experience_ids_round_trip(repo):
    tid = make_tender(repo, experience_ids=[7, 3, 12])
    assert repo.get_tender(tid).experience_ids == [7, 3, 12]   # порядок отбора
    assert make_tender(repo) and Tender().experience_ids == []


# ---- материалы и субподряд тендера (миграция 3) --------------------------------

def test_tender_cost_item_round_trips_with_scale(repo):
    tid = make_tender(repo)
    cid = repo.save_tender_cost_item(TenderCostItem(
        tender_id=tid, kind="materials", position=1,
        name="инвентарь, оборудование", unit="усл.ед.",
        quantity=Decimal("1.000"), price=Decimal("24000.00"), note="прим."))
    (got,) = repo.list_tender_cost_items(tid)
    assert got.id == cid
    assert (got.kind, got.position, got.name, got.unit, got.note) == (
        "materials", 1, "инвентарь, оборудование", "усл.ед.", "прим.")
    assert isinstance(got.quantity, Decimal) and isinstance(got.price, Decimal)
    assert str(got.quantity) == "1.000" and str(got.price) == "24000.00"


def test_tender_cost_item_columns_are_declared_decimal_text(repo):
    cols = {r[1]: r[2] for r in repo.conn.execute("PRAGMA table_info(tender_cost_item)")}
    assert cols["quantity"] == "DECIMAL TEXT" and cols["price"] == "DECIMAL TEXT"


def test_tender_cost_item_keeps_position_order_and_filters_by_kind(repo):
    tid = make_tender(repo)
    for kind, pos, name in [("subcontract", 2, "б"), ("materials", 1, "в"),
                            ("subcontract", 1, "а")]:
        repo.save_tender_cost_item(TenderCostItem(
            tender_id=tid, kind=kind, position=pos, name=name))
    assert [i.name for i in repo.list_tender_cost_items(tid, "subcontract")] == ["а", "б"]
    assert [i.name for i in repo.list_tender_cost_items(tid, "materials")] == ["в"]
    assert len(repo.list_tender_cost_items(tid)) == 3


def test_tender_cost_item_kind_is_restricted(repo):
    tid = make_tender(repo)
    with pytest.raises(sqlite3.IntegrityError):
        repo.save_tender_cost_item(TenderCostItem(tender_id=tid, kind="прочее"))


def test_tender_cost_item_cascades_and_deletes(repo):
    tid = make_tender(repo)
    gone = repo.save_tender_cost_item(TenderCostItem(tender_id=tid, kind="materials"))
    repo.save_tender_cost_item(TenderCostItem(tender_id=tid, kind="materials", name="b"))
    repo.delete_tender_cost_item(gone)
    assert [i.name for i in repo.list_tender_cost_items(tid)] == ["b"]
    repo.delete_tender(tid)
    assert repo.conn.execute("SELECT COUNT(*) FROM tender_cost_item").fetchone()[0] == 0
