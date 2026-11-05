from datetime import date
from decimal import Decimal

import pytest

from app import db
from app.calc import LaborLine, estimate
from app.models import (
    ROLE_CUSTOMER, ROLE_PARTICIPANT, Asset, CustomerProfile, Employee,
    ExperienceRecord, LotItem, Organization, OrgSettings, Tender, TenderLaborRow,
)
from app.render import context as ctx_module
from app.render.context import build_context
from app.render.layouts import Column, TableLayout
from app.repo import Repo

NBSP = " "

# Документы профилей: по ним context выбирает раскладку таблиц.
DOC_EQUIPMENT = "Справка о наличии оборудования"   # раскладка «Коряжма»
DOC_MTR = "Справка о МТР"                          # раскладка «РНПК»

TOP_LEVEL_KEYS = {
    "org", "customer", "tender", "items", "estimate", "labor", "employees",
    "assets", "experience", "declaration_text", "conditions", "today", "money",
}

# КП Арктикгаза: «Эколог-проектировщик, 4 чел.» при одном экологе в штате
ARKTIK = [("Эколог-проектировщик", 4, "1200", "270"),
          ("Начальник отдела проектирования", 1, "360", "420")]


@pytest.fixture
def repo(tmp_path):
    conn = db.connect(tmp_path / "t.db")
    db.migrate(conn)
    yield Repo(conn)
    conn.close()


def add_labor(repo, tender_id, rows=ARKTIK):
    for n, (post, people, hours, rate) in enumerate(rows, start=1):
        repo.save_tender_labor(TenderLaborRow(
            tender_id=tender_id, position=n, post=post, headcount=people,
            hours=Decimal(hours), rate=Decimal(rate)))


def add_tender(repo, customer_id, participant_id, **kw) -> int:
    kw.setdefault("purchase_number", "0556300000225000001")
    return repo.save_tender(Tender(
        customer_id=customer_id, participant_id=participant_id, **kw))


@pytest.fixture
def repo_with_koryazhma(repo):
    """Тендер Коряжмы: участник, заказчик с профилем, лот, справочники."""
    me = repo.save_organization(Organization(
        role=ROLE_PARTICIPANT, inn="5904188041", short_name="ООО «Техносфера»",
        head_name="Рабичко Елена Михайловна"))
    customer = repo.save_organization(Organization(
        role=ROLE_CUSTOMER, inn="2904016130", short_name="МУП «Полигон»"))
    repo.save_customer_profile(CustomerProfile(
        customer_id=customer, documents=[DOC_EQUIPMENT, "Справка об опыте"],
        declaration_text="Декларация Коряжмы"))
    tid = add_tender(
        repo, customer, me, subject="Разработка КЭР",
        price_wo_vat=Decimal("5700000.00"),
        deal_answers={"Аванс": "подтверждаю", "Неустойка": "не подтверждаю"})
    repo.save_lot_item(LotItem(tender_id=tid, position=2, name="Вторая",
                               price=Decimal("1234567.80")))
    repo.save_lot_item(LotItem(tender_id=tid, position=1, name="Первая",
                               price=Decimal("5700000.00")))
    add_labor(repo, tid)
    repo.save_employee(Employee(full_name="Рабичко Е. М.", position="Директор",
                                education="Томский политех", experience_years=30))
    repo.save_employee(Employee(full_name="Мельникова А. А.", position="Эколог"))
    repo.save_asset(Asset(name="Компьютер", maker_model="Intel i3", quantity=3,
                          location="Пермь", ownership_basis="Собственность",
                          ownership_doc="Свидетельство 59-БД"))
    repo.save_experience(ExperienceRecord(
        customer_name="Коряжма", contract_number="1", subject="КЭР",
        amount=Decimal("5700000.00")))
    repo.save_experience(ExperienceRecord(
        customer_name="Полигон", contract_number="2", subject="СЗЗ",
        amount=Decimal("1234567.80")))
    repo.save_experience(ExperienceRecord(
        customer_name="Лишний", contract_number="3", subject="НДВ",
        amount=Decimal("1.00")))
    repo.tender_id = tid
    return repo


def ctx_of(repo, **kw):
    return build_context(repo.tender_id, repo, **kw)


# ---- ключи --------------------------------------------------------------------

def test_context_has_all_documented_keys(repo_with_koryazhma):
    ctx = ctx_of(repo_with_koryazhma)
    assert set(ctx) == TOP_LEVEL_KEYS
    assert callable(ctx["money"])
    assert ctx["org"]["short_name"] == "ООО «Техносфера»"
    assert ctx["org"]["head_name"] == "Рабичко Елена Михайловна"
    assert ctx["customer"]["short_name"] == "МУП «Полигон»"
    assert ctx["tender"]["subject"] == "Разработка КЭР"
    assert ctx["tender"]["price_wo_vat"] == Decimal("5700000.00")
    for table in ("employees", "assets", "experience"):
        assert set(ctx[table]) == {"headers", "rows"}


def test_sparse_tender_still_has_every_key(repo):
    # ни участника, ни заказчика, ни позиций, ни строк ФОТ
    tid = repo.save_tender(Tender())
    ctx = build_context(tid, repo)
    assert set(ctx) == TOP_LEVEL_KEYS
    assert ctx["org"]["short_name"] == "" and ctx["customer"]["inn"] == ""
    assert ctx["items"] == [] and ctx["labor"] == []
    assert ctx["conditions"] == {} and ctx["declaration_text"] == ""
    assert ctx["estimate"]["labor"] == Decimal("0.00")
    for table in ("employees", "assets", "experience"):
        assert ctx[table]["rows"] == []
        assert ctx[table]["headers"]        # шапка есть и без данных


def test_today_is_a_russian_formatted_string(repo_with_koryazhma, monkeypatch):
    monkeypatch.setattr(ctx_module, "_today", lambda: date(2026, 11, 5))
    assert ctx_of(repo_with_koryazhma)["today"] == "05.11.2026"


def test_unknown_tender_id_raises_with_russian_message(repo):
    with pytest.raises(LookupError, match="Тендер.*не найден"):
        build_context(999, repo)


# ---- деньги -------------------------------------------------------------------

def test_money_formats_russian(repo_with_koryazhma):
    money = ctx_of(repo_with_koryazhma)["money"]
    assert money(Decimal("5985000")) == f"5{NBSP}985{NBSP}000,00"
    assert money(Decimal("285000.00")) == f"285{NBSP}000,00"
    assert money(Decimal("0")) == "0,00"


def test_money_uses_non_breaking_space_not_plain_space(repo_with_koryazhma):
    out = ctx_of(repo_with_koryazhma)["money"](Decimal("5985000"))
    assert " " not in out                 # обычный пробел позволил бы разрыв строки
    assert out.count(NBSP) == 2           # две границы разрядов в 5 985 000


def test_money_rounds_half_up_to_kopecks_and_keeps_sign(repo_with_koryazhma):
    money = ctx_of(repo_with_koryazhma)["money"]
    assert money(Decimal("1.005")) == "1,01"
    assert money(Decimal("-1234.5")) == f"-1{NBSP}234,50"
    assert money(Decimal("-0.001")) == "0,00"     # без «-0,00»
    assert money(Decimal("999.999")) == f"1{NBSP}000,00"


def test_money_refuses_float(repo_with_koryazhma):
    with pytest.raises(TypeError):
        ctx_of(repo_with_koryazhma)["money"](5985000.0)


# ---- ФОТ и калькуляция --------------------------------------------------------

def test_labor_rows_keep_position_order_and_decimals(repo_with_koryazhma):
    labor = ctx_of(repo_with_koryazhma)["labor"]
    assert [r["post"] for r in labor] == [a[0] for a in ARKTIK]
    assert [r["number"] for r in labor] == [1, 2]
    assert labor[0]["headcount"] == 4
    assert labor[0]["hours"] == Decimal("1200") and labor[0]["rate"] == Decimal("270")


def test_estimate_is_built_from_stored_labor_rows(repo_with_koryazhma):
    est = ctx_of(repo_with_koryazhma)["estimate"]
    assert est["labor"] == Decimal("1447200.00")      # 4*1200*270 + 1*360*420
    assert est["insurance"] == Decimal("437054.40")   # 30,2 %
    assert est["overhead"] == Decimal("72360.00")
    # материалов и субподряда у тендера в базе нет: калькуляция их не знает
    assert est["materials"] == Decimal("0.00")
    assert est["subcontract"] == Decimal("0.00")


def test_estimate_matches_calculator_for_same_inputs(repo_with_koryazhma):
    repo = repo_with_koryazhma
    lines = [LaborLine(p, n, Decimal(h), Decimal(r)) for p, n, h, r in ARKTIK]
    expected = estimate(lines, Decimal("0"), Decimal("0"), repo.get_settings())
    est = ctx_of(repo)["estimate"]
    assert set(est) == {"materials", "labor", "insurance", "subcontract",
                        "overhead", "profit", "total_wo_vat", "vat",
                        "total_with_vat"}
    assert est == {f: getattr(expected, f) for f in est}


def test_vat_rate_falls_back_tender_then_org_then_settings(repo):
    me = repo.save_organization(Organization(role=ROLE_PARTICIPANT,
                                             vat_rate=Decimal("0.06")))
    no_vat_org = repo.save_organization(Organization(role=ROLE_PARTICIPANT))
    repo.save_settings(OrgSettings(Decimal("0.30"), Decimal("0.002"),
                                   Decimal("0.05"), Decimal("0.10"),
                                   Decimal("0.05")))

    def vat_of(participant, tender_vat):
        tid = add_tender(repo, None, participant, vat_rate=tender_vat)
        add_labor(repo, tid)
        ctx = build_context(tid, repo)
        # ставка в калькуляции -- та же, что показана в тендере
        net = ctx["estimate"]["total_wo_vat"]
        assert ctx["estimate"]["vat"] == (net * ctx["tender"]["vat_rate"]).quantize(
            Decimal("0.01"))
        return ctx["tender"]["vat_rate"]

    assert vat_of(me, Decimal("0.07")) == Decimal("0.07")     # тендер побеждает
    assert vat_of(me, None) == Decimal("0.06")                # организация
    assert vat_of(no_vat_org, None) == Decimal("0.05")        # настройки
    assert vat_of(None, None) == Decimal("0.05")              # участника нет


def test_zero_vat_on_tender_means_exempt_not_fall_through(repo):
    # 0 -- это «освобождён от НДС», а не «не задано»
    me = repo.save_organization(Organization(role=ROLE_PARTICIPANT,
                                             vat_rate=Decimal("0.06")))
    tid = add_tender(repo, None, me, vat_rate=Decimal("0"))
    add_labor(repo, tid)
    ctx = build_context(tid, repo)
    assert ctx["tender"]["vat_rate"] == Decimal("0")
    assert ctx["estimate"]["vat"] == Decimal("0.00")
    assert ctx["estimate"]["total_with_vat"] == ctx["estimate"]["total_wo_vat"]


def test_vat_percent_text(repo):
    me = repo.save_organization(Organization(role=ROLE_PARTICIPANT))
    for rate, text in [("0.05", "5"), ("0.055", "5,5"), ("0.20", "20"), ("0", "0")]:
        tid = add_tender(repo, None, me, vat_rate=Decimal(rate))
        assert build_context(tid, repo)["tender"]["vat_percent"] == text


# ---- позиции лота -------------------------------------------------------------

def test_items_are_ordered_numbered_and_priced_through_money(repo_with_koryazhma):
    items = ctx_of(repo_with_koryazhma)["items"]
    assert [i["name"] for i in items] == ["Первая", "Вторая"]
    assert [i["number"] for i in items] == [1, 2]
    assert items[0]["price"] == f"5{NBSP}700{NBSP}000,00"
    assert items[1]["price"] == f"1{NBSP}234{NBSP}567,80"
    # число для ячейки Excel остаётся числом
    assert items[1]["price_value"] == Decimal("1234567.80")


# ---- таблицы справок ----------------------------------------------------------

def test_assets_table_follows_customer_profile_layout(repo_with_koryazhma):
    repo = repo_with_koryazhma
    kor = ctx_of(repo)["assets"]
    # Коряжма: без нумерации, наименование и марка в разных колонках
    assert len(kor["headers"]) == 7 and kor["headers"][0] == "Наименование"
    assert kor["rows"][0][0:3] == ["Компьютер", "Intel i3", "3"]

    rnpk_customer = repo.save_organization(Organization(
        role=ROLE_CUSTOMER, inn="7708000000", short_name="АО «РНПК»"))
    repo.save_customer_profile(CustomerProfile(
        customer_id=rnpk_customer, documents=[DOC_MTR]))
    tid = add_tender(repo, rnpk_customer, None)
    rnpk = build_context(tid, repo)["assets"]
    assert rnpk["headers"][0] == "№ п/п"
    assert rnpk["rows"][0][0:2] == ["1", "Компьютер Intel i3"]


def test_employees_table_uses_layout_headers_and_all_employees(repo_with_koryazhma):
    emp = ctx_of(repo_with_koryazhma)["employees"]
    assert emp["headers"][:3] == ["№ п/п", "Должность", "ФИО"]
    assert [r[2] for r in emp["rows"]] == ["Рабичко Е. М.", "Мельникова А. А."]
    assert emp["rows"][0][4] == "30"


def test_declaration_text_comes_from_customer_profile(repo_with_koryazhma):
    assert ctx_of(repo_with_koryazhma)["declaration_text"] == "Декларация Коряжмы"


def test_conditions_are_the_tender_deal_answers(repo_with_koryazhma):
    assert ctx_of(repo_with_koryazhma)["conditions"] == {
        "Аванс": "подтверждаю", "Неустойка": "не подтверждаю"}


def test_experience_includes_only_selected_records(repo_with_koryazhma):
    repo = repo_with_koryazhma
    ids = [e.id for e in repo.list_experience()]
    rows = ctx_of(repo, experience_ids=[ids[0], ids[1]])["experience"]["rows"]
    assert [r[2] for r in rows] == ["1", "2"]          # «Лишний» (3) не попал
    assert ctx_of(repo)["experience"]["rows"] == []    # без отбора -- пусто
    # порядок -- как отобрал пользователь
    swapped = ctx_of(repo, experience_ids=[ids[1], ids[0]])["experience"]["rows"]
    assert [r[2] for r in swapped] == ["2", "1"]


def test_unknown_experience_id_raises_with_russian_message(repo_with_koryazhma):
    with pytest.raises(LookupError, match="не найден"):
        ctx_of(repo_with_koryazhma, experience_ids=[999])


def test_experience_amount_is_formatted_by_money(repo_with_koryazhma):
    repo = repo_with_koryazhma
    first = repo.list_experience()[0]
    row = ctx_of(repo, experience_ids=[first.id])["experience"]["rows"][0]
    assert f"5{NBSP}700{NBSP}000,00" in row
    assert "5700000.00" not in row                     # сырого str(Decimal) нет
    # справочник при этом не тронут
    assert repo.get_experience(first.id).amount == Decimal("5700000.00")


def test_any_decimal_field_in_a_table_goes_through_money():
    # формат денег не зависит от того, какую колонку выбрала раскладка
    layout = TableLayout(source="employees", numbered=False, columns=[
        Column("Ставка", ["hourly_rate"]), Column("Стаж", ["experience_years"])])
    emp = Employee(full_name="А", hourly_rate=Decimal("1270.1"), experience_years=5)
    table = ctx_module._table(layout, [emp])
    assert table == {"headers": ["Ставка", "Стаж"],
                     "rows": [[f"1{NBSP}270,10", "5"]]}
    assert emp.hourly_rate == Decimal("1270.1")        # запись не изменена


def test_missing_customer_profile_falls_back_to_default_layouts(repo):
    customer = repo.save_organization(Organization(role=ROLE_CUSTOMER))
    tid = add_tender(repo, customer, None)
    repo.save_asset(Asset(name="Принтер", quantity=1))
    eid = repo.save_experience(ExperienceRecord(customer_name="З", amount=Decimal("1")))
    ctx = build_context(tid, repo, experience_ids=[eid])
    assert ctx["declaration_text"] == ""
    assert ctx["assets"]["headers"][0] == "Наименование"       # «Коряжма»
    assert ctx["assets"]["rows"][0][0] == "Принтер"
    assert ctx["employees"]["headers"][1] == "Должность"
    assert ctx["experience"]["headers"][0] == "№ п/п"
    assert ctx["experience"]["rows"][0][1] == "З"


def test_profile_without_known_documents_uses_default_layouts(repo):
    customer = repo.save_organization(Organization(role=ROLE_CUSTOMER))
    repo.save_customer_profile(CustomerProfile(
        customer_id=customer, documents=["Декларация"]))
    tid = add_tender(repo, customer, None)
    assert build_context(tid, repo)["assets"]["headers"][0] == "Наименование"
