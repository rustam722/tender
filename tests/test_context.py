from datetime import date
from decimal import Decimal

import pytest

from app import db
from app.calc import LaborLine, estimate
from app.models import (
    ROLE_CUSTOMER, ROLE_PARTICIPANT, Asset, CustomerProfile, Employee,
    ExperienceRecord, LotItem, Organization, OrgSettings, Tender, TenderCostItem,
    TenderLaborRow,
)
from app.render import context as ctx_module
from app.render.context import build_context
from app.render.layouts import Column, TableLayout
from app.repo import Repo

NBSP = " "

# Документы профиля; файл раскладки профиль назначает им явно (CustomerProfile.layouts).
DOC_EQUIPMENT = "Справка о наличии оборудования"
DOC_MTR = "Справка о МТР"

TOP_LEVEL_KEYS = {
    "org", "customer", "tender", "items", "estimate", "labor", "employees",
    "assets", "experience", "declaration_text", "conditions", "today", "money",
    "materials", "subcontract", "date_ru", "date_long",
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
        layouts={DOC_EQUIPMENT: "assets_koryazhma.json"},
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
    first = repo.save_experience(ExperienceRecord(
        customer_name="Коряжма", contract_number="1", subject="КЭР",
        amount=Decimal("5700000.00")))
    second = repo.save_experience(ExperienceRecord(
        customer_name="Полигон", contract_number="2", subject="СЗЗ",
        amount=Decimal("1234567.80")))
    repo.save_experience(ExperienceRecord(
        customer_name="Лишний", contract_number="3", subject="НДВ",
        amount=Decimal("1.00")))
    tender = repo.get_tender(tid)
    tender.experience_ids = [first, second]      # отбор хранится при тендере
    repo.save_tender(tender)
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
    assert ctx["materials"] == [] and ctx["subcontract"] == []
    assert ctx["conditions"] == {} and ctx["declaration_text"] == ""
    assert ctx["estimate"]["labor"] == Decimal("0.00")
    assert ctx["date_ru"]("") == "" and ctx["date_long"]("") == ""
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
    # у этого тендера нет строк материалов и субподряда
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


# ---- материалы и субподряд ----------------------------------------------------

def add_arktik_costs(repo, tid):
    rows = [("materials", "инвентарь, оборудование", "24000"),
            ("materials", "расходные материалы", "17686"),
            ("subcontract", "проведение исследований и измерений", "1500000"),
            ("subcontract", "фоновые концентрации и климатические характеристики", "70000"),
            ("subcontract", "экспертиза проектной документации", "50000")]
    for n, (kind, name, price) in enumerate(rows, start=1):
        repo.save_tender_cost_item(TenderCostItem(
            tender_id=tid, kind=kind, position=n, name=name,
            unit="усл.ед." if kind == "materials" else "шт.",
            quantity=Decimal("1"), price=Decimal(price)))


def test_estimate_sums_cost_items_by_kind(repo_with_koryazhma):
    repo = repo_with_koryazhma
    add_arktik_costs(repo, repo.tender_id)
    est = ctx_of(repo)["estimate"]
    assert est["materials"] == Decimal("41686.00")        # 24 000 + 17 686
    assert est["subcontract"] == Decimal("1620000.00")    # 1 500 000 + 70 000 + 50 000
    lines = [LaborLine(p, n, Decimal(h), Decimal(r)) for p, n, h, r in ARKTIK]
    expected = estimate(lines, Decimal("41686"), Decimal("1620000"),
                        repo.get_settings())
    assert est == {f: getattr(expected, f) for f in est}


def test_cost_item_quantity_multiplies_price(repo_with_koryazhma):
    repo = repo_with_koryazhma
    repo.save_tender_cost_item(TenderCostItem(
        tender_id=repo.tender_id, kind="materials", name="бумага",
        quantity=Decimal("2.5"), price=Decimal("100.00")))
    ctx = ctx_of(repo)
    assert ctx["estimate"]["materials"] == Decimal("250.00")
    assert ctx["materials"][0]["amount_value"] == Decimal("250.000")


def test_cost_items_are_exposed_as_two_tables(repo_with_koryazhma):
    repo = repo_with_koryazhma
    add_arktik_costs(repo, repo.tender_id)
    ctx = ctx_of(repo)
    assert [r["name"] for r in ctx["materials"]] == [
        "инвентарь, оборудование", "расходные материалы"]
    assert len(ctx["subcontract"]) == 3
    first = ctx["materials"][0]
    assert first["number"] == 1 and first["unit"] == "усл.ед."
    assert first["quantity"] == Decimal("1")
    assert first["price"] == f"24{NBSP}000,00"           # через money
    assert first["price_value"] == Decimal("24000")      # число для Excel
    assert first["amount"] == f"24{NBSP}000,00"
    assert first["amount_value"] == Decimal("24000")
    assert ctx["materials"][1]["number"] == 2            # нумерация своя у статьи
    assert ctx["subcontract"][2]["number"] == 3


def test_cost_items_of_other_tenders_do_not_leak(repo_with_koryazhma):
    repo = repo_with_koryazhma
    other = add_tender(repo, None, None)
    repo.save_tender_cost_item(TenderCostItem(
        tender_id=other, kind="subcontract", price=Decimal("999")))
    ctx = ctx_of(repo)
    assert ctx["subcontract"] == []
    assert ctx["estimate"]["subcontract"] == Decimal("0.00")


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
        customer_id=rnpk_customer, documents=[DOC_MTR],
        layouts={DOC_MTR: "assets_rnpk.json"}))
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


# ---- опыт ---------------------------------------------------------------------

def test_experience_includes_only_selected_records(repo_with_koryazhma):
    rows = ctx_of(repo_with_koryazhma)["experience"]["rows"]   # отбор тендера
    assert [r[2] for r in rows] == ["1", "2"]          # «Лишний» (3) не попал


def test_experience_selection_is_stored_so_rebuild_is_identical(repo_with_koryazhma):
    first = ctx_of(repo_with_koryazhma)["experience"]
    assert ctx_of(repo_with_koryazhma)["experience"] == first


def test_passed_experience_ids_win_over_the_stored_selection(repo_with_koryazhma):
    repo = repo_with_koryazhma
    ids = [e.id for e in repo.list_experience()]
    rows = ctx_of(repo, experience_ids=[ids[2], ids[0]])["experience"]["rows"]
    assert [r[2] for r in rows] == ["3", "1"]          # порядок -- как передан
    assert ctx_of(repo, experience_ids=[])["experience"]["rows"] == []
    # предпросмотр ничего не сохраняет
    assert repo.get_tender(repo.tender_id).experience_ids == ids[:2]


def test_tender_without_selection_has_empty_experience(repo):
    tid = add_tender(repo, None, None)
    repo.save_experience(ExperienceRecord(customer_name="З"))
    assert build_context(tid, repo)["experience"]["rows"] == []


def test_unknown_experience_id_raises_with_russian_message(repo_with_koryazhma):
    repo = repo_with_koryazhma
    with pytest.raises(LookupError, match="не найден"):
        ctx_of(repo, experience_ids=[999])
    # то же для устаревшего отбора, сохранённого при тендере
    tender = repo.get_tender(repo.tender_id)
    tender.experience_ids = [999]
    repo.save_tender(tender)
    with pytest.raises(LookupError, match="не найден"):
        ctx_of(repo)


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


# ---- раскладки ----------------------------------------------------------------

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


def test_profile_with_empty_layouts_uses_default_layouts(repo):
    customer = repo.save_organization(Organization(role=ROLE_CUSTOMER))
    repo.save_customer_profile(CustomerProfile(
        customer_id=customer, documents=["Декларация"]))
    tid = add_tender(repo, customer, None)
    assert build_context(tid, repo)["assets"]["headers"][0] == "Наименование"


def test_profile_layout_source_is_read_from_the_layout_file(repo):
    # источник (assets/employees/...) определяет сам файл раскладки, поэтому
    # документ в профиле может называться как угодно
    customer = repo.save_organization(Organization(role=ROLE_CUSTOMER))
    repo.save_customer_profile(CustomerProfile(
        customer_id=customer,
        layouts={"Любое название": "assets_rnpk.json",
                 "Другое название": "assets_koryazhma.json"}))
    ctx = build_context(add_tender(repo, customer, None), repo)
    assert ctx["assets"]["headers"][0] == "№ п/п"      # первая в профиле побеждает
    assert ctx["employees"]["headers"][1] == "Должность"   # остальное -- умолчания


def test_new_customer_layout_needs_only_a_layout_file(repo, tmp_path, monkeypatch):
    # обещание проекта: новый заказчик -- описание колонок, без правки кода
    (tmp_path / "assets_novyj.json").write_text(
        '{"source": "assets", "numbered": false, "columns": ['
        '{"header": "Что", "parts": ["name"]}, {"header": "Где", "parts": ["location"]}]}',
        encoding="utf-8")
    monkeypatch.setattr(ctx_module, "LAYOUTS_DIR", tmp_path)
    customer = repo.save_organization(Organization(role=ROLE_CUSTOMER))
    repo.save_customer_profile(CustomerProfile(
        customer_id=customer, layouts={"Перечень техники": "assets_novyj.json"}))
    repo.save_asset(Asset(name="Принтер", location="Пермь"))
    # остальные источники берут умолчания из той же папки: кладём и их
    for name in ctx_module._DEFAULT_LAYOUTS.values():
        (tmp_path / name).write_text(
            (ctx_module.BASE_DIR / "layouts" / name).read_text(encoding="utf-8"),
            encoding="utf-8")
    ctx = build_context(add_tender(repo, customer, None), repo)
    assert ctx["assets"] == {"headers": ["Что", "Где"], "rows": [["Принтер", "Пермь"]]}


# ---- даты ---------------------------------------------------------------------

def test_date_ru_and_date_long(repo_with_koryazhma):
    ctx = ctx_of(repo_with_koryazhma)
    assert ctx["date_ru"]("2025-10-24") == "24.10.2025"
    assert ctx["date_long"]("2025-10-24") == "«24» октября 2025 г."
    assert ctx["date_ru"]("2025-03-05") == "05.03.2025"
    assert ctx["date_long"]("2025-03-05") == "«5» марта 2025 г."


def test_date_long_uses_genitive_month_names(repo_with_koryazhma):
    date_long = ctx_of(repo_with_koryazhma)["date_long"]
    months = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
              "августа", "сентября", "октября", "ноября", "декабря"]
    for number, name in enumerate(months, start=1):
        assert date_long(f"2025-{number:02d}-10") == f"«10» {name} 2025 г."


def test_empty_date_gives_empty_string(repo_with_koryazhma):
    ctx = ctx_of(repo_with_koryazhma)
    assert ctx["date_ru"]("") == "" and ctx["date_long"]("") == ""


def test_malformed_date_raises_with_russian_message(repo_with_koryazhma):
    ctx = ctx_of(repo_with_koryazhma)
    for bad in ("24.10.2025", "вчера", "2025-13-01"):
        with pytest.raises(ValueError, match="ГГГГ-ММ-ДД"):
            ctx["date_ru"](bad)
        with pytest.raises(ValueError, match="ГГГГ-ММ-ДД"):
            ctx["date_long"](bad)


def test_tender_dates_pass_through_the_date_helpers(repo):
    tid = add_tender(repo, None, None, purchase_date="2025-10-24")
    ctx = build_context(tid, repo)
    assert ctx["date_ru"](ctx["tender"]["purchase_date"]) == "24.10.2025"
