"""Заполнение книги заказчика.

Каждая проверка читает результат из готового файла через openpyxl
(load_workbook(data_only=False)), а не из словаря на входе.

Две книги: маленькая синтетическая форма, которую строит тест (на ней
проверяются защита листа и отказы), и настоящая форма ТКП Роснефти --
копия в tests/fixtures, на ней проверяется карта cellmaps/tkp_rosneft_murmansk.json.
"""
import json
import warnings
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Protection

from app import db
from app.models import (
    ROLE_CUSTOMER, ROLE_PARTICIPANT, CustomerProfile, LotItem, Organization,
    Tender,
)
from app.render.context import build_context
from app.render.xlsx_render import (
    RATE_FIELDS, CellMap, CellMapError, RowBlock, fill_workbook, load_cellmap,
    read_rates,
)
from app.repo import Repo

ROOT = Path(__file__).resolve().parent.parent
REAL_FORM = Path(__file__).parent / "fixtures" / "forma_tkp_rosneft_murmansk.xlsx"
REAL_MAP = ROOT / "cellmaps" / "tkp_rosneft_murmansk.json"

# Семь условий сделки -- дословно колонка B формы (B23:B29): по ним
# карта ищет ответ в context["conditions"].
CONDITIONS = [
    "Цена и условия:",
    "Форма, срок и порядок оплаты:",
    "Условия и место оказания услуг:",
    "Срок оказания услуг:",
    "Срок действия коммерческого предложения:",
    "Срок действия договора:",
    "Возможность привлечения субподрядчиков/субисполнителей",
]


# ---- синтетическая форма ------------------------------------------------------

def make_form(path, *, protected=False, positions=11):
    """Форма того же устройства, что у Роснефти: клетки участника C/D 9..,
    формулы E и итоги, заготовки D8, E23.., B31."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"
    ws["A3"] = "Участник заполняет поля, отмеченные желтой заливкой"
    ws["D8"] = "Ставка НДС: ____*%"
    last = 8 + positions
    for n in range(1, positions + 1):
        row = 8 + n
        ws.cell(row, 1, n)
        ws.cell(row, 2, f"Объект {n} (текст заказчика)")
        ws.cell(row, 5, f"=C{row}+D{row}")
    total = last + 1
    ws.cell(total, 2, "Итого в целом по КП:")
    ws.cell(total, 3, f"=SUM(C9:C{last})")
    ws.cell(total, 5, f"=SUM(E9:E{last})")
    ws["B23"], ws["C23"], ws["E23"] = (
        "Цена и условия:", "Условие заказчика", "подтверждаю/не подтверждаю")
    ws["B31"] = "[Руководитель/Должность/ФИО]"
    ws["C31"] = "[Подпись/печать]"
    ws.merge_cells("C23:D23")
    for row in range(9, last + 1):
        for col in (3, 4):
            ws.cell(row, col).protection = Protection(locked=False)
    for address in ("D8", "E23", "B31"):
        ws[address].protection = Protection(locked=False)
    if protected:
        ws.protection.sheet = True
        ws.protection.password = "customer"
    wb.save(path)
    return path


def basic_map():
    return CellMap(
        sheet="Лист1",
        cells={"D8": "Ставка НДС: {tender.vat_percent}%",
               "E23": "{conditions.Цена и условия:}",
               "B31": "{org.head_position} {org.head_name}"},
        rows=[RowBlock(first_row=9, last_row=19, source="items",
                       columns={"C": "{row.price_value}",
                                "D": "{row.vat_value}"})])


def make_items(count, price="100000.00", vat="5000.00"):
    return [{"number": n, "position": n, "name": f"Наш объект {n}",
             "price": "x", "price_value": Decimal(price) + n,
             "vat": "x", "vat_value": Decimal(vat) + n}
            for n in range(1, count + 1)]


def make_context(count=11, **kw):
    context = {
        "items": make_items(count),
        "tender": {"vat_percent": "5", "vat_rate": Decimal("0.05")},
        "conditions": {c: "подтверждаю" for c in CONDITIONS},
        "org": {"head_name": "Рабичко Е. М.", "head_position": "Директор"},
    }
    context.update(kw)
    return context


def cells_of(path, sheet="Лист1", data_only=False):
    ws = load_workbook(path, data_only=data_only)[sheet]
    return {c.coordinate: c.value for row in ws.iter_rows() for c in row
            if c.value is not None}


@pytest.fixture
def form(tmp_path):
    return make_form(tmp_path / "form.xlsx")


# ---- обязательные проверки из брифа -------------------------------------------

def test_writes_scalar_cell(form, tmp_path):
    out = fill_workbook(form, basic_map(), make_context(), tmp_path / "out.xlsx")
    cells = cells_of(out)
    assert cells["D8"] == "Ставка НДС: 5%"
    assert cells["B31"] == "Директор Рабичко Е. М."
    assert cells["E23"] == "подтверждаю"


def test_fills_eleven_lot_rows(form, tmp_path):
    out = fill_workbook(form, basic_map(), make_context(11), tmp_path / "out.xlsx")
    cells = cells_of(out)
    for n in range(1, 12):
        row = 8 + n
        assert cells[f"C{row}"] == 100000 + n
        assert cells[f"D{row}"] == 5000 + n


def test_preserves_existing_formulas(form, tmp_path):
    out = fill_workbook(form, basic_map(), make_context(), tmp_path / "out.xlsx")
    cells = cells_of(out)
    for row in range(9, 20):
        assert cells[f"E{row}"] == f"=C{row}+D{row}"       # формула, не число
    assert cells["C20"] == "=SUM(C9:C19)"
    assert cells["E20"] == "=SUM(E9:E19)"


def test_preserves_sheet_protection(tmp_path):
    protected = make_form(tmp_path / "protected.xlsx", protected=True)
    out = fill_workbook(protected, basic_map(), make_context(),
                        tmp_path / "out.xlsx")
    ws = load_workbook(out)["Лист1"]
    assert ws.protection.sheet is True
    # пароль заказчика (хэш) не потерян: снять защиту без него нельзя
    source = load_workbook(protected)["Лист1"].protection
    assert ws.protection.password and ws.protection.password == source.password
    # клетки участника остались разблокированными, итоги -- заблокированными
    assert ws["C9"].protection.locked is False
    assert ws["E9"].protection.locked is True


def test_does_not_touch_cells_outside_the_map(form, tmp_path):
    before = cells_of(form)
    out = fill_workbook(form, basic_map(), make_context(), tmp_path / "out.xlsx")
    after = cells_of(out)
    mapped = ({"D8", "E23", "B31"}
              | {f"{col}{row}" for col in "CD" for row in range(9, 20)})
    changed = {a for a in before.keys() | after.keys()
               if before.get(a) != after.get(a)}
    assert changed == mapped
    # текст заказчика на месте
    for address in ("A3", "B9", "B19", "C23", "C31", "B20"):
        assert after[address] == before[address]


def test_more_rows_than_template_raises_with_counts(form, tmp_path):
    out = tmp_path / "out.xlsx"
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), make_context(12), out)
    text = str(caught.value)
    assert "12" in text and "11" in text
    assert "строк" in text and "позиц" in text.lower()
    assert not out.exists()                    # половинчатого файла нет


def test_output_opens_in_openpyxl(form, tmp_path):
    out = fill_workbook(form, basic_map(), make_context(), tmp_path / "out.xlsx")
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        wb = load_workbook(out)
    assert wb.sheetnames == ["Лист1"]
    assert out == tmp_path / "out.xlsx"


# ---- числа, строки, пустые хвосты ----------------------------------------------

def test_prices_are_numbers_not_text(form, tmp_path):
    items = make_items(1)
    items[0]["price_value"] = Decimal("5700000.00")
    items[0]["vat_value"] = Decimal("1234567.80")
    out = fill_workbook(form, basic_map(), make_context(items=items),
                        tmp_path / "out.xlsx")
    ws = load_workbook(out)["Лист1"]
    assert ws["C9"].data_type == "n" and ws["D9"].data_type == "n"
    assert isinstance(ws["C9"].value, (int, float))
    assert ws["C9"].value == 5700000
    assert ws["D9"].value == 1234567.8


def test_formatted_price_string_would_not_be_a_number(form, tmp_path):
    # то, чего карта делать не должна: «5 700 000,00» -- текст, SUM его не видит
    block = RowBlock(first_row=9, last_row=19, source="items",
                     columns={"C": "{row.price}"})
    out = fill_workbook(form, CellMap("Лист1", {}, [block]), make_context(1),
                        tmp_path / "out.xlsx")
    assert load_workbook(out)["Лист1"]["C9"].data_type == "s"


def test_fewer_items_leave_surplus_rows_empty(form, tmp_path):
    out = fill_workbook(form, basic_map(), make_context(3), tmp_path / "out.xlsx")
    cells = cells_of(out)
    assert cells["C11"] == 100003
    for row in range(12, 20):
        assert f"C{row}" not in cells and f"D{row}" not in cells
        assert cells[f"E{row}"] == f"=C{row}+D{row}"      # формулы целы


def test_no_items_is_not_an_error(form, tmp_path):
    out = fill_workbook(form, basic_map(), make_context(0), tmp_path / "out.xlsx")
    assert "C9" not in cells_of(out)


def test_template_file_is_never_modified(form, tmp_path):
    before = form.read_bytes()
    fill_workbook(form, basic_map(), make_context(), tmp_path / "out.xlsx")
    assert form.read_bytes() == before


def test_workbook_recalculates_on_open(form, tmp_path):
    # openpyxl не хранит результат формулы; Excel обязан пересчитать итоги сам
    out = fill_workbook(form, basic_map(), make_context(), tmp_path / "out.xlsx")
    assert load_workbook(out).calculation.fullCalcOnLoad is True


def test_formula_cells_have_no_stale_cached_value(form, tmp_path):
    out = fill_workbook(form, basic_map(), make_context(), tmp_path / "out.xlsx")
    cached = cells_of(out, data_only=True)
    # кэша нет вовсе -- а не старый ноль: устаревших цифр читатель не увидит
    for address in ("E9", "E19", "C20", "E20"):
        assert cached.get(address) is None


def test_string_starting_with_equals_is_text_not_formula(form, tmp_path):
    context = make_context(org={"head_name": "=1+1", "head_position": ""})
    cell_map = CellMap("Лист1", {"B31": "{org.head_name}"})
    out = fill_workbook(form, cell_map, context, tmp_path / "out.xlsx")
    ws = load_workbook(out)["Лист1"]
    assert ws["B31"].data_type == "s" and ws["B31"].value == "=1+1"


def test_lone_placeholder_keeps_the_type_of_the_value(form, tmp_path):
    cell_map = CellMap("Лист1", {"D8": "{tender.vat_rate}",
                                 "B31": "{tender.vat_percent}"})
    out = fill_workbook(form, cell_map, make_context(), tmp_path / "out.xlsx")
    ws = load_workbook(out)["Лист1"]
    assert ws["D8"].data_type == "n" and ws["D8"].value == 0.05
    assert ws["B31"].data_type == "s"


# ---- отказы --------------------------------------------------------------------

def test_refuses_to_write_over_a_formula(form, tmp_path):
    cell_map = CellMap("Лист1", {"E9": "{tender.vat_percent}"})
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, cell_map, make_context(), tmp_path / "out.xlsx")
    assert "E9" in str(caught.value) and "формул" in str(caught.value)


def test_refuses_to_write_over_a_formula_in_the_row_block(form, tmp_path):
    block = RowBlock(first_row=9, last_row=19, source="items",
                     columns={"E": "{row.price_value}"})
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, CellMap("Лист1", {}, [block]), make_context(),
                      tmp_path / "out.xlsx")
    assert "E9" in str(caught.value)


def test_refuses_to_write_into_a_merged_cell(form, tmp_path):
    cell_map = CellMap("Лист1", {"D23": "x"})          # C23:D23 слита
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, cell_map, make_context(), tmp_path / "out.xlsx")
    assert "D23" in str(caught.value) and "объединён" in str(caught.value)


def test_refuses_locked_cell_on_protected_sheet(tmp_path):
    protected = make_form(tmp_path / "protected.xlsx", protected=True)
    cell_map = CellMap("Лист1", {"B9": "перезапись"})   # B9 заблокирована
    with pytest.raises(CellMapError) as caught:
        fill_workbook(protected, cell_map, make_context(), tmp_path / "out.xlsx")
    assert "B9" in str(caught.value) and "защит" in str(caught.value)


def test_locked_cell_on_unprotected_sheet_is_fine(form, tmp_path):
    # без защиты листа «заблокирована» -- пустой флаг, как в форме Роснефти
    cell_map = CellMap("Лист1", {"B31": "x"})
    out = fill_workbook(form, cell_map, make_context(), tmp_path / "out.xlsx")
    assert cells_of(out)["B31"] == "x"


def test_refuses_to_overwrite_the_template(form):
    before = form.read_bytes()
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), make_context(), form)
    assert "исходн" in str(caught.value)
    assert form.read_bytes() == before


def test_missing_context_value_names_the_path_and_the_cell(form, tmp_path):
    context = make_context(conditions={})
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), context, tmp_path / "out.xlsx")
    text = str(caught.value)
    assert "conditions.Цена и условия:" in text and "E23" in text
    assert "KeyError" not in text


def test_missing_value_inside_a_row_names_the_row(form, tmp_path):
    items = make_items(3)
    del items[2]["vat_value"]
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), make_context(items=items),
                      tmp_path / "out.xlsx")
    assert "row.vat_value" in str(caught.value) and "D11" in str(caught.value)


def test_unknown_row_source_is_reported(form, tmp_path):
    block = RowBlock(first_row=9, last_row=19, source="nothing",
                     columns={"C": "{row.price_value}"})
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, CellMap("Лист1", {}, [block]), make_context(),
                      tmp_path / "out.xlsx")
    assert "nothing" in str(caught.value)


def test_missing_sheet_is_reported_in_russian(form, tmp_path):
    cell_map = CellMap("Нет такого", {"B31": "x"})
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, cell_map, make_context(), tmp_path / "out.xlsx")
    text = str(caught.value)
    assert "Нет такого" in text and "Лист1" in text


def test_missing_template_is_reported_in_russian(tmp_path):
    with pytest.raises(CellMapError) as caught:
        fill_workbook(tmp_path / "net.xlsx", basic_map(), make_context(),
                      tmp_path / "out.xlsx")
    assert "net.xlsx" in str(caught.value) and "Errno" not in str(caught.value)


def test_garbage_template_is_reported_in_russian(tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_text("это не книга", encoding="utf-8")
    with pytest.raises(CellMapError) as caught:
        fill_workbook(bad, basic_map(), make_context(), tmp_path / "out.xlsx")
    assert "bad.xlsx" in str(caught.value)


def test_unwritable_output_is_reported_in_russian(form, tmp_path):
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), make_context(),
                      tmp_path / "нет-папки" / "out.xlsx")
    assert "out.xlsx" in str(caught.value)


# ---- файл карты ----------------------------------------------------------------

def write_map(tmp_path, data):
    path = tmp_path / "map.json"
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return path


GOOD = {"sheet": "Лист1",
        "cells": {"D8": "Ставка НДС: {tender.vat_percent}%"},
        "rows": [{"first_row": 9, "last_row": 19, "source": "items",
                  "columns": {"C": "{row.price_value}"}}]}


def test_load_cellmap_reads_the_shape(tmp_path):
    cell_map = load_cellmap(write_map(tmp_path, GOOD))
    assert cell_map == CellMap(
        sheet="Лист1", cells={"D8": "Ставка НДС: {tender.vat_percent}%"},
        rows=[RowBlock(first_row=9, last_row=19, source="items",
                       columns={"C": "{row.price_value}"})])


def test_cellmap_without_rows_loads(tmp_path):
    data = {"sheet": "Лист1", "cells": {"B31": "x"}}
    assert load_cellmap(write_map(tmp_path, data)).rows == []


def test_missing_cellmap_file_is_reported_in_russian(tmp_path):
    with pytest.raises(CellMapError) as caught:
        load_cellmap(tmp_path / "net_takoy.json")
    text = str(caught.value)
    assert "net_takoy.json" in text and "не найдена" in text
    assert "No such file" not in text and "Errno" not in text


def test_broken_json_is_reported_in_russian(tmp_path):
    path = tmp_path / "map.json"
    path.write_text("{ это не json", encoding="utf-8")
    with pytest.raises(CellMapError) as caught:
        load_cellmap(path)
    assert "map.json" in str(caught.value)


@pytest.mark.parametrize("mutate, fragment", [
    (lambda d: d.pop("sheet"), "sheet"),
    (lambda d: d.pop("cells"), "cells"),
    (lambda d: d["cells"].update({"d8": "x"}), "d8"),            # не адрес
    (lambda d: d["cells"].update({"8D": "x"}), "8D"),
    (lambda d: d["cells"].update({"B31": 5}), "B31"),            # не строка
    (lambda d: d["cells"].update({"B31": "{org.head_name"}), "B31"),   # скобка
    (lambda d: d["cells"].update({"B31": "}{"}), "B31"),
    (lambda d: d["cells"].update({"C12": "x"}), "C12"),          # внутри блока
    (lambda d: d["rows"][0].pop("last_row"), "last_row"),
    (lambda d: d["rows"][0].update(last_row=8), "last_row"),
    (lambda d: d["rows"][0].update(first_row=0), "first_row"),
    (lambda d: d["rows"][0]["columns"].update({"c": "x"}), "c"),
    (lambda d: d["rows"][0]["columns"].update({"E": 7}), "E"),
    (lambda d: d["rows"][0].update(columns={}), "columns"),
    (lambda d: d.update(rows={"first_row": 9}), "rows"),         # не список
    (lambda d: d.update(rows=[5]), "rows"),
    # второй блок пересекается с первым (колонка C, строки 15-25)
    (lambda d: d["rows"].append({"first_row": 15, "last_row": 25,
                                 "source": "labor",
                                 "columns": {"C": "{row.amount_value}"}}), "C15"),
])
def test_bad_cellmap_is_rejected_with_a_russian_reason(tmp_path, mutate, fragment):
    data = json.loads(json.dumps(GOOD))
    mutate(data)
    with pytest.raises(CellMapError) as caught:
        load_cellmap(write_map(tmp_path, data))
    assert fragment in str(caught.value)


# ---- настоящая форма Роснефти ---------------------------------------------------

def _rosneft_context(tmp_path, *, positions=11, vat_rate=None):
    """Контекст из настоящей базы: тендер Роснефти с позициями лота."""
    conn = db.connect(tmp_path / "t.db")
    db.migrate(conn)
    repo = Repo(conn)
    me = repo.save_organization(Organization(
        role=ROLE_PARTICIPANT, inn="5904188041", short_name="ООО «Техносфера»",
        head_name="Рабичко Елена Михайловна", head_position="Директор"))
    customer = repo.save_organization(Organization(
        role=ROLE_CUSTOMER, inn="5190000000", short_name="АО «НК «Роснефть»"))
    repo.save_customer_profile(CustomerProfile(
        customer_id=customer, documents=["ТКП по форме заказчика"]))
    answers = {c: "подтверждаю" for c in CONDITIONS}
    answers[CONDITIONS[1]] = "не подтверждаю"
    tid = repo.save_tender(Tender(
        customer_id=customer, participant_id=me, purchase_number="2000726410",
        vat_rate=vat_rate, deal_answers=answers))
    for n in range(1, positions + 1):
        repo.save_lot_item(LotItem(tender_id=tid, position=n,
                                   name=f"СЗЗ {n}", price=Decimal(f"{n}00000.10")))
    context = build_context(tid, repo, document="ТКП по форме заказчика")
    conn.close()
    return context


@pytest.fixture
def real_fill(tmp_path):
    context = _rosneft_context(tmp_path)
    out = fill_workbook(REAL_FORM, load_cellmap(REAL_MAP), context,
                        tmp_path / "ТКП.xlsx")
    return out, context


def test_real_fixture_is_the_customers_form():
    cells = cells_of(REAL_FORM)
    assert cells["B7"] == "Лот  №2000726410"
    assert cells["E9"] == "=C9+D9"
    assert "АЗК № 4" in cells["B9"]


def test_real_map_names_only_the_bidders_cells():
    cell_map = load_cellmap(REAL_MAP)
    assert len(cell_map.rows) == 1
    block = cell_map.rows[0]
    assert cell_map.sheet == "Лист1"
    assert set(cell_map.cells) == {"D8", "E23", "E24", "E25", "E26", "E27",
                                   "E28", "E29", "B31"}
    assert block.first_row == 9 and block.last_row == 19
    assert block.source == "items"
    assert set(block.columns) == {"C", "D"}


def test_real_form_formulas_survive(real_fill):
    out, _ = real_fill
    cells = cells_of(out)
    for row in range(9, 20):
        assert cells[f"E{row}"] == f"=C{row}+D{row}"
    assert cells["C20"] == "=SUM(C9:C19)"
    assert cells["E20"] == "=SUM(E9:E19)"


def test_real_form_keeps_customer_text_in_names_and_conditions(real_fill):
    out, _ = real_fill
    before, after = cells_of(REAL_FORM), cells_of(out)
    for row in range(9, 20):
        assert after[f"A{row}"] == before[f"A{row}"]
        assert after[f"B{row}"] == before[f"B{row}"]
    assert "АЗК № 4" in after["B9"] and "Автоколонн" in after["B19"]
    for row in range(23, 30):
        assert after[f"B{row}"] == before[f"B{row}"]
        assert after[f"C{row}"] == before[f"C{row}"]


def test_real_form_changes_exactly_the_bidders_cells(real_fill):
    out, _ = real_fill
    before, after = cells_of(REAL_FORM), cells_of(out)
    changed = {a for a in before.keys() | after.keys()
               if before.get(a) != after.get(a)}
    expected = ({"D8", "B31"}
                | {f"{c}{r}" for c in "CD" for r in range(9, 20)}
                | {f"E{r}" for r in range(23, 30)})
    assert changed == expected


def test_real_form_values(real_fill):
    out, context = real_fill
    cells = cells_of(out)
    assert cells["D8"] == "Ставка НДС: 5%"
    assert cells["B31"] == "Директор Рабичко Елена Михайловна"
    assert [cells[f"E{r}"] for r in range(23, 30)] == [
        "подтверждаю", "не подтверждаю"] + ["подтверждаю"] * 5
    for n, row in enumerate(range(9, 20)):
        item = context["items"][n]
        assert cells[f"C{row}"] == pytest.approx(float(item["price_value"]))
        assert cells[f"D{row}"] == pytest.approx(float(item["vat_value"]))
        # числа, а не текст «100 000,10»
        assert not isinstance(cells[f"C{row}"], str)
    assert cells["C9"] == pytest.approx(100000.10)
    assert cells["D9"] == pytest.approx(5000.01)


def test_real_form_layout_and_signature_block_untouched(real_fill):
    out, _ = real_fill
    src = load_workbook(REAL_FORM)["Лист1"]
    dst = load_workbook(out)["Лист1"]
    assert sorted(map(str, dst.merged_cells.ranges)) == \
        sorted(map(str, src.merged_cells.ranges))
    assert len(dst.merged_cells.ranges) == 12
    assert dst["C31"].value == "[Подпись/печать]"        # подпись остаётся пустой
    assert dst["C9"].number_format == "#,##0.00"
    assert dst["C9"].fill.fgColor.rgb == src["C9"].fill.fgColor.rgb
    for col in "ABCDE":
        assert dst.column_dimensions[col].width == src.column_dimensions[col].width
    assert dst.protection.sheet is False                  # защиты нет, не появилась


def test_real_form_with_exemption_rate_writes_zero_vat(tmp_path):
    context = _rosneft_context(tmp_path, positions=2, vat_rate=Decimal("0"))
    out = fill_workbook(REAL_FORM, load_cellmap(REAL_MAP), context,
                        tmp_path / "ТКП.xlsx")
    cells = cells_of(out)
    assert cells["D8"] == "Ставка НДС: 0%"
    assert cells["D9"] == 0 and cells["D10"] == 0


def test_real_form_refuses_twelve_positions(tmp_path):
    context = _rosneft_context(tmp_path, positions=12)
    with pytest.raises(CellMapError) as caught:
        fill_workbook(REAL_FORM, load_cellmap(REAL_MAP), context,
                      tmp_path / "ТКП.xlsx")
    text = str(caught.value)
    assert "12" in text and "11" in text
    assert not (tmp_path / "ТКП.xlsx").exists()


# ---- несколько блоков строк (подтаблицы КП Арктикгаза) ---------------------------

def make_multi_form(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"
    ws["B34"] = "Итого материалы"
    ws["F34"] = "=SUM(F35:F37)"
    ws["E50"] = "=SUM(E48:E49)"
    wb.save(path)
    return path


def multi_map():
    return CellMap(sheet="Лист1", cells={}, rows=[
        RowBlock(first_row=35, last_row=37, source="materials",
                 columns={"B": "{row.name}", "F": "{row.amount_value}"}),
        RowBlock(first_row=48, last_row=49, source="labor",
                 columns={"B": "{row.post}", "E": "{row.amount_value}"}),
    ])


def multi_context(materials=2, labor=2):
    return {
        "materials": [{"name": f"Материал {n}", "amount_value": Decimal(f"{n}00.50")}
                      for n in range(1, materials + 1)],
        "labor": [{"post": f"Эколог {n}", "amount_value": Decimal(f"{n}000")}
                  for n in range(1, labor + 1)],
    }


def test_several_row_blocks_are_filled_independently(tmp_path):
    form = make_multi_form(tmp_path / "multi.xlsx")
    out = fill_workbook(form, multi_map(), multi_context(3, 2),
                        tmp_path / "out.xlsx")
    cells = cells_of(out)
    assert cells["B35"] == "Материал 1" and cells["F37"] == 300.5
    assert cells["B48"] == "Эколог 1" and cells["E49"] == 2000
    assert "B50" not in cells and "F38" not in cells
    # итоги книги -- её формулы
    assert cells["F34"] == "=SUM(F35:F37)" and cells["E50"] == "=SUM(E48:E49)"


def test_overflow_of_the_second_block_names_its_counts(tmp_path):
    form = make_multi_form(tmp_path / "multi.xlsx")
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, multi_map(), multi_context(2, 3),
                      tmp_path / "out.xlsx")
    text = str(caught.value)
    assert "3" in text and "2" in text and "labor" in text
    assert "строк" in text and "48" in text
    assert not (tmp_path / "out.xlsx").exists()


def test_each_block_has_its_own_capacity(tmp_path):
    # три материала влезают в три строки, хотя во втором блоке строк две
    form = make_multi_form(tmp_path / "multi.xlsx")
    fill_workbook(form, multi_map(), multi_context(3, 2), tmp_path / "out.xlsx")


# ---- пустое значение в клетке ответа -------------------------------------------

def test_missing_deal_answer_refuses_and_names_the_condition(form, tmp_path):
    context = make_context()
    del context["conditions"]["Цена и условия:"]
    out = tmp_path / "out.xlsx"
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), context, out)
    text = str(caught.value)
    assert "Цена и условия:" in text and "E23" in text
    assert not out.exists()


def test_empty_deal_answer_refuses_instead_of_blanking_the_cell(form, tmp_path):
    context = make_context()
    context["conditions"]["Цена и условия:"] = ""
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), context, tmp_path / "out.xlsx")
    text = str(caught.value)
    assert "Цена и условия:" in text and "E23" in text and "пуст" in text


def test_empty_interpolated_text_in_a_scalar_cell_refuses(form, tmp_path):
    context = make_context(org={"head_name": "", "head_position": ""})
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), context, tmp_path / "out.xlsx")
    assert "B31" in str(caught.value) and "пуст" in str(caught.value)


def test_empty_value_in_a_row_block_is_allowed(tmp_path):
    # пустое примечание в строке подтаблицы -- обычное дело, не отказ
    form = make_multi_form(tmp_path / "multi.xlsx")
    context = multi_context(1, 1)
    context["materials"][0]["name"] = ""
    out = fill_workbook(form, multi_map(), context, tmp_path / "out.xlsx")
    assert "B35" not in cells_of(out)


# ---- пробелы, фамилия подписанта, необязательные клетки -------------------------

def test_whitespace_only_deal_answer_refuses(form, tmp_path):
    context = make_context()
    context["conditions"]["Цена и условия:"] = "   "
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), context, tmp_path / "out.xlsx")
    assert "E23" in str(caught.value) and "пуст" in str(caught.value)


def test_non_breaking_space_answer_refuses(form, tmp_path):
    context = make_context()
    context["conditions"]["Цена и условия:"] = " "
    with pytest.raises(CellMapError):
        fill_workbook(form, basic_map(), context, tmp_path / "out.xlsx")


def test_lone_text_value_is_trimmed_like_interpolated_text(form, tmp_path):
    context = make_context()
    context["conditions"]["Цена и условия:"] = "  подтверждаю "
    out = fill_workbook(form, basic_map(), context, tmp_path / "out.xlsx")
    assert cells_of(out)["E23"] == "подтверждаю"


def test_signer_position_without_name_refuses(form, tmp_path):
    context = make_context(org={"head_name": "", "head_position": "Директор"})
    out = tmp_path / "out.xlsx"
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), context, out)
    text = str(caught.value)
    assert "B31" in text and "org.head_name" in text and "пуст" in text
    assert not out.exists()


def test_signer_name_without_position_refuses(form, tmp_path):
    context = make_context(org={"head_name": "Рабичко Е. М.", "head_position": " "})
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, basic_map(), context, tmp_path / "out.xlsx")
    assert "org.head_position" in str(caught.value)


def optional_map(expr="{org.head_name}"):
    return CellMap("Лист1", {"B31": expr}, optional={"B31"})


def test_optional_cell_with_empty_value_is_left_as_the_customer_had_it(form, tmp_path):
    context = make_context(org={"head_name": "  ", "head_position": ""})
    out = fill_workbook(form, optional_map(), context, tmp_path / "out.xlsx")
    assert cells_of(out)["B31"] == "[Руководитель/Должность/ФИО]"


def test_optional_cell_with_a_value_is_written(form, tmp_path):
    out = fill_workbook(form, optional_map(), make_context(), tmp_path / "out.xlsx")
    assert cells_of(out)["B31"] == "Рабичко Е. М."


def test_optional_interpolated_cell_tolerates_empty_parts(form, tmp_path):
    context = make_context(org={"head_name": "Рабичко Е. М.", "head_position": ""})
    cell_map = optional_map("{org.head_position} {org.head_name}")
    out = fill_workbook(form, cell_map, context, tmp_path / "out.xlsx")
    assert cells_of(out)["B31"] == "Рабичко Е. М."


def test_optional_does_not_excuse_a_missing_path(form, tmp_path):
    # «может быть пустым» -- не «может отсутствовать»: опечатка в карте видна
    with pytest.raises(CellMapError) as caught:
        fill_workbook(form, optional_map("{org.head_nmae}"), make_context(),
                      tmp_path / "out.xlsx")
    assert "org.head_nmae" in str(caught.value)


def test_cellmap_file_reads_optional_cells(tmp_path):
    data = {"sheet": "Лист1", "cells": {
        "D88": {"value": "{tender.rounding}", "optional": True},
        "B31": "{org.head_name}",
        "C5": {"value": "x"}}}
    cell_map = load_cellmap(write_map(tmp_path, data))
    assert cell_map.cells == {"D88": "{tender.rounding}", "B31": "{org.head_name}",
                              "C5": "x"}
    assert cell_map.optional == {"D88"}


@pytest.mark.parametrize("cell, fragment", [
    ({"value": "x", "optional": "yes"}, "optional"),          # не логическое
    ({"optional": True}, "value"),                            # нет выражения
    ({"value": "x", "required": False}, "required"),          # неизвестный ключ
    ({"value": 5}, "B31"),
])
def test_bad_optional_cell_is_rejected(tmp_path, cell, fragment):
    data = {"sheet": "Лист1", "cells": {"B31": cell}}
    with pytest.raises(CellMapError) as caught:
        load_cellmap(write_map(tmp_path, data))
    assert fragment in str(caught.value)


def test_real_map_has_no_optional_cells():
    # у Роснефти каждая графа обязательна: пустая -- отказ
    assert load_cellmap(REAL_MAP).optional == set()


# ---- ставки, которые книга хранит у себя ------------------------------------------

RATES_MAP = {"insurance": "C17", "injury": "C18",
             "overhead": "C24", "profit": "C25"}


@pytest.fixture
def rates_form(tmp_path):
    """Форма, которая держит ставки калькуляции в своих клетках, как книга
    Арктикгаза: C17, C18, C24, C25 и формула рядом."""
    path = make_form(tmp_path / "rates.xlsx")
    wb = load_workbook(path)
    ws = wb["Лист1"]
    ws["C17"], ws["C18"], ws["C24"], ws["C25"] = 0.3, 0.002, 0.05, 0.1
    ws["C26"] = "=C17+C18"
    wb.save(path)
    wb.close()
    return path


def test_cellmap_file_reads_the_rates_section(tmp_path):
    data = dict(GOOD, rates={"insurance": "C17", "profit": "C25"})
    assert load_cellmap(write_map(tmp_path, data)).rates == {
        "insurance": "C17", "profit": "C25"}


def test_cellmap_without_rates_keeps_them_empty(tmp_path):
    assert load_cellmap(write_map(tmp_path, GOOD)).rates == {}


def test_real_map_of_rosneft_names_no_rates():
    # в форме ТКП калькуляции нет вовсе: читать ставки неоткуда и незачем
    assert load_cellmap(REAL_MAP).rates == {}


@pytest.mark.parametrize("rates, fragment", [
    ({"страховые": "C17"}, "страховые"),          # имя не из списка
    ({"insurance_rate": "C17"}, "insurance_rate"),
    ({"insurance": "c17"}, "insurance"),          # не адрес
    ({"insurance": "C0"}, "insurance"),
    ({"insurance": 17}, "insurance"),             # не строка
], ids=["русское имя", "имя поля", "строчные", "нулевая строка", "не строка"])
def test_bad_rates_section_is_rejected_with_a_russian_reason(tmp_path, rates,
                                                             fragment):
    data = dict(GOOD, rates=rates)
    with pytest.raises(CellMapError) as caught:
        load_cellmap(write_map(tmp_path, data))
    assert fragment in str(caught.value)


def test_rates_section_must_be_an_object(tmp_path):
    data = dict(GOOD, rates=["C17"])
    with pytest.raises(CellMapError) as caught:
        load_cellmap(write_map(tmp_path, data))
    assert "rates" in str(caught.value)


def test_read_rates_takes_every_rate_from_the_book(rates_form):
    cell_map = CellMap("Лист1", {}, rates=RATES_MAP)
    assert read_rates(rates_form, cell_map) == {
        "insurance": Decimal("0.3"), "injury": Decimal("0.002"),
        "overhead": Decimal("0.05"), "profit": Decimal("0.1")}
    # имена раздела -- ровно поля настроек, с которыми ставки сравниваются
    assert set(RATES_MAP) == set(RATE_FIELDS)


def test_read_rates_without_the_section_does_not_even_open_the_book(tmp_path):
    assert read_rates(tmp_path / "нет-такой-книги.xlsx", basic_map()) == {}


@pytest.mark.parametrize("address, fragment", [
    ("C19", "пуста"),            # клетка пуста
    ("C26", "формул"),           # клетка занята формулой
    ("A3", "не число"),          # в клетке текст заказчика
], ids=["пусто", "формула", "текст"])
def test_read_rates_refuses_a_cell_without_a_rate(rates_form, address, fragment):
    # молча взятая вместо ставки нулевая доля дала бы правдоподобную цену
    # без взносов -- такую заявку никто не перепроверит
    cell_map = CellMap("Лист1", {}, rates={"insurance": address})
    with pytest.raises(CellMapError) as caught:
        read_rates(rates_form, cell_map)
    text = str(caught.value)
    assert fragment in text and address in text and "insurance" in text


def test_read_rates_names_the_missing_sheet(rates_form):
    cell_map = CellMap("Нет такого", {}, rates={"insurance": "C17"})
    with pytest.raises(CellMapError) as caught:
        read_rates(rates_form, cell_map)
    assert "Нет такого" in str(caught.value)


def test_read_rates_reports_a_missing_book_in_russian(tmp_path):
    cell_map = CellMap("Лист1", {}, rates={"insurance": "C17"})
    with pytest.raises(CellMapError) as caught:
        read_rates(tmp_path / "net.xlsx", cell_map)
    assert "net.xlsx" in str(caught.value) and "Errno" not in str(caught.value)
