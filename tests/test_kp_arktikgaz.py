"""Калькуляция к КП Арктикгаза: настоящий шаблон из templates_docs/ и его карта.

Книга считает итоги формулами сама, поэтому проверка идёт так: заполняем книгу
из настоящей базы, вычисляем её формулы здесь же (мини-вычислитель ниже) и
сверяем итог книги с калькулятором программы. Расхождение значило бы, что КП
уйдёт заказчику с ценой, которой в программе нет.
"""
import re
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from app import db
from app.models import Tender, TenderCostItem, TenderLaborRow
from app.render.context import build_context
from app.render.xlsx_render import CellMapError, fill_workbook, load_cellmap
from app.repo import Repo
from scripts.seed import DOC_ESTIMATE, seed

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates_docs" / "Калькуляция к КП Арктикгаз.xlsx"
CELLMAP = ROOT / "cellmaps" / "kp_arktikgaz.json"
SHEET = "Лист1"


# ---- вычислитель формул книги: ровно то, что в ней встречается -------------------

def _dec(value) -> Decimal:
    if value is None or value == "":
        return Decimal(0)
    return Decimal(str(value))


def evaluate(ws, address: str) -> Decimal:
    """Значение клетки; формулы книги считаются рекурсивно (+ - * /, SUM, ROUND)."""
    raw = ws[address].value
    if not (isinstance(raw, str) and raw.startswith("=")):
        return _dec(raw)

    def total(first, last):
        cells = ws[f"{first}:{last}"]
        return sum((evaluate(ws, c.coordinate) for row in cells for c in row),
                   Decimal(0))

    def rounded(value, digits):
        return value.quantize(Decimal(1).scaleb(-int(digits)),
                              rounding=ROUND_HALF_UP)

    expr = raw[1:]
    expr = re.sub(r"SUM\(([A-Z]+\d+):([A-Z]+\d+)\)", r'_sum("\1","\2")', expr)
    expr = expr.replace("ROUND(", "_round(")
    expr = re.sub(r'(?<![A-Za-z_"])([A-Z]{1,2}\d+)(?![\d"])', r'_v("\1")', expr)
    return eval(expr, {"_sum": total, "_round": rounded,
                       "_v": lambda a: evaluate(ws, a)})


# ---- данные ---------------------------------------------------------------------

@pytest.fixture
def repo(tmp_path):
    conn = db.connect(tmp_path / "t.db")
    db.migrate(conn)
    yield Repo(conn)
    conn.close()


def make_tender(repo, *, materials=2, labor=2, subcontract=3) -> int:
    """Тендер Арктикгаза из справочников, которые наполняет seed."""
    seed(repo)
    me = repo.find_organization_by_inn("5904188041")
    customer = next(p.customer_id for p in repo.list_customer_profiles()
                    if DOC_ESTIMATE in p.docs)
    tid = repo.save_tender(Tender(
        customer_id=customer, participant_id=me.id,
        object_address="Ямало-Ненецкий АО, Тазовский район",
        payment_terms="60 дней с даты подписания акта",
        deadline="в течение 90 дней"))
    for n in range(1, labor + 1):
        repo.save_tender_labor(TenderLaborRow(
            tender_id=tid, position=n, post=f"Должность {n}", headcount=n,
            hours=Decimal("120"), rate=Decimal("270.25")))
    for kind, count, price in (("materials", materials, "1234.57"),
                               ("subcontract", subcontract, "70000.40")):
        for n in range(1, count + 1):
            repo.save_tender_cost_item(TenderCostItem(
                tender_id=tid, kind=kind, position=n, name=f"{kind} {n}",
                unit="шт.", quantity=Decimal("2"), price=Decimal(price),
                note="содержание" if kind == "subcontract" else ""))
    return tid


def fill(repo, tmp_path, **sizes):
    tid = make_tender(repo, **sizes)
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context,
                        tmp_path / "КП.xlsx")
    return load_workbook(out)[SHEET], context


# ---- шаблон чист ------------------------------------------------------------------

def test_template_holds_no_data_of_the_submitted_bid():
    ws = load_workbook(TEMPLATE)[SHEET]
    for address in ("A4", "A5", "A6", "A7", "D2", "A69", "E72",
                    "B35", "B36", "B40", "B48", "B49", "B60", "B61", "B62",
                    "F60", "F61", "F62", "E48", "E49", "D48", "D49", "C48", "C49"):
        assert ws[address].value is None, address
    everything = " ".join(str(c.value) for row in ws.iter_rows() for c in row
                          if c.value is not None)
    for leftover in ("Рабичко", "10.02.2026", "АРКТИКГАЗ", "спецодежда",
                     "teh-sfera", "Эколог-проектировщик", "Исследования"):
        assert leftover not in everything, leftover


def test_template_keeps_formulas_and_literals():
    ws = load_workbook(TEMPLATE)[SHEET]
    assert ws["D11"].value == "=D13+D14"
    assert ws["D26"].value == "=ROUND(D11+D23+D24+D25,+D88)"
    assert ws["F34"].value == "=SUM(F35:F38)"
    assert ws["F54"].value == "=SUM(F48:F53)"
    assert ws["G54"].value == "=SUM(G48:G50)"
    assert ws["G48"].value == "=F60*E60"
    assert ws["F35"].value == "=E35*D35"
    assert ws["F48"].value == "=E48*D48*C48"
    for address, literal in (("D19", 0), ("D20", 0), ("D22", 0), ("D27", 1),
                             ("C17", 0.3), ("C24", 0.05)):
        assert ws[address].value == literal, address


# ---- заполнение -------------------------------------------------------------------

def test_scalar_cells_are_filled_from_the_tender(repo, tmp_path):
    ws, _ = fill(repo, tmp_path)
    assert ws["A4"].value == "Наименование заказчика: АО «АРКТИКГАЗ»"
    assert ws["A5"].value == ("Местоположение объекта: "
                              "Ямало-Ненецкий АО, Тазовский район")
    assert ws["A6"].value == "Условия оплаты: 60 дней с даты подписания акта"
    assert ws["A7"].value == "Срок исполнения: в течение 90 дней"
    assert ws["A69"].value == "Директор Рабичко Елена Михайловна"
    assert re.fullmatch(r"Дата: \d\d\.\d\d\.\d{4}г\.", ws["E72"].value)
    header = ws["D2"].value.split("\n")
    assert header[0] == "ООО «Техносфера»"
    assert header[1].startswith("Юридический адрес: ")
    assert header[2].startswith("Почтовый адрес: ")
    assert header[4].startswith("тел. ")
    assert len(header) == 5


def test_rows_are_numbers_not_printed_strings(repo, tmp_path):
    ws, _ = fill(repo, tmp_path, materials=4, labor=6, subcontract=3)
    numeric = ([f"{c}{r}" for r in range(35, 39) for c in "DE"]
               + [f"{c}{r}" for r in range(48, 54) for c in "CDE"]
               + [f"{c}{r}" for r in range(60, 63) for c in "EF"])
    for address in numeric:
        assert isinstance(ws[address].value, (int, float, Decimal)), address
    assert Decimal(str(ws["E35"].value)) == Decimal("1234.57")
    assert Decimal(str(ws["D48"].value)) == Decimal("120")
    assert Decimal(str(ws["F62"].value)) == Decimal("70000.40")
    assert ws["B60"].value == "subcontract 1"
    assert ws["C60"].value == "содержание"
    assert ws["D60"].value == "шт."


def test_second_material_group_stays_empty(repo, tmp_path):
    ws, _ = fill(repo, tmp_path, materials=4)
    for row in (40, 41, 42):
        for letter in "BCDE":
            assert ws[f"{letter}{row}"].value is None, f"{letter}{row}"
    assert evaluate(ws, "F39") == 0


def test_workbook_totals_equal_the_programs_estimate(repo, tmp_path):
    ws, context = fill(repo, tmp_path, materials=4, labor=6, subcontract=3)
    calc = context["estimate"]
    assert calc["total_wo_vat"] > 0              # сверка не на нулях
    assert evaluate(ws, "D13") == calc["materials"]
    assert evaluate(ws, "D16") == calc["labor"]
    assert evaluate(ws, "D21") == calc["subcontract"]
    assert evaluate(ws, "D26") == calc["total_wo_vat"]
    assert evaluate(ws, "D28") == calc["total_wo_vat"]        # объём работ = 1


def test_partial_lists_leave_surplus_rows_as_in_the_form(repo, tmp_path):
    ws, context = fill(repo, tmp_path, materials=1, labor=1, subcontract=1)
    assert ws["B36"].value is None and ws["E38"].value is None
    assert ws["B49"].value is None and ws["B53"].value is None
    assert ws["B61"].value is None and ws["F62"].value is None
    assert evaluate(ws, "D26") == context["estimate"]["total_wo_vat"]


# ---- ёмкость: лишняя позиция -- отказ, а не тихо заниженная цена --------------------

@pytest.mark.parametrize("sizes", [
    {"materials": 5}, {"labor": 7}, {"subcontract": 4}],
    ids=["материалы", "трудозатраты", "субподряд"])
def test_one_item_over_capacity_refuses_the_build(repo, tmp_path, sizes):
    # строки 63-65 субподряда в книге нарисованы, но ни в одну сумму не входят:
    # четвёртая позиция молча не попала бы в итог
    with pytest.raises(CellMapError, match="больше, чем строк"):
        fill(repo, tmp_path, **sizes)


def test_cellmap_blocks_match_the_workbook_sums():
    cell_map = load_cellmap(CELLMAP)
    blocks = {b.source: (b.first_row, b.last_row) for b in cell_map.rows}
    assert blocks == {"materials": (35, 38), "labor": (48, 53),
                      "subcontract": (60, 62)}
    assert "D88" not in cell_map.cells          # разрядность округления не в карте
