"""Калькуляция к КП Арктикгаза: настоящий шаблон из templates_docs/ и его карта.

Книга считает итоги формулами сама, поэтому проверка идёт так: заполняем книгу
из настоящей базы и вычисляем её формулы здесь же (мини-вычислитель ниже).
Результат сверяется с калькулятором программы -- но сходятся они не всегда.
Книга складывает неокруглённое и не округляет страховые взносы; программа
округляет каждую статью до копеек. Поэтому:

  * на данных, где произведения ложатся в копейку и итог не стоит на границе
    рубля (данные Арктикгаза, дробные 1,5 ед. и 25,5 ч с ровными копейками),
    итоги книги и программы равны -- это проверено;
  * на данных с долей копейки в строке или на границе рубля (1 чел. x 1 ч x
    1002,45) итоги расходятся, и это закреплено тестами, как и предупреждение
    сборки, которое о таких данных говорит;
  * `workbook_total` из app.calc сверяется здесь с формулами самой книги.
"""
import re
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from app import db
from app.calc import LaborLine, workbook_total
from app.models import Tender, TenderCostItem, TenderLaborRow
from app.render.context import build_context
from app.render.package import build_package
from app.render.xlsx_render import CellMapError, fill_workbook, load_cellmap
from app.repo import Repo
from scripts.seed import DOC_ESTIMATE, seed

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates_docs" / "Калькуляция к КП Арктикгаз.xlsx"
CELLMAP = ROOT / "cellmaps" / "kp_arktikgaz.json"
SHEET = "Лист1"
NBSP = " "


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


def make_custom_tender(repo, *, labor=(), materials=(), subcontract=()) -> int:
    """Тендер Арктикгаза из справочников, которые наполняет seed.
    labor -- (должность, человек, часов, ставка); materials и subcontract --
    (название, количество, цена)."""
    seed(repo)
    me = repo.find_organization_by_inn("5904188041")
    customer = next(p.customer_id for p in repo.list_customer_profiles()
                    if DOC_ESTIMATE in p.docs)
    tid = repo.save_tender(Tender(
        customer_id=customer, participant_id=me.id,
        object_address="Ямало-Ненецкий АО, Тазовский район",
        payment_terms="60 дней с даты подписания акта",
        deadline="в течение 90 дней"))
    for n, (post, people, hours, rate) in enumerate(labor, start=1):
        repo.save_tender_labor(TenderLaborRow(
            tender_id=tid, position=n, post=post, headcount=people,
            hours=Decimal(hours), rate=Decimal(rate)))
    for kind, rows in (("materials", materials), ("subcontract", subcontract)):
        for n, (name, quantity, price) in enumerate(rows, start=1):
            repo.save_tender_cost_item(TenderCostItem(
                tender_id=tid, kind=kind, position=n, name=name, unit="шт.",
                quantity=Decimal(quantity), price=Decimal(price),
                note="содержание" if kind == "subcontract" else ""))
    return tid


# настоящие данные поданной заявки Арктикгаза: итог без НДС ровно 4 000 000
ARKTIK_LABOR = [("Эколог-проектировщик", 4, "1200", "270"),
                ("Начальник отдела проектирования", 1, "360", "420")]
ARKTIK_MATERIALS = [("инвентарь, оборудование", "1", "24000"),
                    ("расходные материалы", "1", "17686")]
ARKTIK_SUBCONTRACT = [("исследования и измерения", "1", "1500000"),
                      ("фоновые концентрации", "1", "70000"),
                      ("экспертиза", "1", "50000")]


def make_tender(repo, *, materials=2, labor=2, subcontract=3,
                quantity="1.5", price="1234.56", hours="25.5",
                rate="181.20") -> int:
    """Сгенерированный тендер. Множители дробные, но произведения ложатся в
    копейку ровно: 1,5 * 1234,56 = 1851,84; 25,5 * 181,20 = 4620,60. Данные с
    долей копейки -- параметрами (price="1234.57")."""
    return make_custom_tender(
        repo,
        labor=[(f"Должность {n}", n, hours, rate) for n in range(1, labor + 1)],
        materials=[(f"materials {n}", quantity, price)
                   for n in range(1, materials + 1)],
        subcontract=[(f"subcontract {n}", quantity, "70000.40")
                     for n in range(1, subcontract + 1)])


def fill(repo, tmp_path, **sizes):
    tid = make_tender(repo, **sizes)
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context,
                        tmp_path / "КП.xlsx")
    return load_workbook(out)[SHEET], context


# ---- шаблон чист ------------------------------------------------------------------

# каждая клетка, в которой в поданной заявке стояли данные прошлого тендера
# (клетки, пустые и в исходной книге, сюда не входят: они ничего не проверяют)
CLEARED = (["A4", "A5", "A6", "A7", "D2", "A69", "E72"]
           + ["B35", "C35", "D35", "E35", "B36", "C36", "D36", "E36", "B40"]
           + [f"{c}{r}" for r in (48, 49) for c in "BCDE"]
           + ["B60", "D60", "E60", "F60", "B61", "C61", "D61", "E61", "F61",
              "B62", "D62", "E62", "F62"])


def test_template_holds_no_data_of_the_submitted_bid():
    ws = load_workbook(TEMPLATE)[SHEET]
    for address in CLEARED:
        assert ws[address].value is None, address
    everything = " ".join(str(c.value) for row in ws.iter_rows() for c in row
                          if c.value is not None).lower()
    # фразы, дословно стоявшие в очищенных клетках поданной книги
    for leftover in ("рабичко", "10.02.2026", "арктикгаз", "инвентарь, оборудование",
                     "расходные материалы", "спецодежда", "teh-sfera",
                     "эколог-проектировщик", "проведение исследований",
                     "фоновые концентрации и климатические", "экспертиза проектной",
                     "яро-яхинск", "шестидесяти", "декабрь 2026", "куйбышева"):
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
    assert Decimal(str(ws["E35"].value)) == Decimal("1234.56")
    assert Decimal(str(ws["D35"].value)) == Decimal("1.5")
    assert Decimal(str(ws["D48"].value)) == Decimal("25.5")
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
    assert evaluate(ws, "D13") == calc["materials"] == Decimal("7407.36")
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


def test_fractions_of_a_kopeck_make_the_workbook_and_the_program_differ(
        repo, tmp_path):
    # то самое расхождение, о котором предупреждает сборка: книга складывает
    # неокруглённое (4 x 1,5 x 1234,57 = 7407,42), программа -- напечатанные
    # строки (4 x 1851,86 = 7407,44)
    ws, context = fill(repo, tmp_path, materials=4, quantity="1.5",
                       price="1234.57")
    assert evaluate(ws, "D13") == Decimal("7407.42")
    assert context["estimate"]["materials"] == Decimal("7407.44")
    assert len(context["warnings"]) == 1       # одно на документ, с четырьмя строками


# ---- ёмкость: лишняя позиция -- отказ, а не тихо заниженная цена --------------------

@pytest.mark.parametrize("sizes", [
    {"materials": 5}, {"labor": 7}, {"subcontract": 4}],
    ids=["материалы", "трудозатраты", "субподряд"])
def test_one_item_over_capacity_refuses_the_build(repo, tmp_path, sizes):
    # строки 63-65 субподряда в книге нарисованы, но ни в одну сумму не входят:
    # четвёртая позиция молча не попала бы в итог
    with pytest.raises(CellMapError, match="больше, чем строк"):
        fill(repo, tmp_path, **sizes)


def _summed_rows(formula: str, column: str) -> list[int]:
    match = re.fullmatch(rf"=SUM\({column}(\d+):{column}(\d+)\)", formula)
    assert match, formula
    return list(range(int(match[1]), int(match[2]) + 1))


def test_cellmap_capacities_are_derived_from_the_workbook_formulas():
    ws = load_workbook(TEMPLATE)[SHEET]
    cell_map = load_cellmap(CELLMAP)
    blocks = {b.source: list(range(b.first_row, b.last_row + 1))
              for b in cell_map.rows}
    # материалы и ФОТ: подытог суммирует свою подтаблицу
    assert blocks["materials"] == _summed_rows(ws["F34"].value, "F")
    assert blocks["labor"] == _summed_rows(ws["F54"].value, "F")
    # субподряд: итог приходит окольным путём D21 = G54 = SUM(G48:G50), где
    # каждая G -- произведение одной строки субподряда; в сумму входят
    # только эти строки
    assert ws["D21"].value == "=G54"
    reached = []
    for row in _summed_rows(ws["G54"].value, "G"):
        match = re.fullmatch(r"=F(\d+)\*E(\d+)", ws[f"G{row}"].value)
        assert match and match[1] == match[2], ws[f"G{row}"].value
        reached.append(int(match[1]))
    assert blocks["subcontract"] == reached
    assert "D88" not in cell_map.cells          # решение контроллера


# ---- предупреждение: разойдутся ли две бумаги ---------------------------------------

def warnings_of(repo, document=DOC_ESTIMATE, **sizes):
    tid = make_tender(repo, **sizes)
    return build_context(tid, repo, document=document)["warnings"]


def custom_warnings(repo, **lists):
    tid = make_custom_tender(repo, **lists)
    return build_context(tid, repo, document=DOC_ESTIMATE)["warnings"]


def plain(text: str) -> str:
    return text.replace(NBSP, " ")


def test_no_warning_on_the_real_arktikgaz_data(repo, tmp_path):
    # итог книги и программы -- ровно 4 000 000: предупреждения быть не должно.
    # Этот тест не менее важен положительных: сигнал, который горит всегда,
    # через неделю перестают читать
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT)
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    assert context["warnings"] == []
    assert context["estimate"]["total_wo_vat"] == Decimal("4000000.00")
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context, tmp_path / "a.xlsx")
    assert evaluate(load_workbook(out)[SHEET], "D26") == Decimal("4000000")


def test_no_warning_on_whole_kopeck_products_away_from_the_rouble_boundary(repo):
    assert warnings_of(repo) == []
    assert warnings_of(repo, quantity="1", hours="1200", rate="270",
                       price="24000") == []


def test_warns_with_both_totals_when_insurance_rounding_splits_them(repo, tmp_path):
    # случай ревьюера: 1 чел. x 1 ч x 1002,45, произведение ровное, долей копейки
    # нет, а книга и программа разошлись на рубль (страховые 302,7399)
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "1", "1002.45")])
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    (text,) = context["warnings"]
    text = plain(text)
    assert "«Калькуляция к КП»" in text
    assert "итог без НДС 1 507,00" in text and "1 508,00" in text
    assert "Строк с долей копейки нет" in text
    assert "не округляет страховые" in text and "Исправить это нельзя" in text
    # и это не домысел: книга действительно даёт 1507, а программа 1508
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context, tmp_path / "r.xlsx")
    assert evaluate(load_workbook(out)[SHEET], "D26") == Decimal("1507")
    assert context["estimate"]["total_wo_vat"] == Decimal("1508.00")


def test_warns_about_material_row_with_half_a_kopeck(repo):
    (text,) = custom_warnings(repo, materials=[("расходные материалы", "1.5", "1234.57")])
    text = plain(text)
    assert "«Калькуляция к КП»" in text and "«расходные материалы» (материалы)" in text
    assert "1,5 × 1 234,57 = 1 851,855" in text
    assert "в документе напечатано 1 851,86" in text       # то, что видит человек
    assert "количество или цену" in text and "численность" not in text


def test_warns_about_labor_row_and_suggests_all_three_factors(repo):
    (text,) = custom_warnings(repo, labor=[("Эколог", 1, "25.7", "181.26")])
    text = plain(text)
    assert "«Эколог» (трудозатраты)" in text
    assert "1 × 25,7 × 181,26 = 4 658,382" in text
    assert "численность, часы или ставку" in text and "количество или цену" not in text


def test_warns_about_a_row_even_when_the_totals_coincide(repo):
    # четыре строки 4 x 1,5 x 1234,57: итоги до рубля одинаковы (7 407), но
    # подытог книги 7 407,42 не сойдётся с четырьмя напечатанными по 1 851,86
    (text,) = custom_warnings(repo, materials=[(f"м{n}", "1.5", "1234.57")
                                               for n in range(1, 5)])
    text = plain(text)
    assert "итог без НДС в книге Excel и в программе совпадёт (7 407,00)" in text
    assert "подытоги книги не сойдутся" in text
    for n in range(1, 5):
        assert f"«м{n}» (материалы)" in text


def test_warns_about_subcontract_row_by_name(repo):
    (text,) = custom_warnings(repo, subcontract=[("экспертиза", "0.5", "0.01")])
    assert "«экспертиза» (субподряд)" in plain(text)


def test_one_warning_names_every_offending_row(repo):
    (text,) = warnings_of(repo, materials=3, labor=1, quantity="1.5",
                          price="1234.57", hours="25.7", rate="181.26")
    for n in (1, 2, 3):
        assert f"«materials {n}»" in text
    assert "«Должность 1»" in text


def test_no_warning_for_a_document_without_a_cellmap(repo):
    # те же данные, что дают предупреждение, но документ собирается Word
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "1", "1002.45")],
                             materials=[("м", "1.5", "1234.57")])
    assert build_context(tid, repo, document=DOC_ESTIMATE)["warnings"]
    assert build_context(tid, repo, document="Коммерческое предложение"
                         )["warnings"] == []
    assert build_context(tid, repo)["warnings"] == []


def test_no_warning_about_lists_the_workbook_does_not_receive(repo):
    # ТКП Роснефти пишет в книгу только позиции лота: материалы и ФОТ тендера
    # в ней не участвуют, предупреждать о них было бы враньём
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "1", "1002.45")],
                             materials=[("м", "1.5", "1234.57")])
    tender = repo.get_tender(tid)
    tender.customer_id = next(p.customer_id for p in repo.list_customer_profiles()
                              if "ТКП по форме заказчика" in p.docs)
    repo.save_tender(tender)
    assert build_context(tid, repo, document="ТКП по форме заказчика"
                         )["warnings"] == []


@pytest.mark.parametrize("rate", ["1002.45", "1000", "1000.60", "270", "181.26",
                                  "999.99", "1234.56", "0.01"])
def test_warning_fires_exactly_when_the_papers_differ(repo, tmp_path, rate):
    # условие прямое: итоги книги (вычислены по её формулам) и программы
    # расходятся тогда и только тогда, когда есть предупреждение без строк
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "1", rate)])
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context, tmp_path / "x.xlsx")
    book = evaluate(load_workbook(out)[SHEET], "D26")
    differ = book != context["estimate"]["total_wo_vat"]
    assert bool(context["warnings"]) == differ, (rate, book)
    assert workbook_total(
        [LaborLine("Эколог", 1, Decimal("1"), Decimal(rate))], Decimal(0),
        Decimal(0), repo.get_settings()) == book


# ---- сквозная сборка: профиль -> пакет -> файл ----------------------------------------

def test_package_builds_the_estimate_from_the_customer_profile(repo, tmp_path):
    tid = make_tender(repo, materials=1, quantity="1.5", price="1234.57")
    result = build_package(tid, [DOC_ESTIMATE], tmp_path / "out", repo)
    assert result.skipped == []
    assert [p.suffix for p in result.written] == [".xlsx"]
    ws = load_workbook(result.written[0])[SHEET]
    # шаблон и карту выбрал профиль заказчика, а не константы теста
    assert ws["A4"].value == "Наименование заказчика: АО «АРКТИКГАЗ»"
    assert ws["B35"].value == "materials 1"
    assert ws["D35"].value == 1.5
    assert ws["A69"].value == "Директор Рабичко Елена Михайловна"
    assert ws["F35"].value == "=E35*D35"
    # предупреждение дошло до результата сборки вместе с названием документа
    assert len(result.warnings) == 1
    document, text = result.warnings[0]
    assert document == DOC_ESTIMATE and "«materials 1»" in text
