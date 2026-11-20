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
import json
import re
import shutil
from dataclasses import replace
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import docx
import pytest
from openpyxl import load_workbook

from app import config, db
from app.calc import LaborLine, workbook_total
from app.models import LotItem, Tender, TenderCostItem, TenderLaborRow
from app.render.context import build_context, money, required_conditions
from app.render.package import PACKAGE_WIDE, build_package
from app.render.xlsx_render import CellMapError, fill_workbook, load_cellmap
from app.repo import Repo
from scripts.seed import DOC_DECLARATION, DOC_ESTIMATE, DOC_KP, seed

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates_docs" / "Калькуляция к КП Арктикгаз.xlsx"
CELLMAP = ROOT / "cellmaps" / "kp_arktikgaz.json"
SHEET = "Лист1"
# Клетка, где книга держит ставку страховых взносов (C17 = 0,3).
RATE_CELL = "C17"
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


def make_custom_tender(repo, *, labor=(), materials=(), subcontract=(),
                       lot=()) -> int:
    """Тендер Арктикгаза из справочников, которые наполняет seed.
    labor -- (должность, человек, часов, ставка); materials и subcontract --
    (название, количество, цена); lot -- позиции лота (название, цена),
    из них коммерческое предложение печатает свой итог."""
    seed(repo)
    me = repo.find_organization_by_inn("5904188041")
    customer = next(p.customer_id for p in repo.list_customer_profiles()
                    if DOC_ESTIMATE in p.docs)
    # Поля закупки заполнены все: их печатают документы Word (декларация), и
    # на пустой графе сборка теперь предупреждает -- заявку из-за такой графы
    # отклоняют. Пустые графы проверяются своими тестами, а не этим.
    # Так же и текст декларации в профиле: на пустом декларация не собирается вовсе
    # (это проверяет tests/test_package.py), а здесь она лишь соседка по пакету.
    profile = repo.find_customer_profile(customer)
    if not profile.declaration_text.strip():
        profile.declaration_text = "Не находимся в процессе ликвидации"
        repo.save_customer_profile(profile)
    tid = repo.save_tender(Tender(
        customer_id=customer, participant_id=me.id,
        purchase_number="32514010203", purchase_date="2025-10-01",
        method="запрос предложений в электронной форме",
        subject="Разработка КЭР",
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
    for n, (name, price) in enumerate(lot, start=1):
        repo.save_lot_item(LotItem(tender_id=tid, position=n, name=name,
                                   price=Decimal(price)))
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
    # Клетка округления итога книги: карта пишет в неё настройку программы,
    # и тогда книга и Word округляют итог одинаково. Механизм необязательной
    # клетки для этого и задумывался: в чистой форме D88 пуста.
    assert cell_map.cells["D88"] == "{rounding}"
    assert "D88" in cell_map.optional


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


def test_row_kind_is_taken_from_the_data_not_from_the_assembled_text(repo):
    # должность «Инженер (материалы)» -- это трудозатраты, и совет ей положен
    # про численность, часы и ставку; искать вид строки в собранном тексте
    # значит верить тому, что сам же и написал
    (text,) = custom_warnings(
        repo, labor=[("Инженер (материалы)", 1, "25.7", "181.26")])
    text = plain(text)
    assert "«Инженер (материалы)» (трудозатраты)" in text
    assert "исправьте численность, часы или ставку (трудозатраты)" in text
    assert "количество или цену" not in text


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


def test_warns_about_subcontract_row_by_name_and_suggests_its_own_factors(repo):
    (text,) = custom_warnings(repo, subcontract=[("экспертиза", "0.5", "0.01")])
    text = plain(text)
    assert "«экспертиза» (субподряд)" in text
    assert "0,5 × 0,01 = 0,005" in text
    assert "исправьте количество или цену (субподряд)" in text
    assert "численность" not in text and "(материалы)" not in text


def test_one_warning_names_every_offending_row(repo):
    (text,) = warnings_of(repo, materials=3, labor=1, quantity="1.5",
                          price="1234.57", hours="25.7", rate="181.26")
    for n in (1, 2, 3):
        assert f"«materials {n}»" in text
    assert "«Должность 1»" in text


def test_no_warning_for_a_document_without_a_cellmap(repo):
    # те же данные, что дают предупреждение, но документ собирается Word
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "1", "1002.45")],
                             materials=[("м", "1.5", "1234.57")],
                             lot=[("Разработка КЭР", "1000")])
    assert build_context(tid, repo, document=DOC_ESTIMATE)["warnings"]
    assert build_context(tid, repo, document="Коммерческое предложение"
                         )["warnings"] == []
    assert build_context(tid, repo)["warnings"] == []


ROSNEFT_MAP = ROOT / "cellmaps" / "tkp_rosneft_murmansk.json"


def move_to_rosneft(repo, tid):
    """Тот же тендер, но заказчик -- Роснефть: её карта шлёт в книгу только
    позиции лота.

    Ответы по условиям сделки её форма требует все, иначе контекст отказывает;
    сами условия берутся из её карты, а не переписываются сюда списком."""
    tender = repo.get_tender(tid)
    tender.customer_id = next(p.customer_id for p in repo.list_customer_profiles()
                              if "ТКП по форме заказчика" in p.docs)
    tender.deal_answers = {name: "подтверждаю" for name
                           in required_conditions(load_cellmap(ROSNEFT_MAP))}
    repo.save_tender(tender)


@pytest.mark.parametrize("lists", [
    {"labor": [("Эколог", 1, "1", "1002.45")]},
    {"materials": [("м", "1.5", "1234.57")]},
    {"subcontract": [("экспертиза", "0.5", "0.01")]},
], ids=["трудозатраты", "материалы", "субподряд"])
def test_no_warning_about_lists_the_workbook_does_not_receive(repo, lists):
    # ТКП Роснефти пишет в книгу только позиции лота: материалы, ФОТ и
    # субподряд тендера в ней не участвуют, предупреждать о них было бы враньём
    tid = make_custom_tender(repo, lot=[("Разработка КЭР", "1000")], **lists)
    assert build_context(tid, repo, document=DOC_ESTIMATE)["warnings"]
    move_to_rosneft(repo, tid)
    assert build_context(tid, repo, document="ТКП по форме заказчика"
                         )["warnings"] == []


# ---- ставки: книга считает по своим --------------------------------------------

@pytest.mark.parametrize("field, value, label, in_book, in_settings", [
    ("insurance_rate", "0.31", "страховых взносов", "30", "31"),
    # ставку страховых закон может и понизить, и тогда сторож обязан сказать
    # то же самое: сравнение -- на неравенство, а не «настройка выше книжной»
    ("insurance_rate", "0.29", "страховых взносов", "30", "29"),
    ("injury_rate", "0.003", "взносов на травматизм", "0,2", "0,3"),
    ("overhead_rate", "0.10", "накладных расходов", "5", "10"),
    ("profit_rate", "0.11", "прибыли", "10", "11"),
], ids=["страховые-выше", "страховые-ниже", "травматизм", "накладные",
        "прибыль"])
def test_each_rate_of_the_book_is_compared_with_the_settings(
        repo, tmp_path, field, value, label, in_book, in_settings):
    # ставку страховых меняет закон, остальные -- человек в настройках; книга
    # посчитает по своим клеткам C17, C18, C24, C25, что бы мы ни думали
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT)
    repo.save_settings(replace(repo.get_settings(), **{field: Decimal(value)}))
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    totals, note = (plain(w) for w in context["warnings"])
    assert (f"ставка {label} {in_book} %, а в настройках программы "
            f"{in_settings} %") in note
    assert "Книга посчитает по своей" in note
    assert "Исправьте ставку в настройках или в самом шаблоне книги" in note
    # итог книги считается по ЕЁ ставкам и остался прежним, итог программы -- нет
    assert "книга Excel посчитает итог без НДС 4 000 000,00" in totals
    assert context["estimate"]["total_wo_vat"] != Decimal("4000000.00")
    assert plain(money(context["estimate"]["total_wo_vat"])) in totals
    # Чего в тексте быть НЕ должно: причина одна -- ставка, и она исправима.
    # Пообещать «исправить это нельзя» там, где достаточно поправить одно
    # число, -- худшая услуга, чем молчание.
    assert "Исправить это нельзя" not in totals
    assert "не округляет страховые" not in totals
    assert "границе рубля" not in totals
    assert "Строк с долей копейки нет" not in totals
    assert "долей копейки" not in totals
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context, tmp_path / "s.xlsx")
    assert evaluate(load_workbook(out)[SHEET], "D26") == Decimal("4000000")


@pytest.mark.parametrize("field, value, program", [
    ("insurance_rate", "0.31", "4016715.00"),
    ("overhead_rate", "0.10", "4103634.00"),
], ids=["страховые 31 %", "накладные 10 %"])
def test_changed_rate_splits_the_two_papers_by_thousands(repo, field, value,
                                                         program):
    # шестнадцать тысяч рублей разницы в двух бумагах одного пакета -- крупнее
    # всего, что сторож ловил раньше, и раньше он об этом молчал
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT)
    repo.save_settings(replace(repo.get_settings(), **{field: Decimal(value)}))
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    assert context["estimate"]["total_wo_vat"] == Decimal(program)
    assert len(context["warnings"]) == 2
    assert plain(money(Decimal(program))) in plain(context["warnings"][0])
    assert "об этом отдельное предупреждение" in context["warnings"][0]


def test_rates_that_match_the_settings_say_nothing_about_rates(repo):
    # сигнал, который горит всегда, через неделю перестают читать
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT)
    settings = repo.get_settings()
    assert (settings.insurance_rate, settings.injury_rate, settings.overhead_rate,
            settings.profit_rate) == (Decimal("0.30"), Decimal("0.002"),
                                      Decimal("0.05"), Decimal("0.10"))
    assert build_context(tid, repo, document=DOC_ESTIMATE)["warnings"] == []


def test_a_document_without_a_calculation_in_its_book_says_nothing_about_rates(repo):
    # карта Роснефти раздела rates не знает и калькуляции в книгу не шлёт:
    # поведение прежнее, ничего не меняется
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR,
                             lot=[("Разработка КЭР", "1000")])
    repo.save_settings(replace(repo.get_settings(),
                               insurance_rate=Decimal("0.31")))
    move_to_rosneft(repo, tid)
    assert build_context(tid, repo, document="ТКП по форме заказчика"
                         )["warnings"] == []


def cellmap_without_rates(folder: Path) -> Path:
    """Копия карты Арктикгаза без раздела rates, в своей папке карт."""
    data = json.loads(CELLMAP.read_text(encoding="utf-8"))
    del data["rates"]
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / CELLMAP.name
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return path


def test_a_cellmap_without_rates_takes_them_from_the_settings(repo, tmp_path,
                                                              monkeypatch):
    # Карта Роснефти для этого не годится: она не шлёт в книгу списков
    # калькуляции, и workbook_total на ней не считается вовсе. Здесь карта
    # Арктикгаза без раздела rates: калькуляция живая, ставки -- из настроек,
    # и итог книги обязан поехать вместе с настройкой, а не остаться прежним.
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT)
    path = cellmap_without_rates(tmp_path / "cellmaps")
    monkeypatch.setattr(config, "CELLMAPS_DIR", path.parent)
    cell_map = load_cellmap(path)
    assert cell_map.rates == {}
    assert {b.source for b in cell_map.rows} == {"materials", "labor",
                                                 "subcontract"}
    repo.save_settings(replace(repo.get_settings(),
                               insurance_rate=Decimal("0.31")))
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    # ставка 31 % у обоих: и у программы, и у книги, которой карта ставок не дала
    assert context["estimate"]["total_wo_vat"] == Decimal("4016715.00")
    assert context["warnings"] == []


# ---- испорченная клетка ставки: то, что увидит человек -------------------------------

def templates_with_rate_cell(folder: Path, value) -> Path:
    """Копия шаблона Арктикгаза в своей папке шаблонов, с заданным C17.

    Рядом кладётся и шаблон коммерческого предложения: пакет из одного
    документа не показал бы главного -- что испорченная клетка снимает со
    сборки только свой документ, а остальные человек получает."""
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / TEMPLATE.name
    wb = load_workbook(TEMPLATE)
    wb[SHEET][RATE_CELL] = value
    wb.save(path)
    wb.close()
    shutil.copy(TEMPLATE.parent / f"{DOC_KP}.docx", folder / f"{DOC_KP}.docx")
    return path


def package_with_rate_cell(repo, tmp_path, value, monkeypatch):
    """Пакет из калькуляции и КП, собранный с заданным значением в C17 и
    настройкой страховых 31 %: ставки разошлись, сторожу есть о чём сказать.

    Позиция лота нужна КП: без неё коммерческое предложение ушло бы в skipped
    с нулевым итогом, и проверка «остальные документы человек получает»
    проверяла бы пустоту."""
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT,
                             lot=[("Разработка КЭР", "3950000")])
    path = templates_with_rate_cell(tmp_path / "templates", value)
    monkeypatch.setattr(config, "TEMPLATES_DOCS_DIR", path.parent)
    repo.save_settings(replace(repo.get_settings(),
                               insurance_rate=Decimal("0.31")))
    return build_package(tid, [DOC_ESTIMATE, DOC_KP], tmp_path / "out", repo)


@pytest.mark.parametrize("value, fragment", [
    ("=0.3", "формул"),
    ("30 %", "не число"),
    (None, "пуста"),
    (30, "долей единицы"),
], ids=["формула", "текст", "пусто", "проценты-вместо-доли"])
def test_a_broken_rate_cell_skips_the_document_with_a_reason(
        repo, tmp_path, monkeypatch, value, fragment):
    # Худшее, что может сделать сторож, -- молча выключиться: испорченная клетка
    # ставки отключала не только сверку ставок, а всю проверку целиком, и пакет
    # уходил человеку без единого замечания (skipped пуст, warnings пуст).
    result = package_with_rate_cell(repo, tmp_path, value, monkeypatch)
    (document, reason) = result.skipped[0]
    assert len(result.skipped) == 1 and document == DOC_ESTIMATE
    assert RATE_CELL in reason and "insurance" in reason and fragment in reason
    assert "Errno" not in reason and "Traceback" not in reason
    # снят со сборки только свой документ: остальные человек получает
    assert [p.name for p in result.written] == [f"{DOC_KP}.docx"]
    assert [p.name for p in (tmp_path / "out").iterdir()] == [f"{DOC_KP}.docx"]


def test_a_sound_rate_cell_in_the_same_place_builds_and_warns_twice(
        repo, tmp_path, monkeypatch):
    # обратная сторона того же теста: клетка цела -- собраны оба документа,
    # и сторож говорит оба предупреждения (итоги и сама ставка)
    result = package_with_rate_cell(repo, tmp_path, 0.3, monkeypatch)
    assert result.skipped == []
    assert [p.suffix for p in result.written] == [".xlsx", ".docx"]
    # третье предупреждение -- о пакете целиком: КП печатает 3 950 000 из
    # позиции лота, а калькуляция при страховых 31 % даёт другой итог
    assert [d for d, _ in result.warnings] == [DOC_ESTIMATE, DOC_ESTIMATE,
                                               PACKAGE_WIDE]
    totals, note, prices = (plain(t) for _, t in result.warnings)
    assert "книга Excel посчитает итог без НДС 4 000 000,00" in totals
    assert "ставка страховых взносов 30 %, а в настройках программы 31 %" in note
    assert f"«{DOC_KP}» -- 3 950 000,00" in prices


# Данные, на которых видна вся арифметика книги разом: материалы, субподряд и
# три строки ФОТ, причём сырые произведения не ложатся в копейку (1,5 x 1234,57
# = 1851,855; 1,5 x 70 000,41 = 105 000,615; 25,5 x 420,33 = 10 718,415).
BOOK_MATERIALS = [("инвентарь, оборудование", "1.5", "1234.57"),
                  ("расходные материалы", "2", "890.40")]
BOOK_SUBCONTRACT = [("исследования и измерения", "1.5", "70000.41"),
                    ("фоновые концентрации", "2", "35000.20")]


def book_labor(rate):
    return [("Эколог-проектировщик", 4, "25.5", "270"),
            ("Начальник отдела проектирования", 1, "25.5", "420.33"),
            ("Инженер-эколог", 2, "25.5", rate)]


def raw_sum(rows) -> Decimal:
    """Сумма сырых произведений количества на цену -- так считает книга."""
    return sum((Decimal(q) * Decimal(p) for _, q, p in rows), Decimal(0))


@pytest.mark.parametrize("rate, total", [
    # Ставка третьей строки ФОТ подобрана так, чтобы итог книги стоял у границы
    # рубля: тогда лишнее или пропущенное округление внутри workbook_total
    # меняет D26 на рубль, и подмену видно. Что ловит каждый случай:
    ("152.18", "265339"),   # ROUND у накладных (D24): без него книга дала бы 265 340
    ("165.61", "266369"),   # ROUND у прибыли (D25): без него -- 266 370
    ("155.27", "265575"),   # ФОТ (D16) книга НЕ округляет: иначе -- 265 576
    ("183.76", "267760"),   # всё сразу, включая неокруглённый субподряд (D21)
])
def test_workbook_total_reproduces_the_formulas_of_the_real_book(
        repo, tmp_path, rate, total):
    """Итог workbook_total равен D26 настоящей книги, вычисленной по её формулам."""
    labor = book_labor(rate)
    tid = make_custom_tender(repo, labor=labor, materials=BOOK_MATERIALS,
                             subcontract=BOOK_SUBCONTRACT)
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context,
                        tmp_path / "book.xlsx")
    ws = load_workbook(out)[SHEET]
    # сверка не на нулях и не на пустых подтаблицах
    assert evaluate(ws, "D13") == Decimal("3632.655")       # материалы
    assert evaluate(ws, "D21") == Decimal("175001.015")     # субподряд
    assert evaluate(ws, "D16") > 0 and len(labor) == 3      # три строки ФОТ
    book = evaluate(ws, "D26")
    assert book == Decimal(total)
    assert workbook_total(
        [LaborLine(post, people, Decimal(hours), Decimal(money_rate))
         for post, people, hours, money_rate in labor],
        raw_sum(BOOK_MATERIALS), raw_sum(BOOK_SUBCONTRACT),
        repo.get_settings()) == book
    # почему эти данные: неокруглённый итог книги стоит у самой границы рубля
    # (не дальше пяти копеек от неё), и лишнее или пропущенное округление
    # внутри статей меняет D26 на целый рубль
    raw_total = sum(evaluate(ws, a) for a in ("D11", "D23", "D24", "D25"))
    assert abs(abs(raw_total - book) - Decimal("0.5")) < Decimal("0.05")


def test_a_total_exactly_on_the_half_rouble_goes_up_as_excel_rounds(repo, tmp_path):
    # ROUND в Excel округляет половину ВВЕРХ, а не к чётному: неокруглённый итог
    # 4 000 000,50 даёт 4 000 001, хотя целая часть чётная и ROUND_HALF_EVEN
    # оставил бы 4 000 000. Данные Арктикгаза с материалами 41 686,67 --
    # единственный случай, где эта разница видна вообще
    tid = make_custom_tender(
        repo, labor=ARKTIK_LABOR, subcontract=ARKTIK_SUBCONTRACT,
        materials=[("инвентарь, оборудование", "1", "24000"),
                   ("расходные материалы", "1", "17686.67")])
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context,
                        tmp_path / "half.xlsx")
    ws = load_workbook(out)[SHEET]
    raw_total = sum(evaluate(ws, a) for a in ("D11", "D23", "D24", "D25"))
    assert raw_total == Decimal("4000000.5")           # ровно половина
    assert evaluate(ws, "D26") == Decimal("4000001")
    assert workbook_total(
        [LaborLine(post, people, Decimal(hours), Decimal(rate))
         for post, people, hours, rate in ARKTIK_LABOR],
        Decimal("41686.67"), raw_sum(ARKTIK_SUBCONTRACT),
        repo.get_settings()) == Decimal("4000001.00")
    # программа округляет половину так же, поэтому бумаги не расходятся
    assert context["estimate"]["total_wo_vat"] == Decimal("4000001.00")
    assert context["warnings"] == []


def test_the_warning_prints_the_contexts_own_estimate_as_the_programs_total(repo):
    # «программа» в предупреждении -- ровно тот итог без НДС, который контекст
    # отдаёт документам Word, а не пересчитанный заново по сырым трудозатратам:
    # две строки по 25,5 x 181,21 = 4 620,855 печатаются по 4 620,86, и сумма
    # напечатанных (9 241,72) больше округлённой суммы сырых (9 241,71)
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "25.5", "181.21"),
                                          ("Эколог-2", 1, "25.5", "181.21")])
    repo.save_settings(replace(repo.get_settings(), total_rounding_digits=2))
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    (text,) = context["warnings"]
    program = context["estimate"]["total_wo_vat"]
    assert program == Decimal("13897.80")
    assert f"расчёт программы даёт {money(program)}." in text
    assert "13 897,79" not in plain(text)      # итог по сырым трудозатратам
    # «Коммерческое предложение» из текста убрано: оно печатает сумму позиций
    # лота, а не итог расчёта, и сторож не вправе обещать сравнение с ним
    assert "Коммерческое предложение" not in text


def cellmap_without_rounding(folder: Path) -> Path:
    """Копия карты Арктикгаза без клетки округления итога, в своей папке карт.
    Так карта выглядела до того, как D88 в неё вошла: книга оставалась на
    целых рублях, что бы ни стояло в настройке."""
    data = json.loads(CELLMAP.read_text(encoding="utf-8"))
    del data["cells"]["D88"]
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / CELLMAP.name
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return path


def with_rounding(repo, digits, tmp_path=None):
    """Тендер на 1 чел. x 1 ч x 1002,45 при заданной разрядности округления --
    данные, на которых округление итога видно в чистом виде (страховые
    302,7399 стоят на границе)."""
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "1", "1002.45")])
    repo.save_settings(replace(repo.get_settings(),
                               total_rounding_digits=digits))
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    if tmp_path is None:
        return context, None
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context,
                        tmp_path / "d88.xlsx")
    return context, load_workbook(out)[SHEET]


@pytest.mark.parametrize("digits, book", [(0, "1507"), (2, "1507.49"),
                                          (-3, "2000")],
                         ids=["до-рублей", "до-копеек", "до-тысяч"])
def test_the_rounding_cell_carries_the_setting_into_the_book(
        repo, tmp_path, digits, book):
    # Открытое следствие D88 закрыто механизмом необязательной клетки: карта
    # пишет настройку программы в клетку округления книги, и книга считает
    # итог своей же формулой ROUND(..., +D88) -- по нашей разрядности, а не
    # всегда до целых рублей.
    context, ws = with_rounding(repo, digits, tmp_path)
    assert context["rounding"] == digits
    assert ws["D88"].value == digits
    assert evaluate(ws, "D26") == Decimal(book)


def test_rounding_to_thousands_brings_the_two_totals_together(repo, tmp_path):
    # то, ради чего D88 и нужна: округление у книги и у программы одно,
    # разница от страховых в него не попадает -- говорить не о чем
    context, ws = with_rounding(repo, -3, tmp_path)
    assert context["estimate"]["total_wo_vat"] == Decimal("2000.00")
    assert evaluate(ws, "D26") == Decimal("2000")
    assert context["warnings"] == []


def test_the_setting_is_not_blamed_when_the_book_rounds_the_same_way(repo):
    # Разрядность 2: итоги всё ещё расходятся на копейку, но причина уже не
    # настройка -- книга округляет итог так же, как программа, а остаются
    # неокруглённые страховые. Отправить человека менять настройку значило бы
    # отправить его чинить не то.
    context, _ = with_rounding(repo, 2)
    (text,) = context["warnings"]
    text = plain(text)
    assert "итог без НДС 1 507,49" in text and "1 507,50" in text
    assert "не округляет страховые взносы" in text
    assert "на границе округления" in text and "Исправить это нельзя" in text
    assert "Округление итога без НДС" not in text
    assert "целых рублей" not in text and "границе рубля" not in text


@pytest.mark.parametrize("digits, program", [(2, "1 507,50"), (-3, "2 000,00")],
                         ids=["до-копеек", "до-тысяч"])
def test_rounding_setting_is_named_as_the_cause_instead_of_the_rouble_boundary(
        repo, tmp_path, monkeypatch, digits, program):
    # Карта, которая клетку округления НЕ пишет, оставляет книгу на целых
    # рублях: тогда при любой разрядности, кроме нуля, причина расхождения --
    # сама настройка, а не граница рубля. Текст этой причины остаётся нужным:
    # карта любого другого заказчика может клетки округления не иметь.
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "1", "1002.45")])
    path = cellmap_without_rounding(tmp_path / "cellmaps")
    monkeypatch.setattr(config, "CELLMAPS_DIR", path.parent)
    repo.save_settings(replace(repo.get_settings(),
                               total_rounding_digits=digits))
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    (text,) = context["warnings"]
    text = plain(text)
    assert "итог без НДС 1 507,00" in text and program in text
    assert ("«Округление итога без НДС, знаков после запятой», "
            f"сейчас {digits}") in text
    assert "на границе рубля" not in text


def test_warning_names_the_rows_and_the_insurance_when_both_split_the_totals(
        repo, tmp_path):
    # строка материалов с долей копейки И итог у границы рубля из-за страховых:
    # назвать одну причину -- отправить человека чинить не то
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "120", "155.37")],
                             materials=[("расходные материалы", "1.5", "1234.57")])
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    (text,) = context["warnings"]
    text = plain(text)
    assert "итог без НДС 29 889,00" in text and "29 890,00" in text
    assert "«расходные материалы» (материалы)" in text
    assert "Кроме того, книга Excel не округляет страховые взносы" in text
    # строка ФОТ ровная (120 ч x 155,37 = 18 644,40): о ней речи нет
    assert "(трудозатраты)" not in text
    out = fill_workbook(TEMPLATE, load_cellmap(CELLMAP), context, tmp_path / "b.xlsx")
    assert evaluate(load_workbook(out)[SHEET], "D26") == Decimal("29889")


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


# ---- две бумаги с ценой: сличается НАПЕЧАТАННОЕ ---------------------------------------

# Цена «Коммерческого предложения» приходит из позиций лота, цена калькуляции --
# из калькулятора. Это два разных числа, которые человек вводит в разных местах
# карточки, и до этой сверки их не сравнивал никто: пакет уходил в комиссию
# с двумя разными ценами, и ни одного предупреждения в нём не было.

def kp_total(path: Path) -> str:
    """Итог без НДС, НАПЕЧАТАННЫЙ в коммерческом предложении: предпоследняя
    клетка строки «Итого». Читается из готового файла, а не из контекста."""
    row = docx.Document(path).tables[-1].rows[-1]
    return plain(row.cells[-2].text.strip())


def two_papers(repo, tmp_path, lot_price, documents=(DOC_KP, DOC_ESTIMATE)):
    """Пакет Арктикгаза (итог калькулятора ровно 4 000 000) с одной позицией
    лота на lot_price."""
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT,
                             lot=[("Разработка КЭР", lot_price)])
    return build_package(tid, list(documents), tmp_path / "out", repo)


def test_a_package_names_both_papers_and_both_numbers_when_prices_differ(
        repo, tmp_path):
    result = two_papers(repo, tmp_path, "3950000")
    assert result.skipped == []
    written = {p.suffix: p for p in result.written}
    # сперва убеждаемся, что бумаги действительно разные: 50 000 разницы
    assert kp_total(written[".docx"]) == "3 950 000,00"
    assert evaluate(load_workbook(written[".xlsx"])[SHEET], "D26") \
        == Decimal("4000000")
    (document, text), = [w for w in result.warnings if w[0] == PACKAGE_WIDE]
    text = plain(text)
    assert f"«{DOC_KP}» -- 3 950 000,00" in text
    assert f"«{DOC_ESTIMATE}» -- 4 000 000,00" in text


def test_the_same_price_in_both_papers_says_nothing(repo, tmp_path):
    # обратная сторона: сигнал, который горит всегда, через неделю перестают
    # читать. На ровных данных Арктикгаза молчат и прочие сторожа
    result = two_papers(repo, tmp_path, "4000000")
    assert result.skipped == [] and len(result.written) == 2
    written = {p.suffix: p for p in result.written}
    assert kp_total(written[".docx"]) == "4 000 000,00"
    workbook = load_workbook(written[".xlsx"])
    try:
        assert evaluate(workbook[SHEET], "D26") == Decimal("4000000")
    finally:
        workbook.close()
    assert result.warnings == []


def test_one_document_with_a_price_is_compared_with_nothing(repo, tmp_path):
    result = two_papers(repo, tmp_path, "3950000", documents=(DOC_KP,))
    assert [p.name for p in result.written] == [f"{DOC_KP}.docx"]
    assert result.warnings == []


def test_a_document_without_a_price_does_not_join_the_comparison(repo, tmp_path):
    # у декларации цены нет вовсе: назвать её в сверке -- соврать
    result = two_papers(repo, tmp_path, "3950000",
                        documents=(DOC_KP, DOC_DECLARATION))
    assert result.skipped == [] and len(result.written) == 2
    assert result.warnings == []


def test_a_commercial_offer_without_lot_items_is_skipped_not_written_at_zero(
        repo, tmp_path):
    # Коммерческое предложение на 0,00 руб. -- не документ. Раньше оно молча
    # уходило человеку и стояло в списке «Записано» как ни в чём не бывало:
    # слой Excel за пустую обязательную графу отказывается, а Word молчал.
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT)      # позиций лота нет
    result = build_package(tid, [DOC_KP, DOC_ESTIMATE], tmp_path / "out", repo)
    (document, reason), = result.skipped
    assert document == DOC_KP
    assert "позиции лота" in reason and "0,00" in reason
    assert "Traceback" not in reason and "Error" not in reason
    # остальные документы пакета собраны, и нулевого КП в папке нет
    assert [p.name for p in result.written] == [f"{DOC_ESTIMATE}.xlsx"]
    assert [p.name for p in (tmp_path / "out").iterdir()] == [f"{DOC_ESTIMATE}.xlsx"]


def test_one_lot_item_is_enough_to_let_the_offer_through(repo, tmp_path):
    # обратная сторона: напечатанная цена не ноль -- документ проходит
    result = two_papers(repo, tmp_path, "4000000", documents=(DOC_KP,))
    assert result.skipped == []
    assert [p.name for p in result.written] == [f"{DOC_KP}.docx"]


def test_an_offer_whose_lot_items_are_all_priced_at_zero_is_skipped_too(
        repo, tmp_path):
    """Соседняя дверь к тому же нулю: позиции есть, а цены у них нулевые.

    Цена позиции лота по умолчанию ноль, и позиции вставляют списком из
    таблицы заказчика без цен. Отказ держится на НАПЕЧАТАННОЙ цене, а не на
    длине списка, поэтому закрыты обе двери."""
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT,
                             lot=[("Разработка КЭР", "0"), ("Экспертиза", "0")])
    result = build_package(tid, [DOC_KP, DOC_ESTIMATE], tmp_path / "out", repo)
    (document, reason), = result.skipped
    assert document == DOC_KP
    assert "нулевой цене" in reason and "0,00" in reason
    assert "Traceback" not in reason and "Error" not in reason
    # остальные документы пакета собраны, и нулевого КП в папке нет
    assert [p.name for p in result.written] == [f"{DOC_ESTIMATE}.xlsx"]
    assert [p.name for p in (tmp_path / "out").iterdir()] == [f"{DOC_ESTIMATE}.xlsx"]


def test_a_single_priced_item_among_zeroes_lets_the_offer_through(repo, tmp_path):
    """Обратная сторона отказа по нулю: нулевая цена ОТДЕЛЬНОЙ позиции законна
    (позиция «в составе работ»), пока напечатанный итог не ноль. Итог читается
    из готового файла."""
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT,
                             lot=[("Разработка КЭР", "0"),
                                  ("Экспертиза", "4000000")])
    result = build_package(tid, [DOC_KP], tmp_path / "out", repo)
    assert result.skipped == []
    assert [p.name for p in result.written] == [f"{DOC_KP}.docx"]
    assert kp_total(result.written[0]) == "4 000 000,00"


# ---- бумага, оставшаяся в папке: её цена тоже сверяется --------------------------

# Дыра на пути пропуска. Человек правит позиции лота и пересобирает пакет, а
# калькуляция остаётся прежней, потому что «файл уже есть». В папке две бумаги
# с разными ценами, а предупреждений не было ни одного: цена снятого документа
# в сверку не попадала вовсе.

def package_wide(result):
    return [text for document, text in result.warnings if document == PACKAGE_WIDE]


def names(directory: Path) -> list[str]:
    return sorted(p.name for p in directory.iterdir())


def test_a_document_left_in_the_folder_joins_the_price_comparison(repo, tmp_path):
    out = tmp_path / "out"
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT,
                             lot=[("Разработка КЭР", "3950000")])
    first = build_package(tid, [DOC_ESTIMATE], out, repo)
    assert [p.name for p in first.written] == [f"{DOC_ESTIMATE}.xlsx"]
    assert package_wide(first) == []

    result = build_package(tid, [DOC_ESTIMATE, DOC_KP], out, repo)
    assert [d for d, _ in result.skipped] == [DOC_ESTIMATE]      # файл уже есть
    assert names(out) == sorted([f"{DOC_ESTIMATE}.xlsx", f"{DOC_KP}.docx"])
    (text,) = package_wide(result)
    text = plain(text)
    # напечатано сейчас -- это КП; калькуляция названа тем, что она есть
    assert f"«{DOC_KP}» -- 3 950 000,00" in text
    assert "не пересобрала" in text and "по нынешним данным" in text
    assert f"«{DOC_ESTIMATE}» -- 4 000 000,00" in text
    # про файл, который сборка не трогала, нельзя сказать «печатает»: цена
    # в нём прежняя, а 4 000 000 -- то, что он напечатал БЫ сейчас
    assert "Документы пакета печатают разную цену" not in text
    assert "разрешив замену существующих файлов" in text


def test_the_same_price_in_a_left_document_says_nothing(repo, tmp_path):
    # обратная сторона: цена одна на всю папку -- молчание, хотя один
    # документ и не пересобран
    out = tmp_path / "out"
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT,
                             lot=[("Разработка КЭР", "4000000")])
    build_package(tid, [DOC_ESTIMATE], out, repo)
    result = build_package(tid, [DOC_ESTIMATE, DOC_KP], out, repo)
    assert [d for d, _ in result.skipped] == [DOC_ESTIMATE]
    assert [p.name for p in result.written] == [f"{DOC_KP}.docx"]
    assert package_wide(result) == []


def test_the_comparison_does_not_name_a_paper_that_was_never_printed(
        repo, tmp_path, monkeypatch):
    """Зеркало главной находки: цена попадает в сверку только после того, как
    файл лёг в папку под своим именем. Шаблон КП испорчен -- в папке одна
    калькуляция, и назвать цену КП напечатанной значило бы соврать о бумаге,
    которой нет."""
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT,
                             lot=[("Разработка КЭР", "3950000")])
    # папка шаблонов подменяется после наполнения справочников: наполнение
    # проверяет, что каждый шаблон профилей на месте
    templates = tmp_path / "templates"
    templates.mkdir()
    shutil.copy(TEMPLATE, templates / TEMPLATE.name)
    (templates / f"{DOC_KP}.docx").write_bytes(b"not a zip")
    monkeypatch.setattr(config, "TEMPLATES_DOCS_DIR", templates)
    out = tmp_path / "out"
    result = build_package(tid, [DOC_KP, DOC_ESTIMATE], out, repo)
    (document, reason), = result.skipped
    assert document == DOC_KP and "повреждён" in reason
    assert names(out) == [f"{DOC_ESTIMATE}.xlsx"]
    assert package_wide(result) == []


def test_the_price_of_a_left_document_is_the_one_its_context_would_print(repo):
    """Цена снятого документа считается без контекста -- и обязана совпадать
    с тем, что напечатал бы контекст. Два разных счёта разошлись бы молча."""
    from app.render.context import printed_price
    tid = make_custom_tender(repo, labor=ARKTIK_LABOR, materials=ARKTIK_MATERIALS,
                             subcontract=ARKTIK_SUBCONTRACT,
                             lot=[("Разработка КЭР", "3950000")])
    for document, cellmap in ((DOC_ESTIMATE, CELLMAP), (DOC_KP, None)):
        cell_map = load_cellmap(cellmap) if cellmap else None
        context = build_context(tid, repo, document=document)
        assert printed_price(tid, repo, document, cell_map) \
            == context["printed_price_wo_vat"], document
    assert printed_price(tid, repo, DOC_KP, None) == Decimal("3950000.00")
    assert printed_price(tid, repo, DOC_ESTIMATE, load_cellmap(CELLMAP)) \
        == Decimal("4000000.00")
    # документ без цены в сверке не участвует
    assert printed_price(tid, repo, DOC_DECLARATION, None) is None


# ---- докстринг карты: пример должен работать, а не учить неверному ---------------

def test_the_optional_cell_example_from_the_docstring_fills_a_real_cell(
        repo, tmp_path):
    """Докстринг xlsx_render -- то, по чему карту правит не программист.
    Клетка из его примера обязана заполняться настоящим контекстом: карта,
    собранная по документированному примеру, не должна получать отказ
    «в данных нет значения». Документация, которая учит неверному, хуже
    отсутствующей."""
    from app.render import xlsx_render
    from app.render.xlsx_render import CellMap
    found = re.search(r'"D88":\s*(\{.*\})', xlsx_render.__doc__)
    assert found, "в докстринге больше нет примера необязательной клетки"
    example = json.loads(found.group(1))
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "1", "1002.45")])
    repo.save_settings(replace(repo.get_settings(), total_rounding_digits=2))
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    cell_map = CellMap(sheet=SHEET, cells={"D88": example["value"]},
                       optional={"D88"} if example.get("optional") else set())
    out = fill_workbook(TEMPLATE, cell_map, context, tmp_path / "example.xlsx")
    assert load_workbook(out)[SHEET]["D88"].value == 2


# ---- клетка округления -- клетка, а не упоминание в тексте -----------------------

def cellmap_with_rounding_in_text(folder: Path) -> Path:
    """Карта, которая разрядность округления УПОМИНАЕТ в текстовой клетке, а
    клетку округления книги (D88) не пишет. Книга остаётся на целых рублях:
    из текста «Округление: 2 знаков» формула ROUND разрядности не возьмёт."""
    data = json.loads(CELLMAP.read_text(encoding="utf-8"))
    del data["cells"]["D88"]
    data["cells"]["A8"] = "Округление итога: {rounding} знаков"
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / CELLMAP.name
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return path


def test_a_rounding_mentioned_in_a_text_cell_is_not_the_rounding_cell(
        repo, tmp_path, monkeypatch):
    # Сторож смотрел на наличие пути «rounding» в ЛЮБОЙ клетке карты -- и на
    # такой карте решил бы, что книга округляет итог как программа. Причина
    # расхождения была бы названа неверно, то есть человека отправили бы
    # чинить не то.
    tid = make_custom_tender(repo, labor=[("Эколог", 1, "1", "1002.45")])
    path = cellmap_with_rounding_in_text(tmp_path / "cellmaps")
    monkeypatch.setattr(config, "CELLMAPS_DIR", path.parent)
    repo.save_settings(replace(repo.get_settings(), total_rounding_digits=2))
    context = build_context(tid, repo, document=DOC_ESTIMATE)
    (text,) = context["warnings"]
    text = plain(text)
    # книга на целых рублях: 1 507 против 1 507,50 у программы
    assert "итог без НДС 1 507,00" in text and "1 507,50" in text
    assert ("«Округление итога без НДС, знаков после запятой», "
            "сейчас 2") in text
    assert "на границе округления" not in text      # причина не та


def test_the_real_map_still_writes_the_rounding_cell(repo, tmp_path):
    # обратная сторона: настоящая карта пишет D88 отдельной клеткой из одного
    # «{rounding}» -- и причина расхождения остаётся страховыми, а не настройкой
    context, ws = with_rounding(repo, 2, tmp_path)
    assert ws["D88"].value == 2
    (text,) = context["warnings"]
    assert "на границе округления" in plain(text)
    assert "Округление итога без НДС" not in text
