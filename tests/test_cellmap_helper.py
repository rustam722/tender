"""Помощник по формам заказчика: чтение книги и достижимость клеток.

Две настоящие книги из `templates_docs/` (КП Арктикгаза и ТКП Роснефти) только
читаются -- это шаблоны фирмы. Всё, что требует особого устройства книги
(русские формулы, круг в ссылках, кэш значений, два листа), собирается тестом
на месте.

Главное здесь -- достижимость. В форме КП Арктикгаза таблица субподряда
нарисована на шесть строк, а в итог идут три: `D21 = G54`,
`G54 = СУММ(G48:G50)`, и в `G48:G50` попадают только строки 60-62. Четвёртый
субподрядчик молча выпал бы из цены. Это нашли руками, читая формулы книги,
и тесты ниже требуют, чтобы помощник находил это сам -- в том числе через
промежуточную клетку, а не на один шаг.
"""
import json
import re
import threading
import zipfile
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook

from app import config
from app.cellmap_helper import (
    CellView, SheetView, WorkbookError, read_workbook, unreachable_cells,
)

ARKTIKGAZ = config.TEMPLATES_DOCS_DIR / "Калькуляция к КП Арктикгаз.xlsx"
ARKTIKGAZ_MAP = config.CELLMAPS_DIR / "kp_arktikgaz.json"
ROSNEFT = config.TEMPLATES_DOCS_DIR / "ТКП по форме заказчика Роснефть Мурманск.xlsx"
ROSNEFT_MAP = config.CELLMAPS_DIR / "tkp_rosneft_murmansk.json"
SHEET = "Лист1"


@pytest.fixture(autouse=True)
def the_real_forms_are_only_read():
    """Сторож на каждый тест файла: настоящие формы фирмы не меняются.

    Помощник открывает книгу заказчика, чтобы рассказать о ней, -- и обязан
    оставить файл как был: это шаблон, из которого фирма собирает заявки.
    Сторож стоит на каждом тесте, а не в одном из них: тест, читающий
    настоящую форму, своей проверки сохранения не содержит, а поймать такое
    сохранение надо на том тесте, который его вызвал.
    """
    before = {book: book.read_bytes() for book in (ARKTIKGAZ, ROSNEFT)}
    yield
    for book, content in before.items():
        assert book.read_bytes() == content, f"тест изменил форму {book.name}"


# ---- книги, которые собирает тест ----------------------------------------------

def make_book(path, cells, *, sheet=SHEET, other=None, other_cells=None):
    """Книга из перечисленных клеток; `other` -- второй лист, если он нужен."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    for address, value in cells.items():
        ws[address] = value
    if other is not None:
        second = wb.create_sheet(other)
        for address, value in (other_cells or {}).items():
            second[address] = value
    wb.save(path)
    return path


def with_cached_value(path, out, address, cached):
    """Та же книга, но с вычисленным значением формулы, как его сохраняет Excel.

    openpyxl формул не считает и кэш значений не пишет (в его файле у клетки
    с формулой `<v></v>`), а второй проход книги (`data_only=True`) читает
    именно этот кэш. Поэтому значение вписывается в XML здесь -- иначе
    проверить второй проход было бы нечем.
    """
    with zipfile.ZipFile(path) as book:
        items = [(info, book.read(info.filename)) for info in book.infolist()]
    hole = re.compile(rf'(<c r="{address}"[^>]*>(?:<f>[^<]*</f>)?)<v></v>')
    patched = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as book:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                text, patched = hole.subn(rf"\g<1><v>{cached}</v>", data.decode())
                data = text.encode()
            book.writestr(info, data)
    assert patched == 1, f"кэш значения {address} в книгу не вписался"
    return out


def map_addresses(path: Path) -> list[str]:
    """Все адреса, которые размечает написанная руками карта ячеек."""
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    addresses = list(data["cells"]) + list(data.get("rates", {}).values())
    for block in data.get("rows", []):
        for letter in block["columns"]:
            for row in range(block["first_row"], block["last_row"] + 1):
                addresses.append(f"{letter}{row}")
    return addresses


# Клетки написанной руками карты Арктикгаза, в которых стоят числа: количества,
# цены, ставки калькуляции и разрядность округления итога. Все они обязаны
# доходить до итогов книги -- иначе цена считается не по тому, что человек
# ввёл.
ARKTIKGAZ_MONEY = sorted(
    {"C17", "C18", "C24", "C25", "D88"}
    | {f"{letter}{row}" for letter in "DE" for row in range(35, 39)}      # материалы
    | {f"{letter}{row}" for letter in "CDE" for row in range(48, 54)}     # труд
    | {f"{letter}{row}" for letter in "EF" for row in range(60, 63)})     # субподряд

# Клетки той же карты, в которые книга не смотрит ни одной формулой: это её
# текст -- шапка, наименования, единицы, примечания и подпись. Помощник назовёт
# их недостижимыми, и это правда: в цену они не входят, а в документ -- входят.
ARKTIKGAZ_TEXT = sorted(
    {"D2", "A4", "A5", "A6", "A7", "A69", "E72"}
    | {f"{letter}{row}" for letter in "BC" for row in range(35, 39)}
    | {f"B{row}" for row in range(48, 54)}
    | {f"{letter}{row}" for letter in "BCD" for row in range(60, 63)})


# ---- вид книги -----------------------------------------------------------------

def test_the_view_of_a_workbook_shows_the_text_formula_and_address_of_a_cell(
        tmp_path):
    book = make_book(tmp_path / "f.xlsx", {
        "A1": "Субподряд", "E60": 2, "F60": 1002.45, "G48": "=F60*E60"})
    sheet, = read_workbook(book)
    cells = {cell.address: cell for cell in sheet.cells}
    assert sheet.name == SHEET
    assert cells["A1"] == CellView(address="A1", row=1, column="A",
                                   text="Субподряд", value="Субподряд",
                                   formula=None)
    assert cells["F60"].text == "1002,45" and cells["F60"].value == Decimal("1002.45")
    assert cells["G48"].formula == "=F60*E60"
    assert cells["G48"].address == "G48" and cells["G48"].row == 48
    assert cells["G48"].column == "G"


def test_a_cell_that_holds_a_formula_is_marked_as_one(tmp_path):
    book = make_book(tmp_path / "f.xlsx", {
        "A1": "Итого", "E60": 2, "G48": "=F60*E60", "G54": "=СУММ(G48:G50)"})
    sheet, = read_workbook(book)
    marked = {cell.address for cell in sheet.cells if cell.is_formula}
    assert marked == {"G48", "G54"}
    for cell in sheet.cells:
        assert cell.is_formula == (cell.formula is not None), cell


def test_a_number_of_the_book_is_a_decimal_and_is_shown_the_russian_way(tmp_path):
    book = make_book(tmp_path / "f.xlsx", {"C17": 0.3, "C18": 1, "C19": "0,3"})
    sheet, = read_workbook(book)
    cells = {cell.address: cell for cell in sheet.cells}
    assert cells["C17"].value == Decimal("0.3")
    assert isinstance(cells["C17"].value, Decimal)
    assert cells["C17"].text == "0,3"
    assert cells["C18"].value == Decimal(1) and cells["C18"].text == "1"
    # текст «0,3» остаётся текстом: что в клетке лежит, то и показано
    assert cells["C19"].value == "0,3" and not isinstance(cells["C19"].value, Decimal)


def test_the_value_of_a_cell_with_a_formula_is_the_one_excel_saved_in_the_file(
        tmp_path):
    """Второй проход книги (data_only=True) читает кэш значений Excel.

    Без него клетка с формулой осталась бы для человека пустой, и он не увидел
    бы, что книга в ней насчитала."""
    plain = make_book(tmp_path / "f.xlsx", {"E60": 2, "F60": 3, "G48": "=F60*E60"})
    book = with_cached_value(plain, tmp_path / "cached.xlsx", "G48", 6)
    sheet, = read_workbook(book)
    cells = {cell.address: cell for cell in sheet.cells}
    assert cells["G48"].formula == "=F60*E60"
    assert cells["G48"].value == Decimal(6) and cells["G48"].text == "6"


def test_a_cell_with_a_formula_the_book_never_counted_is_shown_without_a_value(
        tmp_path):
    """В книге заказчика кэша значений может не быть вовсе (так в обеих
    настоящих формах): тогда показывается формула, а значение пусто -- врать
    о посчитанном нельзя."""
    book = make_book(tmp_path / "f.xlsx", {"E60": 2, "G48": "=F60*E60"})
    sheet, = read_workbook(book)
    cells = {cell.address: cell for cell in sheet.cells}
    assert cells["G48"].value is None and cells["G48"].text == ""


def test_an_empty_cell_is_not_in_the_view_but_the_size_of_the_sheet_is(tmp_path):
    """Пустых клеток в листе заказчика десятки тысяч, перечислять их незачем;
    а размер листа нужен -- заполнять человек будет как раз пустые клетки."""
    book = make_book(tmp_path / "f.xlsx", {"A1": "Шапка", "C5": 7})
    sheet, = read_workbook(book)
    assert [cell.address for cell in sheet.cells] == ["A1", "C5"]
    assert (sheet.max_row, sheet.max_column) == (5, 3)


def test_the_view_names_every_sheet_of_the_book_in_the_order_of_the_book(tmp_path):
    book = make_book(tmp_path / "f.xlsx", {"A1": "Первый"},
                     other="Расшифровка", other_cells={"B2": "Второй"})
    sheets = read_workbook(book)
    assert [sheet.name for sheet in sheets] == [SHEET, "Расшифровка"]
    assert [cell.text for cell in sheets[1].cells] == ["Второй"]
    assert isinstance(sheets[0], SheetView)


def test_the_view_of_the_arktikgaz_book_shows_its_rate_and_its_total_formula():
    """Настоящая книга: ставка страховых взносов, формула итога и шапка."""
    sheets = read_workbook(ARKTIKGAZ)
    assert [sheet.name for sheet in sheets] == [SHEET]
    cells = {cell.address: cell for cell in sheets[0].cells}
    assert cells["A1"].text == "Коммерческое предложение"
    assert cells["C17"].value == Decimal("0.3") and cells["C17"].text == "0,3"
    assert cells["C17"].formula is None
    assert cells["D26"].formula == "=ROUND(D11+D23+D24+D25,+D88)"
    assert cells["G54"].formula == "=SUM(G48:G50)"
    assert "F63" not in cells             # четвёртая строка субподряда пуста
    assert (sheets[0].max_row, sheets[0].max_column) == (77, 10)


def test_reading_a_workbook_does_not_change_its_file(tmp_path):
    """Шаблон заказчика только читается: помощник открывает его, а не правит.

    Книга здесь своя, а не копия настоящей формы: проверка не должна зависеть
    от того, читал ли настоящую форму какой-нибудь тест до неё."""
    book = make_book(tmp_path / "f.xlsx", {
        "A1": "Субподряд", "E60": 2, "G48": "=F60*E60"})
    before = book.read_bytes()
    read_workbook(book)
    unreachable_cells(book, ["F63"])
    assert book.read_bytes() == before
    copy = tmp_path / "Калькуляция.xlsx"
    copy.write_bytes(ARKTIKGAZ.read_bytes())
    was = copy.read_bytes()
    read_workbook(copy)
    unreachable_cells(copy, ["F63"])
    assert copy.read_bytes() == was


# ---- достижимость --------------------------------------------------------------

def test_a_cell_no_formula_ever_reads_is_named_unreachable():
    """Настоящая форма Арктикгаза: субподряд нарисован на шесть строк, а
    в итог идут три (D21 = G54, G54 = СУММ(G48:G50), G48 = F60*E60).

    Строки 63-65 не складывает ни одна формула книги: записанный туда
    четвёртый субподрядчик молча выпал бы из цены."""
    asked = [f"{letter}{row}" for row in range(60, 66) for letter in "EF"]
    assert unreachable_cells(ARKTIKGAZ, asked) == [
        "E63", "F63", "E64", "F64", "E65", "F65"]


def test_every_money_cell_of_the_hand_written_arktikgaz_map_is_read_by_the_book():
    """Обратная сторона той же проверки: ни одной ложной тревоги на карте,
    которую написали руками и которой фирма пользуется."""
    assert unreachable_cells(ARKTIKGAZ, ARKTIKGAZ_MONEY) == []
    # и тут же -- что пустой ответ получается не всегда: строка за блоком
    # материалов (39) и четвёртая строка субподряда (63) в итог не идут
    assert unreachable_cells(ARKTIKGAZ, ["D39", "E39", "E63", "F63"]) == [
        "D39", "E39", "E63", "F63"]


def test_the_unreachable_cells_of_the_arktikgaz_map_are_exactly_its_text_cells():
    found = unreachable_cells(ARKTIKGAZ, map_addresses(ARKTIKGAZ_MAP))
    assert sorted(found) == ARKTIKGAZ_TEXT


def test_in_the_rosneft_form_a_price_below_the_summed_block_is_unreachable():
    """Вторая настоящая книга: итог складывает C9:C19, и цена, записанная
    строкой ниже блока, в него не попадает."""
    prices = [f"{letter}{row}" for row in range(9, 22) for letter in "CD"]
    assert unreachable_cells(ROSNEFT, prices) == ["C20", "D20", "C21", "D21"]
    assert unreachable_cells(ROSNEFT, map_addresses(ROSNEFT_MAP)) == [
        "D8", "E23", "E24", "E25", "E26", "E27", "E28", "E29", "B31"]


def test_a_cell_inside_a_summed_range_is_not_called_unreachable(tmp_path):
    book = make_book(tmp_path / "f.xlsx", {"G54": "=СУММ(G48:G50)"})
    assert unreachable_cells(book, ["G48", "G49", "G50"]) == []
    assert unreachable_cells(book, ["G51"]) == ["G51"]


def test_an_indirect_sum_through_another_cell_is_followed(tmp_path):
    """D21 = G54, G54 = СУММ(G48:G50): до G49 обход идёт через промежуточную
    клетку. Обход в один шаг (только прямые ссылки итога D21) назвал бы G49
    недостижимой -- то есть соврал бы в самую опасную сторону."""
    book = make_book(tmp_path / "f.xlsx", {
        "D21": "=G54", "G54": "=СУММ(G48:G50)"})
    assert unreachable_cells(book, ["G54", "G48", "G49", "G50"]) == []
    # пустой ответ -- не ответ по умолчанию: клетку за диапазоном помощник
    # называет, и значит, пустота выше получена обходом, а не молчанием
    assert unreachable_cells(book, ["G51", "F60"]) == ["G51", "F60"]


def test_a_chain_of_four_steps_to_a_row_of_the_table_is_followed(tmp_path):
    """Как в книге Арктикгаза целиком: D26 -> D23 -> D21 -> G54 -> G48 -> F60."""
    book = make_book(tmp_path / "f.xlsx", {
        "D26": "=ОКРУГЛ(D23,2)", "D23": "=D21", "D21": "=G54",
        "G54": "=СУММ(G48:G50)", "G48": "=F60*E60"})
    assert unreachable_cells(book, ["F60", "E60"]) == []
    assert unreachable_cells(book, ["F61", "E61"]) == ["F61", "E61"]


def test_a_russian_formula_name_and_an_english_one_are_read_the_same_way(tmp_path):
    """Книги заказчиков приходят и с СУММ, и с SUM (в обеих формах фирмы --
    английские). Помощник смотрит на ссылки, а не на имя функции."""
    russian = make_book(tmp_path / "ru.xlsx", {
        "D21": "=G54", "G54": "=СУММ(G48:G50)*ОКРУГЛ(C24;2)"})
    english = make_book(tmp_path / "en.xlsx", {
        "D21": "=G54", "G54": "=SUM(G48:G50)*ROUND(C24,2)"})
    asked = ["G48", "G49", "G50", "C24", "G51"]
    assert unreachable_cells(russian, asked) == ["G51"]
    assert unreachable_cells(english, asked) == ["G51"]


def test_an_absolute_address_is_followed_like_a_plain_one(tmp_path):
    book = make_book(tmp_path / "f.xlsx", {
        "D21": "=$G$54", "G54": "=СУММ($G$48:$G$50)"})
    assert unreachable_cells(book, ["G48", "G49", "G50", "G54"]) == []
    assert unreachable_cells(book, ["G51"]) == ["G51"]


def test_a_formula_that_reads_itself_still_counts_as_a_total(tmp_path):
    """Битая книга: A1 = A1+B1 (Excel такое зовёт циклической ссылкой).
    Клетка, которая читает себя, остаётся итогом -- иначе B1 назвалась бы
    недостижимой, хотя книга её складывает."""
    book = make_book(tmp_path / "f.xlsx", {"A1": "=A1+B1"})
    assert unreachable_cells(book, ["B1", "C1"]) == ["C1"]


def test_a_cell_read_by_a_formula_of_another_sheet_is_reachable(tmp_path):
    """Итог может лежать на другом листе книги: 'Лист1'!G48 -- та же ссылка."""
    book = make_book(tmp_path / "f.xlsx", {"A1": "Субподряд"},
                     other="Свод",
                     other_cells={"B2": "=СУММ('Лист1'!G48:G50)+Лист1!G60"})
    assert unreachable_cells(book, ["G48", "G49", "G60"], sheet=SHEET) == []
    assert unreachable_cells(book, ["G61"], sheet=SHEET) == ["G61"]


def test_a_whole_column_in_a_formula_covers_the_cells_of_that_column(tmp_path):
    book = make_book(tmp_path / "f.xlsx", {"G1": "=СУММ(F:F)+СУММ(9:9)"})
    assert unreachable_cells(book, ["F1", "F1000", "B9", "C9"]) == []
    assert unreachable_cells(book, ["E8", "G8"]) == ["E8", "G8"]


def test_a_cell_named_only_in_the_text_of_a_formula_is_not_called_read(tmp_path):
    """В кавычках -- текст, а не ссылка: ="Итого F63" клетку F63 не складывает."""
    book = make_book(tmp_path / "f.xlsx", {"A1": '="Итого F63 и "&"E63"'})
    assert unreachable_cells(book, ["F63", "E63"]) == ["F63", "E63"]


def test_a_function_name_with_a_digit_is_not_taken_for_a_cell(tmp_path):
    """LOG10( -- имя функции, а колонка LOG существует: за ссылку её принимать
    нельзя, иначе помощник промолчит о клетке, которую никто не читает."""
    book = make_book(tmp_path / "f.xlsx", {"A1": "=LOG10(B2)"})
    assert unreachable_cells(book, ["LOG10", "B2"]) == ["LOG10"]


def test_a_reference_to_another_workbook_does_not_make_a_local_cell_reachable(
        tmp_path):
    """=[1]Лист1!A1 читает чужую книгу, а не нашу клетку A1 того же имени."""
    book = make_book(tmp_path / "f.xlsx", {"B1": "=[1]Лист1!A1+5"})
    assert unreachable_cells(book, ["A1"]) == ["A1"]


def test_a_circle_of_references_does_not_hang_the_helper(tmp_path):
    """Книга заказчика бывает битой: G48 = G49, а G49 = G48 + F60. Обход
    такого круга обязан закончиться и отдать ответ, а не зависнуть.

    Круг лежит ПОД итогом (D21 -> G54 -> G48:G50), иначе обход в него и не
    зашёл бы: клетки, которых не читает никто, итогами и остаются, и тест на
    круге в стороне ничего бы не доказывал.

    Ответа ждёт отдельный поток: зависший обход должен дать красный тест,
    а не остановленный прогон, по которому не видно, что именно сломалось."""
    book = make_book(tmp_path / "f.xlsx", {
        "D21": "=G54", "G54": "=СУММ(G48:G50)", "G48": "=G49",
        "G49": "=G48+F60"})
    answer = []
    worker = threading.Thread(
        target=lambda: answer.append(
            unreachable_cells(book, ["G49", "F60", "F61"])),
        daemon=True)
    worker.start()
    worker.join(15)
    assert not worker.is_alive(), "обход круга в ссылках не закончился"
    # и ответ -- не пустая отговорка: клетку под кругом помощник нашёл (F60
    # читает G49), а клетку вне книги назвал
    assert answer == [["F61"]]


def test_the_answer_keeps_the_order_of_the_asked_addresses_and_names_each_once(
        tmp_path):
    book = make_book(tmp_path / "f.xlsx", {"G54": "=СУММ(G48:G50)"})
    assert unreachable_cells(book, ["F63", "G49", "E63", "F63"]) == ["F63", "E63"]
    assert unreachable_cells(book, []) == []


# ---- отказы ---------------------------------------------------------------------

# Обрывки английских сообщений: ни одного из них человек видеть не должен.
_ENGLISH_LEAK = re.compile(
    r"Error|Errno|Traceback|must be|expected|invalid|NoneType|BadZip")


def assert_russian_reason(text: str) -> None:
    assert re.search(r"[а-яё]{4,}", text, re.IGNORECASE), text
    assert not _ENGLISH_LEAK.search(text), text


def test_a_workbook_that_cannot_be_opened_is_refused_in_russian(tmp_path):
    missing = tmp_path / "нет.xlsx"
    garbage = tmp_path / "bad.xlsx"
    garbage.write_text("это не книга", encoding="utf-8")
    folder = tmp_path / "папка.xlsx"
    folder.mkdir()
    for book, fragment in ((missing, "нет.xlsx"), (garbage, "bad.xlsx"),
                           (folder, "папка.xlsx")):
        for call in (lambda: read_workbook(book),
                     lambda: unreachable_cells(book, ["A1"])):
            with pytest.raises(WorkbookError) as caught:
                call()
            text = str(caught.value)
            assert fragment in text, text
            assert_russian_reason(text)
    assert "не найдена" in str(
        pytest.raises(WorkbookError, read_workbook, missing).value)


def test_a_book_with_several_sheets_asks_which_sheet_the_cells_are_on(tmp_path):
    """Выбрать лист самому -- значит проверить достижимость не на том листе
    и промолчать о потерянной цене."""
    book = make_book(tmp_path / "f.xlsx", {"G54": "=СУММ(G48:G50)"},
                     other="Свод")
    with pytest.raises(WorkbookError) as caught:
        unreachable_cells(book, ["G49"])
    text = str(caught.value)
    assert SHEET in text and "Свод" in text
    assert_russian_reason(text)
    assert unreachable_cells(book, ["G49"], sheet=SHEET) == []


def test_an_unknown_sheet_is_refused_and_names_the_sheets_the_book_has(tmp_path):
    book = make_book(tmp_path / "f.xlsx", {"G54": "=СУММ(G48:G50)"})
    with pytest.raises(WorkbookError) as caught:
        unreachable_cells(book, ["G49"], sheet="Нет такого")
    text = str(caught.value)
    assert "Нет такого" in text and SHEET in text
    assert_russian_reason(text)


@pytest.mark.parametrize("address", ["g49", "G", "49", "G0", "", "G49:G50",
                                     "ЦЕНА1", "XFE1", 49])
def test_an_address_that_is_not_a_cell_address_is_refused_in_russian(
        tmp_path, address):
    book = make_book(tmp_path / "f.xlsx", {"G54": "=СУММ(G48:G50)"})
    with pytest.raises(WorkbookError) as caught:
        unreachable_cells(book, ["G49", address])
    assert_russian_reason(str(caught.value))
    assert str(address) in str(caught.value)
