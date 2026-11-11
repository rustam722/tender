"""Помощник по новой форме заказчика: чтение его книги Excel.

Карту ячеек (папка `cellmaps/`, формат описан в `app.render.xlsx_render`)
сейчас пишет руками человек, глядя в книгу заказчика в Excel. Это единственное
место, где без программиста не обойтись. Здесь -- первая половина помощника,
который это убирает: она открывает чужую книгу и рассказывает о ней две вещи.

1. **Что в книге лежит** (`read_workbook`): листы, а в них клетки с текстом,
   числом или формулой, каждая со своим адресом. Книга читается двумя
   проходами openpyxl, как в `app.render.xlsx_render`: `data_only=False`
   отдаёт формулы, `data_only=True` -- значения, которые Excel сохранил
   в файле. Сами формулы здесь не вычисляются: считает их Excel, а помощник
   только показывает и объясняет.

2. **Какие из размечаемых клеток книга не читает** (`unreachable_cells`).
   Это самое ценное, что помощник умеет. В форме КП Арктикгаза таблица
   субподряда нарисована на шесть строк, а в итог идут три: `D21 = G54`,
   `G54 = СУММ(G48:G50)`, и в `G48:G50` попадают только строки 60-62.
   Четвёртый субподрядчик, записанный в такую форму, молча выпал бы из цены:
   заявка ушла бы заказчику дешевле, чем считал участник, и увидел бы он это
   от проверяющего. Это нашли руками, читая формулы книги; `unreachable_cells`
   находит такое сама.

Как считается достижимость. Достижимой называется клетка, значение которой
доходит до одного из **итогов книги** -- клеток с формулой, которые дальше не
читает никто (их заказчик и печатает, и пересчитывает). Обход идёт от итогов
вниз по ссылкам и через любое число промежуточных клеток: `D21` ведёт к `G54`,
`G54` -- к диапазону `G48:G50`, оттуда формулы `G48 = F60*E60` ведут к строкам
субподряда. Обход в один шаг (только прямые ссылки итогов) назвал бы `F60`
недостижимой, то есть соврал бы в самую опасную сторону. Пройденные клетки
запоминаются, поэтому круг в ссылках (битая книга заказчика: `A1 = A2`,
`A2 = A1`) обход не вешает -- он заканчивается и отдаёт ответ.

Что из книги в достижимость НЕ входит намеренно: условное форматирование,
проверка данных, диаграммы и ссылки на другие книги (`[1]Лист1!A1`). Они не
складывают цену, а ссылка на чужую книгу в форме, которую заказчик прислал
участнику, -- не то, на что можно положиться.

Отказы -- `WorkbookError`, текст по-русски и говорит, что сделать: карту
правит не программист, и он же прикладывает книгу.
"""
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter


class WorkbookError(Exception):
    """Отказ помощника по формам; текст на русском, для человека."""


# ---- вид книги -----------------------------------------------------------------

@dataclass(frozen=True)
class CellView:
    """Клетка книги заказчика так, как её показывают человеку.

    `text` -- что в клетке видно: текст или число самой клетки, а у клетки
    с формулой -- значение, которое Excel сохранил в файле (его может не быть,
    тогда текст пуст, а формула видна в `formula`). Числа -- `Decimal`:
    `float` в программе запрещён, и 0.3 из книги не должно превратиться
    в 0.30000000000000004 по дороге к человеку.
    """
    address: str                  # 'G54'
    row: int                      # 54
    column: str                   # 'G'
    text: str                     # что видно в клетке ('0,3', 'Итого:', '')
    value: object = None          # значение клетки: Decimal, str, date, bool, None
    formula: str | None = None    # '=СУММ(G48:G50)' или None, если формулы нет

    @property
    def is_formula(self) -> bool:
        """Клетка занята формулой книги: заполнять её нельзя -- затрёт расчёт."""
        return self.formula is not None


@dataclass(frozen=True)
class SheetView:
    """Лист книги: непустые клетки и размер использованной области.

    В `cells` только клетки, в которых что-то есть (текст, число, формула):
    пустых клеток в листе заказчика десятки тысяч, и перечислять их незачем.
    Клетки, которые человек и будет заполнять, как раз пусты, поэтому экран
    рисует сетку по `max_row` и `max_column`, а подписи и формулы берёт
    из `cells`.
    """
    name: str
    cells: list[CellView] = field(default_factory=list)
    max_row: int = 0
    max_column: int = 0


_ADDRESS = re.compile(r"([A-Z]{1,3})([1-9][0-9]*)")
_MAX_COLUMN = 16384             # XFD, последняя колонка Excel


def _number(value: object) -> object:
    """Число из книги -- в Decimal; всё прочее -- как есть."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return value
    return Decimal(str(value))


def _text(value: object) -> str:
    """Значение клетки -> что видит человек. Любой текст -- по-русски."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "ДА" if value else "НЕТ"
    if isinstance(value, Decimal):
        if not value.is_finite():
            return str(value)
        return format(value.normalize(), "f").replace(".", ",")
    if isinstance(value, datetime):
        if (value.hour, value.minute, value.second) == (0, 0, 0):
            return value.strftime("%d.%m.%Y")
        return value.strftime("%d.%m.%Y %H:%M")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    return str(value)


def _formula_text(cell) -> str | None:
    """Формула клетки строкой или None. Формула массива отдаётся своим текстом."""
    if cell.data_type != "f":
        return None
    raw = cell.value
    if isinstance(raw, str):
        return raw
    return getattr(raw, "text", None) or str(raw)


def _open(path: Path, *, values: bool):
    """Книга заказчика двумя проходами: values=False -- формулы, True -- значения."""
    try:
        return load_workbook(path, data_only=values)
    except FileNotFoundError:
        raise WorkbookError(
            f"Книга {path.name} не найдена в папке {path.parent.name}. "
            "Проверьте путь к файлу.") from None
    except IsADirectoryError:
        raise WorkbookError(
            f"{path.name} -- это папка, а не книга Excel (.xlsx). "
            "Выберите файл книги.") from None
    except Exception as exc:        # BadZipFile, KeyError, InvalidFileException...
        raise WorkbookError(
            f"Не удалось открыть книгу {path.name}: файл повреждён или не "
            "является книгой Excel (.xlsx). Если это старый .xls, пересохраните "
            "его в Excel как .xlsx.") from exc


def _sheet_view(formulas_sheet, values_sheet) -> SheetView:
    cells: list[CellView] = []
    for row in formulas_sheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue            # пустая клетка (и не первая клетка объединения)
            formula = _formula_text(cell)
            shown = values_sheet[cell.coordinate].value if formula else cell.value
            value = _number(shown)
            cells.append(CellView(
                address=cell.coordinate, row=cell.row,
                column=get_column_letter(cell.column),
                text=_text(value), value=value, formula=formula))
    return SheetView(name=formulas_sheet.title, cells=cells,
                     max_row=formulas_sheet.max_row,
                     max_column=formulas_sheet.max_column)


def read_workbook(path) -> list[SheetView]:
    """Листы книги заказчика с её непустыми клетками, в порядке листов книги."""
    path = Path(path)
    with_formulas = _open(path, values=False)
    try:
        with_values = _open(path, values=True)
        try:
            return [_sheet_view(with_formulas[name], with_values[name])
                    for name in with_formulas.sheetnames]
        finally:
            with_values.close()
    finally:
        with_formulas.close()


# ---- достижимость --------------------------------------------------------------

# Ссылка в формуле: необязательное имя листа ('Лист 1'!/Лист1!), затем клетка
# (D8, $D$8), диапазон клеток (D8:E9), диапазон колонок (D:E) или строк (8:9).
# Запрет слева и справа отсекает то, что на ссылку только похоже: имя функции
# с цифрой (LOG10), ссылку на другую книгу ([1]Лист1!A1), часть слова.
_REF = re.compile(
    r"(?<![A-Za-z0-9_$.!\[\]Ѐ-ӿ])"
    r"(?:(?:'(?P<quoted>(?:[^']|'')*)'"
    r"|(?P<plain>[A-Za-z_Ѐ-ӿ][A-Za-z0-9_.Ѐ-ӿ]*))!)?"
    r"(?:"
    r"\$?(?P<col1>[A-Za-z]{1,3})\$?(?P<row1>[1-9][0-9]*)"
    r"(?::\$?(?P<col2>[A-Za-z]{1,3})\$?(?P<row2>[1-9][0-9]*))?"
    r"|\$?(?P<wholecol1>[A-Za-z]{1,3}):\$?(?P<wholecol2>[A-Za-z]{1,3})"
    r"|\$?(?P<wholerow1>[1-9][0-9]*):\$?(?P<wholerow2>[1-9][0-9]*)"
    r")"
    r"(?![A-Za-z0-9_(Ѐ-ӿ])")

# Текст в кавычках внутри формулы: ="Итого D21" ссылкой не является.
_STRING = re.compile(r'"(?:[^"]|"")*"')


@dataclass(frozen=True)
class _Rect:
    """Прямоугольник ссылки. None в границе -- вся колонка (A:A) или вся строка."""
    sheet: str
    first_col: int | None
    last_col: int | None
    first_row: int | None
    last_row: int | None

    def holds(self, sheet: str, col: int, row: int) -> bool:
        if sheet != self.sheet:
            return False
        if self.first_col is not None and not self.first_col <= col <= self.last_col:
            return False
        if self.first_row is not None and not self.first_row <= row <= self.last_row:
            return False
        return True


def _column(letter: str) -> int | None:
    try:
        index = column_index_from_string(letter.upper())
    except ValueError:
        return None
    return index if 1 <= index <= _MAX_COLUMN else None


def _sheet_of(match, own: str, names: list[str]) -> str | None:
    """Лист ссылки: свой лист формулы, названный лист книги или None (чужая книга)."""
    raw = match.group("quoted")
    if raw is None:
        raw = match.group("plain")
    if raw is None:
        return own
    name = raw.replace("''", "'")
    if name in names:
        return name
    lowered = {other.lower(): other for other in names}
    return lowered.get(name.lower())


def _references(formula: str, own: str, names: list[str]) -> list[_Rect]:
    """Прямоугольники, на которые смотрит формула."""
    rects: list[_Rect] = []
    for match in _REF.finditer(_STRING.sub("", formula)):
        sheet = _sheet_of(match, own, names)
        if sheet is None:
            continue                # ссылка на другую книгу или неизвестный лист
        if match.group("col1"):
            first_col = _column(match.group("col1"))
            last_col = _column(match.group("col2") or match.group("col1"))
            first_row = int(match.group("row1"))
            last_row = int(match.group("row2") or match.group("row1"))
        elif match.group("wholecol1"):
            first_col = _column(match.group("wholecol1"))
            last_col = _column(match.group("wholecol2"))
            first_row = last_row = None
        else:
            first_col = last_col = None
            first_row = int(match.group("wholerow1"))
            last_row = int(match.group("wholerow2"))
            rects.append(_Rect(sheet=sheet, first_col=None, last_col=None,
                               first_row=min(first_row, last_row),
                               last_row=max(first_row, last_row)))
            continue
        if first_col is None or last_col is None:
            continue                # колонки с таким именем в Excel нет
        if first_col > last_col:
            first_col, last_col = last_col, first_col
        if first_row is not None and first_row > last_row:
            first_row, last_row = last_row, first_row
        rects.append(_Rect(sheet=sheet, first_col=first_col, last_col=last_col,
                           first_row=first_row, last_row=last_row))
    return rects


def _formula_graph(wb) -> dict[tuple[str, int, int], list[_Rect]]:
    """Клетка с формулой -> на что она смотрит. Ключ -- (лист, колонка, строка)."""
    names = list(wb.sheetnames)
    graph: dict[tuple[str, int, int], list[_Rect]] = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                formula = _formula_text(cell)
                if formula is None:
                    continue
                graph[(ws.title, cell.column, cell.row)] = _references(
                    formula, ws.title, names)
    return graph


def _reached(graph: dict[tuple[str, int, int], list[_Rect]]) -> list[_Rect]:
    """Всё, до чего доходит обход от итогов книги вниз по ссылкам.

    Итог -- клетка с формулой, которую не читает ни одна другая формула. От
    каждого итога обход идёт по ссылкам и, попав в клетку с формулой, идёт
    дальше по её ссылкам: `D21` -> `G54` -> `G48:G50` -> `F60`. Без этого
    продолжения строки субподряда назывались бы недостижимыми, хотя их цена
    доходит до итога.
    """
    by_sheet: dict[str, list[tuple[str, int, int]]] = {}
    for sheet, col, row in graph:
        by_sheet.setdefault(sheet, []).append((sheet, col, row))

    def inside(rect: _Rect) -> list[tuple[str, int, int]]:
        return [coord for coord in by_sheet.get(rect.sheet, ())
                if rect.holds(*coord)]

    read_by_formula: set[tuple[str, int, int]] = set()
    for coord, rects in graph.items():
        for rect in rects:
            for other in inside(rect):
                if other != coord:      # ссылка на себя итогом быть не мешает
                    read_by_formula.add(other)

    reached: list[_Rect] = []
    walked: set[tuple[str, int, int]] = set()
    stack = [coord for coord in graph if coord not in read_by_formula]
    while stack:
        coord = stack.pop()
        # Единственная защита от круга в ссылках (A1 = A2, A2 = A1): клетка
        # разворачивается один раз. Второй такой проверки в обходе нет нарочно
        # -- запас, который нельзя сломать тестом, скрывает, чем держится
        # завершение обхода.
        if coord in walked:
            continue
        walked.add(coord)
        for rect in graph[coord]:
            reached.append(rect)
            stack.extend(inside(rect))
    return reached


def _wanted(addresses) -> list[tuple[str, int, int]]:
    """Адреса на проверку -> (адрес, колонка, строка), без повторов, в порядке ввода."""
    wanted: list[tuple[str, int, int]] = []
    seen: set[str] = set()
    for address in addresses:
        if not isinstance(address, str):
            raise WorkbookError(
                f"Адрес клетки должен быть строкой (например D8), а там "
                f"{address!r}.")
        parsed = _ADDRESS.fullmatch(address)
        column = _column(parsed.group(1)) if parsed else None
        if column is None:
            raise WorkbookError(
                f"«{address}» не адрес клетки книги (например D8, заглавными "
                "буквами).")
        if address in seen:
            continue
        seen.add(address)
        wanted.append((address, column, int(parsed.group(2))))
    return wanted


def _sheet_name(wb, sheet: str | None, book: str) -> str:
    names = wb.sheetnames
    shown = ", ".join(f"«{name}»" for name in names)
    if sheet is None:
        if len(names) == 1:
            return names[0]
        raise WorkbookError(
            f"В книге {book} несколько листов: {shown}. Укажите, по какому "
            "листу проверять клетки.")
    if sheet not in names:
        raise WorkbookError(
            f"В книге {book} нет листа «{sheet}». Есть: {shown}.")
    return sheet


def unreachable_cells(path, addresses, sheet: str | None = None) -> list[str]:
    """Те из `addresses`, которые ни одна формула книги не читает.

    Клетка попадает в ответ, если до неё не доходит обход от итогов книги:
    ни напрямую, ни через промежуточные клетки и диапазоны. Такую клетку
    человек заполнит, а книга её не сложит -- так из цены молча выпадает
    субподрядчик. Ответ -- в порядке `addresses`, по одному разу на адрес.

    Пустой ответ значит, что все названные клетки книга читает. Текстовые
    клетки (наименование позиции, единица измерения) в формулы не входят
    и честно называются недостижимыми: что с этим делать, решает человек.
    """
    path = Path(path)
    wanted = _wanted(addresses)
    wb = _open(path, values=False)
    try:
        name = _sheet_name(wb, sheet, path.name)
        reached = _reached(_formula_graph(wb))
    finally:
        wb.close()
    return [address for address, column, row in wanted
            if not any(rect.holds(name, column, row) for rect in reached)]
