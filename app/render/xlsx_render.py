"""Заполнение формы заказчика в его собственной книге Excel.

Крупные заказчики предписывают точную таблицу, отмечают клетки, которые вправе
заполнить участник, и отклоняют заявки, где форма изменена. Поэтому принцип
один: пишутся ТОЛЬКО клетки, отведённые участнику, а всё производное считают
формулы самой книги. Записать вычисленный итог поверх формулы -- сломать
документ, который проверяющий откроет и пересчитает; такая запись отвергается
отказом, а не молча.

Карта ячеек -- небольшой JSON-файл в папке cellmaps/, его правит не программист
(так же устроены раскладки таблиц, layouts/):

    {
      "sheet": "Лист1",
      "cells": {                                   // одиночные клетки
        "D8":  "Ставка НДС: {tender.vat_percent}%",
        "E23": "{conditions.Цена и условия:}",
        "B31": "{org.head_position} {org.head_name}",
        "D88": {"value": "{tender.rounding}", "optional": true}  // пример: может быть пусто
      },
      "rows": [                                    // блоки повторяющихся строк
        {"first_row": 9, "last_row": 19,           // сколько строк в форме
         "source": "items",                        // список из контекста
         "columns": {"C": "{row.price_value}",     // буква колонки -> выражение
                     "D": "{row.vat_value}"}}
      ],
      "rates": {"insurance": "C17", "injury": "C18",   // где книга держит СВОИ
                "overhead": "C24", "profit": "C25"}    // ставки калькуляции
    }

Блоков может быть несколько (книга с подтаблицами: материалы, трудозатраты,
субподряд -- у каждого свои строки, список и колонки); блоки не пересекаются.

Выражение -- строка, в которой `{путь}` подставляется из контекста шаблонов
(задача 8). Путь -- ключи через точку: `org.head_name`, `tender.vat_percent`,
`conditions.Цена и условия:` (ключ может содержать пробелы и двоеточие, но не
точку). Внутри блока строк доступно ещё `row` -- текущий элемент списка.
Это весь язык; арифметики и функций в нём нет -- производное (НДС позиции)
считает контекст и отдаёт готовым полем.

  * Выражение из одного `{путь}` пишет значение как есть: Decimal остаётся
    числом, и ячейка получает число (SUM книги считает только числа).
  * Выражение с текстом вокруг (`Ставка НДС: {..}%`) пишет строку.
  * Одиночная клетка (`cells`) по умолчанию обязательна: если значение пусто
    (нет ответа по условию, одни пробелы) или пуста любая часть текстового
    выражения (должность без фамилии), -- отказ, а не графа, которая выглядит
    пустой и из-за которой заявку отклонят молча. Строки значений обрезаются
    по краям. В строках блоков пустое значение допустимо (примечание).
  * Поле ввода, которое законно может остаться пустым (например разрядность
    округления в книге КП), помечается в самой карте: вместо строки объект
    `{"value": "{путь}", "optional": true}`. Пустое значение такой клетки не
    пишется, клетка остаётся как в форме заказчика. Отсутствующий путь и здесь
    отказ: «может быть пустым» не значит «может отсутствовать».

Раздел `rates` -- необязательный и, в отличие от всего остального, читается, а
не пишется: книга заказчика хранит ставки калькуляции в своих клетках и считает
по ним, что бы ни стояло в настройках программы. Имена ставок -- только
`insurance`, `injury`, `overhead`, `profit` (RATE_FIELDS: ключ -> поле
OrgSettings); значение -- адрес клетки. Адреса свои у каждого заказчика, поэтому
их место здесь, рядом со всеми прочими адресами. `read_rates` достаёт эти ставки
из книги; по ним считает `app.calc.workbook_total`, а расхождение с
настройками называет предупреждение сборки (app.render.context). Раздела нет --
ставки берутся из настроек, как раньше. Клетка, из которой ставку не прочитать,
-- отказ, и этот отказ доходит до человека: документ уходит в
`PackageResult.skipped` с причиной, а не собирается без проверок.

Чего здесь нет намеренно: клетки с названиями, номерами позиций и условиями
заказчика (колонки A, B, C формы Роснефти) в карту не входят -- это его текст.
Строки удалять и добавлять нельзя (форма это запрещает): позиций больше, чем
строк в блоке, -- отказ; меньше -- лишние строки остаются как в чистой форме.
"""
import json
import re
from collections.abc import Mapping
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


# Ставки калькуляции, которые книга заказчика хранит у себя: имя в карте ячеек
# -> поле OrgSettings, с которым ставка сравнивается. Это весь словарь раздела
# `rates`: имени вне этого списка в книге соответствовать нечему, поэтому
# карта с ним отвергается, а не читается наполовину.
RATE_FIELDS = {"insurance": "insurance_rate", "injury": "injury_rate",
               "overhead": "overhead_rate", "profit": "profit_rate"}


class CellMapError(Exception):
    """Ошибка карты ячеек или заполнения книги; текст на русском, для человека."""


@dataclass
class RowBlock:
    first_row: int
    source: str                       # ключ списка в контексте: 'items'
    columns: dict[str, str]           # буква колонки -> выражение
    last_row: int                     # последняя строка формы: ёмкость блока


@dataclass
class CellMap:
    sheet: str
    cells: dict[str, str]             # адрес клетки -> выражение
    rows: list[RowBlock] = field(default_factory=list)
    optional: set[str] = field(default_factory=set)   # клетки, где пусто -- законно
    rates: dict[str, str] = field(default_factory=dict)   # ставка -> адрес, ЧИТАЕТСЯ


_ADDRESS = re.compile(r"[A-Z]{1,3}[1-9][0-9]*")
_COLUMN = re.compile(r"[A-Z]{1,3}")
_PLACEHOLDER = re.compile(r"\{([^{}]*)\}")


# ---- файл карты ----------------------------------------------------------------

def _check_expression(expr: object, label: str, name: str) -> None:
    if not isinstance(expr, str):
        raise CellMapError(
            f"Карта ячеек {name}: выражение для {label} должно быть строкой, "
            f"а там {expr!r}.")
    for path in _PLACEHOLDER.findall(expr):
        if not path.strip():
            raise CellMapError(
                f"Карта ячеек {name}: в выражении для {label} пустые "
                "фигурные скобки {}.")
    rest = _PLACEHOLDER.sub("", expr)
    if "{" in rest or "}" in rest:
        raise CellMapError(
            f"Карта ячеек {name}: в выражении для {label} не закрыта или "
            f"лишняя фигурная скобка: {expr}")


def _row_block(raw: object, name: str) -> RowBlock:
    if not isinstance(raw, dict):
        raise CellMapError(
            f"Карта ячеек {name}: каждый блок в rows должен быть объектом, "
            f"а там {raw!r}.")
    for key in ("first_row", "last_row", "source", "columns"):
        if key not in raw:
            raise CellMapError(
                f"Карта ячеек {name}: в rows не найдено «{key}».")
    first, last = raw["first_row"], raw["last_row"]
    for key, value in (("first_row", first), ("last_row", last)):
        if isinstance(value, bool) or not isinstance(value, int) or value < 1:
            raise CellMapError(
                f"Карта ячеек {name}: {key} должен быть номером строки "
                f"(целое от 1), а там {value!r}.")
    if last < first:
        raise CellMapError(
            f"Карта ячеек {name}: last_row ({last}) меньше first_row ({first}).")
    source = raw["source"]
    if not isinstance(source, str) or not source.strip():
        raise CellMapError(
            f"Карта ячеек {name}: source в rows -- имя списка из контекста.")
    columns = raw["columns"]
    if not isinstance(columns, dict) or not columns:
        raise CellMapError(
            f"Карта ячеек {name}: в rows.columns нужна хотя бы одна колонка.")
    for letter, expr in columns.items():
        if not _COLUMN.fullmatch(letter):
            raise CellMapError(
                f"Карта ячеек {name}: колонка «{letter}» не буква "
                "колонки Excel (A, B, ... заглавными).")
        _check_expression(expr, f"колонки «{letter}»", name)
    return RowBlock(first_row=first, last_row=last, source=source,
                    columns=dict(columns))


def _check_overlaps(cells: dict[str, str], rows: list[RowBlock],
                    name: str) -> None:
    """Одна клетка -- одна запись: ни клетка карты, ни два блока не делят адрес."""
    owner: dict[str, str] = {}
    for address in cells:
        owner[address] = f"клеткой {address} из cells"
    for number, block in enumerate(rows, start=1):
        for letter in block.columns:
            for row in range(block.first_row, block.last_row + 1):
                address = f"{letter}{row}"
                if address in owner:
                    raise CellMapError(
                        f"Карта ячеек {name}: клетка {address} из блока строк "
                        f"№{number} уже занята {owner[address]}.")
                owner[address] = f"блоком строк №{number}"


def load_cellmap(path: Path) -> CellMap:
    path = Path(path)
    name = path.name
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        # имя файла берётся из данных, которые правит пользователь (профиль
        # заказчика), а сообщение ОС по-английски в русский текст не годится
        raise CellMapError(
            f"Карта ячеек {name} не найдена в папке {path.parent.name}. "
            "Проверьте имя файла в профиле заказчика.") from None
    except (OSError, ValueError) as exc:
        raise CellMapError(
            f"Не удалось прочитать карту ячеек {name}: {exc}") from exc
    if not isinstance(data, dict):
        raise CellMapError(f"Карта ячеек {name}: ожидается объект JSON.")
    for key in ("sheet", "cells"):
        if key not in data:
            raise CellMapError(f"Карта ячеек {name}: не найдено «{key}».")
    sheet = data["sheet"]
    if not isinstance(sheet, str) or not sheet.strip():
        raise CellMapError(f"Карта ячеек {name}: sheet -- имя листа книги.")
    raw_cells = data["cells"]
    if not isinstance(raw_cells, dict):
        raise CellMapError(
            f"Карта ячеек {name}: cells должно быть объектом "
            "«адрес клетки -> выражение».")
    cells: dict[str, str] = {}
    optional: set[str] = set()
    for address, raw in raw_cells.items():
        if not _ADDRESS.fullmatch(address):
            raise CellMapError(
                f"Карта ячеек {name}: «{address}» не адрес клетки "
                "(например D8, заглавными).")
        expr = raw
        if isinstance(raw, dict):
            extra = set(raw) - {"value", "optional"}
            if extra:
                raise CellMapError(
                    f"Карта ячеек {name}: у клетки {address} неизвестный "
                    f"ключ «{sorted(extra)[0]}»; допустимы value и optional.")
            if "value" not in raw:
                raise CellMapError(
                    f"Карта ячеек {name}: у клетки {address} не найдено "
                    "«value» (выражение).")
            expr = raw["value"]
            flag = raw.get("optional", False)
            if not isinstance(flag, bool):
                raise CellMapError(
                    f"Карта ячеек {name}: optional у клетки {address} должен "
                    f"быть true или false, а там {flag!r}.")
            if flag:
                optional.add(address)
        _check_expression(expr, f"клетки {address}", name)
        cells[address] = expr
    raw_rows = data.get("rows")
    if raw_rows is None:
        raw_rows = []
    if not isinstance(raw_rows, list):
        raise CellMapError(
            f"Карта ячеек {name}: rows должно быть списком блоков строк "
            "(даже если блок один).")
    rows = [_row_block(raw, name) for raw in raw_rows]
    _check_overlaps(cells, rows, name)
    return CellMap(sheet=sheet, cells=cells, rows=rows, optional=optional,
                   rates=_rates(data.get("rates"), name))


def cell_paths(cell_map: CellMap) -> list[str]:
    """Пути контекста, которые карта берёт из данных, в порядке появления.

    Нужен тем, кто должен знать, ЧТО карта напишет в книгу, книги не открывая:
    экран показывает условия сделки, которых требует форма заказчика, а сборка
    называет недостающие ответы все разом, а не по одному за попытку.

    Пути, начинающиеся с `row.`, в список не входят: они относятся к элементу
    списка внутри блока строк, а не к самому контексту.
    """
    paths: list[str] = []
    expressions = [*cell_map.cells.values()]
    for block in cell_map.rows:
        expressions.extend(block.columns.values())
    for expr in expressions:
        for raw in _PLACEHOLDER.findall(expr):
            path = raw.strip()
            if path == "row" or path.startswith("row."):
                continue
            if path not in paths:
                paths.append(path)
    return paths


def _rates(raw: object, name: str) -> dict[str, str]:
    """Раздел rates: имя ставки -> адрес клетки. Нет раздела -- пустой словарь."""
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise CellMapError(
            f"Карта ячеек {name}: rates должно быть объектом «ставка -> адрес "
            "клетки».")
    known = ", ".join(sorted(RATE_FIELDS))
    rates: dict[str, str] = {}
    for key, address in raw.items():
        if key not in RATE_FIELDS:
            raise CellMapError(
                f"Карта ячеек {name}: в rates неизвестная ставка «{key}»; "
                f"допустимы {known}.")
        if not isinstance(address, str) or not _ADDRESS.fullmatch(address):
            raise CellMapError(
                f"Карта ячеек {name}: у ставки «{key}» в rates должен быть "
                f"адрес клетки (например C17, заглавными), а там {address!r}.")
        rates[key] = address
    return rates


# Ставка -- доля единицы, не проценты. Формулировка дословно та же, что у
# настроек калькуляции (_FRACTIONS в app/web/schema.py): одна и та же фраза
# в двух местах, чтобы человек узнавал правило, а не разбирал его заново.
_RATE_RULE = "ставка задаётся долей единицы, от 0 до 1: например, 0,05 для 5 %"

# Что лежит в клетке, когда там не число: сказать это человеку по-русски, иначе
# он пойдёт искать пустую клетку, а там дата. Падеж винительный: вид клетки
# подставляется в «содержит не ставку, а ...», и «а дата» читается как обрывок.
_CELL_KINDS = ((bool, "логическое значение (ДА/НЕТ)"), (datetime, "дату"),
               (date, "дату"), (time, "время"), (timedelta, "длительность"))
_UNKNOWN_KIND = "значение, которое числом не является"

# Длиннее этого число в сообщение не пишется развёрнуто: 1E+300 развернулся бы
# в триста цифр, и причина утонула бы в них.
_PLAIN_LIMIT = 30


def _plain(value: Decimal) -> str:
    """Число как оно есть, для сообщения: 30 -> «30», 0.3 -> «0,3»."""
    if not value.is_finite():
        return str(value)
    plain = format(value.normalize(), "f").replace(".", ",")
    return plain if len(plain) <= _PLAIN_LIMIT else str(value.normalize())


def _rate_of(value: object, key: str, address: str, book: str) -> Decimal:
    """Значение клетки -> ставка-доля. Всё, что не доля от 0 до 1, -- отказ.

    Нулевая доля, молча взятая вместо непрочитанной ставки, дала бы
    правдоподобную цену без взносов; ставка 30 вместо 0,3 -- цену в сто раз
    больше. И то и другое человек увидит только от проверяющего, поэтому здесь
    отказ, а не догадка."""
    where = f"В книге {book} клетка {address} (ставка «{key}»)"
    shown = value                      # для сообщения: как в клетке, до правок
    if isinstance(value, str):
        text = value.strip().replace(",", ".")
        if text.startswith("="):
            raise CellMapError(
                f"{where} занята формулой: ставку оттуда не прочитать. "
                "Проверьте адрес в карте ячеек.")
        shown = value.strip()
        value = text
    if value is None or value == "":
        raise CellMapError(
            f"{where} пуста: ставку оттуда не прочитать. Проверьте адрес "
            "в карте ячеек или заполните клетку в шаблоне книги.")
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal, str)):
        kind = next((name for cls, name in _CELL_KINDS if isinstance(value, cls)),
                    _UNKNOWN_KIND)
        raise CellMapError(
            f"{where} содержит не ставку, а {kind}. Проверьте адрес в карте "
            "ячеек или клетку в шаблоне книги.")
    try:
        rate = Decimal(str(value))
    except InvalidOperation:
        raise CellMapError(
            f"{where} содержит не число, а «{shown}». Проверьте адрес в карте "
            "ячеек или клетку в шаблоне книги.") from None
    # is_finite первым: сравнение Decimal('NaN') с числом само поднимает
    # InvalidOperation, то есть ошибку не по-русски и не для человека
    if not rate.is_finite() or not Decimal(0) <= rate <= Decimal(1):
        raise CellMapError(
            f"{where} содержит {_plain(rate)}, а {_RATE_RULE}. Исправьте клетку "
            "в шаблоне книги или адрес в карте ячеек.")
    return rate


def read_rates(template: Path, cell_map: CellMap) -> dict[str, Decimal]:
    """Ставки калькуляции из самой книги заказчика: имя ставки -> доля (0.3).

    Книга считает по своим ставкам, что бы ни стояло в настройках программы,
    поэтому читаются они из книги, а не из настроек. Нет раздела rates --
    пустой словарь, книга даже не открывается, и ставки берутся из настроек,
    как раньше.

    Клетка, из которой ставку не прочитать (формула, текст, пусто, дата, доля
    вне 0..1), -- отказ `CellMapError`. Этот отказ обязан дойти до человека:
    ставки не читает больше никто, `fill_workbook` в эти клетки не смотрит, и
    проглоченное здесь исключение выключило бы всю проверку расхождений разом
    (см. `app.render.context._workbook_sums_warnings`).
    """
    if not cell_map.rates:
        return {}
    template = Path(template)
    wb = _open(template)
    try:
        if cell_map.sheet not in wb.sheetnames:
            names = ", ".join(f"«{n}»" for n in wb.sheetnames)
            raise CellMapError(
                f"В книге {template.name} нет листа «{cell_map.sheet}». "
                f"Есть: {names}.")
        ws = wb[cell_map.sheet]
        return {key: _rate_of(ws[address].value, key, address, template.name)
                for key, address in cell_map.rates.items()}
    finally:
        wb.close()


# ---- вычисление выражений --------------------------------------------------------

def _lookup(scope: Mapping, path: str, where: str):
    """Значение по пути из ключей через точку; нет -- CellMapError."""
    node = scope
    for part in path.strip().split("."):
        if isinstance(node, Mapping) and part in node:
            node = node[part]
        else:
            raise CellMapError(
                f"Для {where} в данных нет значения «{path.strip()}» "
                f"(не найдено «{part}»). Задайте его в данных тендера или "
                "поправьте карту ячеек.")
    return node


def _text(value: object) -> str:
    return "" if value is None else str(value)


def _blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _evaluate(expr: str, scope: Mapping, address: str, required: bool = False):
    """required -- обязательная одиночная клетка: пустое значение (в том числе
    из одних пробелов) или пустая часть текста -- отказ, а не пустая графа.
    Не required -- пустое значение даёт None (или пустой текст)."""
    where = f"клетки {address}"
    lone = _PLACEHOLDER.fullmatch(expr)
    if lone:
        path = lone.group(1).strip()
        value = _lookup(scope, path, where)
        if _blank(value):
            if required:
                raise CellMapError(
                    f"Для {where} значение «{path}» пусто. Графа не должна "
                    "остаться незаполненной: задайте значение в данных тендера.")
            return None
        return value.strip() if isinstance(value, str) else value

    def part(match):
        path = match.group(1).strip()
        value = _lookup(scope, path, where)
        if required and _blank(value):
            raise CellMapError(
                f"Для {where} значение «{path}» пусто, а оно входит в текст "
                f"«{expr}». Графа не должна остаться неполной: задайте "
                "значение в данных тендера.")
        return _text(value)

    text = _PLACEHOLDER.sub(part, expr).strip()
    if required and not text:
        raise CellMapError(f"Для {where} из «{expr}» получился пустой текст.")
    return text


# ---- запись ------------------------------------------------------------------------

def _check_writable(ws, address: str, protected: bool) -> None:
    cell = ws[address]
    if isinstance(cell, MergedCell):
        raise CellMapError(
            f"Клетка {address} входит в объединённый диапазон, но не является "
            "его первой клеткой: писать в неё нельзя. Проверьте карту ячеек.")
    if cell.data_type == "f":
        raise CellMapError(
            f"Клетка {address} содержит формулу книги заказчика, записывать в "
            "неё нельзя: итоги книга считает сама. Проверьте карту ячеек.")
    if protected and cell.protection.locked:
        raise CellMapError(
            f"Клетка {address} заблокирована защитой листа: заказчик не отвёл "
            "её участнику. Проверьте карту ячеек.")


def _open(template: Path):
    try:
        return load_workbook(template, data_only=False)
    except FileNotFoundError:
        raise CellMapError(
            f"Книга заказчика {template.name} не найдена в папке "
            f"{template.parent.name}.") from None
    except Exception as exc:        # BadZipFile, KeyError, InvalidFileException...
        raise CellMapError(
            f"Не удалось открыть книгу заказчика {template.name}: файл повреждён "
            "или не является книгой Excel (.xlsx).") from exc


def fill_workbook(template: Path, cell_map: CellMap, context: dict,
                  out: Path) -> Path:
    """Заполняет книгу заказчика по карте и сохраняет её в out.

    Читает книгу с формулами (data_only=False), пишет только клетки карты,
    оформление, формулы и защиту листа остаются как были. Исходный файл
    не изменяется. Сначала проверяются и вычисляются все записи и лишь потом
    пишутся, поэтому отказ не оставляет недописанного файла.
    """
    template, out = Path(template), Path(out)
    if template.resolve() == out.resolve():
        raise CellMapError(
            f"Результат нельзя сохранять поверх исходной книги заказчика "
            f"{template.name}: выберите другое имя файла.")
    wb = _open(template)
    try:
        if cell_map.sheet not in wb.sheetnames:
            names = ", ".join(f"«{n}»" for n in wb.sheetnames)
            raise CellMapError(
                f"В книге {template.name} нет листа «{cell_map.sheet}». "
                f"Есть: {names}.")
        ws = wb[cell_map.sheet]
        protected = bool(ws.protection.sheet)

        writes = []                                   # (адрес, значение)
        for address, expr in cell_map.cells.items():
            required = address not in cell_map.optional
            value = _evaluate(expr, context, address, required=required)
            if not required and _blank(value):
                continue              # законно пусто: клетка как в форме заказчика
            writes.append((address, value))
        for block in cell_map.rows:
            writes.extend(_row_writes(block, context))
        for address, _ in writes:
            _check_writable(ws, address, protected)

        for address, value in writes:
            cell = ws[address]
            cell.value = value
            if isinstance(value, str) and cell.data_type == "f":
                cell.data_type = "s"                  # «=1+1» -- текст, не формула
        # openpyxl не хранит результаты формул (старый кэш -- «0» в E9 -- не
        # переживает сохранение), и Excel должен пересчитать итоги при открытии.
        # openpyxl и так ставит этот флаг по умолчанию; явно -- чтобы корректность
        # файла не зависела от умолчания (тест читает флаг из готового файла)
        wb.calculation.fullCalcOnLoad = True
        try:
            wb.save(out)
        except OSError as exc:
            raise CellMapError(
                f"Не удалось сохранить файл {out.name}: нет такой папки или "
                "файл открыт в другой программе.") from exc
    finally:
        wb.close()
    return out


def _row_writes(block: RowBlock, context: dict) -> list[tuple[str, object]]:
    rows = _lookup(context, block.source,
                   f"блока строк (с {block.first_row}-й строки)")
    if isinstance(rows, (str, Mapping)) or not hasattr(rows, "__len__"):
        raise CellMapError(
            f"«{block.source}» в данных не список, а {type(rows).__name__}.")
    capacity = block.last_row - block.first_row + 1
    if len(rows) > capacity:
        raise CellMapError(
            f"Позиций в «{block.source}» ({len(rows)}) больше, чем строк в форме "
            f"заказчика ({capacity}, строки {block.first_row}-{block.last_row}). "
            "Строки в форме удалять и добавлять нельзя: сократите список "
            "или заполните форму вручную.")
    writes = []
    for offset, item in enumerate(rows):
        row = block.first_row + offset
        scope = {**context, "row": item}
        for letter, expr in block.columns.items():
            address = f"{letter}{row}"
            writes.append((address, _evaluate(expr, scope, address)))
    return writes
