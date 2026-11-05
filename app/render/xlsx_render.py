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
        "B31": "{org.head_position} {org.head_name}"
      },
      "rows": {                                    // повторяющийся блок строк
        "first_row": 9, "last_row": 19,            // сколько строк в форме
        "source": "items",                         // список из контекста
        "columns": {"C": "{row.price_value}",      // буква колонки -> выражение
                    "D": "{row.vat_value}"}
      }
    }

Выражение -- строка, в которой `{путь}` подставляется из контекста шаблонов
(задача 8). Путь -- ключи через точку: `org.head_name`, `tender.vat_percent`,
`conditions.Цена и условия:` (ключ может содержать пробелы и двоеточие, но не
точку). Внутри блока строк доступно ещё `row` -- текущий элемент списка.
Это весь язык; арифметики и функций в нём нет -- производное (НДС позиции)
считает контекст и отдаёт готовым полем.

  * Выражение из одного `{путь}` пишет значение как есть: Decimal остаётся
    числом, и ячейка получает число (SUM книги считает только числа).
  * Выражение с текстом вокруг (`Ставка НДС: {..}%`) пишет строку.

Чего здесь нет намеренно: клетки с названиями, номерами позиций и условиями
заказчика (колонки A, B, C формы Роснефти) в карту не входят -- это его текст.
Строки удалять и добавлять нельзя (форма это запрещает): позиций больше, чем
строк, -- отказ; меньше -- лишние строки остаются как в чистой форме.
"""
import json
import re
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


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
    rows: RowBlock | None = None


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


def _row_block(raw: object, cells: dict[str, str], name: str) -> RowBlock:
    if not isinstance(raw, dict):
        raise CellMapError(f"Карта ячеек {name}: rows должно быть объектом.")
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
    for address in cells:
        column = re.match(r"[A-Z]+", address).group()
        row = int(address[len(column):])
        if column in columns and first <= row <= last:
            raise CellMapError(
                f"Карта ячеек {name}: клетка {address} уже входит в блок "
                f"строк (колонка {column}, строки {first}-{last}).")
    return RowBlock(first_row=first, last_row=last, source=source,
                    columns=dict(columns))


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
    for address, expr in raw_cells.items():
        if not _ADDRESS.fullmatch(address):
            raise CellMapError(
                f"Карта ячеек {name}: «{address}» не адрес клетки "
                "(например D8, заглавными).")
        _check_expression(expr, f"клетки {address}", name)
    rows = (_row_block(data["rows"], raw_cells, name)
            if data.get("rows") is not None else None)
    return CellMap(sheet=sheet, cells=dict(raw_cells), rows=rows)


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
                f"(не найдено «{part}»). Проверьте карту ячеек и данные тендера.")
    return node


def _text(value: object) -> str:
    return "" if value is None else str(value)


def _evaluate(expr: str, scope: Mapping, address: str):
    where = f"клетки {address}"
    lone = _PLACEHOLDER.fullmatch(expr)
    if lone:
        value = _lookup(scope, lone.group(1), where)
        return None if value == "" else value
    text = _PLACEHOLDER.sub(
        lambda m: _text(_lookup(scope, m.group(1), where)), expr)
    return text.strip()


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
            writes.append((address, _evaluate(expr, context, address)))
        block = cell_map.rows
        if block is not None:
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
            f"«{block.source}» в данных не список позиций, а {type(rows).__name__}.")
    capacity = block.last_row - block.first_row + 1
    if len(rows) > capacity:
        raise CellMapError(
            f"Позиций ({len(rows)}) больше, чем строк в форме заказчика "
            f"({capacity}, строки {block.first_row}-{block.last_row}). "
            "Строки в форме удалять и добавлять нельзя: сократите список "
            "позиций или заполните форму вручную.")
    writes = []
    for offset, item in enumerate(rows):
        row = block.first_row + offset
        scope = {**context, "row": item}
        for letter, expr in block.columns.items():
            address = f"{letter}{row}"
            writes.append((address, _evaluate(expr, scope, address)))
    return writes
