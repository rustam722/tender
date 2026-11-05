"""Раскладки колонок табличных справок.

Справочник хранит факты один раз, а каждый заказчик требует свою таблицу.
Раскладка -- небольшой JSON-файл в папке layouts/, его правит не программист:

    {
      "source": "assets",              // employees | assets | experience
      "numbered": true,                // первая колонка -- порядковый номер
      "columns": [
        {"header": "№ п/п"},                                  // колонка номера
        {"header": "Наименование", "parts": ["name", "maker_model"]},
        {"header": "Право собственности",
         "parts": ["ownership_basis", "ownership_doc"], "join": ": "},
        {"header": "Примечание", "parts": []}                 // пустая клетка
      ]
    }

Клетка собирается из значений полей `parts`, соединённых через `join`
(по умолчанию пробел). Пустые части отбрасываются до соединения, `None`
даёт пустую строку. Никакого форматирования: числа и даты выводятся как
есть, фразы из частей не строятся.
"""
import json
from collections.abc import Sequence
from dataclasses import dataclass, field, fields
from pathlib import Path

from app.models import Asset, Employee, ExperienceRecord

SOURCES = {
    "employees": Employee,
    "assets": Asset,
    "experience": ExperienceRecord,
}


class LayoutError(Exception):
    """Ошибка в файле раскладки; текст на русском, для человека."""


@dataclass
class Column:
    header: str
    parts: list[str] = field(default_factory=list)
    join: str = " "


@dataclass
class TableLayout:
    source: str                       # 'employees' | 'assets' | 'experience'
    columns: list[Column]
    numbered: bool                    # columns[0] -- колонка порядкового номера


def _validate(layout: TableLayout) -> None:
    model = SOURCES.get(layout.source)
    if model is None:
        raise LayoutError(
            f"Неизвестный источник данных «{layout.source}». "
            f"Допустимо: {', '.join(SOURCES)}.")
    known = {f.name for f in fields(model)}
    for index, column in enumerate(layout.columns):
        if layout.numbered and index == 0:
            if column.parts:
                raise LayoutError(
                    f"Колонка «{column.header}» -- колонка номера, "
                    f"в ней не должно быть полей.")
            continue
        for name in column.parts:
            if name not in known:
                raise LayoutError(
                    f"В колонке «{column.header}» указано неизвестное поле "
                    f"«{name}». Доступные поля источника "
                    f"«{layout.source}»: {', '.join(sorted(known))}.")
    if layout.numbered and not layout.columns:
        raise LayoutError(
            "Раскладка с нумерацией должна начинаться с колонки номера.")


def load_layout(path: Path) -> TableLayout:
    path = Path(path)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        # имя файла берётся из данных, которые правит пользователь (профиль
        # заказчика), а сообщение ОС по-английски в русский текст не годится
        raise LayoutError(
            f"Файл раскладки {path.name} не найден в папке {path.parent.name}. "
            "Проверьте имя файла в профиле заказчика.") from None
    except (OSError, ValueError) as exc:
        raise LayoutError(
            f"Не удалось прочитать раскладку {path.name}: {exc}") from exc
    try:
        columns = []
        for raw in data["columns"]:
            if "header" not in raw:
                raise LayoutError(
                    f"В раскладке {path.name} у колонки нет заголовка "
                    f"(header): {raw}")
            columns.append(Column(
                header=raw["header"],
                parts=list(raw.get("parts", [])),
                join=raw.get("join", " ")))
        layout = TableLayout(
            source=data["source"],
            columns=columns,
            numbered=bool(data.get("numbered", False)))
    except (KeyError, TypeError) as exc:
        raise LayoutError(
            f"Раскладка {path.name} составлена неверно: "
            f"не найдено {exc}.") from exc
    _validate(layout)
    return layout


def _text(value: object) -> str:
    return "" if value is None else str(value)


def build_rows(layout: TableLayout,
               records: Sequence[object]) -> list[list[str]]:
    """Строки данных без шапки; шапка -- layout.columns."""
    _validate(layout)
    data_columns = layout.columns[1:] if layout.numbered else layout.columns
    rows = []
    for number, record in enumerate(records, start=1):
        row = [str(number)] if layout.numbered else []
        for column in data_columns:
            values = [_text(getattr(record, name)) for name in column.parts]
            row.append(column.join.join(v for v in values if v))
        rows.append(row)
    return rows
