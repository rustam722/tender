"""Предпросмотр одного документа: его содержание до сборки пакета.

Зачем. Человек правит поле и не видит, что стало в документе, пока не соберёт
пакет и не откроет файл в Word. Предпросмотр показывает содержание документа
по нынешним данным тендера -- и, главное, предупреждения: пустые обязательные
графы, расхождение цен, пустую декларацию. Беда видна ДО сборки, а не от
закупочной комиссии.

Чем это НЕ является. Это не вид документа в Word. Шрифтов, отступов, разбиения
на страницы и рамок таблиц здесь нет: рисовать второй Word в браузере --
работа на месяцы, и она разойдётся с настоящим документом в первый же день.
Отдаётся СОДЕРЖАНИЕ: абзацы и таблицы в том порядке, в котором документ их
печатает. Колонтитулы не отдаются вовсе -- в них стоят реквизиты бланка, а не
данные тендера.

Как устроено -- и почему только так. Документ гонится через ТУ ЖЕ отрисовку,
что и сборка: `template_path` выбирает шаблон, `build_context` строит контекст,
`render_docx` заполняет шаблон во временный файл, и уже из файла читаются
абзацы и таблицы. Своей подстановки «похожим способом» здесь нет и быть не
должно: предпросмотр, показывающий не то, что напечатается, хуже своего
отсутствия. По той же причине отказы и предупреждения берутся готовыми у
сборки (`Refusal`, `declaration_refusal`, `blank_fields_warning`), а не
пишутся здесь во второй раз.

Что остаётся после вызова: ничего. Временный файл живёт в своей папке в
`tempfile`, и папка удаляется в `finally` -- после успеха, после отказа и
после исключения в отрисовке. В папку вывода тендера предпросмотр не пишет
никогда: он вообще не знает её пути.

Книги Excel не отрисовываются. Содержание книги -- это сетка, и честно
показать её значит написать третью функцию. Для книги предпросмотр говорит, что
документ собирается в книге заказчика, и называет клетки, которые заполнятся, с
теми значениями, которые в них попадут; значения считает `planned_writes` --
тот самый код, который потом эти клетки и пишет.
"""
import shutil
import tempfile
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from app.render.context import ContextError, blank_fields_warning, build_context
from app.render.docx_render import TemplateError, render_docx
from app.render.layouts import LayoutError
from app.render.package import (
    Refusal, cellmap_path, declaration_refusal, template_path,
)
from app.render.xlsx_render import CellMapError, load_cellmap, planned_writes
from app.repo import Repo


WORD = "word"           # документ Word: показано его содержание
EXCEL = "excel"         # книга заказчика: показано, какие клетки заполнятся
REFUSAL = "refusal"     # документ по этим данным собрать нельзя; текст -- в refusal

_TEMP_PREFIX = "предпросмотр-"


@dataclass
class Preview:
    """Содержание документа для экрана.

    `kind` -- WORD, EXCEL или REFUSAL. `blocks` -- абзацы `{"kind": "p",
    "text": ...}` и таблицы `{"kind": "table", "rows": [[...], ...]}` в порядке
    документа; у отказа блоков нет. `warnings` -- те же предупреждения, что
    показала бы сборка. `refusal` -- причина по-русски или None.
    """
    kind: str
    blocks: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    refusal: str | None = None


def _paragraph(text: str) -> dict:
    return {"kind": "p", "text": text}


def _table(rows: list[list[str]]) -> dict:
    return {"kind": "table", "rows": rows}


def _read_blocks(path: Path) -> list[dict]:
    """Абзацы и таблицы готового файла -- в порядке документа.

    Порядок считается по телу документа, а не по `.paragraphs` и `.tables`
    подряд: те два списка идут мимо друг друга, и таблица в середине документа
    уехала бы в конец -- предпросмотр показывал бы не тот документ, который
    напечатается.

    Абзац без текста пропускается: пустые места между частями документа -- это
    оформление, которого здесь нет, а не содержание. Клетка отдаётся своим
    текстом; клетка, объединённая с соседней, повторится в каждой из них --
    так её видит python-docx, и выдумывать здесь нечего.
    """
    document = docx.Document(path)
    blocks: list[dict] = []
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = Paragraph(child, document).text
            if text.strip():
                blocks.append(_paragraph(text))
        elif child.tag == qn("w:tbl"):
            blocks.append(_table([[cell.text for cell in row.cells]
                                  for row in Table(child, document).rows]))
    return blocks


def _word_preview(template: Path, context: dict) -> tuple[list[dict], list[str]]:
    """Содержание документа Word и пути его пустых граф.

    Файл пишется во временную папку и читается из неё; папка удаляется всегда.
    `ignore_errors` -- потому что настоящая причина (отказ отрисовки) важнее
    неудачной уборки, а оставшийся в системной временной папке файл ничего не
    испортит: папку вывода тендера он не занимает.
    """
    folder = Path(tempfile.mkdtemp(prefix=_TEMP_PREFIX))
    try:
        blanks: list[str] = []
        out = folder / f"предпросмотр{template.suffix}"
        render_docx(template, context, out, blanks)
        return _read_blocks(out), blanks
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def _shown(value: object) -> str:
    """Значение клетки книги так, как оно в неё попадёт.

    Число печатается своими знаками (`format(..., "f")`, без показателя
    степени): разряды и запятую в книге ставит её собственное оформление, и
    показать здесь «100 000,10» значило бы обещать то, чего в клетке не будет.
    """
    if isinstance(value, Decimal):
        return format(value, "f")
    return "" if value is None else str(value)


def _workbook_preview(template: Path, spec, context: dict) -> list[dict]:
    """Клетки книги заказчика, которые заполнятся, и что в них попадёт."""
    cell_map = load_cellmap(cellmap_path(template, spec))
    writes = planned_writes(cell_map, context)
    return [
        _paragraph(
            f"Этот документ собирается в книге заказчика «{template.name}», "
            f"лист «{cell_map.sheet}». Содержание книги предпросмотр не "
            "показывает: заполняются только клетки, отведённые участнику, а "
            "остальное считают формулы самой книги. Ниже -- что именно в них "
            "попадёт."),
        _table([["Клетка", "Что в неё попадёт"]]
               + [[address, _shown(value)] for address, value in writes]),
    ]


def preview_document(tender_id: int, document: str, repo: Repo,
                     experience_ids: list[int] | tuple[int, ...] | None = None
                     ) -> Preview:
    """Содержание документа `document` тендера `tender_id` по нынешним данным.

    Ничего не пишет на диск (кроме временного файла, который сам же удаляет) и
    ничего не сохраняет в базе.

    Неизвестный тендер -- `LookupError`: это неверный запрос, а не отказ по
    данным. Всё, из-за чего документ не собрался бы (нет шаблона или карты,
    пустая декларация, пустая обязательная клетка книги, ошибка в разметке
    шаблона), -- `Preview` с `kind == REFUSAL` и причиной по-русски: отказ
    показывается как отказ, а не пустым экраном.

    `experience_ids` -- как в `build_context`: None берёт отбор, сохранённый
    при тендере.
    """
    tender = repo.get_tender(tender_id)
    profile = (repo.find_customer_profile(tender.customer_id)
               if tender.customer_id is not None else None)
    spec = profile.docs.get(document) if profile else None
    warnings: list[str] = []
    try:
        template = template_path(document, spec.template if spec else "", repo)
        context = build_context(tender_id, repo, document=document,
                                experience_ids=experience_ids)
        # Предупреждения контекста собираются до любого отказа: беда, о которой
        # уже можно сказать, говорится и тогда, когда документ собрать нельзя.
        warnings.extend(context["warnings"])
        if template.suffix.lower() == ".xlsx":
            return Preview(kind=EXCEL,
                           blocks=_workbook_preview(template, spec, context),
                           warnings=warnings)
        refusal = declaration_refusal(document, template, context)
        if refusal:
            raise Refusal(refusal)
        blocks, blanks = _word_preview(template, context)
    except (Refusal, ContextError, TemplateError, CellMapError, LayoutError) as exc:
        return Preview(kind=REFUSAL, warnings=warnings, refusal=str(exc))
    blank_note = blank_fields_warning(document, blanks)
    if blank_note:
        warnings.append(blank_note)
    return Preview(kind=WORD, blocks=blocks, warnings=warnings)
