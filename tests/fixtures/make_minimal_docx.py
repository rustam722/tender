"""Собирает tests/fixtures/minimal.docx -- минимальный шаблон для тестов рендера.

Запуск из корня проекта:  venv/Scripts/python tests/fixtures/make_minimal_docx.py

Шаблон содержит абзац со скаляром и таблицу справки. Колонки таблицы задаются
данными (раскладка заказчика), поэтому в шаблоне на строку приходится одна
клетка-прототип, а размножают её теги docxtpl. Тег удаляет целиком ту клетку
(строку), в которой стоит, поэтому открывающий и закрывающий теги занимают
отдельные клетки (строки), а прототип -- между ними:

    шапка:   [{%tc for h in assets.headers %}] [{{ h }}] [{%tc endfor %}]
    строка:  {%tr for row in assets.rows %}
    данные:  [{%tc for c in row %}] [{{ c }}] [{%tc endfor %}]
    строка:  {%tr endfor %}

В готовом файле остаются клетки по числу колонок раскладки и строки по числу
записей; служебные клетки и строки исчезают. Оформление (жирный шрифт, заливка
шапки) несёт клетка-прототип и переходит на все её копии.
Файл собран python-docx, руками не правится; изменили шаблон -- перезапустите
скрипт и закоммитьте оба файла.
"""
from pathlib import Path

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUT = Path(__file__).with_name("minimal.docx")

HEADER_TAGS = ("{%tc for h in assets.headers %}", "{{ h }}", "{%tc endfor %}")
ROW_OPEN = "{%tr for row in assets.rows %}"
DATA_TAGS = ("{%tc for c in row %}", "{{ c }}", "{%tc endfor %}")
ROW_CLOSE = "{%tr endfor %}"
HEADER_FILL = "D9D9D9"


def _shade(cell, fill: str) -> None:
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def _autowidth(table) -> None:
    """Ширина клетки -- «авто»: число колонок заранее неизвестно, Word разложит сам."""
    for tc in table._tbl.iter(qn("w:tc")):
        tcW = tc.get_or_add_tcPr().get_or_add_tcW()
        tcW.type, tcW.w = "auto", 0


def build() -> docx.document.Document:
    doc = docx.Document()
    doc.add_paragraph("{{ org.short_name }}")
    table = doc.add_table(rows=4, cols=3)
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, HEADER_TAGS):
        cell.text = text
    prototype = table.rows[0].cells[1]
    prototype.text = ""
    prototype.paragraphs[0].add_run(HEADER_TAGS[1]).bold = True
    _shade(prototype, HEADER_FILL)
    table.rows[1].cells[0].text = ROW_OPEN
    for cell, text in zip(table.rows[2].cells, DATA_TAGS):
        cell.text = text
    table.rows[3].cells[0].text = ROW_CLOSE
    _autowidth(table)
    return doc


if __name__ == "__main__":
    build().save(OUT)
    print(f"Записан {OUT}")
