"""Собирает tests/fixtures/minimal.docx -- минимальный шаблон для тестов рендера.

Запуск из корня проекта:  venv/Scripts/python tests/fixtures/make_minimal_docx.py

Шаблон содержит абзац со скаляром и таблицу с циклом по строкам справки.
docxtpl требует, чтобы теги `{%tr ... %}` стояли в отдельных строках таблицы:
строки с тегами при рендере удаляются, остаётся одна строка на элемент цикла.
Файл собран python-docx, руками не правится; изменили шаблон -- перезапустите
скрипт и закоммитьте оба файла.
"""
from pathlib import Path

import docx

OUT = Path(__file__).with_name("minimal.docx")

HEADERS = ("№ п/п", "Наименование", "Примечание")
LOOP_OPEN = "{%tr for r in assets.rows %}"
LOOP_CLOSE = "{%tr endfor %}"
ROW_CELLS = ("{{ r[0] }}", "{{ r[1] }}", "{{ r[2] }}")


def build() -> docx.document.Document:
    doc = docx.Document()
    doc.add_paragraph("{{ org.short_name }}")
    table = doc.add_table(rows=4, cols=len(HEADERS))
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, HEADERS):
        cell.text = text
    table.rows[1].cells[0].text = LOOP_OPEN
    for cell, text in zip(table.rows[2].cells, ROW_CELLS):
        cell.text = text
    table.rows[3].cells[0].text = LOOP_CLOSE
    return doc


if __name__ == "__main__":
    build().save(OUT)
    print(f"Записан {OUT}")
