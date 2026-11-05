"""Рендер Word-шаблонов.

Каждая проверка текста читает его из готового файла через python-docx,
а не из словаря на входе: тест, сверяющий вход с самим собой, ничего не доказывает.
"""
import zipfile
from pathlib import Path

import docx
import pytest
from lxml import etree

from app.render.context import money
from app.render.docx_render import TemplateError, render_docx

FIXTURE = Path(__file__).parent / "fixtures" / "minimal.docx"
NBSP = "\u00a0"


def _context(short_name="ООО «Техносфера»", rows=()):
    return {
        "org": {"short_name": short_name},
        "assets": {"headers": ["№ п/п", "Наименование", "Примечание"],
                   "rows": [list(r) for r in rows]},
    }


def _table(path):
    """Таблица готового файла как список строк из текстов клеток."""
    table = docx.Document(path).tables[0]
    return [[cell.text for cell in row.cells] for row in table.rows]


def _all_text(path):
    d = docx.Document(path)
    parts = [p.text for p in d.paragraphs]
    parts += [c.text for row in d.tables[0].rows for c in row.cells]
    return "\n".join(parts)


def _template_with(tmp_path, paragraph):
    path = tmp_path / "t.docx"
    d = docx.Document()
    d.add_paragraph(paragraph)
    d.save(path)
    return path


def test_substitutes_scalar(tmp_path):
    out = render_docx(FIXTURE, _context("ООО «Техносфера»"), tmp_path / "o.docx")
    assert out == tmp_path / "o.docx"
    assert docx.Document(out).paragraphs[0].text == "ООО «Техносфера»"


def test_expands_table_rows(tmp_path):
    rows = [(str(i), f"Позиция {i}", f"прим. {i}") for i in range(1, 24)]
    out = render_docx(FIXTURE, _context(rows=rows), tmp_path / "o.docx")
    table = _table(out)
    assert len(table) == 1 + 23                       # шапка + 23 позиции
    assert table[0] == ["№ п/п", "Наименование", "Примечание"]
    assert table[1] == ["1", "Позиция 1", "прим. 1"]
    assert table[23] == ["23", "Позиция 23", "прим. 23"]
    assert "{%" not in _all_text(out)                 # служебные строки не просочились


def test_empty_table_keeps_header_and_no_data_rows(tmp_path):
    out = render_docx(FIXTURE, _context(rows=[]), tmp_path / "o.docx")
    assert _table(out) == [["№ п/п", "Наименование", "Примечание"]]
    assert "{%" not in _all_text(out)


def test_nested_quotes_survive(tmp_path):             # Review Focus 2
    name = 'АО «НК «Роснефть»'
    out = render_docx(FIXTURE, _context(name), tmp_path / "o.docx")
    assert 'АО «НК «Роснефть»' in docx.Document(out).paragraphs[0].text


def test_xml_special_chars_are_escaped(tmp_path):
    value = 'Иванов & Партнёры <ООО> "кавычки" \'апострофы\''
    out = render_docx(FIXTURE, _context(value, rows=[("1", value, "")]),
                      tmp_path / "o.docx")
    d = docx.Document(out)                           # разбор без ошибки
    assert d.paragraphs[0].text == value
    assert _table(out)[1][1] == value


def test_braces_in_value_are_not_treated_as_syntax(tmp_path):
    for trap in ("{{ ловушка }}", "{% for x in y %}", "{#комментарий#}"):
        out = render_docx(FIXTURE, _context(trap, rows=[("1", trap, "")]),
                          tmp_path / "o.docx")
        assert docx.Document(out).paragraphs[0].text == trap
        assert _table(out)[1][1] == trap


def test_non_breaking_space_survives(tmp_path):
    amount = money(5985000)                           # «5 985 000,00» с U+00A0
    assert amount == f"5{NBSP}985{NBSP}000,00"
    out = render_docx(FIXTURE, _context(amount, rows=[("1", amount, "")]),
                      tmp_path / "o.docx")
    for text in (docx.Document(out).paragraphs[0].text, _table(out)[1][1]):
        assert text == f"5{NBSP}985{NBSP}000,00"
        assert " " not in text                        # обычного пробела нет


def test_none_renders_as_empty_not_as_word_none(tmp_path):
    out = render_docx(FIXTURE, _context(None, rows=[("1", None, None)]),
                      tmp_path / "o.docx")
    assert "None" not in _all_text(out)
    assert docx.Document(out).paragraphs[0].text == ""


def test_output_opens_in_python_docx(tmp_path):
    rows = [(str(i), f"Позиция & {i}", "<прим.>") for i in range(1, 24)]
    out = render_docx(FIXTURE, _context('АО «НК «Роснефть»', rows=rows),
                      tmp_path / "o.docx")
    with zipfile.ZipFile(out) as z:
        assert z.testzip() is None
        for name in z.namelist():
            if name.endswith((".xml", ".rels")):
                etree.fromstring(z.read(name))        # каждая часть -- корректный XML
    assert docx.Document(out).paragraphs[0].text == 'АО «НК «Роснефть»'


def test_template_file_is_not_modified(tmp_path):
    before = FIXTURE.read_bytes()
    render_docx(FIXTURE, _context(), tmp_path / "o.docx")
    assert FIXTURE.read_bytes() == before


def test_missing_key_names_template_and_key_in_russian(tmp_path):
    with pytest.raises(TemplateError) as exc:
        render_docx(FIXTURE, {"assets": {"headers": [], "rows": []}}, tmp_path / "o.docx")
    message = exc.value.args[0]
    assert "minimal.docx" in message
    assert "org" in message                           # назван недостающий ключ
    assert "undefined" not in message.lower()         # не строка из Jinja
    assert not (tmp_path / "o.docx").exists()         # битый файл не оставлен


def test_missing_nested_key_is_an_error_not_an_empty_cell(tmp_path):
    with pytest.raises(TemplateError) as exc:
        render_docx(FIXTURE, _context() | {"org": {"full_name": "x"}}, tmp_path / "o.docx")
    assert "short_name" in exc.value.args[0]


def test_broken_template_raises_template_error(tmp_path):
    template = _template_with(tmp_path, "Заказчик: {{ org.short_name ")
    with pytest.raises(TemplateError) as exc:
        render_docx(template, _context(), tmp_path / "o.docx")
    assert "t.docx" in exc.value.args[0]
    assert not (tmp_path / "o.docx").exists()


def test_unknown_filter_raises_template_error(tmp_path):
    template = _template_with(tmp_path, "{{ org.short_name | нет_такого }}")
    with pytest.raises(TemplateError) as exc:
        render_docx(template, _context(), tmp_path / "o.docx")
    assert "t.docx" in exc.value.args[0]


def test_missing_template_file_raises_template_error(tmp_path):
    with pytest.raises(TemplateError) as exc:
        render_docx(tmp_path / "нет.docx", _context(), tmp_path / "o.docx")
    assert "нет.docx" in exc.value.args[0]
    assert "найден" in exc.value.args[0]


def test_unreadable_template_raises_template_error(tmp_path):
    template = tmp_path / "мусор.docx"
    template.write_text("это не документ Word", encoding="utf-8")
    with pytest.raises(TemplateError) as exc:
        render_docx(template, _context(), tmp_path / "o.docx")
    assert "мусор.docx" in exc.value.args[0]


def test_unwritable_output_raises_template_error(tmp_path):
    # Путь вывода -- папка: то же поведение, что и у файла, открытого в Word.
    with pytest.raises(TemplateError) as exc:
        render_docx(FIXTURE, _context(), tmp_path)
    assert "сохранить" in exc.value.args[0]


def test_line_break_and_edge_spaces_in_value_survive(tmp_path):
    value = "  г. Москва,\nул. Тверская  "
    out = render_docx(FIXTURE, _context(value, rows=[("1", value, "")]),
                      tmp_path / "o.docx")
    assert docx.Document(out).paragraphs[0].text == value
    assert _table(out)[1][1] == value
