"""Рендер Word-шаблонов.

Каждая проверка текста читает его из готового файла через python-docx,
а не из словаря на входе: тест, сверяющий вход с самим собой, ничего не доказывает.
"""
import errno
import zipfile
from pathlib import Path

import docx
import pytest
from lxml import etree

from app.models import Asset
from app.render.context import money
from app.render.layouts import build_rows, load_layout
from app.render import docx_render
from app.render.docx_render import TemplateError, render_docx

FIXTURE = Path(__file__).parent / "fixtures" / "minimal.docx"
LAYOUTS = Path(__file__).parent.parent / "layouts"
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


def _assets(count):
    return [Asset(name=f"Прибор {i}", maker_model=f"Марка-{i}", quantity=i, year=2000 + i,
                  location=f"Склад {i}", ownership_basis="собственность",
                  ownership_doc=f"св-во № {i}", purpose=f"цель {i}",
                  condition="рабочее")
            for i in range(1, count + 1)]


def _layout_context(path, count):
    layout = load_layout(path)
    return {"org": {"short_name": "ООО «Техносфера»"},
            "assets": {"headers": [c.header for c in layout.columns],
                       "rows": build_rows(layout, _assets(count))}}


@pytest.mark.parametrize("layout_file", ["assets_koryazhma.json", "assets_rnpk.json"])
def test_one_template_renders_any_layout(tmp_path, layout_file):
    """Число колонок задаёт раскладка, а не шаблон: шаблон один."""
    ctx = _layout_context(LAYOUTS / layout_file, 23)
    out = render_docx(FIXTURE, ctx, tmp_path / "o.docx")
    table = _table(out)
    assert len(table) == 1 + 23
    assert table[0] == ctx["assets"]["headers"]
    assert table[1:] == ctx["assets"]["rows"]
    assert {len(row) for row in table} == {7}


def test_same_template_renders_six_and_seven_columns(tmp_path):
    seven = _layout_context(LAYOUTS / "assets_koryazhma.json", 3)
    six = {"org": seven["org"],
           "assets": {"headers": seven["assets"]["headers"][:6],
                      "rows": [r[:6] for r in seven["assets"]["rows"]]}}
    for name, ctx, width in (("six", six, 6), ("seven", seven, 7)):
        out = render_docx(FIXTURE, ctx, tmp_path / f"{name}.docx")
        table = docx.Document(out).tables[0]
        assert [len(r.cells) for r in table.rows] == [width] * 4
        assert _table(out) == [ctx["assets"]["headers"]] + ctx["assets"]["rows"]
        grid = table._tbl.tblGrid.findall(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gridCol")
        assert len(grid) == width                     # сетка таблицы сходится с клетками


def test_header_formatting_follows_every_generated_cell(tmp_path):
    """Оформление несёт клетка-прототип: жирный шрифт и заливка шапки на всех колонках."""
    ctx = _layout_context(LAYOUTS / "assets_koryazhma.json", 2)
    out = render_docx(FIXTURE, ctx, tmp_path / "o.docx")
    header = docx.Document(out).tables[0].rows[0].cells
    assert len(header) == 7
    for cell in header:
        assert all(run.bold for run in cell.paragraphs[0].runs if run.text)
        assert 'w:fill="D9D9D9"' in cell._tc.xml


def test_empty_rows_keep_every_header_cell(tmp_path):
    ctx = _layout_context(LAYOUTS / "assets_rnpk.json", 0)
    out = render_docx(FIXTURE, ctx, tmp_path / "o.docx")
    assert _table(out) == [ctx["assets"]["headers"]]
    assert len(_table(out)[0]) == 7


# ---- причина отказа записи: что можно отличить -- отличаем -----------------------

# Раньше любой OSError получал подсказку «проверьте, что файл не открыт в Word»:
# человек шёл закрывать Word, которого не открывал, а на диске не было места
# или в папку было нельзя записывать. Эти пробы и держат разводку причин.
SHARING_VIOLATION = PermissionError(errno.EACCES, "The process cannot access the file")
SHARING_VIOLATION.winerror = 32

SAVE_FAILURES = [
    (OSError(errno.ENOSPC, "No space left on device"),
     "на диске не осталось места", ("Word", "права")),
    (PermissionError(errno.EACCES, "Permission denied"),
     "в эту папку нельзя записывать", ("в Word", "на диске")),
    (SHARING_VIOLATION,
     "открыт в другой программе", ("на диске", "права на папку")),
    (FileNotFoundError(errno.ENOENT, "No such file or directory"),
     "папки для него больше нет", ("Word", "на диске")),
    (OSError(errno.EIO, "I/O error"),
     "записать его не удалось", ()),
]
SAVE_IDS = ["нет-места", "нет-прав", "занят-другой-программой", "нет-папки",
            "прочее"]


@pytest.mark.parametrize("failure, expected, forbidden", SAVE_FAILURES,
                         ids=SAVE_IDS)
def test_save_failure_names_the_cause_the_os_reported(
        tmp_path, monkeypatch, failure, expected, forbidden):
    def refuse(self, path):
        raise failure

    monkeypatch.setattr(docx_render.DocxTemplate, "save", refuse)
    with pytest.raises(TemplateError) as exc:
        render_docx(FIXTURE, _context(), tmp_path / "o.docx")
    text = exc.value.args[0]
    assert "o.docx" in text and expected in text
    for wrong in forbidden:
        assert wrong not in text, wrong
    # наружу уходит русская фраза, а не текст и код операционной системы
    assert "Errno" not in text and "Traceback" not in text
    assert str(failure.errno) not in text


# ---- пустая графа: что документ напечатал пустым ---------------------------------

# Заявку отклоняют и за одну пустую графу, а Word, в отличие от книги Excel,
# понятия обязательного поля не имел вовсе: заявка ушла с четырьмя пустыми
# строками банковских реквизитов, и человек узнал об этом от комиссии.
# Проверки ниже читают список, который отрисовка отдала о себе, и текст
# готового файла -- не словарь на входе.

def _blanks(tmp_path, paragraph, context, name="o.docx"):
    """Пути полей, напечатанных пустыми, и текст готового документа."""
    found: list[str] = []
    out = render_docx(_template_with(tmp_path, paragraph), context,
                      tmp_path / name, found)
    return found, "\n".join(p.text for p in docx.Document(out).paragraphs)


def test_a_field_printed_empty_is_named_with_its_path(tmp_path):
    found, text = _blanks(
        tmp_path, "{{ org.bank_name }} / {{ org.inn }}",
        {"org": {"bank_name": "", "inn": "5904188041"}})
    assert found == ["org.bank_name"]
    assert text == " / 5904188041"                 # документ собран и подписан


def test_a_filled_field_says_nothing(tmp_path):
    # обратная сторона: сигнал, который горит всегда, через неделю перестают читать
    found, text = _blanks(
        tmp_path, "{{ org.bank_name }} / {{ org.inn }}",
        {"org": {"bank_name": "АО «БАНК»", "inn": "5904188041"}})
    assert found == []
    assert text == "АО «БАНК» / 5904188041"


def test_a_value_of_spaces_alone_counts_as_empty(tmp_path):
    # графа из пробелов выглядит пустой и в заявке, и для комиссии
    found, _ = _blanks(tmp_path, "{{ org.head_basis }}",
                       {"org": {"head_basis": "   "}})
    assert found == ["org.head_basis"]


def test_a_missing_value_of_none_counts_as_empty(tmp_path):
    found, text = _blanks(tmp_path, "{{ tender.vat_rate }}",
                          {"tender": {"vat_rate": None}})
    assert found == ["tender.vat_rate"]
    assert text == "" and "None" not in text


def test_the_optional_filter_is_how_a_legal_blank_is_marked(tmp_path):
    # Факс, которого у фирмы нет, -- не повод пугать человека перед подачей.
    # Законная пустота отмечается в самих данных, то есть в шаблоне.
    found, text = _blanks(tmp_path, "Факс: {{ org.fax|optional }}",
                          {"org": {"fax": ""}})
    assert found == []
    assert text == "Факс: "                        # графа осталась пустой


def test_the_optional_filter_still_prints_a_value_when_there_is_one(tmp_path):
    # обратная сторона: фильтр молчит о пустоте, а не глотает значение
    found, text = _blanks(tmp_path, "Факс: {{ org.fax|optional }}",
                          {"org": {"fax": "8 (342) 20-66-991"}})
    assert found == [] and text == "Факс: 8 (342) 20-66-991"


def test_a_blank_the_template_handles_itself_is_not_named(tmp_path):
    # «{{ a or b }}» печатает b: пустота не напечатана, и говорить не о чем
    found, text = _blanks(tmp_path, "{{ org.website or 'сайта нет' }}",
                          {"org": {"website": ""}})
    assert found == []
    assert text == "сайта нет"


def test_a_blank_only_asked_about_in_a_condition_is_not_named(tmp_path):
    found, text = _blanks(
        tmp_path, "{% if org.website %}{{ org.website }}{% endif %}конец",
        {"org": {"website": ""}})
    assert found == []
    assert text == "конец"


def test_the_same_field_printed_twice_is_named_once(tmp_path):
    found, _ = _blanks(tmp_path, "{{ org.bank_bik }} и ещё {{ org.bank_bik }}",
                       {"org": {"bank_bik": ""}})
    assert found == ["org.bank_bik"]


def test_an_empty_date_keeps_its_field_name_through_date_ru(tmp_path):
    # date_ru и date_long отдают пустую дату как есть именно для этого:
    # «{{ date_ru(tender.doc_date) }}» на пустой дате печатает пустую графу
    from app.render.context import date_long, date_ru
    context = {"tender": {"doc_date": "", "purchase_date": ""},
               "date_ru": date_ru, "date_long": date_long}
    found, text = _blanks(
        tmp_path,
        "{{ date_ru(tender.doc_date) }}|{{ date_long(tender.purchase_date) }}",
        context)
    assert found == ["tender.doc_date", "tender.purchase_date"]
    assert text == "|"


def test_a_filled_date_says_nothing(tmp_path):
    from app.render.context import date_ru
    found, text = _blanks(tmp_path, "{{ date_ru(tender.doc_date) }}",
                          {"tender": {"doc_date": "2025-10-24"},
                           "date_ru": date_ru})
    assert found == [] and text == "24.10.2025"


def test_empty_cells_of_tables_and_subtables_are_not_named(tmp_path):
    """Граница проверки, закреплённая нарочно: в клетке табличной справки и
    в строке подтаблицы пустое значение законно (примечание, документ
    о владении) -- ровно как в блоке строк карты Excel. Иначе предупреждение
    о каждой справке горело бы всегда."""
    context = {"items": [{"name": "", "price": "0,00"}],
               "assets": {"headers": ["Наименование", "Примечание"],
                          "rows": [["Компьютер", ""]]}}
    found, text = _blanks(
        tmp_path,
        "{% for it in items %}{{ it.name }}:{{ it.price }}{% endfor %}"
        "{% for row in assets.rows %}{% for c in row %}[{{ c }}]{% endfor %}"
        "{% endfor %}",
        context)
    assert found == []
    assert text == ":0,00[Компьютер][]"


def test_a_blank_that_went_through_a_filter_is_not_seen(tmp_path):
    """Второй край той же границы: значение, прошедшее через фильтр или
    функцию шаблона, приходит оттуда обычной строкой, и пустота в нём не
    видна. Сказано об этом в начале модуля; тест стоит здесь, чтобы никто
    не принял список пустых граф за полный."""
    found, text = _blanks(tmp_path, "[{{ org.bank_name|upper }}]",
                          {"org": {"bank_name": ""}})
    assert found == [] and text == "[]"


def test_the_source_context_is_not_touched_by_the_check(tmp_path):
    # в книгу Excel и на экран уходит тот же словарь: пометки в нём остаться
    # не должны
    context = {"org": {"bank_name": ""}}
    _blanks(tmp_path, "{{ org.bank_name }}", context)
    assert context == {"org": {"bank_name": ""}}
    assert type(context["org"]["bank_name"]) is str


def test_blanks_are_not_collected_when_the_caller_does_not_ask(tmp_path):
    # список необязателен: экраны и предпросмотр о пустых графах не спрашивают
    out = render_docx(_template_with(tmp_path, "{{ org.bank_name }}"),
                      {"org": {"bank_name": ""}}, tmp_path / "o.docx")
    assert out.is_file()
