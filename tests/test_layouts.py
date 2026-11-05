import json
from dataclasses import replace
from pathlib import Path

import pytest

from app.models import Asset, Employee
from app.render.layouts import (
    Column, LayoutError, TableLayout, build_rows, load_layout,
)

LAYOUTS = Path(__file__).resolve().parent.parent / "layouts"
KOR = LAYOUTS / "assets_koryazhma.json"
RNPK = LAYOUTS / "assets_rnpk.json"
EMPLOYEES = LAYOUTS / "employees_koryazhma.json"

ASSET = Asset(
    name="Компьютер", maker_model="Intel Core i3 2130", quantity=3,
    location="г.Пермь, Куйбышева, 57а",
    ownership_basis="Право собственности",
    ownership_doc="Свидетельство № 59-БД 081725 от 31.10.2013 г.",
)


def test_missing_layout_file_raises_layout_error_in_russian(tmp_path):
    with pytest.raises(LayoutError) as caught:
        load_layout(tmp_path / "net_takogo.json")
    text = str(caught.value)
    assert "net_takogo.json" in text and "не найден" in text
    assert "No such file" not in text and "Errno" not in text


def test_same_asset_renders_differently_under_both_layouts():
    kor = build_rows(load_layout(KOR), [ASSET])[0]
    rnpk = build_rows(load_layout(RNPK), [ASSET])[0]
    assert len(kor) == 7 and len(rnpk) == 7      # у РНПК первая -- нумерация
    # у Коряжмы наименование и марка разнесены, у РНПК -- в одной клетке
    assert kor[0] == "Компьютер" and kor[1] == "Intel Core i3 2130"
    assert rnpk[1] == "Компьютер Intel Core i3 2130"


def test_parts_are_joined_with_the_declared_separator():
    kor = build_rows(load_layout(KOR), [ASSET])[0]
    assert kor[6] == ("Право собственности: "
                      "Свидетельство № 59-БД 081725 от 31.10.2013 г.")


def test_empty_part_leaves_no_dangling_separator():
    a = replace(ASSET, ownership_doc="")
    kor = build_rows(load_layout(KOR), [a])[0]
    assert kor[6] == "Право собственности"       # без ": " на конце


def test_empty_leading_part_leaves_no_leading_separator():
    a = replace(ASSET, ownership_basis="")
    kor = build_rows(load_layout(KOR), [a])[0]
    assert kor[6] == "Свидетельство № 59-БД 081725 от 31.10.2013 г."


def test_column_without_parts_renders_empty():
    kor = build_rows(load_layout(KOR), [ASSET])[0]
    assert kor[4] == ""                          # «Мощность» -- нет данных


def test_numbered_layout_adds_sequence():
    rows = build_rows(load_layout(RNPK), [ASSET, ASSET])
    assert rows[0][0] == "1" and rows[1][0] == "2"


def test_unnumbered_layout_has_no_sequence_cell():
    layout = load_layout(KOR)
    assert layout.numbered is False
    assert build_rows(layout, [ASSET])[0][0] == "Компьютер"


def test_numbered_layout_has_an_own_header_for_the_number_column():
    layout = load_layout(RNPK)
    assert layout.numbered is True
    assert layout.columns[0].header != ""
    # шапка и строки одной ширины: заголовок номера входит в columns
    assert len(layout.columns) == len(build_rows(layout, [ASSET])[0])


def test_unknown_field_raises_layout_error_naming_field_and_column():
    layout = TableLayout(
        source="assets",
        columns=[Column("Наименование", ["name"]),
                 Column("Изготовитель, марка", ["maker_modle"])],
        numbered=False,
    )
    with pytest.raises(LayoutError) as exc:
        build_rows(layout, [ASSET])
    message = str(exc.value)
    assert "maker_modle" in message
    assert "Изготовитель, марка" in message


def test_unknown_field_is_reported_even_without_records():
    layout = TableLayout(
        source="assets", columns=[Column("Колонка", ["nope"])], numbered=False)
    with pytest.raises(LayoutError, match="nope"):
        build_rows(layout, [])


def test_field_of_another_source_is_unknown(tmp_path):
    path = tmp_path / "bad.json"
    path.write_text(json.dumps({
        "source": "assets", "numbered": False,
        "columns": [{"header": "ФИО", "parts": ["full_name"]}],
    }, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(LayoutError, match="full_name"):
        load_layout(path)


def test_unknown_source_raises_layout_error(tmp_path):
    path = tmp_path / "bad.json"
    path.write_text(json.dumps({
        "source": "pets", "numbered": False, "columns": [],
    }), encoding="utf-8")
    with pytest.raises(LayoutError, match="pets"):
        load_layout(path)


def test_malformed_json_raises_layout_error(tmp_path):
    path = tmp_path / "bad.json"
    path.write_text("{ не json", encoding="utf-8")
    with pytest.raises(LayoutError):
        load_layout(path)


def test_column_without_header_raises_layout_error(tmp_path):
    path = tmp_path / "bad.json"
    path.write_text(json.dumps({
        "source": "assets", "numbered": False,
        "columns": [{"parts": ["name"]}],
    }), encoding="utf-8")
    with pytest.raises(LayoutError, match="header"):
        load_layout(path)


def test_none_renders_as_empty_string_not_the_word_none():
    blank = Asset(quantity=None)          # все необязательные поля пусты
    for path in (KOR, RNPK):
        row = build_rows(load_layout(path), [blank])[0]
        assert all("None" not in cell for cell in row)
    kor = build_rows(load_layout(KOR), [blank])[0]
    assert kor == [""] * 7
    assert build_rows(load_layout(RNPK), [blank])[0] == ["1"] + [""] * 6


def test_none_year_is_empty_but_set_year_is_plain_digits():
    layout = TableLayout(
        source="assets", columns=[Column("Год выпуска", ["year"])],
        numbered=False)
    assert build_rows(layout, [Asset(year=None)]) == [[""]]
    assert build_rows(layout, [Asset(year=2013)]) == [["2013"]]


def test_zero_is_not_treated_as_empty():
    layout = TableLayout(
        source="employees", columns=[Column("Стаж", ["experience_years"])],
        numbered=False)
    assert build_rows(layout, [Employee(experience_years=0)]) == [["0"]]


def test_headers_come_from_the_layout_not_from_the_rows():
    layout = load_layout(KOR)
    assert [c.header for c in layout.columns] == [
        "Наименование", "Изготовитель, марка", "Количество", "Год выпуска",
        "Мощность (производительность)", "Нынешнее состояние", "Источник"]
    rows = build_rows(layout, [ASSET])
    assert len(rows) == 1                        # шапки среди строк нет
    assert rows[0][0] != "Наименование"
    assert build_rows(layout, []) == []


def test_employees_layout_has_five_columns():
    layout = load_layout(EMPLOYEES)
    assert layout.source == "employees" and layout.numbered is True
    assert [c.header for c in layout.columns] == [
        "№ п/п", "Должность", "ФИО", "Образование, квалификация",
        "Стаж, лет"]
    e = Employee(
        full_name="Иванова Мария Петровна", position="Главный инженер",
        education="Окончила Томский политехнический институт в 1974 г., "
                  "Химик-технолог",
        graduation_year=1974, specialty="Химик-технолог",
        qualification="Удостоверение № 12", experience_years=30)
    row = build_rows(layout, [e])[0]
    assert row == [
        "1", "Главный инженер", "Иванова Мария Петровна",
        "Окончила Томский политехнический институт в 1974 г., "
        "Химик-технолог, Удостоверение № 12", "30"]


def test_layout_files_are_utf8_with_readable_russian():
    for path in (KOR, RNPK, EMPLOYEES):
        raw = path.read_bytes()
        assert not raw.startswith(b"\xef\xbb\xbf")
        text = raw.decode("utf-8")
        assert r"\u04" not in text              # кириллица не экранирована
        assert any("А" <= ch <= "я" for ch in text)
