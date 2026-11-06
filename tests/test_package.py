"""Сборка пакета документов.

Шаблоны здесь настоящие .docx, построенные python-docx в папке теста; книга
заказчика -- настоящая форма Роснефти из tests/fixtures. Каждая проверка читает
готовые файлы с диска, а не результат вызова.
"""
import shutil
from decimal import Decimal
from pathlib import Path

import docx
import pytest
from openpyxl import load_workbook

from app import db
from app.models import (
    ROLE_CUSTOMER, ROLE_PARTICIPANT, Asset, CustomerProfile, DocSpec,
    ExperienceRecord, LotItem, Organization, Tender,
)
from app.render import package as package_module
from app.render.package import (
    PackageError, PackageResult, build_package, planned_files, safe_filename,
)
from app.repo import Repo

ROOT = Path(__file__).resolve().parent.parent
REAL_FORM = Path(__file__).parent / "fixtures" / "forma_tkp_rosneft_murmansk.xlsx"

DECLARATION = "Декларация соответствия"
EQUIPMENT = "Справка о наличии оборудования"
MTR = "Справка о МТР"
EXPERIENCE = "Справка об опыте"
TKP = "ТКП по форме заказчика"

FORBIDDEN = set(':/\\|?*"<>')
LONG_SUBJECT = ("Разработка проекта санитарно-защитной зоны, проекта нормативов "
                "образования отходов и лимитов на их размещение ") * 3
LONG_SUBJECT = LONG_SUBJECT.strip()[:300]


# ---- помощники ----------------------------------------------------------------

def make_template(path: Path, *lines: str) -> Path:
    """Word-шаблон из абзацев с тегами."""
    doc = docx.Document()
    for line in lines:
        doc.add_paragraph(line)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    return path


def text_of(path: Path) -> list[str]:
    return [p.text for p in docx.Document(path).paragraphs]


@pytest.fixture
def repo(tmp_path):
    conn = db.connect(tmp_path / "t.db")
    db.migrate(conn)
    yield Repo(conn)
    conn.close()


@pytest.fixture
def dirs(tmp_path, monkeypatch):
    """Папки шаблонов и карт -- в папке теста, а не в проекте."""
    templates = tmp_path / "templates_docs"
    cellmaps = tmp_path / "cellmaps"
    templates.mkdir()
    cellmaps.mkdir()
    monkeypatch.setattr(package_module, "TEMPLATES_DOCS_DIR", templates)
    monkeypatch.setattr(package_module, "CELLMAPS_DIR", cellmaps)
    return templates, cellmaps


def make_tender(repo, *, docs=None, documents=(), customer_name="МУП «Полигон»",
                subject="Разработка КЭР", with_profile=True, **tender_kw) -> int:
    me = repo.save_organization(Organization(
        role=ROLE_PARTICIPANT, inn="5904188041", short_name="ООО «Техносфера»",
        head_position="Директор", head_name="Рабичко Елена Михайловна"))
    customer = repo.save_organization(Organization(
        role=ROLE_CUSTOMER, inn="2904016130", short_name=customer_name))
    if with_profile:
        repo.save_customer_profile(CustomerProfile(
            customer_id=customer, documents=list(documents),
            docs=docs or {}, declaration_text="Декларация заказчика"))
    return repo.save_tender(Tender(
        customer_id=customer, participant_id=me, subject=subject,
        purchase_number="0556300000225000001", **tender_kw))


def names(directory: Path) -> list[str]:
    return sorted(p.name for p in directory.iterdir())


# ---- safe_filename ------------------------------------------------------------

def test_safe_filename_strips_forbidden_chars():
    assert safe_filename('АО «НК «Роснефть»: СЗЗ/ТКП №2000726410?') \
        == 'АО «НК «Роснефть» СЗЗ-ТКП №2000726410'


def test_safe_filename_leaves_no_forbidden_char_and_no_trailing_dot():
    result = safe_filename('a:b/c\\d|e?f*g"h<i>j.')
    assert not FORBIDDEN & set(result)
    assert not result.endswith(".")
    assert "\\" not in result and "/" not in result


def test_safe_filename_keeps_cyrillic_guillemets_and_numero():
    name = "Справка о МТР «Техносфера» №5"
    assert safe_filename(name) == name


def test_safe_filename_trims_whitespace_and_trailing_dots_at_both_ends():
    assert safe_filename("  1. Декларация. ") == "1. Декларация"
    assert safe_filename("Декларация . . ") == "Декларация"


def test_safe_filename_collapses_gaps_left_by_removed_characters():
    assert safe_filename("Договор : №5") == "Договор №5"


def test_safe_filename_turns_line_breaks_into_spaces_and_drops_other_controls():
    assert safe_filename("Декларация\n\tсоответствия\x00") == "Декларация соответствия"


def test_safe_filename_truncates_to_windows_limit():
    result = safe_filename("Я" * 400)
    assert len(result) == 255
    # обрезка не оставляет пробел или точку на конце
    cut = safe_filename("а" * 253 + " . Б")
    assert not cut.endswith((" ", ".")) and len(cut) <= 255


def test_safe_filename_honours_a_shorter_limit():
    assert len(safe_filename("Я" * 400, max_length=100)) == 100


def test_safe_filename_never_returns_an_empty_name():
    for hopeless in ("", "   ", "...", '*?"<>|', ".."):
        assert safe_filename(hopeless)


def test_safe_filename_dodges_reserved_windows_device_names():
    for reserved in ("CON", "nul", "COM1", "lpt9", "AUX"):
        assert safe_filename(reserved).upper() != reserved.upper()


def test_safe_filename_is_not_fooled_by_dots_alone_as_a_path():
    result = safe_filename("..\\..\\Windows")
    assert "\\" not in result and "/" not in result and result != ".."


# ---- имена файлов и префиксы --------------------------------------------------

def test_prefixes_come_from_customer_profile(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "Декларация")
    make_template(templates / f"{MTR}.docx", "МТР")

    kor = make_tender(repo, customer_name="Коряжма", documents=[DECLARATION],
                      docs={DECLARATION: DocSpec(prefix="1.")})
    rnpk = make_tender(repo, customer_name="РНПК", documents=[MTR],
                       docs={MTR: DocSpec(prefix="8.1_")})

    first = build_package(kor, [DECLARATION], tmp_path / "kor", repo)
    second = build_package(rnpk, [MTR], tmp_path / "rnpk", repo)

    assert names(tmp_path / "kor") == [f"1.{DECLARATION}.docx"]
    assert names(tmp_path / "rnpk") == [f"8.1_{MTR}.docx"]
    assert first.written == [tmp_path / "kor" / f"1.{DECLARATION}.docx"]
    assert second.skipped == []


def test_spaces_around_prefix_and_name_never_reach_the_filename(
        repo, dirs, tmp_path):
    # layout и prefix приходят из профиля нетронутыми; пробел перед расширением
    # или в начале имени Windows молча съедает или показывает криво
    templates, _ = dirs
    make_template(templates / "t.docx", "x")
    cases = {" 2.": "2.Справка.docx", "3. ": "3. Справка.docx"}
    for number, (prefix, expected) in enumerate(cases.items()):
        tid = make_tender(repo, documents=["Справка "],
                          docs={"Справка ": DocSpec(template="t.docx", prefix=prefix)})
        out = tmp_path / f"out{number}"
        build_package(tid, ["Справка "], out, repo)
        assert names(out) == [expected]


def test_document_without_prefix_gets_a_bare_name(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION],
                      docs={DECLARATION: DocSpec(template="", prefix="")})
    build_package(tid, [DECLARATION], tmp_path / "out", repo)
    assert names(tmp_path / "out") == [f"{DECLARATION}.docx"]


def test_explicit_template_name_from_profile_is_used(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / "Декларация Коряжмы.docx", "Своя форма {{ org.short_name }}")
    tid = make_tender(repo, documents=[DECLARATION], docs={
        DECLARATION: DocSpec(template="Декларация Коряжмы.docx", prefix="1.")})
    result = build_package(tid, [DECLARATION], tmp_path / "out", repo)
    assert result.skipped == []
    assert text_of(result.written[0]) == ["Своя форма ООО «Техносфера»"]


def test_document_absent_from_docs_uses_defaults(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "{{ declaration_text }}")
    # профиль есть, документ в docs не описан
    tid = make_tender(repo, documents=[DECLARATION], docs={})
    result = build_package(tid, [DECLARATION], tmp_path / "out", repo)
    assert result.skipped == []
    assert names(tmp_path / "out") == [f"{DECLARATION}.docx"]
    assert text_of(result.written[0]) == ["Декларация заказчика"]


def test_customer_without_profile_still_gets_a_package(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "{{ customer.short_name }}")
    tid = make_tender(repo, with_profile=False)
    result = build_package(tid, [DECLARATION], tmp_path / "out", repo)
    assert result.skipped == []
    assert text_of(result.written[0]) == ["МУП «Полигон»"]


def test_tender_without_customer_still_gets_a_package(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "[{{ customer.short_name }}]")
    tid = repo.save_tender(Tender(subject="Без заказчика"))
    result = build_package(tid, [DECLARATION], tmp_path / "out", repo)
    assert result.skipped == []
    assert text_of(result.written[0]) == ["[]"]


# ---- пути: ничего не пишется за пределами папки -------------------------------

def test_prefix_with_path_separators_cannot_escape_the_folder(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    out = tmp_path / "deep" / "out"
    for number, prefix in enumerate(("..\\", "../", "..\\..\\", "C:\\Temp\\", "/etc/")):
        tid = make_tender(repo, documents=[DECLARATION],
                          docs={DECLARATION: DocSpec(prefix=prefix)})
        target = out / str(number)
        result = build_package(tid, [DECLARATION], target, repo)
        assert len(result.written) == 1
        assert result.written[0].parent == target
        assert len(names(target)) == 1
    # ничего не появилось рядом с папкой
    assert names(tmp_path / "deep") == ["out"]
    assert sorted(p.name for p in tmp_path.iterdir()) == sorted(
        ["deep", "templates_docs", "cellmaps", "t.db"])


def test_document_name_with_separators_cannot_escape_the_folder(repo, dirs, tmp_path):
    templates, _ = dirs
    evil = "..\\..\\Пакет/Справка"
    # в default-шаблоне имя с разделителем не создать, поэтому шаблон -- явно
    make_template(templates / "t.docx", "x")
    tid = make_tender(repo, documents=[evil],
                      docs={evil: DocSpec(template="t.docx")})
    out = tmp_path / "deep" / "out"
    result = build_package(tid, [evil], out, repo)
    assert len(result.written) == 1
    assert result.written[0].parent == out
    assert names(tmp_path / "deep") == ["out"]


def test_dotdot_only_document_name_is_written_inside_the_folder(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / "t.docx", "x")
    tid = make_tender(repo, documents=[".."], docs={"..": DocSpec(template="t.docx")})
    out = tmp_path / "deep" / "out"
    result = build_package(tid, [".."], out, repo)
    assert len(result.written) == 1 and result.written[0].parent == out


# ---- длинное значение: в документе целиком, в имени -- обрезанное -------------

def test_long_subject_reaches_document_in_full(repo, dirs, tmp_path):
    templates, _ = dirs
    assert len(LONG_SUBJECT) == 300
    make_template(templates / f"{DECLARATION}.docx", "Предмет: {{ tender.subject }}")
    tid = make_tender(repo, documents=[DECLARATION], subject=LONG_SUBJECT)
    result = build_package(tid, [DECLARATION], tmp_path / "out", repo)
    # содержимое: ни одного потерянного символа
    assert text_of(result.written[0]) == ["Предмет: " + LONG_SUBJECT]
    # а имя файла, составленное из того же текста, обрезается и остаётся допустимым
    shortened = safe_filename(LONG_SUBJECT)
    assert len(shortened) <= 255 < len(LONG_SUBJECT)
    assert LONG_SUBJECT.startswith(shortened)


def test_overlong_prefix_truncates_the_filename_but_keeps_the_extension(
        repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "{{ tender.subject }}")
    tid = make_tender(repo, documents=[DECLARATION], subject=LONG_SUBJECT,
                      docs={DECLARATION: DocSpec(prefix=LONG_SUBJECT + " ")})
    out = tmp_path / "out"
    result = build_package(tid, [DECLARATION], out, repo)
    assert len(result.written) == 1
    name = result.written[0].name
    assert name.endswith(".docx") and len(name) <= 255
    assert not FORBIDDEN & set(name)
    # содержимое то же -- в полном объёме
    assert text_of(result.written[0]) == [LONG_SUBJECT]


def test_filename_fits_in_windows_path_limit_for_a_deep_folder(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION],
                      docs={DECLARATION: DocSpec(prefix="П" * 300)})
    out = tmp_path / ("Папка-" + "н" * 40) / ("Вложенная-" + "в" * 40)
    result = build_package(tid, [DECLARATION], out, repo)
    assert len(result.written) == 1
    assert len(str(result.written[0])) <= 259


def test_folder_path_too_long_for_any_filename_is_a_skip_not_a_crash(
        repo, dirs, tmp_path, monkeypatch):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / ("к" * 100) / ("к" * 100) / ("к" * 100)   # длиннее 260
    # такую папку Windows без особой настройки не создаст -- создание подменено,
    # проверяется решение сборки, а не файловая система
    monkeypatch.setattr(Path, "mkdir", lambda *a, **k: None)
    result = build_package(tid, [DECLARATION], out, repo)
    assert result.written == []
    assert [d for d, _ in result.skipped] == [DECLARATION]
    assert "слишком длинный" in result.skipped[0][1]


# ---- один сбой не отменяет остальные ------------------------------------------

def test_failing_document_is_skipped_not_fatal(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "{{ org.short_name }}")
    make_template(templates / f"{MTR}.docx", "{{ нет_такого_ключа }}")
    make_template(templates / f"{EXPERIENCE}.docx", "{{ tender.subject }}")
    tid = make_tender(repo, documents=[DECLARATION, MTR, EXPERIENCE])
    out = tmp_path / "out"
    result = build_package(tid, [DECLARATION, MTR, EXPERIENCE], out, repo)

    assert [p.name for p in result.written] == [
        f"{DECLARATION}.docx", f"{EXPERIENCE}.docx"]
    assert names(out) == sorted([f"{DECLARATION}.docx", f"{EXPERIENCE}.docx"])
    assert len(result.skipped) == 1
    document, reason = result.skipped[0]
    assert document == MTR
    assert "нет_такого_ключа" in reason      # причина названа словами шаблона
    assert "Traceback" not in reason


def test_missing_template_file_is_a_russian_skip(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION, MTR])
    result = build_package(tid, [DECLARATION, MTR], tmp_path / "out", repo)
    assert [p.name for p in result.written] == [f"{DECLARATION}.docx"]
    document, reason = result.skipped[0]
    assert document == MTR
    assert f"{MTR}.docx" in reason and "templates_docs" in reason
    assert "не найден" in reason


def test_missing_explicit_template_names_the_file_from_the_profile(
        repo, dirs, tmp_path):
    tid = make_tender(repo, documents=[DECLARATION], docs={
        DECLARATION: DocSpec(template="Форма заказчика.docx")})
    result = build_package(tid, [DECLARATION], tmp_path / "out", repo)
    assert result.written == []
    assert "Форма заказчика.docx" in result.skipped[0][1]


def test_corrupt_template_is_a_skip(repo, dirs, tmp_path):
    templates, _ = dirs
    (templates / f"{DECLARATION}.docx").write_bytes(b"not a zip")
    tid = make_tender(repo, documents=[DECLARATION])
    result = build_package(tid, [DECLARATION], tmp_path / "out", repo)
    assert result.written == []
    assert "повреждён" in result.skipped[0][1]


def test_unsupported_template_type_is_a_skip(repo, dirs, tmp_path):
    templates, _ = dirs
    (templates / "форма.pdf").write_bytes(b"%PDF")
    tid = make_tender(repo, documents=[DECLARATION], docs={
        DECLARATION: DocSpec(template="форма.pdf")})
    result = build_package(tid, [DECLARATION], tmp_path / "out", repo)
    assert result.written == []
    assert ".pdf" in result.skipped[0][1]


def test_unexpected_error_in_one_document_does_not_cost_the_others(
        repo, dirs, tmp_path, monkeypatch):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    make_template(templates / f"{MTR}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION, MTR])
    real = package_module.build_context

    def flaky(tender_id, repo_, *, document=None, **kw):
        if document == DECLARATION:
            raise RuntimeError("boom")
        return real(tender_id, repo_, document=document, **kw)

    monkeypatch.setattr(package_module, "build_context", flaky)
    result = build_package(tid, [DECLARATION, MTR], tmp_path / "out", repo)
    assert [p.name for p in result.written] == [f"{MTR}.docx"]
    document, reason = result.skipped[0]
    assert document == DECLARATION
    assert "RuntimeError" in reason and "boom" not in reason   # без сырого текста
    assert "Traceback" not in reason


def test_bad_layout_file_skips_only_its_document(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{EQUIPMENT}.docx", "x")
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[EQUIPMENT, DECLARATION], docs={
        EQUIPMENT: DocSpec(layout="нет_такой_раскладки.json")})
    result = build_package(tid, [EQUIPMENT, DECLARATION], tmp_path / "out", repo)
    assert [p.name for p in result.written] == [f"{DECLARATION}.docx"]
    assert result.skipped[0][0] == EQUIPMENT
    assert "нет_такой_раскладки.json" in result.skipped[0][1]


def test_failed_document_leaves_no_partial_file(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{MTR}.docx", "{{ нет_такого_ключа }}")
    tid = make_tender(repo, documents=[MTR])
    out = tmp_path / "out"
    build_package(tid, [MTR], out, repo)
    assert names(out) == []


def test_unknown_tender_is_refused_before_anything_is_written(repo, dirs, tmp_path):
    out = tmp_path / "out"
    with pytest.raises(LookupError):
        build_package(999, [DECLARATION], out, repo)
    assert not out.exists()


def test_empty_document_list_writes_nothing(repo, dirs, tmp_path):
    tid = make_tender(repo)
    result = build_package(tid, [], tmp_path / "out", repo)
    assert result == PackageResult(written=[], skipped=[])


# ---- контекст на каждый документ ----------------------------------------------

def test_each_document_gets_its_own_context(repo, dirs, tmp_path):
    # обе справки берут справочник МТР, но раскладки разные: контекст на весь
    # пакет отдал бы одной из них чужие колонки
    templates, _ = dirs
    repo.save_asset(Asset(name="Компьютер", maker_model="Intel i3", quantity=3))
    line = "{{ assets.headers | join('|') }}"
    make_template(templates / f"{EQUIPMENT}.docx", line)
    make_template(templates / f"{MTR}.docx", line)
    tid = make_tender(repo, documents=[EQUIPMENT, MTR], docs={
        EQUIPMENT: DocSpec(layout="assets_koryazhma.json", prefix="1."),
        MTR: DocSpec(layout="assets_rnpk.json", prefix="2.")})
    out = tmp_path / "out"
    result = build_package(tid, [EQUIPMENT, MTR], out, repo)
    assert result.skipped == []
    kor = text_of(out / f"1.{EQUIPMENT}.docx")[0]
    rnpk = text_of(out / f"2.{MTR}.docx")[0]
    assert kor != rnpk
    assert kor.startswith("Наименование") and rnpk.startswith("№ п/п")


def test_context_warnings_are_collected_per_document(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{EQUIPMENT}.docx", "x")
    make_template(templates / f"{MTR}.docx", "x")
    make_template(templates / f"{DECLARATION}.docx", "x")
    # у справки об оборудовании раскладка не назначена, у МТР назначена
    tid = make_tender(repo, documents=[EQUIPMENT, MTR, DECLARATION], docs={
        MTR: DocSpec(layout="assets_rnpk.json")})
    result = build_package(tid, [EQUIPMENT, MTR, DECLARATION], tmp_path / "out", repo)
    assert result.skipped == []
    assert [d for d, _ in result.warnings] == [EQUIPMENT]
    assert "раскладка" in result.warnings[0][1]


def test_warnings_of_a_skipped_document_are_not_lost(repo, dirs, tmp_path):
    # документ не собран, но предупреждение о его данных пользователю нужно
    templates, _ = dirs
    make_template(templates / f"{EQUIPMENT}.docx", "{{ нет_ключа }}")
    tid = make_tender(repo, documents=[EQUIPMENT])
    result = build_package(tid, [EQUIPMENT], tmp_path / "out", repo)
    assert result.written == [] and len(result.skipped) == 1
    assert [d for d, _ in result.warnings] == [EQUIPMENT]


# ---- папка и существующие файлы -----------------------------------------------

def test_creates_output_dir_if_absent(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "нет" / "такой" / "папки"
    assert not out.exists()
    result = build_package(tid, [DECLARATION], out, repo)
    assert out.is_dir() and len(result.written) == 1


def test_existing_output_dir_is_reused(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    out.mkdir()
    (out / "чужой файл.txt").write_text("мой", encoding="utf-8")
    build_package(tid, [DECLARATION], out, repo)
    assert names(out) == sorted(["чужой файл.txt", f"{DECLARATION}.docx"])


def test_output_path_that_is_a_file_is_refused_in_russian(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    blocker = tmp_path / "файл"
    blocker.write_text("x", encoding="utf-8")
    with pytest.raises(PackageError) as caught:
        build_package(tid, [DECLARATION], blocker, repo)
    assert "папк" in str(caught.value).lower()
    assert "файл" in str(caught.value).lower()


def test_existing_file_is_not_silently_overwritten(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "новый")
    make_template(templates / f"{MTR}.docx", "новый")
    tid = make_tender(repo, documents=[DECLARATION, MTR])
    out = tmp_path / "out"
    out.mkdir()
    existing = out / f"{DECLARATION}.docx"
    existing.write_bytes(b"signed by hand")

    result = build_package(tid, [DECLARATION, MTR], out, repo)

    assert existing.read_bytes() == b"signed by hand"          # не тронут
    assert [p.name for p in result.written] == [f"{MTR}.docx"]  # остальные собраны
    document, reason = result.skipped[0]
    assert document == DECLARATION
    assert f"{DECLARATION}.docx" in reason and "уже есть" in reason
    assert names(out) == sorted([f"{DECLARATION}.docx", f"{MTR}.docx"])


def test_second_build_into_the_same_folder_skips_everything_and_changes_nothing(
        repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    first = build_package(tid, [DECLARATION], out, repo)
    before = first.written[0].read_bytes()
    second = build_package(tid, [DECLARATION], out, repo)
    assert second.written == [] and len(second.skipped) == 1
    assert first.written[0].read_bytes() == before


def test_document_listed_twice_is_built_once(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    result = build_package(tid, [DECLARATION, DECLARATION], tmp_path / "out", repo)
    assert len(result.written) == 1
    assert len(result.skipped) == 1
    assert result.skipped[0][0] == DECLARATION
    assert "дважды" in result.skipped[0][1]
    assert "обработан один раз" in result.skipped[0][1]


def test_two_documents_with_the_same_final_name_do_not_overwrite_each_other(
        repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / "a.docx", "первый")
    make_template(templates / "b.docx", "второй")
    # после очистки от запрещённых символов имена совпадают
    tid = make_tender(repo, documents=["Справка: 1", "Справка? 1"], docs={
        "Справка: 1": DocSpec(template="a.docx"),
        "Справка? 1": DocSpec(template="b.docx")})
    out = tmp_path / "out"
    result = build_package(tid, ["Справка: 1", "Справка? 1"], out, repo)
    assert len(result.written) == 1 and len(result.skipped) == 1
    assert text_of(result.written[0]) == ["первый"]
    assert "совпало" in result.skipped[0][1]


# ---- книга Excel --------------------------------------------------------------

CONDITIONS = [
    "Цена и условия:", "Форма, срок и порядок оплаты:",
    "Условия и место оказания услуг:", "Срок оказания услуг:",
    "Срок действия коммерческого предложения:", "Срок действия договора:",
    "Возможность привлечения субподрядчиков/субисполнителей",
]


def rosneft_tender(repo, positions=2):
    tid = make_tender(
        repo, documents=[TKP], customer_name='АО «НК «Роснефть»',
        docs={TKP: DocSpec(template="forma_tkp_rosneft_murmansk.xlsx",
                           cellmap="tkp_rosneft_murmansk.json", prefix="3.")},
        deal_answers={c: "подтверждаю" for c in CONDITIONS},
        vat_rate=Decimal("0.05"))
    for n in range(1, positions + 1):
        repo.save_lot_item(LotItem(tender_id=tid, position=n, name=f"Позиция {n}",
                                   price=Decimal("100000.10")))
    return tid


def test_excel_form_is_filled_from_the_customers_own_workbook(repo, dirs, tmp_path):
    templates, cellmaps = dirs
    shutil.copy(REAL_FORM, templates / "forma_tkp_rosneft_murmansk.xlsx")
    shutil.copy(ROOT / "cellmaps" / "tkp_rosneft_murmansk.json", cellmaps)
    tid = rosneft_tender(repo)
    out = tmp_path / "out"
    result = build_package(tid, [TKP], out, repo)

    assert result.skipped == []
    assert names(out) == [f"3.{TKP}.xlsx"]
    ws = load_workbook(result.written[0])["Лист1"]
    assert ws["D8"].value == "Ставка НДС: 5%"
    assert ws["E9"].value == "=C9+D9"               # формулы заказчика целы
    # точное сравнение: приближённое спрятало бы регресс округления денег
    assert Decimal(str(ws["C9"].value)) == Decimal("100000.10")


def test_refused_excel_form_leaves_no_file_and_spares_the_word_documents(
        repo, dirs, tmp_path):
    templates, cellmaps = dirs
    shutil.copy(REAL_FORM, templates / "forma_tkp_rosneft_murmansk.xlsx")
    shutil.copy(ROOT / "cellmaps" / "tkp_rosneft_murmansk.json", cellmaps)
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = rosneft_tender(repo, positions=12)        # форма вмещает 11
    out = tmp_path / "out"
    result = build_package(tid, [TKP, DECLARATION], out, repo)
    assert [p.name for p in result.written] == [f"{DECLARATION}.docx"]
    assert result.skipped[0][0] == TKP
    assert "12" in result.skipped[0][1] and "11" in result.skipped[0][1]
    assert names(out) == [f"{DECLARATION}.docx"]


def test_missing_cellmap_file_is_a_russian_skip(repo, dirs, tmp_path):
    templates, _ = dirs
    shutil.copy(REAL_FORM, templates / "forma_tkp_rosneft_murmansk.xlsx")
    tid = rosneft_tender(repo)                      # карты в папке нет
    result = build_package(tid, [TKP], tmp_path / "out", repo)
    assert result.written == []
    assert "tkp_rosneft_murmansk.json" in result.skipped[0][1]


def test_half_written_file_of_a_failed_document_is_removed(
        repo, dirs, tmp_path, monkeypatch):
    # сохранение могло оборваться посреди записи: огрызок файла не оставляем,
    # иначе пересборка сочтёт его «уже существующим»
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])

    def crash_after_writing(template, context, out):
        Path(out).write_bytes(b"half")
        raise package_module.TemplateError("Не удалось сохранить.")

    monkeypatch.setattr(package_module, "render_docx", crash_after_writing)
    out = tmp_path / "out"
    result = build_package(tid, [DECLARATION], out, repo)
    assert len(result.skipped) == 1
    assert names(out) == []


# ---- пересборка: overwrite ----------------------------------------------------

def test_overwrite_false_is_the_default_and_refuses(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "новый")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    out.mkdir()
    existing = out / f"{DECLARATION}.docx"
    existing.write_bytes(b"old")
    result = build_package(tid, [DECLARATION], out, repo, overwrite=False)
    assert result.written == [] and len(result.skipped) == 1
    assert existing.read_bytes() == b"old"


def test_overwrite_true_replaces_the_existing_file(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "новый {{ org.short_name }}")
    make_template(templates / f"{MTR}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION, MTR])
    out = tmp_path / "out"
    out.mkdir()
    existing = out / f"{DECLARATION}.docx"
    existing.write_bytes(b"old")
    result = build_package(tid, [DECLARATION, MTR], out, repo, overwrite=True)
    assert result.skipped == []
    assert sorted(p.name for p in result.written) == sorted(
        [f"{DECLARATION}.docx", f"{MTR}.docx"])
    assert text_of(existing) == ["новый ООО «Техносфера»"]
    assert names(out) == sorted([f"{DECLARATION}.docx", f"{MTR}.docx"])  # без остатков


def test_rebuild_after_a_data_fix_shows_the_new_value(repo, dirs, tmp_path):
    # главный сценарий: собрал, увидел не ту цену, поправил, собрал снова
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "{{ tender.subject }}")
    tid = make_tender(repo, documents=[DECLARATION], subject="Старый предмет")
    out = tmp_path / "out"
    build_package(tid, [DECLARATION], out, repo)
    tender = repo.get_tender(tid)
    tender.subject = "Исправленный предмет"
    repo.save_tender(tender)
    result = build_package(tid, [DECLARATION], out, repo, overwrite=True)
    assert result.skipped == []
    assert text_of(out / f"{DECLARATION}.docx") == ["Исправленный предмет"]


def test_overwrite_true_keeps_the_old_file_when_the_new_one_cannot_be_built(
        repo, dirs, tmp_path):
    # пересборка не должна стоить пользователю прежнего файла
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "{{ нет_ключа }}")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    out.mkdir()
    existing = out / f"{DECLARATION}.docx"
    existing.write_bytes(b"old")
    result = build_package(tid, [DECLARATION], out, repo, overwrite=True)
    assert result.written == [] and len(result.skipped) == 1
    assert existing.read_bytes() == b"old"
    assert names(out) == [f"{DECLARATION}.docx"]


def test_overwrite_true_keeps_the_old_file_when_saving_breaks_halfway(
        repo, dirs, tmp_path, monkeypatch):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    out.mkdir()
    existing = out / f"{DECLARATION}.docx"
    existing.write_bytes(b"old")

    def crash_after_writing(template, context, target):
        Path(target).write_bytes(b"half")
        raise package_module.TemplateError("Не удалось сохранить.")

    monkeypatch.setattr(package_module, "render_docx", crash_after_writing)
    result = build_package(tid, [DECLARATION], out, repo, overwrite=True)
    assert len(result.skipped) == 1
    assert existing.read_bytes() == b"old"
    assert names(out) == [f"{DECLARATION}.docx"]


def test_overwrite_true_with_a_file_open_in_word_is_a_skip_and_keeps_it(
        repo, dirs, tmp_path, monkeypatch):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    out.mkdir()
    existing = out / f"{DECLARATION}.docx"
    existing.write_bytes(b"old")

    def locked(src, dst):
        raise PermissionError("locked")

    monkeypatch.setattr(package_module.os, "replace", locked)
    result = build_package(tid, [DECLARATION], out, repo, overwrite=True)
    assert result.written == []
    assert "открыт" in result.skipped[0][1]
    assert existing.read_bytes() == b"old"
    assert names(out) == [f"{DECLARATION}.docx"]       # временный файл убран


def test_overwrite_true_does_not_let_two_documents_share_one_name(
        repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / "a.docx", "первый")
    make_template(templates / "b.docx", "второй")
    tid = make_tender(repo, documents=["Справка: 1", "Справка? 1"], docs={
        "Справка: 1": DocSpec(template="a.docx"),
        "Справка? 1": DocSpec(template="b.docx")})
    out = tmp_path / "out"
    result = build_package(tid, ["Справка: 1", "Справка? 1"], out, repo,
                           overwrite=True)
    assert len(result.written) == 1 and len(result.skipped) == 1
    assert text_of(result.written[0]) == ["первый"]


def test_overwrite_true_replaces_an_excel_form_too(repo, dirs, tmp_path):
    templates, cellmaps = dirs
    shutil.copy(REAL_FORM, templates / "forma_tkp_rosneft_murmansk.xlsx")
    shutil.copy(ROOT / "cellmaps" / "tkp_rosneft_murmansk.json", cellmaps)
    tid = rosneft_tender(repo)
    out = tmp_path / "out"
    build_package(tid, [TKP], out, repo)
    result = build_package(tid, [TKP], out, repo, overwrite=True)
    assert result.skipped == [] and len(result.written) == 1
    assert names(out) == [f"3.{TKP}.xlsx"]


# ---- нет папки шаблонов: одна ошибка, а не по записи на документ ---------------

def test_missing_templates_folder_is_one_error_not_one_skip_per_document(
        repo, tmp_path, monkeypatch):
    monkeypatch.setattr(package_module, "TEMPLATES_DOCS_DIR",
                        tmp_path / "нет_такой_папки")
    monkeypatch.setattr(package_module, "CELLMAPS_DIR", tmp_path / "cellmaps")
    tid = make_tender(repo, documents=[DECLARATION, MTR])
    out = tmp_path / "out"
    with pytest.raises(PackageError) as caught:
        build_package(tid, [DECLARATION, MTR], out, repo)
    assert "нет_такой_папки" in str(caught.value) or "templates" in str(caught.value)
    assert "папк" in str(caught.value).lower()
    assert not out.exists()                      # ничего не создано


def test_templates_folder_that_is_a_file_is_the_same_single_error(
        repo, tmp_path, monkeypatch):
    blocker = tmp_path / "templates_docs"
    blocker.write_text("x", encoding="utf-8")
    monkeypatch.setattr(package_module, "TEMPLATES_DOCS_DIR", blocker)
    tid = make_tender(repo, documents=[DECLARATION])
    with pytest.raises(PackageError):
        build_package(tid, [DECLARATION], tmp_path / "out", repo)


def test_missing_templates_folder_with_nothing_to_build_is_not_an_error(
        repo, tmp_path, monkeypatch):
    monkeypatch.setattr(package_module, "TEMPLATES_DOCS_DIR", tmp_path / "нет")
    tid = make_tender(repo)
    result = build_package(tid, [], tmp_path / "out", repo)
    assert result == PackageResult(written=[], skipped=[])


def test_one_missing_template_in_an_existing_folder_stays_a_skip(
        repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION, MTR])
    result = build_package(tid, [DECLARATION, MTR], tmp_path / "out", repo)
    assert len(result.written) == 1
    assert [d for d, _ in result.skipped] == [MTR]


# ---- planned_files: что будет записано, до записи -----------------------------

def test_planned_files_equal_what_build_package_then_writes(repo, dirs, tmp_path):
    templates, cellmaps = dirs
    shutil.copy(REAL_FORM, templates / "forma_tkp_rosneft_murmansk.xlsx")
    shutil.copy(ROOT / "cellmaps" / "tkp_rosneft_murmansk.json", cellmaps)
    for name in (DECLARATION, MTR, EQUIPMENT):
        make_template(templates / f"{name}.docx", "x")
    docs = {
        DECLARATION: DocSpec(prefix="1."),
        MTR: DocSpec(prefix="8.1_ "),
        EQUIPMENT: DocSpec(prefix="П" * 300),            # имя обрежется
        TKP: DocSpec(template="forma_tkp_rosneft_murmansk.xlsx",
                     cellmap="tkp_rosneft_murmansk.json", prefix="3."),
    }
    tid = make_tender(repo, documents=list(docs), docs=docs,
                      deal_answers={c: "подтверждаю" for c in CONDITIONS})
    repo.save_lot_item(LotItem(tender_id=tid, position=1, name="П",
                               price=Decimal("1.00")))
    documents = [DECLARATION, MTR, EQUIPMENT, TKP, DECLARATION]
    out = tmp_path / "out"

    plan = planned_files(tid, documents, out, repo)
    result = build_package(tid, documents, out, repo)

    assert [path for _, path in plan] == result.written
    assert [doc for doc, _ in plan] == [DECLARATION, MTR, EQUIPMENT, TKP]


def test_planned_files_writes_nothing_and_creates_no_folder(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "нет" / "такой"
    plan = planned_files(tid, [DECLARATION], out, repo)
    assert plan == [(DECLARATION, out / f"{DECLARATION}.docx")]
    assert not (tmp_path / "нет").exists()


def test_planned_files_include_files_that_already_exist(repo, dirs, tmp_path):
    # именно по этому списку экран выбирает, о чём спросить про замену
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    make_template(templates / f"{MTR}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION, MTR])
    out = tmp_path / "out"
    out.mkdir()
    (out / f"{MTR}.docx").write_bytes(b"old")
    plan = planned_files(tid, [DECLARATION, MTR], out, repo)
    existing = [doc for doc, path in plan if path.exists()]
    assert existing == [MTR]
    assert [doc for doc, _ in plan] == [DECLARATION, MTR]


def test_planned_files_leave_out_documents_that_will_certainly_be_skipped(
        repo, dirs, tmp_path):
    templates, _ = dirs
    shutil.copy(REAL_FORM, templates / "forma_tkp_rosneft_murmansk.xlsx")
    make_template(templates / f"{DECLARATION}.docx", "x")
    (templates / "форма.pdf").write_bytes(b"%PDF")
    docs = {
        TKP: DocSpec(template="forma_tkp_rosneft_murmansk.xlsx",
                     cellmap="нет_карты.json"),               # карты нет
        EXPERIENCE: DocSpec(template="форма.pdf"),            # не Word и не Excel
    }
    tid = make_tender(repo, documents=[DECLARATION, MTR, TKP, EXPERIENCE],
                      docs=docs)
    # МТР -- шаблона нет вовсе
    plan = planned_files(tid, [DECLARATION, MTR, TKP, EXPERIENCE], tmp_path / "o", repo)
    assert [doc for doc, _ in plan] == [DECLARATION]
    built = build_package(tid, [DECLARATION, MTR, TKP, EXPERIENCE], tmp_path / "o", repo)
    assert [doc for doc, _ in built.skipped] == [MTR, TKP, EXPERIENCE]


def test_planned_files_do_not_promise_two_documents_one_name(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / "a.docx", "первый")
    make_template(templates / "b.docx", "второй")
    tid = make_tender(repo, documents=["Справка: 1", "Справка? 1"], docs={
        "Справка: 1": DocSpec(template="a.docx"),
        "Справка? 1": DocSpec(template="b.docx")})
    documents = ["Справка: 1", "Справка? 1"]
    plan = planned_files(tid, documents, tmp_path / "out", repo)
    assert [doc for doc, _ in plan] == ["Справка: 1"]
    assert [p for _, p in plan] == build_package(
        tid, documents, tmp_path / "out", repo).written


def test_planned_files_unknown_tender_is_lookup_error(repo, dirs, tmp_path):
    with pytest.raises(LookupError):
        planned_files(999, [DECLARATION], tmp_path / "out", repo)


def test_planned_files_missing_templates_folder_is_package_error(
        repo, tmp_path, monkeypatch):
    monkeypatch.setattr(package_module, "TEMPLATES_DOCS_DIR", tmp_path / "нет")
    tid = make_tender(repo, documents=[DECLARATION])
    with pytest.raises(PackageError):
        planned_files(tid, [DECLARATION], tmp_path / "out", repo)


def test_planned_files_for_an_empty_package_is_empty(repo, dirs, tmp_path):
    tid = make_tender(repo)
    assert planned_files(tid, [], tmp_path / "out", repo) == []


# ---- круг ревью: временные файлы, experience_ids, мелочи ----------------------

def test_leftover_temp_files_of_a_killed_build_are_swept_at_the_start(
        repo, dirs, tmp_path):
    # питание отключили посреди пересборки: остаток не должен уйти заказчику
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    out.mkdir()
    ours = ["~pkg-deadbeef.docx", "~pkg-0123abcd.xlsx"]
    foreign = ["~$Декларация соответствия.docx",     # замок Word
               "~pkg-zzzzzzzz.docx", "~pkg-deadbeef.pdf", "~заметки.docx",
               "чужой.docx"]
    for name in ours + foreign:
        (out / name).write_bytes(b"x")
    build_package(tid, [DECLARATION], out, repo)
    assert names(out) == sorted(foreign + [f"{DECLARATION}.docx"])


def test_planned_files_does_not_sweep(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    out.mkdir()
    (out / "~pkg-deadbeef.docx").write_bytes(b"x")
    planned_files(tid, [DECLARATION], out, repo)
    assert names(out) == ["~pkg-deadbeef.docx"]


def test_temp_name_is_checked_for_being_taken(repo, dirs, tmp_path, monkeypatch):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "новый")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    out.mkdir()
    (out / f"{DECLARATION}.docx").write_bytes(b"old")
    stale = out / "~pkg-aaaaaaaa.docx"
    stale.write_bytes(b"stale")
    ids = iter(["a" * 32, "b" * 32])

    class Fake:
        def __init__(self, hex):
            self.hex = hex

    monkeypatch.setattr(package_module.uuid, "uuid4", lambda: Fake(next(ids)))
    monkeypatch.setattr(package_module, "_sweep_temp_files", lambda folder: None)
    result = build_package(tid, [DECLARATION], out, repo, overwrite=True)
    assert result.skipped == [] and len(result.written) == 1
    assert stale.read_bytes() == b"stale"            # чужой остаток не затронут
    assert text_of(out / f"{DECLARATION}.docx") == ["новый"]


def test_failure_to_remove_the_temp_file_does_not_hide_the_real_reason(
        repo, dirs, tmp_path, monkeypatch):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION])
    out = tmp_path / "out"
    out.mkdir()
    (out / f"{DECLARATION}.docx").write_bytes(b"old")

    def crash_after_writing(template, context, target):
        Path(target).write_bytes(b"half")
        raise package_module.TemplateError("Настоящая причина сбоя.")

    real_unlink = Path.unlink

    def stubborn(self, *args, **kwargs):
        if self.name.startswith("~pkg-"):
            raise PermissionError("held by antivirus")
        return real_unlink(self, *args, **kwargs)

    monkeypatch.setattr(package_module, "render_docx", crash_after_writing)
    monkeypatch.setattr(Path, "unlink", stubborn)
    result = build_package(tid, [DECLARATION], out, repo, overwrite=True)
    assert result.written == []
    assert "Настоящая причина сбоя" in result.skipped[0][1]
    assert "PermissionError" not in result.skipped[0][1]


def test_duplicate_message_does_not_claim_the_first_one_was_built(
        repo, dirs, tmp_path):
    # первый экземпляр не собрался (нет шаблона) -- «собран один раз» было бы ложью
    tid = make_tender(repo, documents=[MTR])
    result = build_package(tid, [MTR, MTR], tmp_path / "out", repo)
    assert result.written == [] and len(result.skipped) == 2
    second = result.skipped[1][1]
    assert "дважды" in second and "собран" not in second


def test_target_that_is_a_folder_gets_the_right_advice(repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "x")
    make_template(templates / f"{MTR}.docx", "x")
    tid = make_tender(repo, documents=[DECLARATION, MTR])
    out = tmp_path / "out"
    (out / f"{DECLARATION}.docx").mkdir(parents=True)
    assert planned_files(tid, [DECLARATION, MTR], out, repo) == [
        (MTR, out / f"{MTR}.docx")]
    for overwrite in (False, True):
        result = build_package(tid, [DECLARATION, MTR], out, repo,
                               overwrite=overwrite)
        reason = dict(result.skipped)[DECLARATION]
        assert "папк" in reason and "Word" not in reason
        assert (out / f"{DECLARATION}.docx").is_dir()
        (out / f"{MTR}.docx").unlink(missing_ok=True)


def test_second_document_is_written_when_replacing_the_first_is_refused(
        repo, dirs, tmp_path, monkeypatch):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "новый")
    make_template(templates / f"{MTR}.docx", "новый")
    tid = make_tender(repo, documents=[DECLARATION, MTR])
    out = tmp_path / "out"
    out.mkdir()
    for name in (DECLARATION, MTR):
        (out / f"{name}.docx").write_bytes(b"old")
    real_replace = package_module.os.replace

    def locked_first(src, dst):
        if Path(dst).name == f"{DECLARATION}.docx":
            raise PermissionError("open in Word")
        return real_replace(src, dst)

    monkeypatch.setattr(package_module.os, "replace", locked_first)
    result = build_package(tid, [DECLARATION, MTR], out, repo, overwrite=True)
    assert [p.name for p in result.written] == [f"{MTR}.docx"]
    assert result.skipped[0][0] == DECLARATION and "открыт" in result.skipped[0][1]
    assert (out / f"{DECLARATION}.docx").read_bytes() == b"old"
    assert text_of(out / f"{MTR}.docx") == ["новый"]
    assert names(out) == sorted([f"{DECLARATION}.docx", f"{MTR}.docx"])


EXPERIENCE_LINE = "{% for r in experience.rows %}{{ r[1] }};{% endfor %}"


def experience_tender(repo):
    first = repo.save_experience(ExperienceRecord(customer_name="Первый"))
    second = repo.save_experience(ExperienceRecord(customer_name="Второй"))
    third = repo.save_experience(ExperienceRecord(customer_name="Третий"))
    tid = make_tender(repo, documents=[EXPERIENCE])
    tender = repo.get_tender(tid)
    tender.experience_ids = [first]               # отбор, сохранённый при тендере
    repo.save_tender(tender)
    return tid, (first, second, third)


def test_experience_ids_choose_the_contracts_for_the_experience_reference(
        repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{EXPERIENCE}.docx", EXPERIENCE_LINE)
    tid, (first, second, third) = experience_tender(repo)
    out = tmp_path / "out"
    build_package(tid, [EXPERIENCE], out, repo, experience_ids=[third, second])
    # те, что выбрал пользователь, в его порядке; сохранённый отбор не применён
    text = text_of(out / f"{EXPERIENCE}.docx")[0]
    assert text.index("Третий") < text.index("Второй")
    assert "Первый" not in text


def test_experience_ids_none_takes_the_selection_stored_with_the_tender(
        repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{EXPERIENCE}.docx", EXPERIENCE_LINE)
    tid, _ = experience_tender(repo)
    out = tmp_path / "out"
    build_package(tid, [EXPERIENCE], out, repo)
    text = text_of(out / f"{EXPERIENCE}.docx")[0]
    assert "Первый" in text and "Второй" not in text


def test_empty_experience_ids_mean_no_experience_not_the_stored_selection(
        repo, dirs, tmp_path):
    templates, _ = dirs
    make_template(templates / f"{EXPERIENCE}.docx", "[" + EXPERIENCE_LINE + "]")
    tid, _ = experience_tender(repo)
    out = tmp_path / "out"
    build_package(tid, [EXPERIENCE], out, repo, experience_ids=[])
    assert text_of(out / f"{EXPERIENCE}.docx")[0] == "[]"
