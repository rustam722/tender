"""Предпросмотр документа.

Шаблоны здесь настоящие .docx, построенные python-docx в папке теста; книга
заказчика -- настоящая форма Роснефти из tests/fixtures. Главное, что
проверяется: предпросмотр показывает то, что НАПЕЧАТАЕТСЯ (гонит документ через
ту же отрисовку, что сборка), говорит те же слова, что сборка, и не оставляет
после себя ни временного файла, ни файла в папке тендера.
"""
import shutil
import tempfile
from decimal import Decimal
from pathlib import Path

import docx
import pytest

from app import config, db
from app.models import (
    ROLE_CUSTOMER, ROLE_PARTICIPANT, CustomerProfile, DocSpec, LotItem,
    Organization, Tender,
)
from app.render import preview as preview_module
from app.render.context import blank_fields_warning, build_context
from app.render.package import build_package
from app.render.preview import Preview, preview_document
from app.repo import Repo

ROOT = Path(__file__).resolve().parent.parent
REAL_FORM = Path(__file__).parent / "fixtures" / "forma_tkp_rosneft_murmansk.xlsx"

DECLARATION = "Декларация соответствия"
KP = "Коммерческое предложение"
TKP = "ТКП по форме заказчика"
CONDITIONS = [
    "Цена и условия:", "Форма, срок и порядок оплаты:",
    "Условия и место оказания услуг:", "Срок оказания услуг:",
    "Срок действия коммерческого предложения:", "Срок действия договора:",
    "Возможность привлечения субподрядчиков/субисполнителей",
]


# ---- помощники ----------------------------------------------------------------

def make_template(path: Path, *lines: str) -> Path:
    doc = docx.Document()
    for line in lines:
        doc.add_paragraph(line)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    return path


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
    monkeypatch.setattr(config, "TEMPLATES_DOCS_DIR", templates)
    monkeypatch.setattr(config, "CELLMAPS_DIR", cellmaps)
    return templates, cellmaps


@pytest.fixture
def temp_home(tmp_path, monkeypatch):
    """Папка временных файлов -- в папке теста: по ней видно, что предпросмотр
    за собой убрал. Пустая папка после вызова -- и есть утверждение."""
    home = tmp_path / "temp"
    home.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(home))
    return home


def make_tender(repo, *, docs=None, documents=(), declaration="Декларация заказчика",
                with_profile=True, **tender_kw) -> int:
    me = repo.save_organization(Organization(
        role=ROLE_PARTICIPANT, inn="5904188041", short_name="ООО «Техносфера»",
        head_position="Директор", head_name="Рабичко Елена Михайловна"))
    customer = repo.save_organization(Organization(
        role=ROLE_CUSTOMER, inn="2904016130", short_name="МУП «Полигон»"))
    if with_profile:
        repo.save_customer_profile(CustomerProfile(
            customer_id=customer, documents=list(documents), docs=docs or {},
            declaration_text=declaration))
    return repo.save_tender(Tender(
        customer_id=customer, participant_id=me, subject="Разработка КЭР",
        purchase_number="0556300000225000001", **tender_kw))


def paragraphs(result: Preview) -> list[str]:
    return [block["text"] for block in result.blocks if block["kind"] == "p"]


def tables(result: Preview) -> list[list[list[str]]]:
    return [block["rows"] for block in result.blocks if block["kind"] == "table"]


def rosneft_tender(repo, positions=2, answers=None):
    tid = make_tender(
        repo, documents=[TKP],
        docs={TKP: DocSpec(template="forma_tkp_rosneft_murmansk.xlsx",
                           cellmap="tkp_rosneft_murmansk.json", prefix="3.")},
        deal_answers={c: "подтверждаю" for c in CONDITIONS} if answers is None else answers,
        vat_rate=Decimal("0.05"))
    for n in range(1, positions + 1):
        repo.save_lot_item(LotItem(tender_id=tid, position=n, name=f"Позиция {n}",
                                   price=Decimal("100000.10")))
    return tid


# ---- что показывает предпросмотр ----------------------------------------------

def test_a_preview_shows_the_paragraphs_and_tables_the_document_prints(
        repo, dirs, temp_home):
    templates, _ = dirs
    doc = docx.Document()
    doc.add_paragraph("Предмет закупки: {{ tender.subject }}")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Должность"
    table.cell(0, 1).text = "ФИО"
    table.cell(1, 0).text = "{{ org.head_position }}"
    table.cell(1, 1).text = "{{ org.head_name }}"
    doc.add_paragraph("Директор: {{ org.head_name }}")
    (templates / f"{DECLARATION}.docx").parent.mkdir(parents=True, exist_ok=True)
    doc.save(templates / f"{DECLARATION}.docx")
    tid = make_tender(repo, documents=[DECLARATION])

    result = preview_document(tid, DECLARATION, repo)

    assert result.kind == "word"
    assert result.refusal is None
    # значения тендера подставлены, а не остались метками шаблона
    assert "Предмет закупки: Разработка КЭР" in paragraphs(result)
    assert "Директор: Рабичко Елена Михайловна" in paragraphs(result)
    assert tables(result) == [[["Должность", "ФИО"],
                               ["Директор", "Рабичко Елена Михайловна"]]]
    # порядок -- как в документе: абзац, таблица, абзац
    assert [block["kind"] for block in result.blocks] == ["p", "table", "p"]


def test_a_preview_is_rendered_by_the_real_renderer_not_by_its_own_substitution(
        repo, dirs, temp_home):
    """Таблица, которую строит цикл docxtpl, и сумма, которую печатает money():
    своя подстановка «похожим способом» не сделала бы ни того, ни другого, и
    предпросмотр показал бы не то, что напечатается."""
    templates, _ = dirs
    doc = docx.Document()
    doc.add_paragraph("Итого без НДС: {{ money(estimate.total_wo_vat) }}")
    table = doc.add_table(rows=3, cols=1)
    table.cell(0, 0).text = "{%tr for row in items %}"
    table.cell(1, 0).text = "{{ row.name }}"
    table.cell(2, 0).text = "{%tr endfor %}"
    doc.save(templates / f"{KP}.docx")
    tid = make_tender(repo, documents=[KP])
    for n in (1, 2, 3):
        repo.save_lot_item(LotItem(tender_id=tid, position=n, name=f"Объект {n}",
                                   price=Decimal("1000")))

    result = preview_document(tid, KP, repo)

    assert result.kind == "word", result.refusal
    # цикл шаблона развернулся в три строки -- это сделал docxtpl, не предпросмотр
    assert tables(result) == [[["Объект 1"], ["Объект 2"], ["Объект 3"]]]
    # money() напечатал сумму с неразрывными пробелами, как в документе
    context = build_context(tid, repo, document=KP)
    assert f"Итого без НДС: {context['money'](context['estimate']['total_wo_vat'])}" \
        in paragraphs(result)


def test_a_preview_names_the_blank_mandatory_fields_like_the_package_does(
        repo, dirs, temp_home):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx",
                  "Банк: {{ org.bank_name }}", "Счёт: {{ org.bank_account }}")
    tid = make_tender(repo, documents=[DECLARATION])

    result = preview_document(tid, DECLARATION, repo)

    assert result.kind == "word", result.refusal
    expected = blank_fields_warning(DECLARATION, ["org.bank_name", "org.bank_account"])
    assert expected in result.warnings
    # имена граф -- русские, из словарей app/models.py, а не пути полей
    assert "Банк, Расчётный счёт" in expected and "org.bank_name" not in expected


def test_the_warnings_of_the_context_reach_the_preview(repo, dirs, temp_home):
    """Предупреждение контекста -- главная польза предпросмотра: человек видит
    беду до сборки. Здесь это пропавшая запись опыта."""
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "Декларация")
    tid = make_tender(repo, documents=[DECLARATION])

    result = preview_document(tid, DECLARATION, repo, experience_ids=[999])

    assert result.kind == "word", result.refusal
    assert any("id=999" in text for text in result.warnings), result.warnings


def test_a_document_that_cannot_be_built_is_a_refusal_not_an_empty_preview(
        repo, dirs, temp_home):
    make_tender(repo, documents=[TKP])
    tid = make_tender(repo, documents=[TKP])

    result = preview_document(tid, TKP, repo)

    assert result.kind == "refusal"
    assert result.blocks == []
    assert "собирается по форме заказчика" in result.refusal
    assert "форма для него не назначена" in result.refusal


def test_a_refusal_of_a_preview_says_the_same_words_as_a_refusal_of_the_package(
        repo, dirs, temp_home, tmp_path):
    """Одни и те же слова на один и тот же отказ: два текста про одну беду
    расходятся на первой же правке."""
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx",
                  "{% for line in declaration_text.splitlines() %}{{ line }}{% endfor %}")
    tid = make_tender(repo, documents=[DECLARATION], declaration="")
    out = tmp_path / "out"

    result = preview_document(tid, DECLARATION, repo)
    built = build_package(tid, [DECLARATION], out, repo)

    assert result.kind == "refusal"
    assert [(DECLARATION, result.refusal)] == built.skipped


def test_an_unknown_tender_is_a_lookup_error_and_not_a_refusal(repo, dirs, temp_home):
    with pytest.raises(LookupError):
        preview_document(10 ** 6, DECLARATION, repo)


# ---- после предпросмотра ничего не остаётся -----------------------------------

def test_the_temporary_file_of_a_preview_is_always_removed(
        repo, dirs, temp_home, monkeypatch):
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "Предмет: {{ tender.subject }}")
    tid = make_tender(repo, documents=[DECLARATION])

    assert preview_document(tid, DECLARATION, repo).kind == "word"
    assert list(temp_home.iterdir()) == [], "после успеха"

    broken = make_tender(repo, documents=[DECLARATION], declaration="")
    make_template(templates / f"{DECLARATION}.docx",
                  "{% for line in declaration_text.splitlines() %}{{ line }}{% endfor %}")
    assert preview_document(broken, DECLARATION, repo).kind == "refusal"
    assert list(temp_home.iterdir()) == [], "после отказа"

    make_template(templates / f"{DECLARATION}.docx", "Предмет: {{ tender.subject }}")

    def boom(template, context, out, blanks=None):
        Path(out).write_text("недописанный файл", encoding="utf-8")
        raise RuntimeError("отрисовка упала посреди работы")

    monkeypatch.setattr(preview_module, "render_docx", boom)
    with pytest.raises(RuntimeError):
        preview_document(tid, DECLARATION, repo)
    assert list(temp_home.iterdir()) == [], "после исключения в отрисовке"


def test_a_preview_never_writes_into_the_output_folder_of_the_tender(
        repo, dirs, temp_home, tmp_path, monkeypatch):
    """Папка вывода уходит заказчику, и временный файл предпросмотра в ней не
    имеет права появиться даже на миг: оборванный запрос оставил бы его там
    навсегда. Поэтому папка смотрится и ВО ВРЕМЯ отрисовки, а не только после:
    уборка за собой -- не оправдание для записи не туда."""
    templates, _ = dirs
    make_template(templates / f"{DECLARATION}.docx", "Предмет: {{ tender.subject }}")
    out = tmp_path / "выход тендера"
    out.mkdir()
    tid = make_tender(repo, documents=[DECLARATION], output_dir=str(out))
    real = preview_module.render_docx
    seen = []

    def watching(template, context, target, blanks=None):
        seen.append((Path(target), sorted(p.name for p in out.iterdir())))
        return real(template, context, target, blanks)

    monkeypatch.setattr(preview_module, "render_docx", watching)

    assert preview_document(tid, DECLARATION, repo).kind == "word"

    assert len(seen) == 1, "документ отрисован один раз"
    target, during = seen[0]
    assert during == [], f"в папке вывода во время отрисовки появилось: {during}"
    assert out not in target.parents, f"документ отрисован в папку вывода: {target}"
    assert list(out.iterdir()) == []


# ---- книга заказчика ----------------------------------------------------------

def test_an_excel_document_is_not_rendered_but_names_the_cells_it_would_fill(
        repo, dirs, temp_home):
    templates, cellmaps = dirs
    shutil.copy(REAL_FORM, templates / "forma_tkp_rosneft_murmansk.xlsx")
    shutil.copy(ROOT / "cellmaps" / "tkp_rosneft_murmansk.json", cellmaps)
    tid = rosneft_tender(repo)

    result = preview_document(tid, TKP, repo)

    assert result.kind == "excel", result.refusal
    assert result.refusal is None
    # сказано, что это книга заказчика, и названа она сама
    assert any("forma_tkp_rosneft_murmansk.xlsx" in text for text in paragraphs(result))
    rows = tables(result)[0]
    filled = {row[0]: row[1] for row in rows}
    # клетки карты названы вместе с тем, что в них попадёт
    assert filled["D8"] == "Ставка НДС: 5%"
    assert filled["E23"] == "подтверждаю"
    assert filled["B31"] == "Директор Рабичко Елена Михайловна"
    assert filled["C9"] == "100000.10"
    assert filled["C10"] == "100000.10"
    assert "C11" not in filled, "пустые строки блока книга оставляет как в форме"
    # книга не открывалась на запись: временного файла нет и ничего не создано
    assert list(temp_home.iterdir()) == []


def test_an_excel_document_whose_form_cannot_be_filled_is_a_refusal(
        repo, dirs, temp_home):
    templates, cellmaps = dirs
    shutil.copy(REAL_FORM, templates / "forma_tkp_rosneft_murmansk.xlsx")
    shutil.copy(ROOT / "cellmaps" / "tkp_rosneft_murmansk.json", cellmaps)
    tid = rosneft_tender(repo, positions=20)

    result = preview_document(tid, TKP, repo)

    assert result.kind == "refusal"
    assert result.blocks == []
    assert "больше, чем строк в форме заказчика" in result.refusal


def test_an_excel_document_without_a_cellmap_in_the_profile_is_a_refusal(
        repo, dirs, temp_home):
    templates, _ = dirs
    shutil.copy(REAL_FORM, templates / "forma_tkp_rosneft_murmansk.xlsx")
    tid = make_tender(
        repo, documents=[TKP],
        docs={TKP: DocSpec(template="forma_tkp_rosneft_murmansk.xlsx",
                           cellmap="нет_такой_карты.json")})

    result = preview_document(tid, TKP, repo)

    assert result.kind == "refusal"
    assert "нет_такой_карты.json" in result.refusal
