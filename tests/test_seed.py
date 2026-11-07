"""Наполнение справочников (задача 14): данные берутся из поданных заявок."""
import hashlib
import zipfile
from decimal import Decimal
from pathlib import Path

import pytest

from app import config, db
from app.models import ROLE_PARTICIPANT, LotItem, Tender
from app.render.context import DOCUMENTS, TABLE_DOCUMENTS, build_context
from app.render.docx_render import render_docx
from app.repo import Repo
from scripts.seed import SeedError, seed


@pytest.fixture
def repo(tmp_path):
    conn = db.connect(tmp_path / "t.db")
    db.migrate(conn)
    yield Repo(conn)
    conn.close()


def counts(repo) -> dict:
    return {
        "organizations": len(repo.list_organizations()),
        "employees": len(repo.list_employees()),
        "assets": len(repo.list_assets()),
        "services": len(repo.list_services()),
        "experience": len(repo.list_experience()),
        "profiles": len(repo.list_customer_profiles()),
    }


def test_seed_creates_five_employees(repo):
    seed(repo)
    names = [e.full_name for e in repo.list_employees()]
    assert len(names) == 5
    assert [n.split()[0] for n in names] == [
        "Рабичко", "Мельникова", "Ивонина", "Четина", "Шилоносова"]


def test_seed_creates_twenty_plus_assets(repo):
    seed(repo)
    assets = repo.list_assets()
    assert len(assets) >= 20
    office = [a for a in assets if a.kind == "premises"]
    assert len(office) == 1
    assert "59-БД 081725" in office[0].ownership_doc


def test_seed_is_idempotent(repo):
    seed(repo)
    first = counts(repo)
    seed(repo)
    assert counts(repo) == first
    assert all(first.values())


def test_technosfera_requisites_match_registry(repo):
    seed(repo)
    org = repo.find_organization_by_inn("5904188041")
    assert (org.kpp, org.ogrn) == ("590401001", "1085904010224")
    assert org.head_name == "Рабичко Елена Михайловна"
    assert org.role == ROLE_PARTICIPANT
    assert org.msp_status == "micro"


def test_every_seeded_profile_has_existing_template_files(repo):
    seed(repo)
    profiles = repo.list_customer_profiles()
    assert len(profiles) == 3
    for profile in profiles:
        for name, spec in profile.docs.items():
            assert name in DOCUMENTS, f"{name!r} -- не каноническое имя"
            template = spec.template or f"{name}.docx"
            assert (config.TEMPLATES_DOCS_DIR / template).is_file(), template
            if spec.cellmap:
                assert (config.CELLMAPS_DIR / spec.cellmap).is_file()
            if spec.layout:
                assert (config.LAYOUTS_DIR / spec.layout).is_file()


def test_profile_documents_are_canonical_and_described(repo):
    seed(repo)
    for profile in repo.list_customer_profiles():
        assert profile.documents
        assert set(profile.documents) <= set(DOCUMENTS)
        assert set(profile.docs) <= set(profile.documents)
        # документ с таблицей без раскладки в профиле давал бы предупреждение
        for name in profile.documents:
            if name in TABLE_DOCUMENTS:
                assert profile.docs[name].layout, name


def test_service_catalogue_has_all_codes(repo):
    seed(repo)
    codes = {s.code for s in repo.list_services()}
    assert codes == {"КЭР", "СЗЗ", "НДВ", "НДС", "НООЛР", "ППЭЭ", "ПЭК",
                     "инвентаризация ИЗАВ"}
    ker = repo.find_service_by_code("КЭР")
    assert ker.stage_wording
    assert ker.norms["Эколог-проектировщик"] > 0


def test_hourly_rates_only_where_the_source_shows_them(repo):
    seed(repo)
    rates = {e.position: e.hourly_rate for e in repo.list_employees()}
    assert rates["Эколог-проектировщик"] == Decimal("270")
    assert rates["Начальник отдела проектирования"] == Decimal("420")
    assert rates["Директор"] == Decimal("0")
    assert rates["Заместитель директора по развитию"] == Decimal("0")
    assert rates["Начальник отдела нормирования"] == Decimal("0")


def test_bank_details_are_left_empty(repo):
    seed(repo)
    org = repo.find_organization_by_inn("5904188041")
    assert (org.bank_name, org.bank_bik, org.bank_account,
            org.bank_corr_account) == ("", "", "", "")


def test_seed_does_not_overwrite_user_edits(repo):
    seed(repo)
    org = repo.find_organization_by_inn("5904188041")
    org.phones = "правка пользователя"
    repo.save_organization(org)
    seed(repo)
    assert repo.find_organization_by_inn("5904188041").phones == "правка пользователя"


def test_missing_layout_file_fails_in_russian(repo, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "LAYOUTS_DIR", tmp_path / "нет")
    with pytest.raises(SeedError, match="раскладки"):
        seed(repo)
    assert counts(repo)["profiles"] == 0


def test_missing_template_fails_in_russian(repo, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "TEMPLATES_DOCS_DIR", tmp_path / "нет")
    with pytest.raises(SeedError, match="шаблон"):
        seed(repo)


def test_missing_cellmap_fails_in_russian(repo, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "CELLMAPS_DIR", tmp_path / "нет")
    with pytest.raises(SeedError, match="Не найдена карта ячеек") as caught:
        seed(repo)
    assert "шаблон" not in str(caught.value)        # ловим карту, а не шаблон
    assert counts(repo)["profiles"] == 0


# ---- сторож: в шаблонах нет подписи и печати -----------------------------------

# Фирменный бланк «Техносферы» (логотип) в двух вариантах: .jpg и .jpeg. Белый
# список -- хеши, а не имена: «image2.png» ничего не говорит, хеш отличает
# одобренный логотип от скана подписи с печатью.
APPROVED_IMAGES = {
    "3f50c493245a6839e5e8f77dc8e2f76b2974782eb7028f5478bb9d9cf35d1a40":
        "логотип (image1.jpg)",
    "dd65e0bf7ff2277b8225252dae28322c8d807a4e9bbdf2875e42837a151ff8eb":
        "логотип (image1.jpeg)",
}


# Шаблоном может быть и книга с макросами, и заготовка Word (.dotx): все они --
# архивы с картинками внутри. Старые двоичные .doc/.xls сторож прочесть не может,
# поэтому такой шаблон тоже отвергается.
TEMPLATE_SUFFIXES = {".docx", ".docm", ".dotx", ".dotm", ".xlsx", ".xlsm",
                     ".xltx", ".xltm"}
LEGACY_SUFFIXES = {".doc", ".dot", ".xls", ".xlt"}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".jpe", ".gif", ".bmp", ".tif", ".tiff",
                  ".emf", ".wmf", ".wdp", ".jxr", ".svg", ".webp", ".ico"}


def _is_picture(part: str) -> bool:
    """Картинка по месту (media, embeddings) или по расширению где угодно:
    миниатюра docProps/thumbnail.jpeg несёт первую страницу подписанного
    документа и в media не лежит."""
    suffix = "." + part.rsplit(".", 1)[-1].lower() if "." in part else ""
    return ("/media/" in part or "/embeddings/" in part
            or suffix in IMAGE_SUFFIXES)


def _images_outside_whitelist(templates_dir) -> list[str]:
    problems = []
    for path in sorted(Path(templates_dir).glob("*")):
        suffix = path.suffix.lower()
        if path.name.startswith("~$"):
            continue
        if suffix in LEGACY_SUFFIXES:
            problems.append(
                f"Шаблон «{path.name}» в старом двоичном формате: картинки в нём "
                "проверить нельзя. Сохраните его как .docx или .xlsx.")
            continue
        if suffix not in TEMPLATE_SUFFIXES:
            continue
        with zipfile.ZipFile(path) as archive:
            for part in archive.namelist():
                if not _is_picture(part):
                    continue
                digest = hashlib.sha256(archive.read(part)).hexdigest()
                if digest not in APPROVED_IMAGES:
                    problems.append(
                        f"В шаблоне «{path.name}» картинка {part}, которой нет в "
                        "белом списке. Если это подпись или печать -- уберите её: "
                        "программа отдаёт документы с пустыми подписными блоками. "
                        "Если это новая часть бланка -- добавьте её хеш в белый "
                        f"список осознанно (sha256 {digest}).")
    return problems


def test_templates_carry_no_signature_or_stamp_images():
    # пустая или не та папка дала бы зелёный тест, который ничего не проверил
    found = {p.suffix.lower() for p in Path(config.TEMPLATES_DOCS_DIR).glob("*")}
    assert {".docx", ".xlsx"} <= found, (
        f"в папке шаблонов {config.TEMPLATES_DOCS_DIR} не нашлось ни .docx, ни .xlsx")
    assert not _images_outside_whitelist(config.TEMPLATES_DOCS_DIR), \
        "\n".join(_images_outside_whitelist(config.TEMPLATES_DOCS_DIR))


def test_signature_guard_catches_a_scan_and_names_the_part(tmp_path):
    # сторож, который никогда не срабатывает, ничего не стережёт
    template = tmp_path / "Справка с печатью.docx"
    with zipfile.ZipFile(template, "w") as archive:
        archive.writestr("word/document.xml", "<w:document/>")
        archive.writestr("word/media/image2.png", "скан подписи и печати".encode())
    problems = _images_outside_whitelist(tmp_path)
    assert len(problems) == 1
    assert "«Справка с печатью.docx»" in problems[0]
    assert "word/media/image2.png" in problems[0]
    assert "подпись или печать" in problems[0]


def test_signature_guard_also_reads_workbooks(tmp_path):
    book = tmp_path / "Форма.xlsx"
    with zipfile.ZipFile(book, "w") as archive:
        archive.writestr("xl/media/image1.png", "печать".encode())
    assert "xl/media/image1.png" in _images_outside_whitelist(tmp_path)[0]


def test_every_seeded_word_template_renders_with_a_real_context(repo, tmp_path):
    """Шаблоны после правки (убрана подпись, итоги КП) по-прежнему собираются."""
    seed(repo)
    me = repo.find_organization_by_inn("5904188041")
    rendered = []
    for profile in repo.list_customer_profiles():
        tid = repo.save_tender(Tender(customer_id=profile.customer_id,
                                      participant_id=me.id))
        repo.save_lot_item(LotItem(tender_id=tid, position=1, name="Услуга",
                                   price=Decimal("1000.00")))
        for name in profile.documents:
            template = profile.docs[name].template or f"{name}.docx"
            if not template.endswith(".docx"):
                continue
            context = build_context(tid, repo, document=name)
            out = render_docx(config.TEMPLATES_DOCS_DIR / template, context,
                              tmp_path / f"{len(rendered)}.docx")
            assert out.is_file()
            rendered.append(template)
    assert "Письмо о крупности сделки.docx" in rendered
    assert "Справка о кадровых ресурсах.docx" in rendered
    assert "Коммерческое предложение.docx" in rendered


@pytest.mark.parametrize("name, part", [
    ("Заготовка.dotx", "word/media/image2.png"),
    ("С макросом.docm", "word/media/image2.png"),
    ("Книга.xlsm", "xl/media/image1.png"),
    ("Миниатюра.docx", "docProps/thumbnail.jpeg"),         # не в media
    ("Рисунок.docx", "word/embeddings/oleObject1.bin"),
    ("Запасной.docx", "word/hdphoto1.wdp"),
])
def test_signature_guard_covers_other_formats_and_places(tmp_path, name, part):
    with zipfile.ZipFile(tmp_path / name, "w") as archive:
        archive.writestr(part, "скан".encode())
    problems = _images_outside_whitelist(tmp_path)
    assert len(problems) == 1 and part in problems[0] and name in problems[0]


def test_signature_guard_rejects_legacy_binary_formats(tmp_path):
    (tmp_path / "Старый.doc").write_bytes(bytes.fromhex("d0cf11e0"))
    problems = _images_outside_whitelist(tmp_path)
    assert len(problems) == 1 and "старом двоичном формате" in problems[0]
