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
from scripts.seed import SeedError, seed, seed_if_empty


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


# ---- первый запуск: справочники наполняются сами ---------------------------------

def test_seed_if_empty_fills_an_empty_base(repo):
    # фирма на новом компьютере не должна открыть «Справочники» и увидеть
    # пустоту: всё, что здесь записано, уже написано
    assert counts(repo) == {k: 0 for k in counts(repo)}
    assert seed_if_empty(repo) is True
    assert all(counts(repo).values())


def test_seed_if_empty_does_not_touch_a_base_that_already_has_data(repo):
    # Повторное наполнение вернуло бы запись, которую человек СТЁР осознанно
    # (лишний прибор, уволившийся сотрудник). Поэтому условие -- пустота базы,
    # а не идемпотентность seed.
    seed(repo)
    victim = repo.list_employees()[-1]
    repo.delete_employee(victim.id)
    before = counts(repo)
    assert seed_if_empty(repo) is False
    assert counts(repo) == before
    assert victim.full_name not in [e.full_name for e in repo.list_employees()]


def test_one_surviving_record_is_enough_to_count_as_not_empty(repo):
    # пустота проверяется по каждому справочнику: одна организация в базе --
    # уже не первый запуск
    from app.models import ROLE_CUSTOMER, Organization
    repo.save_organization(Organization(role=ROLE_CUSTOMER, inn="2904016130",
                                        short_name="МУП «Полигон»"))
    assert seed_if_empty(repo) is False
    assert len(repo.list_employees()) == 0


def test_the_program_fills_the_references_at_its_first_start(tmp_path, capsys):
    # то, что увидит человек: наполнение делает сама программа, а не докстринг
    from app import db as db_module
    from app.web import __main__ as entry
    path = tmp_path / "first.db"
    conn = db_module.connect(path)
    db_module.migrate(conn)
    conn.close()
    entry.fill_references(path)
    out = capsys.readouterr().out
    assert "Справочники заполнены" in out and "Traceback" not in out
    conn = db_module.connect(path)
    try:
        assert all(counts(Repo(conn)).values())
    finally:
        conn.close()
    entry.fill_references(path)                 # второй запуск молчит
    assert "Справочники заполнены" not in capsys.readouterr().out


# ---- живой случай: форма заказчика печатает пустые реквизиты ----------------------

# Заявку по форме ПМУП «Автоспецтранс» собрали на живой работе, и в анкете
# участника четыре строки банковских реквизитов и основание полномочий
# оказались пустыми: в справочнике их нет намеренно (см. выше). Программа
# отдала документ и ничего не сказала -- узнать об этом человек мог только
# от закупочной комиссии.

ANKETA = "Анкета участника"
AUTOSPETSTRANS_FORM = "Заявка на участие Автоспецтранс.docx"


def autospetstrans_tender(repo, **requisites) -> int:
    from app.models import ROLE_CUSTOMER, CustomerProfile, DocSpec, Organization
    seed(repo)
    me = repo.find_organization_by_inn("5904188041")
    if requisites:
        from dataclasses import replace as dc_replace
        repo.save_organization(dc_replace(me, **requisites))
    customer = repo.save_organization(Organization(
        role=ROLE_CUSTOMER, short_name="ПМУП «Автоспецтранс»", inn="5903004541"))
    repo.save_customer_profile(CustomerProfile(
        customer_id=customer, documents=[ANKETA],
        docs={ANKETA: DocSpec(template=AUTOSPETSTRANS_FORM, prefix="1.")}))
    tid = repo.save_tender(Tender(
        customer_id=customer, participant_id=me.id,
        purchase_number="32514010203", subject="Разработка КЭР",
        deadline="90 дней с даты заключения договора", doc_date="2025-11-10"))
    repo.save_lot_item(LotItem(tender_id=tid, position=1, name="Разработка КЭР",
                               price=Decimal("1000000")))
    return tid


BANK = {"bank_name": "Филиал «Нижегородский» АО «АЛЬФА-БАНК»",
        "bank_bik": "042202824", "bank_account": "40702810122250001234",
        "bank_corr_account": "30101810200000000824", "head_basis": "Устав"}


def test_the_form_of_the_new_customer_names_its_empty_requisites(repo, tmp_path):
    from app.render.package import build_package
    tid = autospetstrans_tender(repo)
    result = build_package(tid, [ANKETA], tmp_path / "out", repo)
    assert [p.name for p in result.written] == [f"1.{ANKETA}.docx"]
    (document, text), = result.warnings
    assert document == ANKETA
    for name in ("Банк", "БИК", "Расчётный счёт", "Корреспондентский счёт",
                 "Основание полномочий"):
        assert name in text, name
    assert "справочник «Организации»" in text      # где их заполняют
    assert "org.bank_name" not in text             # имя поля, а не путь


def test_the_same_form_with_the_requisites_filled_says_nothing(repo, tmp_path):
    # обратная сторона: заполнили -- и предупреждения нет ни одного
    from app.render.package import build_package
    tid = autospetstrans_tender(repo, **BANK)
    result = build_package(tid, [ANKETA], tmp_path / "out", repo)
    assert [p.name for p in result.written] == [f"1.{ANKETA}.docx"]
    assert result.warnings == []


def test_the_form_prints_the_requisites_it_was_given(repo, tmp_path):
    """Третья сторона: предупреждение говорит о НАПЕЧАТАННОМ. Реквизиты,
    которые документ печатает, читаются из готового файла."""
    import docx
    from app.render.package import build_package
    tid = autospetstrans_tender(repo, **BANK)
    result = build_package(tid, [ANKETA], tmp_path / "out", repo)
    document = docx.Document(result.written[0])
    text = "\n".join(
        [p.text for p in document.paragraphs]
        + [c.text for t in document.tables for row in t.rows for c in row.cells])
    assert BANK["bank_account"] in text and BANK["bank_bik"] in text
    assert "Устав" in text


# ---- сторож рабочей базы: он один на весь набор и включается сам ------------------

def test_the_guard_refuses_the_working_database_of_the_firm(tmp_path):
    """Рабочая база фирмы для набора тестов закрыта (tests/conftest.py).
    Сторож, которого можно забыть, ничего не стережёт: он автоматический,
    смотрит на путь и падает по-русски."""
    from app import config
    from app import db as db_module
    existed = Path(config.DB_PATH).exists()
    with pytest.raises(pytest.fail.Exception) as caught:
        db_module.connect(config.DB_PATH)
    text = str(caught.value)
    assert "рабочую базу" in text and str(config.DB_PATH) in text
    assert "tmp_path" in text                      # что делать вместо этого
    # база не создана и папка data не появилась: сторож не пропустил вызов
    assert Path(config.DB_PATH).exists() == existed
    # а временные базы открываются как раньше -- иначе сторожил бы всё подряд
    conn = db_module.connect(tmp_path / "t.db")
    conn.close()


def test_the_guard_looks_at_the_path_not_at_the_name(tmp_path, monkeypatch):
    # Модуль может взять DB_PATH копией имени, и подмена config.DB_PATH до него
    # не дойдёт: настоящий путь сторож запомнил при загрузке.
    from app import config
    from app import db as db_module
    monkeypatch.setattr(config, "DB_PATH", tmp_path / "подменённая.db")
    with pytest.raises(pytest.fail.Exception):
        db_module.connect(Path(__file__).resolve().parent.parent
                          / "data" / "tender.db")


def test_main_works_on_the_database_it_is_given(tmp_path, monkeypatch, capsys):
    """`main(db_path=...)` -- вторая дверь к рабочей базе, закрытая у входа:
    тест передаёт временный путь и ничего не подменяет. Проверяется то, что
    видит человек: база создана, справочники наполнены."""
    from app.web import __main__ as entry
    path = tmp_path / "своя.db"

    def busy(app, port=entry.PORT):
        raise entry.BindError("занято")

    monkeypatch.setattr(entry, "build_server", busy)
    assert entry.main(db_path=path) == 1            # до сервера дело дошло
    out = capsys.readouterr().out
    assert "Справочники заполнены" in out and "Traceback" not in out
    assert path.is_file()
    conn = db.connect(path)
    try:
        assert all(counts(Repo(conn)).values())
    finally:
        conn.close()
