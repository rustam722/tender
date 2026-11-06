"""Наполнение справочников (задача 14): данные берутся из поданных заявок."""
from decimal import Decimal

import pytest

from app import db
from app.models import ROLE_CUSTOMER, ROLE_PARTICIPANT
from app.render.context import DOCUMENTS, TABLE_DOCUMENTS
from app.repo import Repo
from scripts import seed as seed_module
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
            assert (seed_module.TEMPLATES_DIR / template).is_file(), template
            if spec.cellmap:
                assert (seed_module.CELLMAPS_DIR / spec.cellmap).is_file()
            if spec.layout:
                assert (seed_module.LAYOUTS_DIR / spec.layout).is_file()


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
    monkeypatch.setattr(seed_module, "LAYOUTS_DIR", tmp_path / "нет")
    with pytest.raises(SeedError, match="раскладки"):
        seed(repo)
    assert counts(repo)["profiles"] == 0


def test_missing_template_fails_in_russian(repo, tmp_path, monkeypatch):
    monkeypatch.setattr(seed_module, "TEMPLATES_DIR", tmp_path / "нет")
    with pytest.raises(SeedError, match="шаблон"):
        seed(repo)


def test_missing_cellmap_fails_in_russian(repo, tmp_path, monkeypatch):
    monkeypatch.setattr(seed_module, "CELLMAPS_DIR", tmp_path / "нет")
    with pytest.raises(SeedError, match="карт"):
        seed(repo)
