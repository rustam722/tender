from dataclasses import replace
from decimal import Decimal
from zipfile import ZipFile

from docx import Document

from app import db
from app.models import Asset, Employee, ExperienceRecord, LotItem, Organization, Tender
from app.render import source_forms as forms
from app.render.package import build_package
from app.repo import Repo
from scripts.install_customer_templates import install


def test_source_customer_packages_fill_forms_without_identifying_first_part(tmp_path):
    connection = db.connect(tmp_path / "test.db")
    try:
        db.migrate(connection)
        repo = Repo(connection)
        install(repo)
        install(repo)
        participant = repo.save_organization(Organization(
            short_name="Организация УНИКАЛЬНАЯ", inn="5900000000",
            head_name="Подписант УНИКАЛЬНЫЙ", head_position="Руководитель",
            legal_address="Адрес УНИКАЛЬНЫЙ"))
        employee = repo.save_employee(Employee(full_name="Работник по лоту", position="Инженер"))
        asset = repo.save_asset(Asset(name="Прибор по лоту", quantity=2))
        experience = repo.save_experience(ExperienceRecord(
            customer_name="Прошлый заказчик", subject="Работы по договору", amount=Decimal("25.50")))
        for source, package in forms.CUSTOMERS:
            customer = next(org for org in repo.list_organizations("customer")
                            if org.short_name == source.short_name)
            conditions = {key: "Заданное значение" for form in package for key in form.conditions}
            conditions.update({form.confirmation: "Подтверждаю" for form in package if form.confirmation})
            conditions[forms.ROLE] = "Участник закупки"
            record = Tender(customer_id=customer.id, participant_id=participant,
                            purchase_number="TEST-1", subject="Описание работ",
                            employee_ids=[employee], asset_ids=[asset],
                            experience_ids=[experience], deal_answers=conditions)
            tender_id = repo.save_tender(record)
            if package is forms.IP_FORMS:
                incomplete = build_package(tender_id, [forms.BID], tmp_path / "без цены", repo)
                assert not incomplete.written and "укажите позиции лота" in incomplete.skipped[0][1]
            repo.save_lot_item(LotItem(tender_id=tender_id, position=1, name="Работы по лоту",
                                      price=Decimal("123456.78")))
            out = tmp_path / str(customer.id)
            result = build_package(tender_id, [form.document for form in package], out, repo)
            assert not result.skipped and len(result.written) == len(package), result.skipped
            for form, file in zip(package, result.written, strict=True):
                doc = Document(file)
                text = "\n".join([p.text for p in doc.paragraphs] +
                                 [cell.text for table in doc.tables for r in table.rows for cell in r.cells])
                assert text.startswith("ЧЕРНОВИК.") and "{{" not in text and "{%" not in text
                assert all(conditions[key] in text for key in form.conditions)
                if package is forms.ZVKS_FORMS and form.document == forms.OFFER:
                    assert "УНИКАЛЬН" not in text and "5900000000" not in text and "123" not in text
                    with ZipFile(file) as z:
                        assert not any("УНИКАЛЬН".encode() in z.read(n) for n in z.namelist() if n.endswith(".xml"))
                if package is forms.IP_FORMS and form.document in (forms.BID, forms.COMMERCIAL):
                    assert "123 456,78" in text
                if form.document == forms.STAFF:
                    assert "Работник по лоту" in text
                if form.document == forms.ASSETS:
                    assert "Прибор по лоту" in text
                if form.document == forms.EXPERIENCE:
                    assert "Работы по договору" in text
            if package is forms.IP_FORMS:
                conditions[forms.APPROVE_OFFER] = "Не подтверждаю"
                repo.save_tender(replace(record, id=tender_id, deal_answers=conditions))
                result = build_package(tender_id, [forms.BID], tmp_path / "без согласия", repo)
                assert not result.written and "содержит заявление" in result.skipped[0][1]
    finally:
        connection.close()
