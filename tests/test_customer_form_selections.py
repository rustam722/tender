from dataclasses import replace

from docx import Document

from app import db
from app.models import Organization, Tender
from app.render import gpn_forms as gpn
from app.render.package import build_package
from app.repo import Repo
from app.render.context import GPN_ASSESSMENT, GPN_ASSESSMENT_ANSWERS
from app.web import create_app
from scripts.install_customer_templates import install


def test_customer_forms_use_only_selected_resources_and_confirmed_answers(tmp_path):
    path = tmp_path / "test.db"
    client = create_app(path).test_client()

    def post(url, body):
        response = client.post(url, json=body)
        assert response.status_code in (200, 201), response.get_json()
        return response.get_json()

    def create(entity, **fields):
        return post(f"/api/{entity}", fields)["id"]

    conn = db.connect(path)
    try:
        install(Repo(conn))
    finally:
        conn.close()
    customer = next(row["id"] for row in client.get("/api/organizations").get_json()
                    if "Хантос" in row["short_name"])
    participant = create("organizations", role="participant", short_name="Исполнитель",
                         head_name="Петров П. П.", head_position="Директор")
    staff = create("employees", full_name="Петров П. П.", position="Эколог")
    create("employees", full_name="Не выбранный сотрудник")
    equipment = create("assets", name="Выбранный анализатор", maker_model="Модель 1",
                       quantity="2", ownership_basis="Аренда")
    create("assets", name="Не выбрано")
    experience = create("experience", subject="Проект СЗЗ", customer_name="Другой заказчик",
                        amount="123456.78", term="2025, 100%")
    tender = create("tenders", customer_id=customer, participant_id=participant,
                    purchase_number="TEST-1", subject="Подготовка проекта СЗЗ",
                    experience_ids=[experience])
    base = f"/api/tenders/{tender}"
    staff_doc = "Справка о кадровых ресурсах"
    assets_doc = "Справка о МТР"
    experience_doc = "Справка об опыте"

    assert "Номер лота" in post(f"{base}/preview", {"document": staff_doc})["refusal"]
    answers = {"Номер лота": "LOT-1"}
    post("/api/tenders", {"id": tender, "deal_answers": answers})
    assert "отметьте сотрудников" in post(f"{base}/preview", {"document": staff_doc})["refusal"]
    assert GPN_ASSESSMENT_ANSWERS[0] in post(
        f"{base}/preview", {"document": GPN_ASSESSMENT})["refusal"]
    shown_conditions = client.get(f"{base}/conditions").get_json()["rows"]
    assert all(any(row["condition"] == name and GPN_ASSESSMENT in row["documents"]
                   for row in shown_conditions) for name in GPN_ASSESSMENT_ANSWERS)

    answers.update({name: f"Проверенный ответ {n}"
                    for n, name in enumerate(GPN_ASSESSMENT_ANSWERS, 1)})
    post("/api/tenders", {"id": tender, "deal_answers": answers, "employee_ids": [staff, staff],
                          "asset_ids": [equipment]})
    stored = client.get(base).get_json()
    assert stored["employee_ids"] == [staff] and stored["asset_ids"] == [equipment]
    documents = [staff_doc, assets_doc, experience_doc, GPN_ASSESSMENT]
    result = post(f"{base}/build", {"documents": documents, "output_dir": str(tmp_path / "out")})
    assert result["skipped"] == []
    docs = {name: Document(file["path"]) for name, file in zip(documents, result["written"], strict=True)}

    staff_rows = docs[staff_doc].tables[0].rows
    assert len(staff_rows) == 2 and staff_rows[1].cells[2].text == "Петров П. П."
    asset_rows = docs[assets_doc].tables[0].rows
    assert len(asset_rows) == 4 and asset_rows[1].cells[2].text == "Выбранный анализатор"
    assert asset_rows[1].cells[-1].text == "2"
    cells = [cell.text for cell in docs[experience_doc].tables[0].rows[1].cells]
    assert cells[2:] == ["Проект СЗЗ", "Другой заказчик", "123,45678", "", "2025, 100%"]
    assert [row.cells[1].text for row in docs[GPN_ASSESSMENT].tables[0].rows[1:]] == [
        f"Проверенный ответ {n}" for n in range(1, 6)]


def test_additional_forms_are_installed_and_rendered_with_explicit_answers(tmp_path):
    path = tmp_path / "test.db"
    app = create_app(path)
    conn = db.connect(path)
    try:
        repo = Repo(conn)
        install(repo)
        install(repo)
        participant = repo.save_organization(Organization(
            short_name="ООО «Исполнитель»", full_name="Общество «Исполнитель»",
            inn="5900000000", ogrn="1000000000000", kpp="590001001",
            head_name="Петров П. П.", head_position="Директор", head_basis="Устав",
            legal_address="Тестовый адрес"))
        with app.test_client() as client:
            for label in ("ННГ", "Хантос", "СН-МНГ"):
                customer = next(org for org in repo.list_organizations("customer")
                                if label in org.short_name)
                specs = repo.find_customer_profile(customer.id)
                configured = {name: gpn.extra_word_form(name, spec)
                              for name, spec in specs.docs.items()}
                forms = {name: item for name, item in configured.items() if item}
                assert all(name not in specs.documents for name in forms)
                answers = {key: f"Тестовое значение {i}" for i, key in enumerate({
                    field for form in forms.values() for field in form.conditions})}
                answers.update({form.confirmation: "Не подтверждаю"
                                for form in forms.values() if form.confirmation})
                record = Tender(customer_id=customer.id, participant_id=participant,
                                purchase_number="TEST-2", subject="Предмет закупки",
                                deal_answers=answers)
                tender_id = repo.save_tender(record)
                base = f"/api/tenders/{tender_id}"
                shown = client.get(f"{base}/conditions").get_json()["rows"]
                assert all(any(row["condition"] == name for row in shown) for name in answers)
                for form in forms.values():
                    if form.confirmation:
                        sample = client.post(f"{base}/preview", json={"document": form.document})
                        assert "содержит заявление" in sample.get_json()["refusal"]
                        answers[form.confirmation] = "Подтверждаю"
                repo.save_tender(replace(record, id=tender_id, deal_answers=answers))
                output = build_package(tender_id, list(forms), tmp_path / label, repo)
                assert not output.skipped and len(output.written) == len(forms), output.skipped
                for name, file in zip(forms, output.written, strict=True):
                    doc = Document(file)
                    body = [p.text for p in doc.paragraphs]
                    body.extend(cell.text for table in doc.tables
                                for row in table.rows for cell in row.cells)
                    merged = "\n".join(body)
                    assert body[0].startswith("ЧЕРНОВИК.") and "{{" not in merged
                    for condition in forms[name].conditions:
                        assert answers[condition] in merged, (label, name, condition)
    finally:
        conn.close()


def test_gpn_application_prints_one_confirmed_variant_and_attachment_list(tmp_path):
    path = tmp_path / "test.db"
    create_app(path)
    conn = db.connect(path)
    try:
        repo = Repo(conn)
        install(repo)
        customer = next(org for org in repo.list_organizations("customer")
                        if "ННГ" in org.short_name)
        participant = repo.save_organization(Organization(
            short_name="Заявитель", inn="5900000000",
            head_name="Петров П. П.", head_position="Директор"))
        bid = "Заявка на участие"
        record = Tender(customer_id=customer.id, participant_id=participant,
                        purchase_number="РНП-102", subject="Подготовка проекта СЗЗ")
        answers = {gpn.LOT_NUMBER: "ЛОТ-1", gpn.BID_AGREEMENT: "2",
                   gpn.BID_APPROVAL: "0", gpn.BID_ATTACHMENTS: "Справка о МТР — 2 страницы"}
        tender_id = repo.save_tender(replace(record, deal_answers=answers))
        result = build_package(tender_id, [bid], tmp_path / "out", repo)
        assert not result.written and "Допустимые значения" in result.skipped[0][1]

        answers[gpn.BID_APPROVAL] = "2"
        repo.save_tender(replace(record, id=tender_id, deal_answers=answers))
        result = build_package(tender_id, [bid], tmp_path / "out", repo)
        (output,) = result.written
        text = "\n".join(p.text for p in Document(output).paragraphs)
        assert "ЛОТ-1" in text and "Справка о МТР — 2 страницы" in text
        assert "за исключением арбитражной оговорки" in text
        assert "приложено к настоящей заявке" in text
        assert "будет предоставлено до момента заключения договора" not in text
        assert "данная организация не подпадает" not in text
    finally:
        conn.close()
