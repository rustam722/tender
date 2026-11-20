from docx import Document

from app import db
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
