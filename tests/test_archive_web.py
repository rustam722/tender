import pytest

from app import db
from app.web import create_app


def test_save_selected_archive_entries_round_trips_and_can_be_retried(tmp_path):
    db_path = tmp_path / "tender.db"
    conn = db.connect(db_path)
    try:
        for migration in db._MIGRATIONS[:6]:
            conn.executescript(migration)
        conn.execute("INSERT INTO tender (subject) VALUES ('Текущая закупка')")
        conn.commit()
    finally:
        conn.close()

    root = tmp_path / "Архив"
    tender = root / "2025 год" / "Энергосбыт КЭР до 14.03.2025 выиграли"
    bid = tender / "Заявка"
    bid.mkdir(parents=True)
    paper = bid / "Коммерческое предложение.docx"
    paper.write_bytes(b"invalid word file")
    (root / "Незнакомый заказчик").mkdir()
    client = create_app(db_path).test_client()
    preview = client.post("/api/archive/preview", json={"folder": str(root)})
    assert preview.status_code == 200
    entries = preview.get_json()
    assert len(entries) == 2
    chosen = next(entry for entry in entries if entry["outcome"] == "won")
    assert chosen["deadline"] == "2025-03-14"
    assert chosen["services"] == ["КЭР"]
    assert chosen["bid_papers"] == [
        {"path": "Заявка/Коммерческое предложение.docx",
         "kinds": ["proposal"],
         "kind_labels": ["Предложение / коммерческое предложение"]},
    ]
    assert client.get("/api/archive").get_json() == []

    payload = {"folder": str(root), "selected": [chosen["key"]]}
    stale = client.post("/api/archive/import", json={
        **payload, "selected": [chosen["key"], "../За пределами архива"]})
    assert stale.status_code == 409
    assert client.get("/api/archive").get_json() == []

    for _ in range(2):
        response = client.post("/api/archive/import", json=payload)
        assert response.status_code == 200
        assert response.get_json() == [chosen]
    assert create_app(db_path).test_client().get("/api/archive").get_json() == [chosen]
    assert client.get("/api/tenders").get_json()[0]["subject"] == "Текущая закупка"
    assert paper.read_bytes() == b"invalid word file"


@pytest.mark.parametrize("folder", ["", "архив/2025", "//server/share", 1, None])
def test_archive_preview_requires_a_local_absolute_folder(tmp_path, folder):
    client = create_app(tmp_path / "tender.db").test_client()
    response = client.post("/api/archive/preview", json={"folder": folder})
    assert response.status_code == 400
    assert response.get_json()["field"] == "folder"


def test_archive_preview_skips_linked_tender_directories(tmp_path):
    root = tmp_path / "Архив"
    root.mkdir()
    outside = tmp_path / "Другая папка"
    outside.mkdir()
    try:
        (root / "Ссылка").symlink_to(outside, target_is_directory=True)
    except (NotImplementedError, OSError):
        pytest.skip("Создание ссылок недоступно")
    client = create_app(tmp_path / "tender.db").test_client()
    response = client.post("/api/archive/preview", json={"folder": str(root)})
    assert response.status_code == 200
    assert response.get_json() == []
