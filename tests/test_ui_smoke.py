"""Интерфейс: что можно проверить без браузера.

Эти тесты доказывают СТРОЕНИЕ, а не вид и не удобство: что три экрана есть,
что страница находит свои файлы, что каждая запись уходит с нужным типом
содержимого, что ответы разбираются, что предупреждение `warning` не приводит
к повторной отправке, что кнопка сборки не даёт начать вторую сборку, что вопрос
о замене называет документы. Как это выглядит и удобно ли этим пользоваться,
они не знают: это проверяет человек в задаче 15.

Логика экрана лежит в `app/web/static/app.js` чистыми функциями и выносится
наружу, когда нет страницы (`module.exports`). Эти функции гоняет Node.js
(`tests/ui_logic.test.js`, только встроенные модули, ни npm, ни сборки): сначала
с подставным `fetch`, потом против настоящего сервера Werkzeug на свободном порту
и настоящей базы во временной папке -- через настоящий HTTP, с настоящими отказами
сервера (415 без типа содержимого, 400 с русской фразой). Нет Node -- эти тесты
пропускаются, а не молча считаются зелёными: в отчёте это видно.
"""
import json
import os
import re
import shutil
import subprocess
import threading
from pathlib import Path

import docx
import pytest
from werkzeug.serving import make_server

from app.config import BASE_DIR
from app.models import CustomerProfile, DocSpec
from app.render import package as package_module
from app.repo import Repo
from app.web import create_app

STATIC = BASE_DIR / "app" / "web" / "static"
NODE = shutil.which("node")
NEEDS_NODE = pytest.mark.skipif(NODE is None, reason="Node.js не найден: логику экрана гонять нечем")


@pytest.fixture
def client(tmp_path):
    return create_app(tmp_path / "ui.db").test_client()


# ---- страница и файлы ----------------------------------------------------------

def test_page_references_all_three_screens(client):
    html = client.get("/").get_data(as_text=True)
    for anchor in ("Справочники", "Карточка тендера", "Сборка"):
        assert anchor in html


def test_page_loads_its_own_assets(client):
    for path, mimetype in (("/static/app.js", "text/javascript"), ("/static/app.css", "text/css")):
        response = client.get(path, buffered=True)
        assert response.status_code == 200, path
        assert response.mimetype in (mimetype, "application/javascript"), (path, response.mimetype)
        assert len(response.get_data()) > 500, f"{path}: файл пуст"


def test_page_points_at_the_assets_it_ships(client):
    """Страница ссылается ровно на те файлы, что отдаёт сервер, и больше ни на
    что: ни на сторонние скрипты, ни на шрифты из сети (работа без интернета)."""
    html = client.get("/").get_data(as_text=True)
    refs = re.findall(r'(?:src|href)="([^"#]+)"', html)
    local = [r for r in refs if r.startswith("/static/")]
    assert "/static/app.js" in local and "/static/app.css" in local
    for ref in refs:
        assert not re.match(r"(https?:)?//", ref), f"внешняя ссылка на странице: {ref}"
        assert client.get(ref, buffered=True).status_code == 200, ref


def test_page_is_russian_and_says_so_without_scripts(client):
    html = client.get("/").get_data(as_text=True)
    assert '<html lang="ru"' in html
    title = re.search(r"<title>(.*?)</title>", html, re.S).group(1)
    assert re.search("[А-Яа-я]", title)
    noscript = re.search(r"<noscript>(.*?)</noscript>", html, re.S)
    assert noscript and re.search("[А-Яа-я]", noscript.group(1))      # без скриптов -- фраза, не пустота


def test_page_has_no_inline_handlers_or_external_code(client):
    """Весь код -- в app.js: встроенные обработчики и чужие скрипты не нужны."""
    html = client.get("/").get_data(as_text=True)
    assert not re.search(r"\son[a-z]+\s*=", html)
    assert "<script src=\"/static/app.js\"" in html.replace("'", '"')


# ---- грубые запреты на текст скрипта -------------------------------------------

def _code_without_comments() -> str:
    source = (STATIC / "app.js").read_text(encoding="utf-8")
    source = re.sub(r"/\*.*?\*/", "", source, flags=re.S)
    return re.sub(r"(?m)^\s*//.*$", "", source)


def test_script_never_parses_or_formats_a_number():
    """В JavaScript нет Decimal: всё, что превращает строку в число или число в
    строку с округлением, -- это float. Показывается то, что прислал сервер.
    (Грубая проверка текста; настоящая гарантия -- что сервер отдаёт готовые
    строки, это проверяет test_web.py.)"""
    code = _code_without_comments()
    for forbidden in (r"\bparseFloat\b", r"\bparseInt\b", r"\bNumber\(", r"\.toFixed\b",
                      r"\.toLocaleString\b", r"\bIntl\.", r"\bMath\.(round|floor|ceil|trunc)\b",
                      r"\bBigInt\b"):
        assert not re.search(forbidden, code), f"в app.js найдено запрещённое: {forbidden}"


def test_every_visible_string_is_russian():
    """Всё, что видит человек, по-русски: ни в разметке, ни в строках скрипта нет
    английских фраз. (Грубо: строковый литерал с двумя латинскими словами и без
    единой кириллицы считается английской фразой, если эти слова не названия
    CSS-классов: списки классов вроде «card inline» -- не текст для человека.)"""
    html = (BASE_DIR / "app" / "web" / "templates" / "index.html").read_text(encoding="utf-8")
    html = re.sub(r"(?s)<(script|style)\b.*?</\1>", "", html)
    assert not re.search(r"[A-Za-z]{3,}", re.sub(r"<[^>]+>", " ", html))
    css = (STATIC / "app.css").read_text(encoding="utf-8")
    technical = set(re.findall(r"\.([A-Za-z][\w-]*)", css)) | {"use", "strict"}
    english = []
    for match in re.finditer(r"""'([^'
]*)'|"([^"
]*)"|`([^`]*)`""", _code_without_comments()):
        literal = re.sub(r"\$\{[^}]*\}", " ", next(g for g in match.groups() if g is not None))
        words = re.findall(r"[A-Za-z][A-Za-z-]*", literal)
        if (re.search(r"[A-Za-z]{3,}\s+[A-Za-z]{2,}", literal) and not re.search("[А-Яа-яЁё]", literal)
                and not all(word in technical for word in words)):
            english.append(literal)
    assert not english, english


def test_every_request_goes_through_one_function():
    """Один `fetch` на весь скрипт: тип содержимого JSON ставится в одном месте,
    а не в двадцати. Забытый заголовок у одной из ручек дал бы 415 только при
    ручном обходе."""
    code = _code_without_comments()
    assert len(re.findall(r"\bfetch\(", code)) == 1


# ---- логика экрана через Node ----------------------------------------------------

def _run_node(*args, env=None, timeout=120):
    merged = {**os.environ, **(env or {})}
    return subprocess.run([NODE, *args], capture_output=True, text=True, encoding="utf-8",
                          env=merged, timeout=timeout, cwd=BASE_DIR)


@NEEDS_NODE
def test_screen_logic_unit_tests():
    """Чистые функции экрана с подставным fetch: заголовки, ошибки, разбор
    вставки, выбор документов, замок кнопки сборки, вопрос о замене."""
    done = _run_node("tests/ui_logic.test.js")
    assert done.returncode == 0, done.stdout + done.stderr


@pytest.fixture
def live_server(tmp_path, monkeypatch):
    """Настоящий сервер на свободном порту, настоящая база, шаблон и профиль."""
    templates = tmp_path / "templates_docs"
    templates.mkdir()
    monkeypatch.setattr(package_module, "TEMPLATES_DOCS_DIR", templates)
    monkeypatch.setattr(package_module, "CELLMAPS_DIR", tmp_path / "cellmaps")
    for name in ("Справка о МТР", "Декларация соответствия"):
        document = docx.Document()
        document.add_paragraph(f"{name}: {{{{ tender.subject }}}}")
        document.save(templates / f"{name}.docx")

    app = create_app(tmp_path / "live.db")
    server = make_server("127.0.0.1", 0, app, threaded=True)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield {"base": f"http://127.0.0.1:{server.server_address[1]}", "app": app,
               "tmp": tmp_path}
    finally:
        server.shutdown()
        server.server_close()
        thread.join(5)


@NEEDS_NODE
def test_screen_logic_against_the_real_server(live_server):
    """Те же функции, но через настоящий HTTP: запись целиком возвращается
    на сервер и не теряет точность, правка не теряется, проценты уходят
    процентами, вставка списком заменяет лот, подбор часов и расхождение
    приходят готовыми строками, сборка спрашивает о замене по именам файлов и
    не даёт начать вторую."""
    out_dir = live_server["tmp"] / "Заявка"
    done = _run_node("tests/ui_logic.test.js", env={
        "TENDER_BASE": live_server["base"], "TENDER_OUT_DIR": str(out_dir),
        "TENDER_REPORT_REQUESTS": "1"})
    assert done.returncode == 0, done.stdout + done.stderr
    assert out_dir.is_dir() and any(out_dir.iterdir())       # сборка действительно писала
    # страховка от пустого прогона: сценарий обязан дойти до сервера
    counted = re.search(r"REQUESTS_SENT=(\d+)", done.stdout)
    assert counted and int(counted.group(1)) > 25, done.stdout[-500:]
