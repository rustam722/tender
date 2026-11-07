"""Интерфейс: что можно проверить без браузера -- и чего нельзя.

Четыре слоя, от грубого к точному:

1. Текст страницы и скрипта (Python): три экрана названы, файлы отдаются,
   страница не ходит во внешнюю сеть, в скрипте один `fetch`, нет разбора и
   форматирования чисел, нет арифметических операторов, все тексты для человека
   по-русски.
2. Чистые функции (`tests/ui_logic.test.js`, Node): запросы с типом содержимого
   JSON, разбор отказа, разбор вставки, галочки документов, реестр сборок и
   замок, который переживает пересборку экрана. Сначала с подставным `fetch`,
   затем против настоящего сервера Werkzeug и настоящей базы.
3. Исполнение кода отрисовки (`tests/ui_render.test.js`, Node): весь код
   страницы запускается поверх заглушки DOM (`tests/fake_dom.js`) и работает
   с настоящим сервером. Проверяется, что построение таблиц, отрисовка расчёта и
   результата сборки, переходы между экранами доходят до конца и не падают на
   ответах без итогов и с `warning`, и что на экране нет «undefined»/«null»/«NaN».
4. Вид, удобство, настоящее поведение браузера (фокус, диалог, буфер обмена,
   раскрытие блоков, прокрутка) -- НЕ проверяются ничем из этого файла. Это
   делает человек в задаче 15. Заглушка DOM -- не браузер: она не знает ни CSS,
   ни вёрстки.

Нет Node -- слои 2 и 3 пропускаются (с причиной), а не считаются зелёными.
"""
import os
import re
import shutil
import subprocess
import threading

import docx
import pytest
from werkzeug.serving import make_server

from app import config
from app.config import BASE_DIR
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
    for forbidden in (r"(?<!JSON\.)\bparse\w*", r"\bNumber\b", r"\.toFixed\b", r"\.toPrecision\b",
                      r"\.toLocale\w*", r"\bIntl\b", r"\bMath\.\w+", r"\bBigInt\b",
                      r"\bDecimal\b", r"\bisNaN\b", r"\bisFinite\b", r"\.valueOf\b"):
        assert not re.search(forbidden, code), f"в app.js найдено запрещённое: {forbidden}"


def _strip_literals(code: str) -> str:
    """Код без строк и регулярных выражений: остаются только операции."""
    code = re.sub(r"`(?:\\.|\$\{[^}]*\}|[^`\\])*`", "``", code)
    code = re.sub(r"'(?:\\.|[^'\\\n])*'", "''", code)
    code = re.sub(r'"(?:\\.|[^"\\\n])*"', '""', code)
    code = re.sub(r"(?<=[(,=:!&|?{;])\s*/(?![/*])(?:\\.|\[(?:\\.|[^\]\\\n])*\]|[^/\\\n\[])+/[a-z]*", " /r/", code)
    return re.sub(r"//[^\n]*", "", code)       # хвостовые комментарии: литералы уже убраны


def test_script_has_no_arithmetic_operators_on_anything():
    """Ни умножения, ни деления, ни вычитания, ни сложения: ни `* 100`, ни `a - b`.
    Единственное разрешённое исключение -- счётчик отрисовок `renderToken += 1`
    (не число на экране: защита от гонки при быстрой смене экранов). Конкатенация
    строк тоже не нужна: тексты собираются шаблонными строками."""
    code = _strip_literals(_code_without_comments())
    code = code.replace("renderToken += 1", "")
    hits = []
    for number, line in enumerate(code.split("\n"), 1):
        for match in re.finditer(r"[\w)\]]\s*(\+\+|--|\+=|-=|\*=|/=|%=|\*\*|[*%])|[\w)\]]\s+[+\-/]\s+[\w(]|[\w)\]][+\-/][\w(]", line):
            hits.append((number, line.strip()[:100]))
    assert not hits, hits


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


def test_every_literal_text_shown_to_the_user_is_russian():
    """Строже общей проверки выше: любой литерал, который уходит человеку --
    `text: '…'`, `textContent = '…'`, `.show('вид', '…')`, `title:`, `message:`,
    `placeholder:`, `ok:`/`cancel:` -- содержит кириллицу или вовсе не содержит
    латинских букв. Односложное «Save» и `text: 'OK'` здесь не проходят."""
    code = _code_without_comments()
    shown = []
    patterns = (r"\b(?:text|title|message|placeholder|ok|cancel)\s*[:=]\s*(?:'((?:\\.|[^'\\\n])*)'|\"((?:\\.|[^\"\\\n])*)\"|`((?:\\.|[^`\\])*)`)",
                r"\.show\(\s*'[a-z]+'\s*,\s*(?:'((?:\\.|[^'\\\n])*)'|`((?:\\.|[^`\\])*)`)")
    for pattern in patterns:
        for match in re.finditer(pattern, code):
            shown.extend(g for g in match.groups() if g is not None)
    assert len(shown) > 80, f"проверка не нашла тексты ({len(shown)}): шаблон устарел"
    english = []
    for literal in shown:
        literal = re.sub(r"\$\{[^}]*\}", " ", literal)
        if literal.startswith("http"):
            continue                              # пример адреса в подсказке поля
        if re.search("[A-Za-z]{2,}", literal) and not re.search("[А-Яа-яЁё]", literal):
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
    monkeypatch.setattr(config, "TEMPLATES_DOCS_DIR", templates)
    monkeypatch.setattr(config, "CELLMAPS_DIR", tmp_path / "cellmaps")
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


@NEEDS_NODE
def test_screens_render_against_the_real_server(live_server):
    """Код страницы ИСПОЛНЯЕТСЯ в Node поверх заглушки DOM (tests/fake_dom.js) и
    разговаривает с настоящим сервером: справочники, карточка тендера, калькулятор,
    сборка и переходы между экранами рисуются без падений и без «undefined» на
    экране -- на полных ответах, без итогов и с предупреждением. Это проверка
    того, что код отрисовки доходит до конца; как это выглядит, она не знает."""
    out_dir = live_server["tmp"] / "render"
    out_dir.mkdir()
    done = _run_node("tests/ui_render.test.js", env={
        "TENDER_BASE": live_server["base"], "TENDER_OUT_DIR": str(out_dir),
        "TENDER_REPORT_REQUESTS": "1"})
    assert done.returncode == 0, done.stdout + done.stderr
    counted = re.search(r"REQUESTS_SENT=(\d+)", done.stdout)
    assert counted and int(counted.group(1)) > 80, done.stdout[-500:]
