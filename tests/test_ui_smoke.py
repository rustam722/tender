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


def _skip_quoted(code: str, i: int) -> int:
    """Индекс сразу за закрывающей кавычкой строки, начинающейся на i."""
    quote = code[i]
    i += 1
    while i < len(code):
        if code[i] == "\\":
            i += 2
            continue
        if code[i] == "\n":           # незакрытая строка: перенос отдаём разбору обратно
            return i
        if code[i] == quote:
            return i + 1
        i += 1
    return i


def _skip_regex(code: str, i: int) -> int:
    """Индекс сразу за регулярным выражением (с флагами), начинающимся на i.
    Внутри класса символов `[…]` косая черта не закрывает выражение: так написано
    `.replace(/[\\/]+/g, '/')` в app.js."""
    i += 1
    in_class = False
    while i < len(code):
        ch = code[i]
        if ch == "\\":
            i += 2
            continue
        if ch == "\n":
            return i
        if in_class:
            if ch == "]":
                in_class = False
        elif ch == "[":
            in_class = True
        elif ch == "/":
            i += 1
            while i < len(code) and code[i].isalpha():
                i += 1
            return i
        i += 1
    return i


# После этих знаков косая черта начинает регулярное выражение, а не делит. Ни
# одной регулярки после ключевого слова (`return /…/`) в app.js нет; появится --
# сторож посчитает её делением и пожалуется, то есть ошибётся в сторону шума.
_EXPRESSION_START = set("([{,;:=!&|?+-*/%<>~^") | {""}


def _strip_literals(code: str) -> str:
    """Код без ТЕКСТА строк, комментариев и регулярных выражений -- но С
    содержимым подстановок `${…}` шаблонных строк.

    Подстановка -- ровно то место, где экран собирает выводимое: «${rate * 100}»
    и есть находка, из-за которой запрет появился, поэтому её содержимое обязано
    дойти до проверки. Текст вокруг подстановки -- не код: «2 * 2» в предложении
    и «/items» в адресе ничего не считают, и их видеть не нужно.

    Разбор посимвольный, а не регуляркой: подстановки вкладываются друг в друга
    (`${items.map((x) => `${x}`)}`), и `\\$\\{[^}]*\\}` на первой же `}` обрывается.
    Переносы строк сохраняются -- номер строки в находке должен быть настоящим."""
    out = []
    prev = ""                 # последний значимый (не пробельный) выданный знак
    stack = []                # "tpl" -- текст шаблонной строки; int -- глубина скобок в ${}

    def emit(text):
        nonlocal prev
        out.append(text)
        stripped = text.strip()
        if stripped:
            prev = stripped[-1]

    i, n = 0, len(code)
    while i < n:
        ch = code[i]
        if stack and stack[-1] == "tpl":                      # внутри текста шаблонной строки
            if ch == "\\":
                i += 2
            elif ch == "`":
                stack.pop()
                emit("`")
                i += 1
            elif ch == "$" and code[i + 1:i + 2] == "{":
                stack.append(0)
                emit(" ")
                i += 2
            else:
                out.append("\n" if ch == "\n" else "")        # текст выбрасывается, строки -- нет
                i += 1
            continue
        if ch == "`":
            stack.append("tpl")
            emit("`")
            i += 1
        elif ch in "'\"":
            i = _skip_quoted(code, i)
            emit(ch * 2)
        elif code[i:i + 2] == "//":
            while i < n and code[i] != "\n":
                i += 1
        elif code[i:i + 2] == "/*":
            end = code.find("*/", i + 2)
            end = n if end < 0 else end + 2
            out.append("\n" * code.count("\n", i, end))
            i = end
        elif ch == "/" and prev in _EXPRESSION_START:
            i = _skip_regex(code, i)
            emit(" rx ")                                      # имя без знаков операций
        elif ch == "{" and stack and isinstance(stack[-1], int):
            stack[-1] += 1
            emit(ch)
            i += 1
        elif ch == "}" and stack and isinstance(stack[-1], int):
            if stack[-1] == 0:
                stack.pop()                                   # подстановка закрылась -- снова текст
                emit(" ")
            else:
                stack[-1] -= 1
                emit(ch)
            i += 1
        else:
            emit(ch)
            i += 1
    return "".join(out)


# Единственное разрешённое исключение, названное и объяснённое: счётчик отрисовок.
# Это не число на экране, а защита от гонки при быстрой смене экранов.
ALLOWED_ARITHMETIC = "renderToken += 1"

_OPERATORS = re.compile(r"\*\*=?|\+\+|--|[+\-*/%]=|[+\-*/%~]")
_OPERAND_END = re.compile(r"[\w$)\]'\"`]")        # чем может кончаться операнд
_OPERAND_START = re.compile(r"[A-Za-z_$(\['\"`]")  # имя, вызов, скобка или текст


def _is_arithmetic(token: str, prev: str, nxt: str) -> bool:
    """Помечать ли операцию, зная её знак и ближайшие значимые знаки вокруг.

    `*`, `/`, `%`, `**`, `++`, `--` и присваивания с операцией -- всегда счёт.
    `+`, `-`, `~` смотрят по сторонам: между операндами (с любыми пробелами и
    через перенос строки) -- счёт; без операнда слева -- унарная форма, и она
    счёт перед ИМЕНЕМ, вызовом или текстом (`+value`, `~index`, `+'5'` -- всё
    это приведение к float), но не перед числом: `cells.at(-1)` и
    `cells.slice(0, -1)` -- индекс с конца, а не арифметика."""
    if token not in ("+", "-", "~"):
        return True
    if not nxt:
        return False
    if prev and _OPERAND_END.fullmatch(prev):
        return bool(_OPERAND_START.fullmatch(nxt)) or nxt.isdigit()
    return bool(_OPERAND_START.fullmatch(nxt))


def _arithmetic_hits(source: str) -> list:
    """Находки сторожа: (номер строки, сама строка) на каждую операцию."""
    code = _strip_literals(source).replace(ALLOWED_ARITHMETIC, "")
    lines = code.split("\n")
    hits = []
    for match in _OPERATORS.finditer(code):
        before = code[:match.start()].rstrip()
        after = code[match.end():].lstrip()
        if not _is_arithmetic(match.group(0), before[-1:], after[:1]):
            continue
        number = code.count("\n", 0, match.start()) + 1
        hits.append((number, lines[number - 1].strip()[:100], match.group(0)))
    return hits


# Пробы сторожа: на каждое написание -- и образец, который он ОБЯЗАН пометить,
# и близкий к нему, который метить НЕЛЬЗЯ. Жадный сторож, метящий всё подряд,
# бесполезен ровно так же, как слепой: его просто выключат.
ARITHMETIC_PROBES = (
    ("умножение на сто внутри подстановки шаблонной строки",
     "const t = `ставка ${rate * 100} %`;",
     "const t = `ставка ${rate} %, и 2 * 2 в тексте`;"),
    ("сложение внутри подстановки шаблонной строки",
     "const t = `итого ${a + b} руб.`;",
     "const t = `итого ${a}+${b} руб.`;"),
    ("операция внутри ВЛОЖЕННОЙ подстановки",
     "const t = `${items.map((x) => `${x * 2}`)}`;",
     "const t = `${items.map((x) => `${x}`)}`;"),
    ("голое умножение вне строки",
     "const t = a * 100;",
     "/* в комментарии\n   a * 100\n   ничего не считает */"),
    ("пробел только слева",
     "const t = a +b;",
     "const t = `a +b -- это текст`;"),
    ("пробел только справа",
     "const t = a+ b;",
     "const t = `a+ b -- это тоже текст`;"),
    ("унарный плюс перед именем",
     "const n = +value;",
     "const n = cells.at(-1);"),
    ("унарный минус перед именем",
     "const n = -value;",
     "const n = cells.slice(0, -1);"),
    ("унарная тильда перед именем",
     "const n = ~index;",
     "const t = `~${index}`;"),
    ("приведение текста к числу унарным плюсом",
     "const n = +'5';",
     "const t = ['5', '7'];"),
    ("операция в начале строки (перенос выражения)",
     "const t = a\n  + b;",
     "const t = a\n  // + b\n  ;"),
    ("деление",
     "const t = total / count;",
     "const ok = /^\\d+[.)]?$/.test(cells[0]);"),
    ("деление рядом с регуляркой с косой чертой внутри класса",
     "const t = folder.replace(/[\\\\/]+/g, '/') / 2;",
     "const t = folder.replace(/[\\\\/]+/g, '/').replace(/\\/$/, '');"),
    ("приращение счётчика",
     "shownCount += 1;",
     "renderToken += 1;"),
    ("операция в теле стрелочной функции",
     "const f = (x) => x * 2;",
     "const f = (x) => x;"),
    ("конкатенация строк",
     "const t = 'первая ' + 'вторая';",
     "const t = `первая вторая`;"),
    ("остаток от деления",
     "const t = index % 2;",
     "const t = `${index} %`;"),
)


def test_the_arithmetic_guard_catches_every_spelling_it_claims_to():
    """Сторож проверяет сам себя, и проверка настоящая: на каждое написание из
    таблицы ARITHMETIC_PROBES -- образец, который обязан быть помечен, и близкий
    к нему, который метить нельзя. Без обеих половин проба ничего не значит:
    сторож, метящий всё подряд, зелёный на положительных образцах."""
    assert len(ARITHMETIC_PROBES) >= 12, "таблица проб обеднела"
    missed = [(name, sample) for name, sample, _ in ARITHMETIC_PROBES if not _arithmetic_hits(sample)]
    overzealous = [(name, sample, _arithmetic_hits(sample))
                   for name, _, sample in ARITHMETIC_PROBES if _arithmetic_hits(sample)]
    assert not missed, f"сторож не увидел арифметику: {missed}"
    assert not overzealous, f"сторож пометил то, что арифметикой не является: {overzealous}"


def test_script_has_no_arithmetic_operators_on_anything():
    """Ни умножения, ни деления, ни вычитания, ни сложения: ни `* 100`, ни `a - b`,
    ни `${rate * 100}` внутри шаблонной строки, ни конкатенации текстов (тексты
    собираются шаблонными строками). Единственное разрешённое исключение --
    счётчик отрисовок `renderToken += 1` (не число на экране: защита от гонки при
    быстрой смене экранов), он назван в ALLOWED_ARITHMETIC. Унарный минус перед
    числом -- не исключение, а не арифметика: `cells.at(-1)` -- индекс с конца.
    Что именно сторож ловит и чего не ловит -- в таблице ARITHMETIC_PROBES."""
    hits = _arithmetic_hits((STATIC / "app.js").read_text(encoding="utf-8"))
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


def _skip_template(code: str, i: int) -> int:
    """Индекс сразу за шаблонной строкой, начинающейся на i: с учётом подстановок,
    вложенных шаблонных строк и обычных строк внутри них."""
    i += 1
    while i < len(code):
        ch = code[i]
        if ch == "\\":
            i += 2
            continue
        if ch == "`":
            return i + 1
        if ch == "$" and code[i + 1:i + 2] == "{":
            depth = 1
            i += 2
            while i < len(code) and depth:
                inner = code[i]
                if inner in "'\"":
                    i = _skip_quoted(code, i)
                    continue
                if inner == "`":
                    i = _skip_template(code, i)
                    continue
                if inner == "{":
                    depth += 1
                elif inner == "}":
                    depth -= 1
                i += 1
            continue
        i += 1
    return i


def _call_args(code: str, open_paren: int) -> list:
    """Аргументы вызова, чья открывающая скобка стоит на open_paren: по куску
    исходного текста на аргумент, со счётом скобок и без заглядывания в строки."""
    args = []
    depth = 0
    start = open_paren + 1
    i = open_paren
    while i < len(code):
        ch = code[i]
        if ch in "'\"":
            i = _skip_quoted(code, i)
            continue
        if ch == "`":
            i = _skip_template(code, i)
            continue
        if ch in "([{":
            depth += 1
        elif ch in ")]}":
            depth -= 1
            if depth == 0:
                args.append(code[start:i])
                return args
        elif ch == "," and depth == 1:
            args.append(code[start:i])
            start = i + 1
        i += 1
    return args


def _open_literals(fragment: str) -> list:
    """Литералы фрагмента, не спрятанные внутрь вызова или объекта:
    `isNew ? 'Новый тендер' : 'Шапка тендера'` даёт оба (человек увидит любой),
    а `h('span', { class: 'sub', text: r.sub })` -- ни одного: там имя тега и
    название класса, а не текст для человека."""
    out = []
    depth = 0
    i = 0
    while i < len(fragment):
        ch = fragment[i]
        if ch in "'\"`":
            end = _skip_template(fragment, i) if ch == "`" else _skip_quoted(fragment, i)
            if depth == 0:
                out.append(fragment[i + 1:end - 1])
            i = end
            continue
        if ch in "({":
            depth += 1
        elif ch in ")}":
            depth -= 1
        i += 1
    return out


def _child_literals(code: str) -> list:
    """Литералы, которые уходят человеку ДЕТЬМИ, а не через `text:`.
    `h('h2', {}, 'Настройки калькуляции ', tag)` печатает ровно так же, как
    `text:`, и такой литерал обязан быть русским. Разбор по месту вызова: у `h(`
    дети начинаются с третьего аргумента (тег и атрибуты -- не текст), у
    `append(` и `replaceChildren(` -- с первого."""
    out = []
    for match in re.finditer(r"(?:\bh|\.append|\.replaceChildren)\s*\(", code):
        args = _call_args(code, match.end() - 1)
        for arg in args[2 if match.group(0).startswith("h") else 0:]:
            out.extend(_open_literals(arg))
    return out


def test_every_literal_text_shown_to_the_user_is_russian():
    """Строже общей проверки выше: любой литерал, который уходит человеку --
    `text: '…'`, `textContent = '…'`, `.show('вид', '…')`, `title:`, `message:`,
    `placeholder:`, `ok:`/`cancel:`, `'aria-label':` и литерал, переданный ДЕТЯМИ
    в `h(…)`/`append(…)` -- содержит кириллицу или вовсе не содержит латинских
    букв. Односложное «Save» и `text: 'OK'` здесь не проходят, и в `aria-label`
    тоже: программа экранного доступа читает человеку то же самое."""
    code = _code_without_comments()
    shown = []
    patterns = (r"\b(?:text|title|message|placeholder|ok|cancel)\s*[:=]\s*(?:'((?:\\.|[^'\\\n])*)'|\"((?:\\.|[^\"\\\n])*)\"|`((?:\\.|[^`\\])*)`)",
                r"['\"]aria-label['\"]\s*:\s*(?:'((?:\\.|[^'\\\n])*)'|\"((?:\\.|[^\"\\\n])*)\"|`((?:\\.|[^`\\])*)`)",
                r"\.show\(\s*'[a-z]+'\s*,\s*(?:'((?:\\.|[^'\\\n])*)'|`((?:\\.|[^`\\])*)`)")
    for pattern in patterns:
        for match in re.finditer(pattern, code):
            shown.extend(g for g in match.groups() if g is not None)
    labels = len(shown)
    children = _child_literals(code)
    shown.extend(children)
    assert labels > 80, f"проверка не нашла подписи ({labels}): шаблон устарел"
    # разбор детей обязан доходить до тех мест, ради которых он написан
    assert len(children) >= 12, f"проверка не нашла дочерние литералы ({len(children)}): разбор устарел"
    for anchor in ('Цена без НДС', 'Папка: ', 'Шапка тендера', 'Все организации'):
        assert anchor in children, f"разбор дочерних литералов не дошёл до «{anchor}»"
    assert 'Показать' in shown, "подписи aria-label до проверки не доходят"
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
