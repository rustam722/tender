"""Подстановка контекста в Word-шаблон.

Функция делает одно: берёт шаблон, подставляет в него контекст и пишет файл.
Контекст строит `app.render.context`, выбор шаблона и имени файла -- забота
вызывающего кода.

Для автора шаблона (синтаксис docxtpl / Jinja2):

    {{ org.short_name }}                  значение; пустое поле -- пустое место, не «None»
    {{ money(estimate.total_with_vat) }}  сумма: «5 985 000,00» с неразрывными пробелами
    {{ date_long(tender.doc_date) }}       дата: «24» октября 2025 г.

    таблица справки -- число колонок задаёт раскладка заказчика, а не шаблон.
    Тег удаляет целиком клетку (строку), где стоит, поэтому парные теги
    занимают отдельные клетки (строки), а прототип стоит между ними:
        шапка:    [{%tc for h in X.headers %}] [{{ h }}] [{%tc endfor %}]
        строка:   {%tr for row in X.rows %}
        данные:   [{%tc for c in row %}] [{{ c }}] [{%tc endfor %}]
        строка:   {%tr endfor %}
    Оформление клетки-прототипа (шрифт, заливка, границы) получают все колонки.
    Ширину клеток ставить «авто»: заранее число колонок неизвестно.
    Пустой список строк даёт таблицу из одной шапки.
    Подтаблицы с dict-строками (items, labor, ...) -- обычный цикл по строкам:
        {%tr for it in items %} / {{ it.name }} | {{ it.price }} | / {%tr endfor %}

Значения передаются строками, и docxtpl сам экранирует `& < >`; текст значения
не разбирается как шаблон, так что `{{ ... }}` внутри данных остаётся текстом.
Обращение к ключу, которого нет в контексте, -- ошибка, а не пустое место:
документ, молча потерявший реквизит, в тендер не уходит.

Пустая графа. Ключ, который ЕСТЬ в контексте, но пуст, документ печатает
пустым местом -- и заявку из-за такой графы отклоняют молча (так ушла заявка
с четырьмя пустыми строками банковских реквизитов). Поэтому отрисовка
запоминает поля, напечатанные пустыми, и отдаёт их списком в `blanks`;
сборка пакета превращает список в предупреждение с именами полей. Правило то
же, что у одиночных клеток карты Excel: пусто -- повод сказать, а законная
пустота отмечается в самих данных, то есть в шаблоне:

    {{ org.website|optional }}      графа законно может остаться пустой

Что видно, а что нет. Видна пустота значения, напечатанного напрямую
(`{{ org.bank_name }}`). Значение, прошедшее через функцию или фильтр
(`{{ money(...) }}`, `{{ x|upper }}`), приходит оттуда обычной строкой, и
пустота в нём не видна; исключение -- `date_ru`/`date_long`, которые отдают
пустоту как есть. Клетки табличных справок (`{{ c }}` из `assets.rows`) и
строки подтаблиц (`{{ it.name }}` из `items`) не проверяются вовсе: там
пустое значение законно (примечание, документ о владении) -- ровно так же,
как в блоках строк карты Excel.

Не видна и пустота, которую шаблон пропускает через цикл по значению
(`declaration_text.splitlines()`): пустой текст декларации даёт документ без
единого требования. Для него стоит отдельный отказ в сборке пакета
(`template_uses` и `_build_one` в `app.render.package`).
"""
import re
from collections.abc import Mapping
from pathlib import Path

import jinja2
from docxtpl import DocxTemplate

from app.render import Blank, save_cause


class TemplateError(Exception):
    """Шаблон не удалось заполнить; текст на русском, с именем шаблона и причиной."""


_UNDEFINED_NAME = re.compile(r"'(\w+)' is undefined|has no attribute '(\w+)'")


def _missing_key(exc: jinja2.UndefinedError) -> str:
    match = _UNDEFINED_NAME.search(str(exc))
    if match:
        return f"в данных нет значения «{match.group(1) or match.group(2)}»"
    return "шаблон обращается к значению, которого нет в данных"


def _watched(value: object, path: str = ""):
    """Копия контекста, в которой пустое значение помечено своим путём.

    Помечаются значения словарей контекста (`org.bank_name`, `tender.deadline`)
    на любой глубине. Элементы списков не помечаются: в клетке табличной
    справки и в строке подтаблицы пустое значение законно, как в блоке строк
    карты Excel. Исходный контекст не меняется -- в книгу Excel и на экран
    уходит он же."""
    if isinstance(value, Mapping):
        return {key: _watched(item, f"{path}.{key}" if path else key)
                for key, item in value.items()}
    if path and (value is None or (isinstance(value, str) and not value.strip())):
        return Blank(value, path)
    return value


def _optional(value: object):
    """Фильтр шаблона: эта графа законно может остаться пустой.

    Снимает пометку пустоты, и предупреждение о пустых графах про неё
    молчит. Факс, которого у фирмы нет, -- не повод пугать человека перед
    подачей; сигнал, который горит всегда, через неделю перестают читать."""
    if value is None:
        return ""
    return "" + value if isinstance(value, str) else value     # пометка снята


def _finalize(value: object, blanks: list[str]):
    """Что печатается на месте `{{ ... }}`: None -- пустота (иначе
    необязательное поле превратилось бы в слово «None»), помеченная пустота --
    пустота, и путь поля запоминается."""
    if isinstance(value, Blank):
        if value.path not in blanks:
            blanks.append(value.path)
        return ""
    return "" if value is None else value


def _environment(blanks: list[str]) -> jinja2.Environment:
    # StrictUndefined: пропавший ключ -- ошибка.
    env = jinja2.Environment(
        undefined=jinja2.StrictUndefined,
        finalize=lambda value: _finalize(value, blanks),
    )
    env.filters["optional"] = _optional
    return env


def template_uses(template: Path, name: str) -> bool:
    """Обращается ли шаблон к значению контекста `name` (`declaration_text`).

    Смотрит на разметку, а не на имя документа: у одного заказчика форма декларации
    печатает требования из профиля, у другого -- свои, и метки в ней нет. Разметку
    читает сам docxtpl: он склеивает куски текста, на которые Word режет метку при
    правке, чего поиск строки по XML не сделает; заглядывает и в колонтитулы.
    Шаблон, который не открывается или не разбирается, -- «нет»: честную причину
    назовёт `render_docx`."""
    try:
        return name in DocxTemplate(Path(template)).get_undeclared_template_variables()
    except Exception:
        return False


def render_docx(template: Path, context: dict, out: Path,
                blanks: list[str] | None = None) -> Path:
    """Подставить `context` в `template` и сохранить результат в `out`.

    Возвращает `out`. Любой сбой -- `TemplateError` с именем шаблона и причиной.

    `blanks` -- список, в который дописываются пути полей, напечатанных
    пустыми (`org.bank_name`), в порядке появления и без повторов. Что в него
    попадает, а что нет, сказано в начале модуля; имена полей по этим путям
    собирает `app.render.context.blank_fields_warning`.
    """
    template, out = Path(template), Path(out)
    name = template.name

    if not template.is_file():
        raise TemplateError(f"Шаблон «{name}» не найден: {template}.")

    try:
        tpl = DocxTemplate(template)
        tpl.get_docx()      # DocxTemplate читает файл лениво -- читаем здесь, а не при рендере
    except Exception as exc:
        raise TemplateError(
            f"Шаблон «{name}» не открывается: файл повреждён или не является "
            f"документом Word (.docx).") from exc

    found: list[str] = []
    try:
        tpl.render(_watched(context), jinja_env=_environment(found),
                   autoescape=True)
    except jinja2.UndefinedError as exc:
        raise TemplateError(f"Шаблон «{name}»: {_missing_key(exc)}.") from exc
    except jinja2.TemplateError as exc:
        raise TemplateError(
            f"Шаблон «{name}» содержит ошибку в разметке: проверьте, что все "
            f"{{{{ }}}} и {{% %}} закрыты, а циклы {{%tr for %}} завершены "
            f"{{%tr endfor %}}.") from exc
    except Exception as exc:
        raise TemplateError(
            f"Шаблон «{name}» не удалось заполнить: непредвиденная ошибка "
            f"({type(exc).__name__}).") from exc

    try:
        tpl.save(out)
    except OSError as exc:
        # Причина называется та, которую сообщила ОС: нехватка места и запрет
        # записи в папку не лечатся закрытием Word (см. app.render.save_cause).
        raise TemplateError(
            f"Не удалось сохранить «{out.name}»: {save_cause(exc)}") from exc
    if blanks is not None:
        blanks.extend(found)
    return out
