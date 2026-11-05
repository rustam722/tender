"""Подстановка контекста в Word-шаблон.

Функция делает одно: берёт шаблон, подставляет в него контекст и пишет файл.
Контекст строит `app.render.context`, выбор шаблона и имени файла -- забота
вызывающего кода.

Для автора шаблона (синтаксис docxtpl / Jinja2):

    {{ org.short_name }}                  значение; пустое поле -- пустое место, не «None»
    {{ money(estimate.total_with_vat) }}           сумма: «5 985 000,00» с неразрывными пробелами
    {{ date_long(tender.doc_date) }}      дата: «24» октября 2025 г.

    таблица -- три строки таблицы Word, теги стоят в первой и последней:
        {%tr for r in assets.rows %}      эта строка в готовом файле исчезает
        {{ r[0] }} | {{ r[1] }} | ...     размножается по числу записей
        {%tr endfor %}                    и эта исчезает
    Пустой список даёт таблицу из одной шапки.

Значения передаются строками, и docxtpl сам экранирует `& < >`; текст значения
не разбирается как шаблон, так что `{{ ... }}` внутри данных остаётся текстом.
Обращение к ключу, которого нет в контексте, -- ошибка, а не пустое место:
документ, молча потерявший реквизит, в тендер не уходит.
"""
import re
from pathlib import Path

import jinja2
from docxtpl import DocxTemplate


class TemplateError(Exception):
    """Шаблон не удалось заполнить; текст на русском, с именем шаблона и причиной."""


_UNDEFINED_NAME = re.compile(r"'(\w+)' is undefined|has no attribute '(\w+)'")


def _missing_key(exc: jinja2.UndefinedError) -> str:
    match = _UNDEFINED_NAME.search(str(exc))
    if match:
        return f"в данных нет значения «{match.group(1) or match.group(2)}»"
    return "шаблон обращается к значению, которого нет в данных"


def _environment() -> jinja2.Environment:
    # StrictUndefined: пропавший ключ -- ошибка. finalize: None печатается пустотой,
    # иначе необязательное поле (год, ставка) превратилось бы в слово «None».
    return jinja2.Environment(
        undefined=jinja2.StrictUndefined,
        finalize=lambda value: "" if value is None else value,
    )


def render_docx(template: Path, context: dict, out: Path) -> Path:
    """Подставить `context` в `template` и сохранить результат в `out`.

    Возвращает `out`. Любой сбой -- `TemplateError` с именем шаблона и причиной.
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

    try:
        tpl.render(context, jinja_env=_environment(), autoescape=True)
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
        raise TemplateError(
            f"Не удалось сохранить «{out.name}»: проверьте, что файл не открыт "
            f"в Word и что папка существует.") from exc
    return out
