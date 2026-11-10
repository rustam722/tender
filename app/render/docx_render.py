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
"""
import re
from pathlib import Path

import jinja2
from docxtpl import DocxTemplate

from app.render import save_cause


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
        # Причина называется та, которую сообщила ОС: нехватка места и запрет
        # записи в папку не лечатся закрытием Word (см. app.render.save_cause).
        raise TemplateError(
            f"Не удалось сохранить «{out.name}»: {save_cause(exc)}") from exc
    return out
