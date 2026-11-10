"""Отказы веб-слоя: что пользователь видит вместо исключения.

Правило одно: наружу уходит `{"error": "<русская фраза>"}` и больше ничего из
исключения -- ни текста, ни имени класса, ни трассировки. Подробности
непредвиденного сбоя пишутся в журнал сервера (окно программы).

Таблица «исключение -> статус»:

    ApiError                          свой статус   текст отказа проверки ввода
    LookupError (ровно этот класс)    404           текст репозитория, по-русски
    InnNotFound                       400           отказ домена
    CaptchaRequired, SourceUnavailable,
    InnLookupError                    502           внешний источник не ответил
    TargetTooLow, PackageError,
    TemplateError, CellMapError,
    LayoutError, ContextError         400           отказ домена, текст по-русски
    sqlite3.IntegrityError            400           см. translate_integrity
    sqlite3.OperationalError          503           база занята или недоступна
    decimal.DecimalException          400           числа слишком велики для счёта
    HTTPException (404, 405, 413...)  свой код      фиксированная русская фраза
    любое другое                      500           общая фраза, деталь -- в журнал

KeyError и IndexError -- потомки LookupError, поэтому проверяется точный класс:
иначе баг в коде выглядел бы как «запись не найдена» с английским текстом.
Прочие ValueError тоже не пропускаются как есть: их текст может быть английским,
а русские ValueError (DocSpec, дата) перехватываются в тех местах, где они
известны.
"""
import re
import sqlite3
from decimal import DecimalException

from flask import current_app, jsonify, request
from werkzeug.exceptions import HTTPException

from app.calc import TargetTooLow
from app.inn.fns import CaptchaRequired, InnLookupError, InnNotFound, SourceUnavailable
from app.render.context import ContextError
from app.render.docx_render import TemplateError
from app.render.layouts import LayoutError
from app.render.package import PackageError
from app.render.xlsx_render import CellMapError

UNEXPECTED = ("Произошла непредвиденная ошибка. Повторите действие; если она "
              "повторяется, покажите разработчику чёрное окно программы: "
              "подробности записаны там.")

_HTTP_TEXT = {
    400: "Запрос не удалось разобрать.",
    403: "Запрос отклонён: программа отвечает только на локальном адресе.",
    404: "Такого адреса в программе нет.",
    405: "Для этого адреса такой способ запроса не поддерживается.",
    413: "Запрос слишком большой: программа принимает до 2 МБ данных за раз.",
    415: "Тело запроса должно быть в формате JSON.",
}
_HTTP_FALLBACK = "Запрос не может быть выполнен."


class ApiError(Exception):
    """Отказ с готовой русской фразой. path -- где в теле запроса ошибка:
    имена полей, ключи словарей и номера элементов списка, от внешнего к
    внутреннему; экран по нему подсвечивает нужное поле (field -- первый шаг)."""

    def __init__(self, message: str, status: int = 400, path: list | None = None):
        super().__init__(message)
        self.message = message
        self.status = status
        self.path = path

    def extra(self) -> dict:
        if not self.path:
            return {}
        return {"field": self.path[0], "path": list(self.path)}


_UNIQUE = re.compile(r"UNIQUE constraint failed: (.+)")
_NOT_NULL = re.compile(r"NOT NULL constraint failed: (.+)")
_COLUMN = re.compile(r"(?:\w+\.)?(\w+)$")


def _columns(text: str) -> list[str]:
    """«organization.inn, organization.kpp» -> ['inn', 'kpp']."""
    found = []
    for part in text.split(","):
        match = _COLUMN.search(part.strip())
        if match:
            found.append(match.group(1))
    return found


def _quoted(labels: dict, columns: list[str]) -> str | None:
    """Названия колонок по-русски, если известны ВСЕ; иначе None: английское
    имя колонки пользователю не показываем."""
    names = [labels.get(c) for c in columns]
    if not columns or not all(names):
        return None
    return " и ".join(f"«{n}»" for n in names)


def _missing_references(entity, obj, repo) -> list[str]:
    """Поля-ссылки записи, на которые в справочнике нет записи. Проверка идёт
    после отказа базы: сама база называет только «FOREIGN KEY constraint
    failed», какое поле -- не говорит."""
    missing = []
    for field, (label, getter) in entity.refs.items():
        value = getattr(obj, field, None)
        if value is None:
            continue
        try:
            getattr(repo, getter)(value)
        except LookupError:
            missing.append(field)
    return missing


def translate_integrity(exc: sqlite3.IntegrityError, entity=None, operation: str = "save",
                        obj=None, repo=None) -> ApiError:
    """Нарушение ограничения базы -> русский отказ.

    Сообщения SQLite английские и называют таблицу и колонку; пользователю они
    ничего не говорят. Поле называется по-русски там, где ограничение его
    выдаёт (UNIQUE, NOT NULL, CHECK) или где его можно выяснить (внешний ключ
    при сохранении); иначе -- прямая фраза «значение конфликтует с данными».
    entity -- описание справочника, откуда пришёл запрос (None -- неизвестно)."""
    text = str(exc)
    labels = entity.labels if entity else {}
    if "FOREIGN KEY" in text:
        if operation == "delete":
            return ApiError(entity.in_use if entity and entity.in_use else
                            "Эту запись нельзя удалить: она используется в других данных.")
        if entity and obj is not None and repo is not None:
            missing = _missing_references(entity, obj, repo)
            if missing:
                names = " и ".join(f"«{entity.refs[f][0]}»" for f in missing)
                return ApiError(
                    f"Не сохранено: в поле {names} указана запись, которой нет в "
                    "справочнике. Выберите существующую.", path=[missing[0]])
        return ApiError("Не сохранено: запись ссылается на данные, которых больше нет. "
                        "Обновите страницу и повторите.")
    if match := _UNIQUE.search(text):
        columns = _columns(match.group(1))
        names = _quoted(labels, columns)
        path = [columns[0]] if columns and columns[0] in labels else None
        if names:
            return ApiError(f"Не сохранено: значение в поле {names} уже есть в "
                            "справочнике, а оно не должно повторяться.", path=path)
        return ApiError("Не сохранено: такая запись уже есть.")
    if match := _NOT_NULL.search(text):
        columns = _columns(match.group(1))
        names = _quoted(labels, columns)
        path = [columns[0]] if columns and columns[0] in labels else None
        if names:
            return ApiError(f"Не сохранено: поле {names} обязательно для заполнения.",
                            path=path)
        return ApiError("Не сохранено: не заполнено обязательное поле.")
    if "CHECK" in text:
        for column, label in labels.items():
            if re.search(rf"\b{re.escape(column)}\b", text):
                return ApiError(f"Не сохранено: недопустимое значение в поле «{label}».",
                                path=[column])
        return ApiError("Не сохранено: значение одного из полей недопустимо.")
    return ApiError("Не сохранено: значение конфликтует с уже имеющимися данными.")


def _reply(message: str, status: int, extra: dict | None = None):
    body = {"error": message}
    if extra:
        body.update(extra)
    response = jsonify(body)
    response.status_code = status
    return response


def handle_exception(exc: Exception):
    """Единственный обработчик: всё, что вылетело из ручки, проходит здесь."""
    if isinstance(exc, ApiError):
        return _reply(exc.message, exc.status, exc.extra())
    if isinstance(exc, HTTPException):
        if exc.code is not None and exc.code < 400:
            return exc                                    # перенаправление
        return _reply(_HTTP_TEXT.get(exc.code, _HTTP_FALLBACK), exc.code or 500)
    if type(exc) is LookupError:
        return _reply(str(exc), 404)
    if isinstance(exc, InnNotFound):
        return _reply(str(exc), 400)
    if isinstance(exc, (CaptchaRequired, SourceUnavailable, InnLookupError)):
        return _reply(str(exc), 502)
    if isinstance(exc, (TargetTooLow, PackageError, TemplateError, CellMapError,
                        LayoutError, ContextError)):
        return _reply(str(exc), 400)
    if isinstance(exc, sqlite3.IntegrityError):
        error = translate_integrity(exc)
        return _reply(error.message, error.status)
    if isinstance(exc, sqlite3.OperationalError):
        current_app.logger.error("Ошибка базы данных: %s %s", request.method, request.path,
                                 exc_info=exc)
        return _reply("База данных сейчас занята или недоступна. Подождите несколько "
                      "секунд и повторите действие.", 503)
    if isinstance(exc, DecimalException):
        return _reply("Числа в расчёте слишком велики или не образуют допустимого "
                      "значения. Проверьте часы, ставки и цены.", 400)
    current_app.logger.error("Непредвиденная ошибка: %s %s", request.method, request.path,
                             exc_info=exc)
    return _reply(UNEXPECTED, 500)
