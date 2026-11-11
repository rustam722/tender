"""Ручки веба. Проверяют ввод, вызывают готовые функции, переводят отказы.

Арифметики и логики документов здесь нет: счёт делает app.calc, цепочку ставки
НДС и округлённые построчные суммы -- app.render.context (поэтому цифры экрана
совпадают с цифрами в документах), имена файлов и запись -- app.render.package.

Конвенция чисел в ответах -- как в контексте документов: голое имя -- готовая
строка (money: «5 985 000,00»), <поле>_value -- сырой Decimal строкой; ставки --
сырая доля под своим именем и <поле>_percent. См. app.web.texts.

Соединение с SQLite -- своё на каждый запрос (см. get_repo). Все ответы --
JSON, кроме главной страницы; все отказы -- {"error": "<по-русски>"}, см. errors.

Ручки:

    GET    /                                страница
    GET    /api/<справочник>                список (organizations?role=,
                                            experience?service_id=, profiles?customer_id=)
    GET    /api/<справочник>/<id>           одна запись
    POST   /api/<справочник>                создать; с "id" в теле -- изменить
    DELETE /api/<справочник>/<id>           удалить
        справочники: organizations employees assets experience services profiles tenders
    GET    /api/documents                   канонические имена документов пакета
    PUT    /api/tenders/<id>/items|labor    замена всего списка одной транзакцией
    GET/POST /api/settings                  настройки калькуляции
    GET/POST/DELETE /api/tenders/<id>/items|labor|costs[/<row_id>]
                                            позиции лота, трудозатраты, материалы и
                                            субподряд (costs?kind=materials|subcontract)
    GET    /api/tenders/<id>/conditions    условия сделки, которых требуют формы
                                            заказчика, с нынешними ответами
    GET    /api/inn/<inn>                   черновик карточки и предупреждения
    POST   /api/tenders/<id>/estimate       прямой счёт по сохранённым строкам
    POST   /api/tenders/<id>/solve          подбор часов под цену без НДС (не сохраняет)
    POST   /api/tenders/<id>/plan           какие файлы запишет сборка и какие уже есть
    POST   /api/tenders/<id>/build          сборка пакета
    POST   /api/tenders/<id>/preview        содержание одного документа до сборки
                                            (отказ по данным -- 200 с kind refusal)

Каждый POST обязан нести Content-Type: application/json: так чужая страница в
браузере не сможет «простым» запросом (без предварительной проверки) записать
данные в программу.
"""
import json
import re
import sqlite3
from dataclasses import replace
from decimal import Decimal
from pathlib import Path
from urllib.parse import urlsplit

from flask import current_app, g, jsonify, render_template, request

from app import config, db
from app.calc import (
    Estimate, LaborLine, estimate, money_total, price_gap, printed_labor,
    solve_hours,
)
from app.inn.lookup import build_draft
from app.render.context import (
    DOCUMENTS, build_context, labor_rows, money, required_conditions,
)
from app.render.package import build_package, planned_files
from app.render.preview import preview_document
from app.render.xlsx_render import CellMapError, load_cellmap
from app.repo import Repo
from app.web import errors, schema
from app.web.errors import ApiError
from app.web.texts import present, texts_of
from app.web.schema import (
    CHILDREN, COSTS, EXPERIENCE, LABOR, LOT_ITEMS, ORGANIZATIONS, PROFILES, SETTINGS, SOLVE, TENDERS, TOP_LEVEL,
    Entity, plain,
)

_LOCAL_HOSTS = {"127.0.0.1", "localhost", "::1"}
_INN = re.compile(r"[0-9]{10}|[0-9]{12}")
MAX_BODY = 2 * 1024 * 1024


# ---- соединение с базой -------------------------------------------------------

def get_repo() -> Repo:
    """Репозиторий этого запроса. Соединение открывается при первом обращении и
    закрывается по окончании запроса (teardown).

    Почему своё соединение на запрос, а не одно общее: сервер Flask обрабатывает
    запросы в разных потоках, а соединение SQLite нельзя использовать из потока,
    в котором оно не создано (ProgrammingError -- и только при настоящей работе,
    когда два запроса идут одновременно). Общее соединение с check_same_thread=False
    потребовало бы замка вокруг каждого обращения и общей транзакции на всех:
    чужой запрос, не дошедший до commit, оказался бы в нашей. Свежее соединение
    стоит доли миллисекунды -- для одного пользователя это ничто, а PRAGMA
    foreign_keys, включённый в db.connect, действует на каждое из них."""
    if "repo" not in g:
        g.repo = Repo(db.connect(current_app.config["DB_PATH"]))
    return g.repo


def _close_repo(_exc):
    repo = g.pop("repo", None)
    if repo is not None:
        repo.conn.close()


# ---- защита локального сервера ------------------------------------------------

def _guard():
    """Сервер слушает только 127.0.0.1, но браузер пользователя -- общий. Чужая
    страница может отправить запрос на 127.0.0.1:8765 (подмена DNS меняет только
    имя узла, форма-POST шлёт «простой» запрос без проверки). Поэтому:
    имя узла -- только локальное, источник записи (Origin) -- только эта же
    страница, тело POST -- только JSON (такой запрос чужой странице отправить
    нельзя без разрешения сервера, а мы его не даём)."""
    host = urlsplit("//" + request.host).hostname
    if host not in _LOCAL_HOSTS:
        raise ApiError("Запрос отклонён: программа отвечает только на локальном адресе "
                       "(127.0.0.1).", 403)
    if request.method in ("POST", "PUT", "PATCH", "DELETE"):
        origin = request.headers.get("Origin")
        if origin is not None:
            parts = urlsplit(origin)
            if parts.scheme != "http" or parts.netloc.lower() != request.host.lower():
                raise ApiError("Запрос отклонён: он отправлен не со страницы программы.", 403)
    if request.method in ("POST", "PUT", "PATCH") and not request.is_json:
        raise ApiError("Тело запроса должно быть в формате JSON.", 415)


def _protect(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    if request.path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store"
    return response


# ---- разбор запроса -----------------------------------------------------------

def _reject_constant(name):
    raise ValueError(name)


def json_body() -> dict:
    """Тело POST как словарь; пустое тело -- {}. Числа с точкой читаются сразу
    как Decimal, минуя float; NaN и Infinity отвергаются."""
    raw = request.get_data()
    if not raw.strip():
        return {}
    try:
        data = json.loads(raw.decode("utf-8"), parse_float=Decimal,
                          parse_constant=_reject_constant)
    except (ValueError, RecursionError):
        raise ApiError("Тело запроса не разобрано: ожидается корректный JSON.") from None
    if not isinstance(data, dict):
        raise ApiError("Тело запроса должно быть JSON-объектом: {\"поле\": значение}.")
    return data


def _only(data: dict, allowed: tuple, what: str):
    for key in data:
        if key not in allowed:
            raise ApiError(f"Неизвестное поле «{key}» в запросе «{what}».", path=[key])


def _record_id(entity: Entity, value) -> int:
    """Номер записи из адреса или тела. Больше, чем хранит SQLite, -- такой
    записи заведомо нет (и OverflowError наружу не уходит)."""
    if isinstance(value, bool) or not isinstance(value, int):
        raise ApiError("Поле «id» должно быть целым числом.", path=["id"])
    if value > schema.MAX_INT or value < 1:
        raise ApiError(f"{entity.title}: запись с id={value} не найдена", 404)
    return value


def _query_int(name: str) -> int | None:
    raw = request.args.get(name)
    if raw is None:
        return None
    if not re.fullmatch(r"[0-9]{1,18}", raw):
        raise ApiError(f"Параметр «{name}» должен быть целым числом.", path=[name])
    return int(raw)


def _not_found(entity: Entity, row_id: int):
    return ApiError(f"{entity.title}: запись с id={row_id} не найдена", 404)


# ---- чтение и запись справочников ---------------------------------------------

def _find(entity: Entity, repo: Repo, row_id: int, tender_id: int | None = None):
    """Запись по номеру; строка тендера -- только этого тендера."""
    row_id = _record_id(entity, row_id)
    if entity.get:
        found = getattr(repo, entity.get)(row_id)
        if tender_id is not None and found.tender_id != tender_id:
            raise _not_found(entity, row_id)
        return found
    for row in getattr(repo, entity.list_)(tender_id):
        if row.id == row_id:
            return row
    raise _not_found(entity, row_id)


def _labor_view(rows) -> list[dict]:
    """Строки ФОТ (сохранённые или подобранные) в одной форме: сырые значения,
    готовые строки и сумма строки -- её считает контекст документов, как для
    печатной калькуляции."""
    view = []
    for row, printed in zip(rows, labor_rows(rows)):
        body = present(LABOR, row)
        body["amount"], body["amount_value"] = printed["amount"], plain(printed["amount_value"])
        view.append(body)
    return view


def _child_view(entity: Entity, repo: Repo, tender_id: int, kind: str | None = None) -> dict:
    """Строки тендера (позиции лота, ФОТ, материалы и субподряд) и итог под ними.

    Суммы строк и итоги -- не наши: суммы ФОТ и материалов берутся из контекста
    документов (округление и порядок те же, что в печатной форме, итоги ФОТ,
    материалов и субподряда -- статьи расчёта), итог позиций лота --
    money_total. Экран только показывает."""
    ctx = build_context(tender_id, repo)
    if entity is LOT_ITEMS:
        # НДС позиции -- как в контексте документов (округлён на строке); итог НДС --
        # сумма напечатанных строк, как складывает колонку форма заказчика
        items = repo.list_lot_items(tender_id)
        rows = []
        for item, printed in zip(items, ctx["items"]):
            body = present(entity, item)
            body["vat"], body["vat_value"] = printed["vat"], plain(printed["vat_value"])
            rows.append(body)
        total = money_total(i.price for i in items)
        vat_total = money_total(printed["vat_value"] for printed in ctx["items"])
        return {"rows": rows, "total": money(total), "total_value": plain(total),
                "vat_total": money(vat_total), "vat_total_value": plain(vat_total),
                "vat_percent": ctx["tender"]["vat_percent"]}
    if entity is LABOR:
        total = ctx["estimate"]["labor"]
        return {"rows": _labor_view(repo.list_tender_labor(tender_id)),
                "total": money(total), "total_value": plain(total)}
    # материалы и субподряд: строки обеих статей, итоги каждой
    if kind is not None and kind not in COSTS.choices["kind"]:
        raise ApiError("Параметр «kind»: допустимо materials или subcontract.", path=["kind"])
    printed = {}
    for name in COSTS.choices["kind"]:
        for row, line in zip(repo.list_tender_cost_items(tender_id, name), ctx[name]):
            printed[row.id] = line
    rows = []
    for row in repo.list_tender_cost_items(tender_id, kind):
        body = present(entity, row)
        line = printed[row.id]
        body["amount"], body["amount_value"] = line["amount"], plain(line["amount_value"])
        rows.append(body)
    totals = {name: ctx["estimate"][name] for name in COSTS.choices["kind"]}
    return {"rows": rows, "totals": texts_of(totals), "totals_value": plain(totals)}


def _list(entity: Entity, repo: Repo):
    if entity is ORGANIZATIONS:
        role = request.args.get("role")
        if role is not None and role not in ORGANIZATIONS.choices["role"]:
            raise ApiError("Параметр «role»: допустимо participant или customer.",
                           path=["role"])
        return repo.list_organizations(role)
    if entity is EXPERIENCE:
        return repo.list_experience(_query_int("service_id"))
    if entity is PROFILES:
        customer_id = _query_int("customer_id")
        if customer_id is None:
            return repo.list_customer_profiles()
        found = repo.find_customer_profile(customer_id)
        return [found] if found else []
    return getattr(repo, entity.list_)()


def _save(entity: Entity, tender_id: int | None = None):
    repo = get_repo()
    data = json_body()
    raw_id = data.pop("id", None)
    existing = _find(entity, repo, raw_id, tender_id) if raw_id is not None else None
    fixed = None
    if tender_id is not None:
        sent = data.pop("tender_id", None)
        if sent is not None and sent != tender_id:
            raise ApiError("Строка относится к тому тендеру, в адресе которого она "
                           "сохраняется; поле «tender_id» менять нельзя.",
                           path=["tender_id"])
        fixed = {"tender_id": tender_id}
    obj = schema.parse_object(entity, data, existing, fixed)
    if entity.check:
        entity.check(obj, repo)
    try:
        saved_id = getattr(repo, entity.save)(obj)
    except sqlite3.IntegrityError as exc:
        raise errors.translate_integrity(exc, entity, "save", obj, repo) from exc
    saved = _find(entity, repo, saved_id, tender_id)
    if tender_id is None:
        body = present(entity, saved)
        warning = _notice_after_write(entity, repo, saved)
        if warning:
            body["warning"] = warning
    else:       # у строки есть сумма или НДС: их знает только расчёт
        body = _row_after_write(entity, repo, saved, tender_id)
    response = jsonify(body)
    response.status_code = 200 if existing is not None else 201
    return response


def _notice_after_write(entity: Entity, repo: Repo, saved) -> str | None:
    """Предупреждение к уже сохранённой записи (у тендера -- «такой номер закупки уже
    есть»). Запись зафиксирована, поэтому сбой при его поиске -- не ошибка запроса:
    человек повторил бы сохранение и получил второй тендер. Без предупреждения -- не страшно."""
    if entity.notice is None:
        return None
    try:
        return entity.notice(saved, repo)
    except Exception:
        current_app.logger.exception("Не удалось проверить запись после сохранения")
        return None


WRITTEN_BUT_NOT_RECALCULATED = ("Запись сохранена, но пересчитать суммы не удалось. "
                                "Обновите страницу, чтобы увидеть их.")


def _row_after_write(entity: Entity, repo: Repo, saved, tender_id: int) -> dict:
    """Ответ на сохранение строки тендера. Запись уже зафиксирована, поэтому сбой при
    построении вида (суммы считает контекст документов) -- не ошибка запроса: иначе
    человек повторит сохранение и получит дубль. Отвечаем успехом с предупреждением."""
    try:
        return next(row for row in _child_view(entity, repo, tender_id)["rows"]
                    if row["id"] == saved.id)
    except Exception:
        current_app.logger.exception("Не удалось построить вид сохранённой строки")
        return {**present(entity, saved), "warning": WRITTEN_BUT_NOT_RECALCULATED}


def _delete(entity: Entity, row_id: int, tender_id: int | None = None):
    repo = get_repo()
    found = _find(entity, repo, row_id, tender_id)
    try:
        getattr(repo, entity.delete)(found.id)
    except sqlite3.IntegrityError as exc:
        raise errors.translate_integrity(exc, entity, "delete") from exc
    return jsonify({"deleted": found.id})


def _tender(tender_id: int):
    """Тендер или 404; проверка нужна и строкам тендера, и счёту."""
    return _find(TENDERS, get_repo(), tender_id)


# ---- сборка -------------------------------------------------------------------

_PACKAGE_FIELDS = ("documents", "output_dir", "overwrite", "experience_ids")


def _document_name(value, label: str, where: list) -> str:
    """Имя документа из тела. Это ключ шаблонов («Справка о МТР»), а не путь:
    разделители, диск и управляющие знаки в имени -- отказ: шаблон ищется как
    `<имя>.docx` в папке шаблонов, и имя с «../» читало бы вне неё."""
    if not isinstance(value, str):
        raise ApiError(f"{label}: название документа должно быть текстом.", path=where)
    if (not value.strip() or any(ch in value for ch in "/\\:")
            or any(ord(ch) < 32 for ch in value)):
        raise ApiError(f"{label}: название документа пустое или содержит "
                       "знаки пути (/ \\ :) или управляющие знаки.", path=where)
    return value.strip()


def _document_names(value) -> list[str]:
    """Имена документов состава пакета из тела."""
    if not isinstance(value, list):
        raise ApiError("Поле «Состав пакета»: должно быть списком названий документов.",
                       path=["documents"])
    return [_document_name(name, f"Состав пакета, №{i + 1}", ["documents", i])
            for i, name in enumerate(value)]


def _package_request(tender_id: int, data: dict):
    """Общая часть планирования и сборки: что собирать и куда."""
    repo = get_repo()
    tender = _tender(tender_id)
    _only(data, _PACKAGE_FIELDS, "сборка пакета")

    if data.get("documents") is not None:
        documents = _document_names(data["documents"])
    else:
        profile = (repo.find_customer_profile(tender.customer_id)
                   if tender.customer_id is not None else None)
        documents = list(profile.documents) if profile else []
    if not documents:
        raise ApiError("Не выбрано ни одного документа: отметьте документы пакета "
                       "или задайте состав пакета в профиле заказчика.",
                       path=["documents"])

    folder = data.get("output_dir")
    if folder is None or folder == "":
        folder = tender.output_dir
    if not folder:
        raise ApiError("Не указана папка, в которую собирать пакет.", path=["output_dir"])
    out_dir = Path(schema.check_out_dir(folder, TENDERS, ["output_dir"]))

    experience_ids = data.get("experience_ids")
    if experience_ids is not None:
        experience_ids = schema.coerce(experience_ids, list[int], TENDERS, ["experience_ids"])
    overwrite = data.get("overwrite", False)
    if not isinstance(overwrite, bool):
        raise ApiError("Поле «overwrite» должно быть true или false.", path=["overwrite"])
    return repo, documents, out_dir, overwrite, experience_ids


def _conditions(tender_id: int):
    """Условия сделки, ответы на которые требуют формы заказчика этого тендера.

    Семь условий формы Роснефти лежат в её собственных клетках, ключи дословные
    (у шестерых двоеточие на конце, у седьмого нет), и угадывать их человек не
    должен: экран показывает список готовым, с нынешними ответами. Умолчания
    нет -- без ответа документ не собирается, и об этом говорит сборка.

    Карту, которую не разобрать, не проглатываем: она попадает в `warnings`
    фразой по-русски. Сборка этого документа всё равно отказала бы, и молчать
    здесь значило бы показать человеку список без одной формы.
    """
    repo = get_repo()
    tender = _tender(tender_id)
    profile = (repo.find_customer_profile(tender.customer_id)
               if tender.customer_id is not None else None)
    answers = tender.deal_answers
    rows: dict[str, list[str]] = {}         # условие -> документы, что его требуют
    warnings: list[str] = []
    for document, spec in (profile.docs.items() if profile else ()):
        if not spec.cellmap:
            continue
        try:
            cell_map = load_cellmap(config.CELLMAPS_DIR / spec.cellmap)
        except CellMapError as exc:
            warnings.append(
                f"Условия документа «{document}» прочитать не удалось: {exc}")
            continue
        for name in required_conditions(cell_map):
            rows.setdefault(name, []).append(document)
    return jsonify({
        "rows": [{"condition": name, "answer": answers.get(name, ""),
                  "documents": documents}
                 for name, documents in rows.items()],
        "warnings": warnings})


def _preview(tender_id: int):
    """Содержание одного документа по нынешним данным тендера.

    Отказ по данным («документ собирается по форме заказчика, форма не
    назначена в профиле») отдаётся кодом 200 с `kind: "refusal"`, а не 400:
    запрос верен, неверны данные тендера, и эта фраза -- содержимое
    предпросмотра, а не ошибка обращения. 400 здесь остаётся за неверным
    запросом: нет имени документа, имя не текст, имя со знаками пути.
    """
    repo = get_repo()
    _tender(tender_id)
    data = json_body()
    _only(data, ("document",), "предпросмотр документа")
    if "document" not in data:
        raise ApiError("Не указан документ, содержание которого показать.",
                       path=["document"])
    document = _document_name(data["document"], "Предпросмотр", ["document"])
    result = preview_document(tender_id, document, repo)
    return jsonify({"document": document, "kind": result.kind,
                    "blocks": result.blocks, "warnings": result.warnings,
                    "refusal": result.refusal})


def _plan(tender_id: int):
    repo, documents, out_dir, _, _ = _package_request(tender_id, json_body())
    files = planned_files(tender_id, documents, out_dir, repo)
    return jsonify({"output_dir": str(out_dir), "files": [
        {"document": document, "name": path.name, "path": str(path), "exists": path.exists()}
        for document, path in files]})


def _build(tender_id: int):
    repo, documents, out_dir, overwrite, experience_ids = _package_request(
        tender_id, json_body())
    result = build_package(tender_id, documents, out_dir, repo, overwrite=overwrite,
                           experience_ids=experience_ids)
    return jsonify({
        "output_dir": str(out_dir),
        "written": [{"name": path.name, "path": str(path)} for path in result.written],
        "skipped": [{"document": d, "reason": r} for d, r in result.skipped],
        "warnings": [{"document": d, "text": t} for d, t in result.warnings]})


# ---- расчёт -------------------------------------------------------------------

def _target(data: dict, required: bool):
    """Целевая цена и признак «включает НДС» из тела. Приведение цены и разницу
    считает app.calc; здесь только проверка ввода."""
    includes_vat = data.get("target_includes_vat", False)
    if not isinstance(includes_vat, bool):
        raise ApiError("Поле «Цена включает НДС» должно быть true или false.",
                       path=["target_includes_vat"])
    if "target" not in data:
        if required or includes_vat:
            raise ApiError("Не указана целевая цена.", path=["target"])
        return None, False
    target = schema.decimal_value(data["target"], SOLVE, ["target"])
    if target == 0:
        raise ApiError("Поле «Цена»: цена должна быть больше нуля.", path=["target"])
    return target, includes_vat


def _gap_fields(target: Decimal, includes_vat: bool, result: Estimate) -> dict:
    """Цель и расхождение итога с ней -- числом (строкой Decimal) и готовой
    строкой. Разницу считает calc.price_gap по итогу расчёта, не по статьям."""
    gap = price_gap(target, result, includes_vat)
    return {"target": money(target), "target_value": plain(target),
            "target_includes_vat": includes_vat,
            "difference": money(gap), "difference_value": plain(gap)}


def _estimate(tender_id: int):
    """Прямой счёт. Берётся из контекста документов, а не считается заново:
    ставка НДС выбирается по цепочке тендер -> организация -> настройки, суммы
    строк округляются так же, как в документе, -- на экране те же цифры, что
    будут напечатаны."""
    data = json_body()
    _only(data, ("target", "target_includes_vat"), "прямой счёт")
    target, includes_vat = _target(data, required=False)
    _tender(tender_id)
    ctx = build_context(tender_id, get_repo())
    lines = {name: [row["amount_value"] for row in ctx[name]]
             for name in ("labor", "materials", "subcontract")}
    body = {
        "vat_rate": plain(ctx["tender"]["vat_rate"]),
        "vat_percent": ctx["tender"]["vat_percent"],
        "estimate": texts_of(ctx["estimate"]),
        "estimate_value": plain(ctx["estimate"]),
        "lines": {name: [money(v) for v in values] for name, values in lines.items()},
        "lines_value": plain(lines)}
    if target is not None:
        body.update(_gap_fields(target, includes_vat, Estimate(**ctx["estimate"])))
    return jsonify(body)


def _solve(tender_id: int):
    """Подбор часов под цену без НДС. Результат -- предложение: в базе часы не
    меняются, пользователь применяет их сам (POST строк трудозатрат).

    ФОТ считается по напечатанным строкам -- так же, как в документах и в
    ручке /estimate: доктрина напечатанной строки одна на весь проект."""
    data = json_body()
    _only(data, ("target", "target_includes_vat"), "подбор часов")
    target, includes_vat = _target(data, required=True)
    repo = get_repo()
    _tender(tender_id)
    ctx = build_context(tender_id, repo)
    settings = replace(repo.get_settings(), vat_rate=ctx["tender"]["vat_rate"])
    rows = repo.list_tender_labor(tender_id)
    lines = [LaborLine(r.post, r.headcount, r.hours, r.rate) for r in rows]
    materials, subcontract = ctx["estimate"]["materials"], ctx["estimate"]["subcontract"]
    solved = solve_hours(target, lines, materials, subcontract, settings,
                         target_includes_vat=includes_vat)
    # Итог -- по напечатанным строкам, как его считают документы и экран
    # калькуляции (context._printed). По сырым строкам подбор показывал
    # расхождение в копейках там, где документ попадал в цель ровно.
    result = estimate(printed_labor(solved), materials, subcontract, settings)
    body = {
        # та же форма, что у сохранённых строк ФОТ: с суммой строки
        "labor": _labor_view([replace(row, hours=line.hours)
                              for row, line in zip(rows, solved)]),
        "estimate": texts_of(result),
        "estimate_value": plain(result),
        # ставка, по которой считан подбор: экран не должен брать её из соседнего ответа
        "vat_rate": plain(ctx["tender"]["vat_rate"]),
        "vat_percent": ctx["tender"]["vat_percent"]}
    body.update(_gap_fields(target, includes_vat, result))
    return jsonify(body)


MAX_LOT_ITEMS = 2000


def _replace_rows(tender_id: int, entity: Entity, body_key: str, noun: str, replacer: str,
                  what: str):
    """Замена всего списка строк тендера присланным -- одной транзакцией (вставка
    списка лота из буфера, применение подобранных часов). Не прошла проверка
    хотя бы одной строки -- ничего не записано, а в ответе путь [body_key, номер,
    поле]. id и tender_id внутри строк допускаются (список можно отправить обратно
    как получен) и игнорируются: id назначаются заново, прежние номера строк
    недействительны. Записанный список уже зафиксирован, поэтому сбой при
    построении вида -- не ошибка запроса, а успех с предупреждением."""
    repo = get_repo()
    _tender(tender_id)
    data = json_body()
    _only(data, (body_key,), what)
    raw = data.get(body_key)
    if not isinstance(raw, list):
        raise ApiError(f"Поле «{body_key}»: должно быть списком.", path=[body_key])
    if len(raw) > MAX_LOT_ITEMS:
        raise ApiError(f"Слишком много строк: не больше {MAX_LOT_ITEMS} за раз.",
                       path=[body_key])
    parsed = []
    for i, row in enumerate(raw):
        where = [body_key, i]
        if not isinstance(row, dict):
            raise ApiError(f"{noun} №{i + 1}: должна быть набором полей.", path=where)
        row = dict(row)
        row.pop("id", None)
        sent = row.pop("tender_id", None)
        if sent is not None and sent != tender_id:
            raise ApiError(f"{noun} №{i + 1}: поле «tender_id» не совпадает с тендером "
                           "в адресе.", path=where + ["tender_id"])
        try:
            parsed.append(schema.parse_object(entity, row, None, {"tender_id": tender_id}))
        except ApiError as exc:
            raise ApiError(f"{noun} №{i + 1}. {exc.message}", exc.status,
                           where + list(exc.path or [])) from None
    try:
        getattr(repo, replacer)(tender_id, parsed)
    except sqlite3.IntegrityError as exc:
        raise errors.translate_integrity(exc, entity, "save") from exc
    try:
        return jsonify(_child_view(entity, repo, tender_id))
    except Exception:       # список уже заменён: отвечаем успехом, а не ошибкой
        current_app.logger.exception("Не удалось построить вид заменённого списка")
        saved = getattr(repo, entity.list_)(tender_id)
        return jsonify({"rows": [present(entity, r) for r in saved],
                        "warning": WRITTEN_BUT_NOT_RECALCULATED})


def _replace_items(tender_id: int):
    """Вставка списка лота из буфера."""
    return _replace_rows(tender_id, LOT_ITEMS, "items", "Позиция", "replace_lot_items",
                         "замена позиций лота")


def _replace_labor(tender_id: int):
    """Применение подобранных часов: весь ФОТ заменяется разом."""
    return _replace_rows(tender_id, LABOR, "rows", "Строка", "replace_tender_labor",
                         "замена строк трудозатрат")


def _documents():
    """Канонические имена документов пакета: профиль заказчика отмечает и
    упорядочивает нужные, но не заменяет список."""
    return jsonify(list(DOCUMENTS))


def _backup():
    """Копия базы по просьбе человека: перед правкой справочников, перед
    отправкой заявки -- когда он сам решил, что момент важный.

    Автоматическая копия снимается при каждом запуске (app/web/__main__.py),
    но между запусками программа может работать неделями, и за это время
    справочники успевают измениться. Ответ называет полный путь файла: копия,
    о которой человек не знает, где лежит, не копия.
    """
    try:
        copy = db.backup(current_app.config["DB_PATH"])
    except (OSError, sqlite3.Error) as exc:
        raise ApiError(
            f"Копия базы не сохранена: {exc}. Проверьте, есть ли место на диске "
            "и права на запись в папку рядом с базой.", status=500) from exc
    if copy is None:
        raise ApiError("Копировать нечего: базы данных ещё нет. Она появится "
                       "после первой записи в справочники.")
    return jsonify({"path": str(copy), "name": copy.name})


def _inn(inn: str):
    if not _INN.fullmatch(inn):
        raise ApiError("ИНН должен состоять из 10 цифр (организация) или из 12 "
                       "(предприниматель).", path=["inn"])
    url = request.args.get("rusprofile_url") or None
    if url is not None and len(url) > 500:
        raise ApiError("Ссылка на rusprofile слишком длинная.", path=["rusprofile_url"])
    organization, warnings = build_draft(inn, rusprofile_url=url)
    return jsonify({"organization": present(ORGANIZATIONS, organization),
                    "warnings": warnings})


# ---- настройки ----------------------------------------------------------------

def _settings_save():
    repo = get_repo()
    settings = schema.parse_object(SETTINGS, json_body(), repo.get_settings())
    repo.save_settings(settings)
    return jsonify(present(SETTINGS, repo.get_settings()))


# ---- регистрация --------------------------------------------------------------

def register(app) -> None:
    app.config["MAX_CONTENT_LENGTH"] = MAX_BODY
    app.json.ensure_ascii = False
    app.json.sort_keys = False
    app.before_request(_guard)
    app.after_request(_protect)
    app.teardown_appcontext(_close_repo)
    app.register_error_handler(Exception, errors.handle_exception)

    app.add_url_rule("/", "index", lambda: render_template("index.html"))
    app.add_url_rule("/api/inn/<inn>", "inn", _inn, methods=["GET"])
    app.add_url_rule("/api/backup", "backup", _backup, methods=["POST"])
    app.add_url_rule("/api/tenders/<int:tender_id>/items", "items_replace", _replace_items,
                     methods=["PUT"])
    app.add_url_rule("/api/tenders/<int:tender_id>/labor", "labor_replace", _replace_labor,
                     methods=["PUT"])
    app.add_url_rule("/api/documents", "documents", _documents, methods=["GET"])
    app.add_url_rule("/api/settings", "settings_get",
                     lambda: jsonify(present(SETTINGS, get_repo().get_settings())),
                     methods=["GET"])
    app.add_url_rule("/api/settings", "settings_save", _settings_save, methods=["POST"])

    for ent in TOP_LEVEL:
        base = f"/api/{ent.path}"
        app.add_url_rule(base, f"{ent.path}_list",
                         lambda ent=ent: jsonify([present(ent, o) for o in _list(ent, get_repo())]),
                         methods=["GET"])
        app.add_url_rule(base, f"{ent.path}_save", lambda ent=ent: _save(ent),
                         methods=["POST"])
        app.add_url_rule(f"{base}/<int:row_id>", f"{ent.path}_get",
                         lambda row_id, ent=ent: jsonify(present(ent, _find(ent, get_repo(), row_id))),
                         methods=["GET"])
        app.add_url_rule(f"{base}/<int:row_id>", f"{ent.path}_delete",
                         lambda row_id, ent=ent: _delete(ent, row_id), methods=["DELETE"])

    for ent in CHILDREN:
        base = f"/api/tenders/<int:tender_id>/{ent.path}"

        def listing(tender_id, ent=ent):
            _tender(tender_id)
            return jsonify(_child_view(ent, get_repo(), tender_id, request.args.get("kind")))

        def saving(tender_id, ent=ent):
            _tender(tender_id)
            return _save(ent, tender_id)

        def deleting(tender_id, row_id, ent=ent):
            _tender(tender_id)
            return _delete(ent, row_id, tender_id)

        app.add_url_rule(base, f"{ent.path}_list", listing, methods=["GET"])
        app.add_url_rule(base, f"{ent.path}_save", saving, methods=["POST"])
        app.add_url_rule(f"{base}/<int:row_id>", f"{ent.path}_delete", deleting,
                         methods=["DELETE"])

    app.add_url_rule("/api/tenders/<int:tender_id>/conditions", "conditions",
                     _conditions, methods=["GET"])
    for name, view in (("estimate", _estimate), ("solve", _solve), ("plan", _plan),
                       ("build", _build), ("preview", _preview)):
        app.add_url_rule(f"/api/tenders/<int:tender_id>/{name}", name, view, methods=["POST"])
