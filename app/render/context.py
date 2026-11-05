"""Контекст шаблонов: один словарь, из которого читают все документы.

Читает данные через Repo и раскладывает их по ключам; документы не пишет.
Правил, кроме цепочки НДС и форматирования денег, здесь нет.

Ключи (значения отдаются как есть, Decimal остаётся Decimal -- форматирует
шаблон через `money`, кроме клеток таблиц, см. ниже):
    org, customer      поля организации (пустые строки, если организации нет)
    tender             поля тендера; vat_rate -- итоговая ставка по цепочке,
                       vat_percent -- она же текстом: «5», «5,5»
    items              позиции лота: number, position, name, price (текст
                       через money), price_value (Decimal -- для ячейки Excel)
    estimate           статьи калькуляции, Decimal
    materials,         строки материалов и субподряда: number, name, unit,
    subcontract        quantity (Decimal), price и amount (текст через money),
                       price_value и amount_value (Decimal), note; их суммы
                       идут в estimate
    labor              строки ФОТ: number, post, headcount, hours, rate
    employees, assets, experience
                       {"headers": [...], "rows": [[...]]} под раскладку
    declaration_text   текст декларации из профиля заказчика
    conditions         ответы по условиям сделки: условие -> ответ
    today              дата сборки, «ДД.ММ.ГГГГ»
    money              функция: Decimal -> «5 985 000,00»
    date_ru, date_long функции: ISO-строка -> «24.10.2025» и «24» октября 2025 г.;
                       пустая строка -> пустая строка

Деньги в клетках таблиц. Раскладки (задача 7) ничего не форматируют, поэтому
перед build_rows каждое Decimal-поле записи заменяется своим текстом через
money -- в копии записи, справочник не трогается. Так формат не зависит от
того, какую колонку выбрала раскладка: форматируется любое денежное поле
(сумма контракта, часовая ставка), а не список имён, который пришлось бы
держать в согласии с правкой раскладок.
"""
from dataclasses import asdict, fields, replace
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from app.calc import LaborLine, estimate
from app.config import BASE_DIR
from app.models import (
    COST_MATERIALS, COST_SUBCONTRACT, ROLE_CUSTOMER, ROLE_PARTICIPANT,
    CustomerProfile, Organization, TenderCostItem,
)
from app.render.layouts import TableLayout, build_rows, load_layout
from app.repo import Repo

LAYOUTS_DIR = BASE_DIR / "layouts"

NBSP = " "
_KOPECK = Decimal("0.01")
_ZERO = Decimal("0")

# Раскладка по умолчанию -- для заказчика без профиля или с профилем, где
# нет документа, который выбирает другую раскладку.
_DEFAULT_LAYOUTS = {
    "employees": "employees_koryazhma.json",
    "assets": "assets_koryazhma.json",
    "experience": "experience_default.json",
}

_MONTHS_GENITIVE = ("января", "февраля", "марта", "апреля", "мая", "июня",
                    "июля", "августа", "сентября", "октября", "ноября", "декабря")


def money(value: Decimal) -> str:
    """5985000 -> «5 985 000,00»: разряды через неразрывный пробел U+00A0
    (обычный позволил бы Word разорвать число переносом), запятая -- десятичная.
    Округление до копеек ROUND_HALF_UP, как в калькуляторе."""
    if not isinstance(value, (Decimal, int)) or isinstance(value, bool):
        raise TypeError(f"money принимает Decimal, а не {type(value).__name__}")
    rounded = Decimal(value).quantize(_KOPECK, rounding=ROUND_HALF_UP)
    if not rounded:
        rounded = Decimal("0.00")        # без «-0,00»
    return f"{rounded:,.2f}".translate({ord(","): NBSP, ord("."): ","})


def _parse_date(value: str) -> date | None:
    """ISO-строка -> date; пустая строка -> None; прочее -- ValueError по-русски."""
    if not value:
        return None
    try:
        return date.fromisoformat(value)
    except ValueError:
        raise ValueError(
            f"Дата «{value}» не распознана: ожидается ГГГГ-ММ-ДД.") from None


def date_ru(value: str) -> str:
    """2025-10-24 -> 24.10.2025."""
    parsed = _parse_date(value)
    return parsed.strftime("%d.%m.%Y") if parsed else ""


def date_long(value: str) -> str:
    """2025-10-24 -> «24» октября 2025 г.; день без ведущего нуля."""
    parsed = _parse_date(value)
    if not parsed:
        return ""
    return f"«{parsed.day}» {_MONTHS_GENITIVE[parsed.month - 1]} {parsed.year} г."


def _today() -> date:
    return date.today()


def _percent(rate: Decimal) -> str:
    """0.05 -> «5», 0.055 -> «5,5»."""
    return format((rate * 100).normalize(), "f").replace(".", ",")


def _vat_rate(tender, participant: Organization | None, settings) -> Decimal:
    """Тендер -> участник -> настройки; первое значение, не равное None.
    Ноль -- осознанное «без НДС», в следующее звено не проваливается."""
    for rate in (tender.vat_rate,
                 participant.vat_rate if participant else None):
        if rate is not None:
            return rate
    return settings.vat_rate


def _layouts(profile: CustomerProfile | None) -> dict[str, TableLayout]:
    """Раскладка на каждый источник. Профиль называет файлы явно (документ ->
    файл), а какому источнику файл служит, говорит он сам. Первая раскладка
    источника в профиле побеждает; остальное -- умолчания."""
    layouts = {}
    for name in (profile.layouts.values() if profile else ()):
        layout = load_layout(Path(LAYOUTS_DIR) / name)
        layouts.setdefault(layout.source, layout)
    for source, name in _DEFAULT_LAYOUTS.items():
        if source not in layouts:
            layouts[source] = load_layout(Path(LAYOUTS_DIR) / name)
    return layouts


def _table(layout: TableLayout, records: list) -> dict:
    """Шапка из раскладки, строки данных; Decimal-поля -- через money."""
    shown = [replace(r, **{f.name: money(getattr(r, f.name))
                           for f in fields(r)
                           if isinstance(getattr(r, f.name), Decimal)})
             for r in records]
    return {"headers": [c.header for c in layout.columns],
            "rows": build_rows(layout, shown)}


def _organization(repo: Repo, org_id: int | None, role: str):
    if org_id is None:
        return None, asdict(Organization(role=role))
    org = repo.get_organization(org_id)
    return org, asdict(org)


def _cost_rows(items: list[TenderCostItem]) -> list[dict]:
    return [{"number": n, "name": i.name, "unit": i.unit, "quantity": i.quantity,
             "price": money(i.price), "price_value": i.price,
             "amount": money(i.quantity * i.price),
             "amount_value": i.quantity * i.price, "note": i.note}
            for n, i in enumerate(items, start=1)]


def build_context(tender_id: int, repo: Repo, *,
                  experience_ids: list[int] | tuple[int, ...] | None = None) -> dict:
    """Собирает контекст тендера. Неизвестный id -- LookupError по-русски.

    experience_ids -- записи опыта для справки, в нужном порядке. None --
    взять отбор, сохранённый при тендере (повторная сборка даёт тот же
    результат); переданный список, даже пустой, побеждает и ничего не
    сохраняет -- так делается предпросмотр с другим отбором.
    """
    tender = repo.get_tender(tender_id)
    participant, org = _organization(repo, tender.participant_id,
                                     ROLE_PARTICIPANT)
    _, customer = _organization(repo, tender.customer_id, ROLE_CUSTOMER)
    profile = (repo.find_customer_profile(tender.customer_id)
               if tender.customer_id is not None else None)

    settings = repo.get_settings()
    rate = _vat_rate(tender, participant, settings)
    labor_rows = repo.list_tender_labor(tender_id)
    materials = repo.list_tender_cost_items(tender_id, COST_MATERIALS)
    subcontract = repo.list_tender_cost_items(tender_id, COST_SUBCONTRACT)
    calc = estimate(
        [LaborLine(r.post, r.headcount, r.hours, r.rate) for r in labor_rows],
        sum((i.quantity * i.price for i in materials), _ZERO),
        sum((i.quantity * i.price for i in subcontract), _ZERO),
        replace(settings, vat_rate=rate))

    tender_fields = asdict(tender)
    tender_fields["vat_rate"] = rate
    tender_fields["vat_percent"] = _percent(rate)

    layouts = _layouts(profile)
    if experience_ids is None:
        experience_ids = tender.experience_ids
    selected = [repo.get_experience(i) for i in experience_ids]

    return {
        "org": org,
        "customer": customer,
        "tender": tender_fields,
        "items": [{"number": n, "position": i.position, "name": i.name,
                   "price": money(i.price), "price_value": i.price}
                  for n, i in enumerate(repo.list_lot_items(tender_id), start=1)],
        "estimate": asdict(calc),
        "materials": _cost_rows(materials),
        "subcontract": _cost_rows(subcontract),
        "labor": [{"number": n, "post": r.post, "headcount": r.headcount,
                   "hours": r.hours, "rate": r.rate}
                  for n, r in enumerate(labor_rows, start=1)],
        "employees": _table(layouts["employees"], repo.list_employees()),
        "assets": _table(layouts["assets"], repo.list_assets()),
        "experience": _table(layouts["experience"], selected),
        "declaration_text": profile.declaration_text if profile else "",
        "conditions": dict(tender.deal_answers),
        "today": _today().strftime("%d.%m.%Y"),
        "money": money,
        "date_ru": date_ru,
        "date_long": date_long,
    }
