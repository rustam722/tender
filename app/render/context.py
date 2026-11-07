"""Контекст шаблонов: словарь, из которого читает один документ пакета.

Читает данные через Repo и раскладывает их по ключам; документы не пишет.
Правил, кроме цепочки НДС и форматирования денег, здесь нет.

Ключи (значения отдаются как есть, Decimal остаётся Decimal -- форматирует
шаблон через `money`, кроме клеток таблиц, см. ниже).

Единый контракт денег в списках items, materials, subcontract и labor:
текст через money под естественным именем (price, amount, rate) и Decimal
под именем с суффиксом _value (price_value, amount_value, rate_value).
Исключение -- количества: quantity у материалов и hours у ФОТ остаются
сырым Decimal, это не деньги, и разряды с запятой им не положены.
Построчные суммы (amount) округляются до копеек один раз, и в калькуляцию
идёт сумма именно этих округлённых значений: каждая напечатанная строка
воспроизводится от напечатанных соседей, подтаблица сходится со своим итогом.

    org, customer      поля организации (пустые строки, если организации нет)
    tender             поля тендера; vat_rate -- итоговая ставка по цепочке,
                       vat_percent -- она же текстом: «5», «5,5»
    items              позиции лота: number, position, name, price (текст
                       через money), price_value (Decimal -- для ячейки Excel);
                       vat и vat_value -- НДС позиции по ставке тендера,
                       округлённый до копеек (ROUND_HALF_UP) на строке
    estimate           статьи калькуляции, Decimal
    materials,         строки материалов и субподряда: number, name, unit,
    subcontract        quantity (Decimal), price и amount (текст через money),
                       price_value и amount_value (Decimal), note; их суммы
                       идут в estimate
    labor              строки ФОТ: number, post, headcount, hours (Decimal),
                       rate и amount (headcount * hours * rate) -- текст,
                       rate_value и amount_value -- Decimal
    employees, assets, experience
                       {"headers": [...], "rows": [[...]]} под раскладку
    declaration_text   текст декларации из профиля заказчика
    conditions         ответы по условиям сделки: условие -> ответ
    today              дата сборки, «ДД.ММ.ГГГГ»
    money              функция: Decimal -> «5 985 000,00»
    money_total        функция: список Decimal -> Decimal, сумма строк, каждая из
                       которых округлена до копеек так, как напечатана; для
                       итогов под столбцами (app.calc.money_total)
    date_ru, date_long функции: ISO-строка -> «24.10.2025» и «24» октября 2025 г.;
                       пустая строка -> пустая строка
    warnings           list[str]: что пропущено или подставлено по умолчанию
                       (пропавший опыт, документ без назначенной раскладки;
                       для документа с картой ячеек -- итог без НДС в книге
                       заказчика (app.calc.workbook_total) расходится с итогом
                       программы, или есть строки с долей копейки)

Деньги в клетках таблиц. Раскладки (задача 7) ничего не форматируют, поэтому
перед build_rows каждое Decimal-поле записи заменяется своим текстом через
money -- в копии записи, справочник не трогается. Внимание: форматируется
КАЖДОЕ Decimal-поле как деньги. Сегодня это верно (у Employee и ExperienceRecord
Decimal -- только деньги, у Asset Decimal нет), но Decimal-поле часов или
количества, добавленное в справочник, молча получило бы разряды и округление
до копеек во всех документах -- такое поле придётся исключить здесь. Так формат не зависит от
того, какую колонку выбрала раскладка: форматируется любое денежное поле
(сумма контракта, часовая ставка), а не список имён, который пришлось бы
держать в согласии с правкой раскладок.
"""
from dataclasses import asdict, fields, replace
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from app import config
from app.calc import Estimate, LaborLine, estimate, money_total, workbook_total
from app.models import (
    COST_MATERIALS, COST_SUBCONTRACT, ROLE_CUSTOMER, ROLE_PARTICIPANT,
    CustomerProfile, DocSpec, LotItem, Organization, TenderCostItem, TenderLaborRow,
)
from app.render.layouts import TableLayout, build_rows, load_layout
from app.render.xlsx_render import CellMapError, load_cellmap
from app.repo import Repo


NBSP = " "
_KOPECK = Decimal("0.01")
_ZERO = Decimal("0")


def _round(value: Decimal) -> Decimal:
    """До копеек, ROUND_HALF_UP -- как статьи калькулятора."""
    return value.quantize(_KOPECK, rounding=ROUND_HALF_UP)

# Канонические имена документов пакета -- список из начала плана реализации, в
# порядке подачи. Эти строки -- ключи CustomerProfile.documents и .docs, имена
# шаблонов и seed. Экран сборки показывает их все; профиль заказчика лишь
# отмечает и упорядочивает нужные, а не заменяет список (у заказчика без профиля
# сборка не должна быть пустой).
DOCUMENTS = (
    "Коммерческое предложение",
    "Калькуляция к КП",
    "ТКП по форме заказчика",
    "Декларация соответствия",
    "Справка о кадровых ресурсах",
    "Справка о наличии оборудования",
    "Справка о МТР",
    "Письмо о крупности сделки",
    "Справка об опыте",
    "Анкета участника",
)

# Документы, несущие табличную справку, -- имена из списка документов
# раздела 8 описания; они же ключи CustomerProfile.documents и .docs.
# Только для них отсутствие раскладки в профиле -- повод предупредить: у
# декларации или письма о крупности сделки таблицы нет, и предупреждение
# о них было бы шумом, который приучает пролистывать настоящие.
# Контекст не знает, какой шаблон его отрендерит, поэтому список держится
# вручную: его нужно держать в согласии с шаблонами задачи 9 и профилями,
# которые засевает задача 14 (новый документ с таблицей -- добавить сюда).
TABLE_DOCUMENTS = frozenset({
    "Справка о кадровых ресурсах",
    "Справка о наличии оборудования",
    "Справка о МТР",
    "Справка об опыте",
})

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
    rounded = _round(Decimal(value))
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


def percent(rate: Decimal) -> str:
    """Ставка-доля -> проценты строкой: 0.05 -> «5», 0.002 -> «0,2», 0.055 -> «5,5».
    Лежит рядом с money: форматирование денег и процентов -- в одном месте, чтобы
    экран не умножал ставку на 100 сам (в JavaScript нет Decimal)."""
    return format((rate * 100).normalize(), "f").replace(".", ",")


def rate_from_percent(value: Decimal) -> Decimal:
    """Обратное к percent: проценты -> ставка-доля, 20 -> 0.20, 0.7 -> 0.007.
    Сдвиг порядка, а не деление: результат точен (в JavaScript 0.7 / 100 даёт
    0.007000000000000001, а Decimal.scaleb -- ровно 0.007)."""
    return value.scaleb(-2)


def number_ru(value: Decimal) -> str:
    """Не деньги -- часы и количества: 12.50 -> «12,5», 360 -> «360». Запятая
    десятичная, разрядов нет."""
    return format(value.normalize(), "f").replace(".", ",")


def _vat_rate(tender, participant: Organization | None, settings) -> Decimal:
    """Тендер -> участник -> настройки; первое значение, не равное None.
    Ноль -- осознанное «без НДС», в следующее звено не проваливается."""
    for rate in (tender.vat_rate,
                 participant.vat_rate if participant else None):
        if rate is not None:
            return rate
    return settings.vat_rate


def _doc_spec(profile: CustomerProfile | None,
              document: str | None) -> DocSpec | None:
    """Описание документа в профиле; None, если профиля, имени или записи нет."""
    return profile.docs.get(document) if profile and document else None


def _layouts(profile: CustomerProfile | None,
             document: str | None) -> dict[str, TableLayout]:
    """Раскладка на каждый источник таблиц. Умолчания; поверх них -- раскладка,
    которую профиль назначает именно этому документу (docs[документ].layout).
    Какому справочнику служит файл, говорит он сам (поле source): один контекст
    собирается под один документ, поэтому двух раскладок на один источник
    не бывает. Документ без раскладки в профиле остаётся на умолчаниях."""
    layouts = {source: load_layout(Path(config.LAYOUTS_DIR) / name)
               for source, name in _DEFAULT_LAYOUTS.items()}
    spec = _doc_spec(profile, document)
    name = spec.layout if spec else ""
    if name:
        chosen = load_layout(Path(config.LAYOUTS_DIR) / name)
        layouts[chosen.source] = chosen
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
    rows = []
    for n, i in enumerate(items, start=1):
        amount = _round(i.quantity * i.price)      # один раз, на строке
        rows.append({"number": n, "name": i.name, "unit": i.unit,
                     "quantity": i.quantity,
                     "price": money(i.price), "price_value": i.price,
                     "amount": money(amount), "amount_value": amount,
                     "note": i.note})
    return rows


def _exact(value: Decimal) -> str:
    """Число как оно есть, без округления: 1851.855 -> «1 851,855»."""
    text = f"{value.normalize():,f}"
    return text.translate({ord(","): NBSP, ord("."): ","})


def _inexact_rows(label: str, factors: list[tuple[str, tuple[Decimal, ...]]]
                  ) -> list[str]:
    """Строки, где произведение множителей не ложится в копейку ровно: что
    получилось в книге и что напечатано в документе."""
    found = []
    for name, numbers in factors:
        product = numbers[0]
        for number in numbers[1:]:
            product *= number
        if product != _round(product):
            found.append(
                f"строка «{name}» ({label}): "
                f"{' × '.join(_exact(n) for n in numbers)} = {_exact(product)}, "
                f"в документе напечатано {money(_round(product))}")
    return found


def _estimate_rows(labor: list[dict], materials: list[dict],
                   subcontract: list[dict], settings) -> Estimate:
    """Калькуляция по напечатанным (округлённым) строкам. ФОТ идёт строками
    «1 чел. * сумма строки * 1 руб.»: estimate округляет ФОТ сам, и от сырых
    произведений сумма расходилась бы с напечатанными строками."""
    return estimate(
        [LaborLine(row["post"], 1, row["amount_value"], Decimal("1"))
         for row in labor],
        sum((row["amount_value"] for row in materials), _ZERO),
        sum((row["amount_value"] for row in subcontract), _ZERO),
        settings)


_KINDS = {"materials": "материалы", "subcontract": "субподряд",
          "labor": "трудозатраты"}
_FIX = {"материалы": "количество или цену", "субподряд": "количество или цену",
        "трудозатраты": "численность, часы или ставку"}


def _workbook_sums_warnings(document: str, spec: DocSpec | None, labor, materials,
                            subcontract, settings) -> list[str]:
    """Документ с книгой Excel заказчика: разойдутся ли две бумаги.

    Программа считает итог по напечатанным строкам и округляет страховые до
    копеек, книга складывает неокруглённое (формулу менять нельзя: арифметика
    книги -- заказчика). Поэтому итог считается дважды: как считает программа
    (`estimate`) и как посчитает книга (`workbook_total`); не сошлись --
    предупреждение. Это меряет следствие, а не перебирает причины. Отдельно,
    даже при равных итогах, предупреждаем о строках с долей копейки: подытог
    книги не сойдётся с напечатанными над ним строками.
    Отказ был бы неправ: цена бывает из прайса заказчика, а страховые человек
    не исправит; предупреждение сообщает, а не требует. Смотрим только те
    списки, которые карта этого документа действительно отправляет в книгу."""
    if not (spec and spec.cellmap):
        return []
    try:
        sources = {b.source for b in
                   load_cellmap(Path(config.CELLMAPS_DIR) / spec.cellmap).rows}
    except CellMapError:
        return []          # пропавшую или битую карту назовёт сборка пакета
    labor = labor if "labor" in sources else []
    materials = materials if "materials" in sources else []
    subcontract = subcontract if "subcontract" in sources else []

    program = _estimate_rows(labor_rows(labor), _cost_rows(materials),
                             _cost_rows(subcontract), settings).total_wo_vat
    book = workbook_total(
        [LaborLine(r.post, r.headcount, r.hours, r.rate) for r in labor],
        sum((i.quantity * i.price for i in materials), _ZERO),
        sum((i.quantity * i.price for i in subcontract), _ZERO), settings)
    rows = (_inexact_rows(_KINDS["materials"], [
                (i.name, (i.quantity, i.price)) for i in materials])
            + _inexact_rows(_KINDS["subcontract"], [
                (i.name, (i.quantity, i.price)) for i in subcontract])
            + _inexact_rows(_KINDS["labor"], [
                (r.post, (Decimal(r.headcount), r.hours, r.rate))
                for r in labor]))
    if book == program and not rows:
        return []

    if book != program:
        text = (f"В «{document}» книга Excel посчитает итог без НДС "
                f"{money(book)}, а программа и «Коммерческое предложение» -- "
                f"{money(program)}.")
    else:
        text = (f"В «{document}» итог без НДС в книге Excel и в программе "
                f"совпадёт ({money(program)}), но подытоги книги не сойдутся "
                "с напечатанными над ними строками.")
    if rows:
        kinds = [k for k in _FIX if any(f"({k})" in r for r in rows)]
        fix = "; ".join(f"{_FIX[k]} ({k})" for k in kinds)
        text += (" Книга складывает неокруглённые произведения строк, а "
                 "программа и документы Word -- напечатанные строки: "
                 + "; ".join(rows) + ". Если эти цифры не из прайса заказчика, "
                 f"исправьте {fix} так, чтобы произведение было целым числом "
                 "копеек.")
    else:
        text += (" Строк с долей копейки нет: книга Excel не округляет "
                 "страховые взносы и округляет итог до целых рублей, а "
                 "программа округляет страховые до копеек, поэтому на границе "
                 "рубля итоги расходятся. Исправить это нельзя, предупреждение "
                 "только сообщает.")
    return [text]


def labor_rows(rows: list[TenderLaborRow]) -> list[dict]:
    result = []
    for n, r in enumerate(rows, start=1):
        amount = _round(r.headcount * r.hours * r.rate)   # один раз, на строке
        result.append({"number": n, "post": r.post, "headcount": r.headcount,
                       "hours": r.hours,
                       "rate": money(r.rate), "rate_value": r.rate,
                       "amount": money(amount), "amount_value": amount})
    return result


def _lot_rows(items: list[LotItem], rate: Decimal) -> list[dict]:
    """Позиции лота; НДС позиции округляется до копеек один раз, на строке,
    как amount у материалов, -- форма заказчика берёт его готовым числом."""
    rows = []
    for n, i in enumerate(items, start=1):
        vat = _round(i.price * rate)
        rows.append({"number": n, "position": i.position, "name": i.name,
                     "price": money(i.price), "price_value": i.price,
                     "vat": money(vat), "vat_value": vat})
    return rows


def build_context(tender_id: int, repo: Repo, *, document: str | None = None,
                  experience_ids: list[int] | tuple[int, ...] | None = None) -> dict:
    """Собирает контекст тендера. Неизвестный id -- LookupError по-русски.

    document -- имя документа пакета, под который собирается контекст: таблицы
    строятся по раскладке, которую профиль заказчика назначает этому документу.
    Без имени -- по раскладкам по умолчанию. Пакет вызывает сборку на каждый
    документ: «Справка об оборудовании» и «Справка о МТР» берут один справочник
    МТР, но с разными колонками.

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
    stored_labor = repo.list_tender_labor(tender_id)
    materials = repo.list_tender_cost_items(tender_id, COST_MATERIALS)
    subcontract = repo.list_tender_cost_items(tender_id, COST_SUBCONTRACT)
    labor = labor_rows(stored_labor)
    materials_table, subcontract_table = _cost_rows(materials), _cost_rows(subcontract)
    # В калькуляцию уходят суммы напечатанных (округлённых) строк.
    priced = replace(settings, vat_rate=rate)
    calc = _estimate_rows(labor, materials_table, subcontract_table, priced)

    tender_fields = asdict(tender)
    tender_fields["vat_rate"] = rate
    tender_fields["vat_percent"] = percent(rate)

    layouts = _layouts(profile, document)
    spec = _doc_spec(profile, document)
    warnings = []
    if (profile and document in TABLE_DOCUMENTS
            and document in profile.documents
            and not (spec and spec.layout)):
        warnings.append(
            f"Для документа «{document}» в профиле заказчика не назначена "
            "раскладка таблиц: колонки взяты по умолчанию. Если у заказчика "
            "своя форма таблицы, добавьте раскладку в профиль.")
    warnings.extend(_workbook_sums_warnings(
        document or "", spec, stored_labor, materials, subcontract, priced))
    if experience_ids is None:
        experience_ids = tender.experience_ids
    selected = []
    for exp_id in experience_ids:
        try:
            selected.append(repo.get_experience(exp_id))
        except LookupError:
            # Сборку не роняем, но и молча терять опыт нельзя: он влияет
            # на оценку заявки.
            warnings.append(
                f"Запись опыта с id={exp_id} не найдена в справочнике, "
                "в справку об опыте не попала. Проверьте отбор опыта.")

    return {
        "org": org,
        "customer": customer,
        "tender": tender_fields,
        "items": _lot_rows(repo.list_lot_items(tender_id), rate),
        "estimate": asdict(calc),
        "materials": materials_table,
        "subcontract": subcontract_table,
        "labor": labor,
        "employees": _table(layouts["employees"], repo.list_employees()),
        "assets": _table(layouts["assets"], repo.list_assets()),
        "experience": _table(layouts["experience"], selected),
        "declaration_text": profile.declaration_text if profile else "",
        "conditions": dict(tender.deal_answers),
        "today": _today().strftime("%d.%m.%Y"),
        "money": money,
        "money_total": money_total,
        "date_ru": date_ru,
        "date_long": date_long,
        "warnings": warnings,
    }
