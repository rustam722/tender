"""Описание справочников для веба: подписи полей, разбор тела запроса, вывод.

Тело запроса разбирается по аннотациям датаклассов из app.models: поле str,
int, Decimal, `X | None`, словарь, список или DocSpec. Разбор строгий -- число
в текстовом поле, `true` вместо числа, неизвестное поле, отрицательные часы
отказываются с русской фразой, которая называет поле. Вложенные структуры
(описания документов в профиле, нормы услуги, ответы по условиям) дают в
ошибке путь `path`: имена полей, ключи словарей и номера элементов списка.

Деньги и часы -- Decimal. Тело запроса читается с parse_float=Decimal, поэтому
число `5985000.10` в JSON не проходит через float и не превращается в
5985000.1000000002.
"""
import os
import re
import types
import typing
from dataclasses import dataclass, field, fields, is_dataclass, replace
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

from app.config import BASE_DIR
from app.models import (
    Asset, CustomerProfile, DocSpec, Employee, ExperienceRecord, LotItem, Organization,
    OrgSettings, Service, Tender, TenderCostItem, TenderLaborRow,
)
from app.web.errors import ApiError

MAX_DECIMAL = Decimal(10) ** 12      # триллион: правдоподобный предел для денег и часов
MAX_INT = 2 ** 63 - 1                # больше SQLite не хранит
# Пределы поменьше там, где значения перемножаются (численность * часы * ставка,
# количество * цена): контекст Decimal держит 28 значащих цифр, и произведение трёх
# чисел по 10^12 уже не округлить до копеек (InvalidOperation). С этими пределами
# любое произведение укладывается, и сохранённая строка не ломает расчёт и список.
_LIMITS = {"hours": Decimal(10) ** 6, "rate": Decimal(10) ** 6, "quantity": Decimal(10) ** 9}
_INT_LIMITS = {"headcount": 10 ** 6}
MAX_NUMBER_TEXT = 40

_DOC_FIELDS = {"template": "Шаблон", "cellmap": "Карта ячеек", "layout": "Раскладка",
               "prefix": "Префикс"}
_ISO_DATE = re.compile(r"\d{4}-\d{2}-\d{2}")
_DRIVE_PATH = re.compile(r"[A-Za-z]:[\\/]")


@dataclass(frozen=True)
class Entity:
    """Как справочник выглядит снаружи.

    save/get/list_/delete -- имена методов Repo (get может отсутствовать: тогда
    запись ищется в list). labels -- русские подписи полей, для сообщений.
    refs -- поля-ссылки: поле -> (подпись, метод Repo, которым проверить, что
    запись есть). in_use -- что сказать, когда запись удалить нельзя.
    parent -- поле, которым строка привязана к родителю (из адреса, не из тела)."""
    path: str
    cls: type
    title: str
    labels: dict[str, str]
    save: str = ""
    get: str = ""
    list_: str = ""
    delete: str = ""
    refs: dict = field(default_factory=dict)
    choices: dict = field(default_factory=dict)
    dates: tuple = ()
    dirs: tuple = ()
    money: tuple = ()        # денежные поля: в ответе рядом <поле>_text
    numbers: tuple = ()      # часы и количества: <поле>_text с запятой
    percents: tuple = ()     # ставки-доли: <поле>_percent
    in_use: str = ""
    parent: str = ""
    check: typing.Callable | None = None     # (obj, repo) -> None или ApiError


def _one_profile_per_customer(profile, repo):
    """У заказчика один профиль: сборка берёт первый найденный, второй молча
    остался бы невидимым."""
    if profile.customer_id is None:
        return
    found = repo.find_customer_profile(profile.customer_id)
    if found is not None and found.id != profile.id:
        raise ApiError("У этого заказчика уже есть профиль. Откройте его и измените, "
                       "а не создавайте второй.", path=["customer_id"])


_ROLES = ("participant", "customer")
_MSP = ("", "micro", "small", "medium")
_COST_KINDS = ("materials", "subcontract")

ORGANIZATIONS = Entity(
    "organizations", Organization, "Организация",
    labels={"role": "Роль", "inn": "ИНН", "kpp": "КПП", "ogrn": "ОГРН",
            "short_name": "Краткое наименование", "full_name": "Полное наименование",
            "legal_address": "Юридический адрес", "postal_address": "Почтовый адрес",
            "phones": "Телефоны", "email": "E-mail", "website": "Сайт", "bank_name": "Банк",
            "bank_bik": "БИК", "bank_account": "Расчётный счёт",
            "bank_corr_account": "Корреспондентский счёт", "head_name": "ФИО руководителя",
            "head_position": "Должность руководителя", "head_basis": "Основание полномочий",
            "okpo": "ОКПО", "oktmo": "ОКТМО", "okved": "ОКВЭД", "vat_rate": "Ставка НДС",
            "msp_status": "Статус МСП"},
    save="save_organization", get="get_organization", list_="list_organizations",
    delete="delete_organization",
    choices={"role": _ROLES, "msp_status": _MSP}, percents=("vat_rate",),
    in_use="Эту организацию нельзя удалить: она указана в тендерах как заказчик или "
           "участник. Сначала удалите эти тендеры или выберите в них другую организацию.")

EMPLOYEES = Entity(
    "employees", Employee, "Сотрудник",
    labels={"full_name": "ФИО", "position": "Должность", "education": "Образование",
            "graduation_year": "Год окончания", "specialty": "Специальность",
            "qualification": "Квалификация", "experience_years": "Стаж, лет",
            "hourly_rate": "Ставка, руб/ч"},
    save="save_employee", get="get_employee", list_="list_employees", delete="delete_employee",
    money=("hourly_rate",))

ASSETS = Entity(
    "assets", Asset, "Ресурс",
    labels={"name": "Наименование", "kind": "Вид", "maker_model": "Изготовитель и марка",
            "quantity": "Количество", "year": "Год", "location": "Местонахождение",
            "ownership_basis": "Основание владения", "ownership_doc": "Документ о владении",
            "purpose": "Назначение", "condition": "Состояние"},
    save="save_asset", get="get_asset", list_="list_assets", delete="delete_asset")

EXPERIENCE = Entity(
    "experience", ExperienceRecord, "Опыт",
    labels={"customer_name": "Заказчик", "contract_number": "Номер контракта",
            "contract_date": "Дата контракта", "subject": "Предмет", "amount": "Сумма",
            "term": "Срок", "status": "Статус исполнения", "service_id": "Услуга",
            "contract_scan": "Скан контракта", "act_scan": "Скан акта"},
    save="save_experience", get="get_experience", list_="list_experience",
    delete="delete_experience",
    refs={"service_id": ("Услуга", "get_service")}, money=("amount",))

SERVICES = Entity(
    "services", Service, "Услуга",
    labels={"code": "Код", "name": "Название", "stage_wording": "Формулировки этапов",
            "norms": "Нормы часов"},
    save="save_service", get="get_service", list_="list_services", delete="delete_service")

PROFILES = Entity(
    "profiles", CustomerProfile, "Профиль заказчика",
    labels={"customer_id": "Заказчик", "documents": "Состав пакета",
            "docs": "Описание документов", "declaration_text": "Текст декларации"},
    save="save_customer_profile", get="get_customer_profile", list_="list_customer_profiles",
    delete="delete_customer_profile",
    refs={"customer_id": ("Заказчик", "get_organization")},
    check=_one_profile_per_customer)

TENDERS = Entity(
    "tenders", Tender, "Тендер",
    labels={"customer_id": "Заказчик", "participant_id": "Участник",
            "purchase_number": "Номер закупки", "purchase_date": "Дата закупки",
            "method": "Способ закупки", "subject": "Предмет закупки", "object_name": "Объект",
            "object_address": "Адрес объекта", "price_wo_vat": "Цена без НДС",
            "vat_rate": "Ставка НДС", "deadline": "Сроки исполнения",
            "payment_terms": "Условия оплаты", "doc_date": "Дата документов",
            "platform": "Площадка", "output_dir": "Папка пакета",
            "deal_answers": "Ответы по условиям сделки", "experience_ids": "Отбор опыта"},
    save="save_tender", get="get_tender", list_="list_tenders", delete="delete_tender",
    refs={"customer_id": ("Заказчик", "get_organization"),
          "participant_id": ("Участник", "get_organization")},
    dates=("purchase_date", "doc_date"), dirs=("output_dir",),
    money=("price_wo_vat",), percents=("vat_rate",))

LOT_ITEMS = Entity(
    "items", LotItem, "Позиция лота",
    labels={"position": "Порядок", "name": "Наименование", "price": "Цена"},
    save="save_lot_item", get="get_lot_item", list_="list_lot_items", delete="delete_lot_item",
    parent="tender_id", money=("price",))

LABOR = Entity(
    "labor", TenderLaborRow, "Строка трудозатрат",
    labels={"position": "Порядок", "post": "Должность", "headcount": "Численность",
            "hours": "Часы", "rate": "Ставка, руб/ч"},
    save="save_tender_labor", list_="list_tender_labor", delete="delete_tender_labor",
    parent="tender_id", money=("rate",), numbers=("hours",))

COSTS = Entity(
    "costs", TenderCostItem, "Строка материалов или субподряда",
    labels={"kind": "Статья", "position": "Порядок", "name": "Наименование",
            "unit": "Единица", "quantity": "Количество", "price": "Цена", "note": "Примечание"},
    save="save_tender_cost_item", list_="list_tender_cost_items",
    delete="delete_tender_cost_item", choices={"kind": _COST_KINDS}, parent="tender_id",
    money=("price",), numbers=("quantity",))

SETTINGS = Entity(
    "settings", OrgSettings, "Настройки",
    labels={"insurance_rate": "Страховые взносы", "injury_rate": "Травматизм",
            "overhead_rate": "Накладные", "profit_rate": "Прибыль", "vat_rate": "НДС",
            "total_rounding_digits": "Округление итога"},
    percents=("insurance_rate", "injury_rate", "overhead_rate", "profit_rate", "vat_rate"))

# Не справочник: подписи полей запроса подбора часов, чтобы отказ назвал поле.
SOLVE = Entity("solve", Tender, "Подбор часов",
               labels={"target": "Цена", "target_includes_vat": "Цена включает НДС"})

TOP_LEVEL = (ORGANIZATIONS, EMPLOYEES, ASSETS, EXPERIENCE, SERVICES, PROFILES, TENDERS)
CHILDREN = (LOT_ITEMS, LABOR, COSTS)

# Ставки-доли: 0,05 -- это 5 %. Не по суффиксу «_rate»: hourly_rate -- рубли в час.
_FRACTIONS = {"vat_rate", "insurance_rate", "injury_rate", "overhead_rate", "profit_rate"}

# Поля, где допустимо отрицательное целое число. Остальные Decimal и int -- от нуля.
_SIGNED = {"total_rounding_digits"}
_ROUNDING_RANGE = (-9, 2)        # знаков после запятой; 0 -- рубли, 2 -- копейки, -3 -- тысячи


# ---- разбор значений ----------------------------------------------------------

def _where(entity: Entity, path: list) -> str:
    """«Поле «Описание документов» → «Справка о МТР» → «Раскладка»»."""
    parts = []
    for i, step in enumerate(path):
        if i == 0:
            parts.append(f"«{entity.labels.get(step, step)}»")
        elif isinstance(step, int):
            parts.append(f"№{step + 1}")
        elif i == 2 and path[0] == "docs":
            parts.append(f"«{_DOC_FIELDS.get(step, step)}»")
        else:
            parts.append(f"«{step}»")
    return "Поле " + " → ".join(parts)


def fail(entity: Entity, path: list, problem: str):
    raise ApiError(f"{_where(entity, path)}: {problem}", path=list(path))


def decimal_value(value, entity, path):
    if isinstance(value, bool):
        fail(entity, path, "введите число, например 12 или 12,5.")
    if isinstance(value, Decimal):
        number = value
    elif isinstance(value, int):
        number = Decimal(value)
    elif isinstance(value, str):
        text = re.sub(r"[\s\u00a0\u202f]", "", value).replace(",", ".")
        if not text or len(text) > MAX_NUMBER_TEXT:
            fail(entity, path, "введите число, например 12 или 12,5.")
        try:
            number = Decimal(text)
        except InvalidOperation:
            fail(entity, path, "введите число, например 12 или 12,5.")
    else:
        fail(entity, path, "введите число, например 12 или 12,5.")
    if not number.is_finite():
        fail(entity, path, "введите обычное число.")
    if number < 0:
        fail(entity, path, "число не может быть отрицательным.")
    if number > _LIMITS.get(path[-1], MAX_DECIMAL):
        fail(entity, path, "число слишком велико.")
    number = abs(number)                                  # «-0» -> «0»
    if number.as_tuple().exponent > 0:                    # 1E+5 -> 100000
        number = number.quantize(Decimal(1))
    if path[-1] in _FRACTIONS and number > 1:
        fail(entity, path, "ставка задаётся долей единицы, от 0 до 1: "
                           "например, 0,05 для 5 %.")
    return number


def _int(value, entity, path):
    problem = "введите целое число."
    if isinstance(value, bool):
        fail(entity, path, problem)
    if isinstance(value, Decimal):
        if value != value.to_integral_value():
            fail(entity, path, problem)
        number = int(value)
    elif isinstance(value, int):
        number = value
    elif isinstance(value, str) and re.fullmatch(r"\s*-?\d{1,19}\s*", value):
        number = int(value)
    else:
        fail(entity, path, problem)
    if abs(number) > _INT_LIMITS.get(path[-1], MAX_INT):
        fail(entity, path, "число слишком велико.")
    if number < 0 and path[-1] not in _SIGNED:
        fail(entity, path, "число не может быть отрицательным.")
    return number


def _str(value, entity, path):
    if not isinstance(value, str):
        fail(entity, path, "должно быть текстом (может быть пустым).")
    return value


def _doc_spec(value, entity, path):
    if not isinstance(value, dict):
        fail(entity, path, "должно быть набором полей: "
                           + ", ".join(f"«{n}»" for n in _DOC_FIELDS.values()) + ".")
    for key, item in value.items():
        if key not in _DOC_FIELDS:
            fail(entity, path + [key], "такого поля у описания документа нет; допустимы "
                 + ", ".join(f"«{n}»" for n in _DOC_FIELDS.values()) + ".")
        if not isinstance(item, str):
            fail(entity, path + [key], "должно быть текстом (может быть пустым).")
    try:
        return DocSpec(**value)
    except ValueError as exc:         # русский отказ модели: книга без карты ячеек
        fail(entity, path, str(exc))


def coerce(value, hint, entity, path):
    origin = typing.get_origin(hint)
    if origin in (typing.Union, types.UnionType):
        inner = [a for a in typing.get_args(hint) if a is not type(None)]
        nullable = len(inner) < len(typing.get_args(hint))
        # Пустое поле формы приходит как "": для необязательного числа или ссылки
        # это «не задано», а не текст, который надо разбирать как число.
        blank = isinstance(value, str) and not value.strip() and inner[0] is not str
        if value is None or (nullable and blank):
            if nullable:
                return None
            fail(entity, path, "значение не может быть пустым.")
        return coerce(value, inner[0], entity, path)
    if origin is dict:
        key_type, value_type = typing.get_args(hint)
        if not isinstance(value, dict):
            fail(entity, path, "должно быть набором «название -> значение».")
        return {key: coerce(item, value_type, entity, path + [key])
                for key, item in value.items()}
    if origin is list:
        (item_type,) = typing.get_args(hint)
        if not isinstance(value, list):
            fail(entity, path, "должно быть списком.")
        return [coerce(item, item_type, entity, path + [i]) for i, item in enumerate(value)]
    if hint is DocSpec:
        return _doc_spec(value, entity, path)
    if hint is Decimal:
        return decimal_value(value, entity, path)
    if hint is int:
        return _int(value, entity, path)
    if hint is str:
        return _str(value, entity, path)
    raise TypeError(f"Тип поля {hint!r} не описан в разборе веба")   # ошибка программиста


def check_out_dir(value, entity: Entity, path: list) -> str:
    """Папка пакета: абсолютный локальный путь с диском, не в папке программы.

    Строка приходит из браузера, а по ней сборка создаёт папку и пишет файлы.
    Поэтому отказываем: относительным путям (неясно, от чего они считаются),
    сетевым и служебным (`\\\\сервер\\ресурс`, `\\\\?\\`, `\\\\.\\`), путям с «:» после
    имени диска (потоки NTFS), управляющим знакам, а также самой папке программы
    и всему внутри неё: сборка с заменой файлов могла бы затереть шаблоны.
    `..` не отвергается, а раскрывается: выбранная папка -- та, на которую путь
    указывает на самом деле."""
    if not isinstance(value, str):
        fail(entity, path, "должно быть текстом -- путём к папке.")
    text = value.strip()
    if not text:
        fail(entity, path, "папка не указана.")
    if any(ord(ch) < 32 for ch in text):
        fail(entity, path, "в пути есть управляющие знаки.")
    if text.startswith(("\\\\", "//")) or not _DRIVE_PATH.match(text):
        fail(entity, path, "укажите полный путь к папке на этом компьютере, начиная с "
                           "диска, например C:\\Тендеры\\Заявка. Сетевые папки и "
                           "относительные пути не принимаются.")
    if ":" in text[2:]:
        fail(entity, path, "в пути не должно быть знака «:» после имени диска.")
    folder = os.path.normpath(text)
    resolved = Path(os.path.realpath(folder))
    program = Path(os.path.realpath(BASE_DIR))
    if resolved == program or program in resolved.parents:
        fail(entity, path, "эта папка относится к самой программе. Выберите другую, "
                           "например папку тендера.")
    return folder


def _check_field(entity: Entity, name: str, value, path: list):
    if name in entity.choices and value is not None and value not in entity.choices[name]:
        fail(entity, path, "недопустимое значение.")
    if name in entity.dates and value:
        if not _ISO_DATE.fullmatch(value):
            fail(entity, path, "дата должна быть в виде ГГГГ-ММ-ДД, например 2025-10-24.")
        try:
            date.fromisoformat(value)
        except ValueError:
            fail(entity, path, "такой даты нет в календаре. Формат: ГГГГ-ММ-ДД.")
    if name in entity.dirs and value:
        return check_out_dir(value, entity, path)
    if name == "total_rounding_digits":
        low, high = _ROUNDING_RANGE
        if not low <= value <= high:
            fail(entity, path, f"допустимы значения от {low} до {high}: 0 -- целые рубли, "
                               "2 -- копейки, -3 -- тысячи.")
    return value


def output_only_keys(entity: Entity) -> set[str]:
    """Ключи, которые ответ добавляет к записи сам (готовые строки и суммы строк).
    Экран может отправить запись обратно целиком, как получил: эти ключи при
    разборе игнорируются. Любые другие неизвестные ключи -- по-прежнему отказ."""
    keys = {name + "_text" for name in entity.money + entity.numbers}
    keys |= {name + "_percent" for name in entity.percents}
    if entity.cls is Service:
        keys.add("norms_text")
    if entity in (LABOR, COSTS):
        keys |= {"amount", "amount_text"}
    return keys


def parse_object(entity: Entity, data: dict, existing=None, fixed: dict | None = None):
    """Тело запроса -> датакласс. Присланные поля заменяют поля existing, не
    присланные остаются как были: частичная правка не стирает остальное.
    fixed -- значения, которые задаёт адрес, а не тело (родитель строки)."""
    hints = typing.get_type_hints(entity.cls)
    names = {f.name for f in fields(entity.cls)} - {"id"} - ({entity.parent} if entity.parent else set())
    values = {}
    skip = output_only_keys(entity)
    for name, raw in data.items():
        if name in skip:
            continue
        if name not in names:
            raise ApiError(f"Неизвестное поле «{name}»: такого поля у записи "
                           f"«{entity.title}» нет. Проверьте имя.", path=[name])
        value = coerce(raw, hints[name], entity, [name])
        values[name] = _check_field(entity, name, value, [name])
    values.update(fixed or {})
    if existing is not None:
        return replace(existing, **values)
    return entity.cls(**values)


# ---- вывод --------------------------------------------------------------------

def plain(obj):
    """Датаклассы, Decimal и пути -- в то, что отдаёт JSON. Деньги идут строкой
    ровно с теми знаками, что хранятся: '5985000.10', не 5985000.1."""
    if isinstance(obj, Decimal):
        return format(obj, "f")
    if is_dataclass(obj) and not isinstance(obj, type):
        return {f.name: plain(getattr(obj, f.name)) for f in fields(obj)}
    if isinstance(obj, dict):
        return {key: plain(value) for key, value in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [plain(item) for item in obj]
    if isinstance(obj, Path):
        return str(obj)
    return obj
