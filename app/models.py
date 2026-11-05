"""Модели справочников. Только датаклассы, без логики.

Деньги и часы -- только Decimal. Даты хранятся строкой ISO (ГГГГ-ММ-ДД):
встроенные адаптеры дат в sqlite3 объявлены устаревшими.
`id is None` означает запись, ещё не сохранённую в базе.
"""
from dataclasses import dataclass, field
from decimal import Decimal

ZERO = Decimal("0")

ROLE_PARTICIPANT = "participant"
ROLE_CUSTOMER = "customer"


@dataclass
class Organization:
    """Своё юрлицо (role='participant') или заказчик (role='customer')."""
    role: str = ROLE_PARTICIPANT
    inn: str = ""
    kpp: str = ""
    ogrn: str = ""
    short_name: str = ""
    full_name: str = ""
    legal_address: str = ""
    postal_address: str = ""
    phones: str = ""
    email: str = ""
    website: str = ""
    bank_name: str = ""
    bank_bik: str = ""
    bank_account: str = ""
    bank_corr_account: str = ""
    head_name: str = ""
    head_position: str = ""
    head_basis: str = ""              # основание полномочий
    okpo: str = ""                    # коды статистики
    oktmo: str = ""
    okved: str = ""
    vat_rate: Decimal | None = None   # None -- брать из настроек организации
    msp_status: str = ""              # статус МСП: '', 'micro', 'small', 'medium'
    id: int | None = None


@dataclass
class Employee:
    full_name: str = ""
    position: str = ""
    education: str = ""
    graduation_year: int | None = None
    specialty: str = ""
    qualification: str = ""           # квалификация и сертификаты
    experience_years: int | None = None
    hourly_rate: Decimal = ZERO       # руб/ч, нужна калькуляции
    id: int | None = None


@dataclass
class Asset:
    """Материально-технический ресурс."""
    name: str = ""
    kind: str = ""                    # 'premises' / 'equipment' / 'software'
    maker_model: str = ""             # изготовитель и марка
    quantity: int = 1
    year: int | None = None
    location: str = ""
    ownership_basis: str = ""         # основание владения
    ownership_doc: str = ""           # реквизиты документа
    purpose: str = ""
    condition: str = ""
    id: int | None = None


@dataclass
class ExperienceRecord:
    customer_name: str = ""
    contract_number: str = ""
    contract_date: str = ""
    subject: str = ""
    amount: Decimal = ZERO
    term: str = ""
    status: str = ""                  # статус исполнения
    service_id: int | None = None     # услуга из каталога, для отбора
    contract_scan: str = ""           # путь к скану контракта
    act_scan: str = ""                # путь к скану акта
    id: int | None = None


@dataclass
class Service:
    """Услуга каталога. norms -- часы по должностям: {должность: часы}."""
    code: str = ""                    # КЭР, СЗЗ, НДВ, ...
    name: str = ""
    stage_wording: str = ""           # типовые формулировки этапов
    norms: dict[str, Decimal] = field(default_factory=dict)
    id: int | None = None


@dataclass
class OrgSettings:
    """Проценты калькуляции. Значения по умолчанию задаёт репозиторий."""
    insurance_rate: Decimal = ZERO
    injury_rate: Decimal = ZERO
    overhead_rate: Decimal = ZERO
    profit_rate: Decimal = ZERO
    vat_rate: Decimal = ZERO


@dataclass
class CustomerProfile:
    """Состав пакета и правила оформления у конкретного заказчика."""
    customer_id: int | None = None    # organization.id заказчика
    documents: list[str] = field(default_factory=list)            # состав пакета
    file_prefixes: dict[str, str] = field(default_factory=dict)   # документ -> префикс
    declaration_text: str = ""
    id: int | None = None


@dataclass
class Tender:
    customer_id: int | None = None
    participant_id: int | None = None
    purchase_number: str = ""
    purchase_date: str = ""
    method: str = ""                  # способ закупки
    subject: str = ""
    object_name: str = ""
    object_address: str = ""
    price_wo_vat: Decimal = ZERO
    vat_rate: Decimal | None = None   # None -- брать из организации/настроек
    deadline: str = ""                # сроки исполнения
    payment_terms: str = ""
    doc_date: str = ""
    platform: str = ""
    output_dir: str = ""              # папка, куда собирается пакет
    deal_answers: dict[str, str] = field(default_factory=dict)    # условие -> ответ
    id: int | None = None


@dataclass
class LotItem:
    tender_id: int | None = None
    position: int = 0                 # порядок в списке
    name: str = ""
    price: Decimal = ZERO
    id: int | None = None
