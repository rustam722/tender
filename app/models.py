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
    """Проценты калькуляции. У ставок значений по умолчанию нет намеренно:
    нулевые ставки дали бы правдоподобное КП без взносов, накладных, прибыли
    и НДС. Умолчания ставок живут в repo.DEFAULT_SETTINGS.

    total_rounding_digits -- до какого знака округляется итог без НДС, как
    второй аргумент ROUND в книге КП.xlsx (0 -- целые рубли, 2 -- копейки,
    -3 -- тысячи). Здесь умолчание безопасно: 0 -- именно то, что делает книга."""
    insurance_rate: Decimal
    injury_rate: Decimal
    overhead_rate: Decimal
    profit_rate: Decimal
    vat_rate: Decimal
    total_rounding_digits: int = 0


@dataclass(frozen=True)
class DocSpec:
    """Как у конкретного заказчика оформляется один документ пакета.

    Любое поле может быть пустым:
      template -- файл шаблона в templates_docs/; пусто -- `<имя документа>.docx`;
      cellmap  -- файл карты ячеек в cellmaps/; нужен книге .xlsx, Word-шаблону
                  не нужен;
      layout   -- файл раскладки таблицы в layouts/; пусто -- раскладка
                  по умолчанию для справочника документа;
      prefix   -- префикс имени файла; пусто -- имя без префикса.

    Неизменяем: проверка выполняется при создании, и присвоением поля её
    нельзя обойти."""
    template: str = ""
    cellmap: str = ""
    layout: str = ""
    prefix: str = ""

    def __post_init__(self):
        for name in ("template", "cellmap", "layout", "prefix"):
            if not isinstance(getattr(self, name), str):
                raise ValueError(
                    f"Поле «{name}» описания документа должно быть строкой "
                    "(может быть пустой).")
        # Пробелы по краям имён файлов -- опечатка, а не часть имени: без обрезки
        # « форма.xlsx » прошёл бы мимо проверки расширения.
        object.__setattr__(self, "template", self.template.strip())
        object.__setattr__(self, "cellmap", self.cellmap.strip())
        # Книгу заполняют по карте ячеек: без неё заполнять нечем.
        if self.template.lower().endswith(".xlsx") and not self.cellmap:
            raise ValueError(
                f"Для шаблона «{self.template}» не указана карта ячеек: "
                "без неё книгу Excel заполнить нечем. Укажите файл карты "
                "в описании документа.")


@dataclass
class CustomerProfile:
    """Состав пакета и правила оформления у конкретного заказчика.

    docs -- по одному описанию на документ; ключ -- каноническое имя документа,
    то же, что в documents. Документа нет в docs -- для него всё по умолчанию."""
    customer_id: int | None = None    # organization.id заказчика
    documents: list[str] = field(default_factory=list)            # состав пакета
    docs: dict[str, DocSpec] = field(default_factory=dict)        # документ -> описание
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
    experience_ids: list[int] = field(default_factory=list)       # отобранный опыт, по порядку
    # Услуга закупки из каталога: по ней подбираются нормы часов и фильтруется
    # опыт. Одна услуга на тендер (КЭР; НДВ, НДС, НООЛР, ППЭЭ и ПЭК -- её этапы).
    service_id: int | None = None
    id: int | None = None


@dataclass
class LotItem:
    tender_id: int | None = None
    position: int = 0                 # порядок в списке
    name: str = ""
    price: Decimal = ZERO
    id: int | None = None


@dataclass
class TenderLaborRow:
    """Строка трудозатрат КП тендера. Численность и часы -- плановые величины
    под конкретный тендер, а не состав штата, поэтому хранятся при тендере."""
    tender_id: int | None = None
    position: int = 0                 # порядок в списке
    post: str = ""                    # должность
    headcount: int = 1                # человек
    hours: Decimal = ZERO             # часов на человека
    rate: Decimal = ZERO              # руб/ч
    id: int | None = None


COST_MATERIALS = "materials"
COST_SUBCONTRACT = "subcontract"


@dataclass
class TenderCostItem:
    """Строка материалов или субподряда в калькуляции тендера. Хранится
    разбивкой: форма калькуляции печатает обе статьи подтаблицами."""
    tender_id: int | None = None
    kind: str = COST_MATERIALS        # 'materials' | 'subcontract'
    position: int = 0                 # порядок внутри статьи
    name: str = ""
    unit: str = ""                    # единица измерения: «усл.ед.», «шт.»
    quantity: Decimal = Decimal("1")
    price: Decimal = ZERO             # цена за единицу, руб.
    note: str = ""
    id: int | None = None
