"""Бланки из приложений к полученной документации ЗВКС и Инфраструктурных проектов."""
from dataclasses import dataclass

from app.models import ROLE_CUSTOMER, DocSpec, Organization


BID = "Заявка на участие"
OFFER = "ТКП по форме заказчика"
PROFILE = "Анкета участника"
INVENTORY = "Опись документов"
SCHEDULE = "Календарный график"
AFFILIATION = "Справка об аффилированности"
STAFF = "Справка о кадровых ресурсах"
ASSETS = "Справка о МТР"
EXPERIENCE = "Справка об опыте"
COMMERCIAL = "Коммерческое предложение"

OFFER_NUMBER = "Номер письма о подаче оферты"
LOT = "Номер лота"
BEGIN = "Начало оказания услуг по графику"
END = "Окончание оказания услуг по графику"
APPROVE_OFFER = "Оферта — подтверждение условий"
APPROVE_BID = "ЗВКС — подтверждение условий заявки"
APPROVE_AFFILIATION = "Аффилированность — подтверждение сведений"
AFFILIATION_DETAILS = "Аффилированность — сведения или отсутствие признаков"
ROLE = "Оферта — роль участника"
TECH_DETAILS = "Предложение — описание работ и объём по ТЗ"


@dataclass(frozen=True)
class SourceForm:
    document: str
    filename: str
    prefix: str
    conditions: tuple[str, ...] = ()
    confirmation: str = ""


ZVKS_FORMS = (
    SourceForm(BID, "ЗВКС — форма 1 Заявка.docx", "Часть 2 — форма 1. ", confirmation=APPROVE_BID),
    SourceForm(OFFER, "ЗВКС — форма 2 Техническое предложение.docx", "Часть 1 — форма 2. ",
               (TECH_DETAILS,)),
    SourceForm(PROFILE, "ЗВКС — форма 3 Сведения об участнике.docx", "Часть 2 — форма 3. "),
    SourceForm(INVENTORY, "ЗВКС — форма 5 Опись.docx", "Часть 2 — форма 5. "),
)
IP_OFFER_TEMPLATE = "Инфраструктурные проекты — форма 2 Оферта.docx"
IP_STAFF_TEMPLATE = "Инфраструктурные проекты — форма 9 Кадры.docx"
IP_ASSETS_TEMPLATE = "Инфраструктурные проекты — форма 8 МТР.docx"
IP_FORMS = (
    SourceForm(INVENTORY, "Инфраструктурные проекты — форма 1 Опись.docx", "Форма 1. "),
    SourceForm(BID, IP_OFFER_TEMPLATE, "Форма 2. ",
               (OFFER_NUMBER, LOT, ROLE), APPROVE_OFFER),
    SourceForm(COMMERCIAL, "Инфраструктурные проекты — форма 3 Коммерческое предложение.docx",
               "Форма 3. ", (OFFER_NUMBER, LOT)),
    SourceForm(OFFER, "Инфраструктурные проекты — форма 4 Техническое предложение.docx",
               "Форма 4. ", (OFFER_NUMBER, LOT, TECH_DETAILS)),
    SourceForm(SCHEDULE, "Инфраструктурные проекты — форма 5 График.docx",
               "Форма 5. ", (OFFER_NUMBER, LOT, BEGIN, END)),
    SourceForm(PROFILE, "Инфраструктурные проекты — форма 6 Анкета.docx", "Форма 6. ",
               (OFFER_NUMBER,)),
    SourceForm(EXPERIENCE, "Инфраструктурные проекты — форма 7 Опыт.docx", "Форма 7. ",
               (OFFER_NUMBER,)),
    SourceForm(ASSETS, IP_ASSETS_TEMPLATE, "Форма 8. ", (OFFER_NUMBER,)),
    SourceForm(STAFF, IP_STAFF_TEMPLATE, "Форма 9. ", (OFFER_NUMBER,)),
    SourceForm(AFFILIATION, "Инфраструктурные проекты — форма 10 Аффилированность.docx",
               "Форма 10. ", (OFFER_NUMBER, AFFILIATION_DETAILS), APPROVE_AFFILIATION),
)
_FORMS = {entry.filename: entry for entry in (*ZVKS_FORMS, *IP_FORMS)}

CUSTOMERS = (
    (Organization(role=ROLE_CUSTOMER, short_name="АО «ЗВКС»",
                  full_name="Акционерное общество «ЗВКС»",
                  legal_address="422541, Республика Татарстан, г. Зеленодольск, ул. Озерная, д. 48"),
     ZVKS_FORMS),
    (Organization(role=ROLE_CUSTOMER, short_name="АО «Инфраструктурные проекты»",
                  full_name="Акционерное общество «Инфраструктурные проекты»"),
     IP_FORMS),
)


def word_form(document: str, spec: DocSpec) -> SourceForm | None:
    form = _FORMS.get(spec.template)
    return form if form and form.document == document else None


def customer_specs():
    for org, forms in CUSTOMERS:
        yield org, {
            form.document: DocSpec(template=form.filename, prefix=f"Черновик {form.prefix}")
            for form in forms
        }
