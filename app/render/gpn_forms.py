"""Настройки дополнительных Word-форм из полученных закупок ГПН."""
from dataclasses import dataclass

from app.models import DocSpec


OWNER_CONSENT = "Согласие собственника на обработку ПД"
REPRESENTATIVE_CONSENT = "Согласие представителя на обработку ПД"
OWNERSHIP = "Информация о цепочке собственников"
COMMITMENTS = "Справка об обязательствах в текущем году"
PROXY = "Доверенность на участие в отборе"
SME_DECLARATION = "Декларация СМСП"
BID_CONSENT = "Согласие участника закупки"
AUDIT = "Готовность к техническому аудиту"
CONFIDENTIALITY = "Соглашение о конфиденциальности"
DOCUMENT_NAMES = (
    OWNER_CONSENT, REPRESENTATIVE_CONSENT, OWNERSHIP, COMMITMENTS,
    PROXY, SME_DECLARATION, BID_CONSENT, AUDIT, CONFIDENTIALITY,
)

LOT_NUMBER = "Номер лота"
BID_AGREEMENT = "Заявка — условия договора (1: согласие; 2: кроме арбитражной оговорки)"
BID_APPROVAL = "Заявка — одобрение сделки (1: не требуется; 2: приложено; 3: до договора)"
BID_ATTACHMENTS = "Заявка — наименования приложений и число страниц"
BID_TEMPLATE_NAMES = frozenset({
    "ННГ — форма 1 Заявка.docx",
    "Хантос — форма 1 Заявка.docx",
    "Славнефть-МНГ — форма 1 Заявка.docx",
})
OPERATOR = "ПД — полное наименование оператора"
OPERATOR_ADDRESS = "ПД — юридический адрес оператора"
PD_CUSTOMER = "ПД — полное наименование заказчика"
PD_CUSTOMER_ADDRESS = "ПД — юридический адрес заказчика"
ORGANIZER = "Наименование организатора отбора"
AGREE_BID = "Согласие участника — для сборки введите «Подтверждаю»"
AGREE_CONFIDENTIALITY = "Соглашение — подтверждение условий и роли участника"
NDA_OTHER_PARTY = "Соглашение — полное наименование другой стороны"
AGREE_AUDIT = "Готовность к аудиту — для сборки введите «Подтверждаю»"
AGREE_SME = "Принадлежность к СМСП — для сборки введите «Подтверждаю»"
SME_CATEGORY = "СМСП — категория по действующим сведениям"
AUDIT_ADDRESS = "Технический аудит — адрес базы"
AUDIT_RESOURCES = "Технический аудит — ресурсы по методике"
AUDIT_CONTACT = "Технический аудит — контактное лицо и контакты"
SIMILAR_SUBJECT = "Обязательства — что считается аналогичным предметом"


@dataclass(frozen=True)
class ExtraWordForm:
    document: str
    filename: str
    form_number: str
    customers: tuple[str, ...]
    conditions: tuple[str, ...] = ()
    confirmation: str = ""


_ALL = ("ННГ", "Хантос", "Славнефть-МНГ")
EXTRA_FORMS = (
    ExtraWordForm(OWNER_CONSENT, "ННГ — форма 3 Согласие собственника.docx", "3",
                  ("ННГ",), (PD_CUSTOMER, PD_CUSTOMER_ADDRESS)),
    ExtraWordForm(OWNER_CONSENT, "Хантос — форма 3 Согласие собственника.docx", "3",
                  ("Хантос",), (PD_CUSTOMER, PD_CUSTOMER_ADDRESS)),
    ExtraWordForm(OWNER_CONSENT, "Славнефть-МНГ — форма 3 Согласие собственника.docx", "3",
                  ("Славнефть-МНГ",), (OPERATOR, OPERATOR_ADDRESS)),
    ExtraWordForm(REPRESENTATIVE_CONSENT, "Славнефть-МНГ — форма 3.1 Согласие представителя.docx",
                  "3.1", ("Славнефть-МНГ",), (OPERATOR, OPERATOR_ADDRESS)),
    ExtraWordForm(OWNERSHIP, "ГПН — форма 4 Цепочка собственников.docx", "4", _ALL),
    ExtraWordForm(COMMITMENTS, "ГПН — форма 7 Текущие обязательства.docx", "7",
                  ("ННГ", "Славнефть-МНГ"), (LOT_NUMBER, SIMILAR_SUBJECT)),
    ExtraWordForm(PROXY, "ГПН — форма 9 Доверенность.docx", "9", _ALL, (ORGANIZER,)),
    ExtraWordForm(SME_DECLARATION, "ННГ — форма 11 Декларация СМСП.docx", "11",
                  ("ННГ",), (SME_CATEGORY,), AGREE_SME),
    ExtraWordForm(SME_DECLARATION, "Хантос — форма 11 Декларация СМСП.docx", "11",
                  ("Хантос",), (SME_CATEGORY,), AGREE_SME),
    ExtraWordForm(BID_CONSENT, "ННГ и Хантос — форма 15 Согласие участника.docx", "15",
                  ("ННГ", "Хантос"), (ORGANIZER,), AGREE_BID),
    ExtraWordForm(BID_CONSENT, "Славнефть-МНГ — форма 12 Согласие участника.docx", "12",
                  ("Славнефть-МНГ",), (ORGANIZER,), AGREE_BID),
    ExtraWordForm(CONFIDENTIALITY, "Хантос — форма 10 Соглашение.docx", "10",
                  ("Хантос",), (NDA_OTHER_PARTY,), AGREE_CONFIDENTIALITY),
    ExtraWordForm(CONFIDENTIALITY, "ННГ и Славнефть — форма 10 Контрагент-резидент.docx", "10",
                  ("ННГ", "Славнефть-МНГ"), (NDA_OTHER_PARTY,), AGREE_CONFIDENTIALITY),
    ExtraWordForm(AUDIT, "ННГ — форма 19 Технический аудит.docx", "19",
                  ("ННГ",), (LOT_NUMBER, AUDIT_ADDRESS, AUDIT_RESOURCES, AUDIT_CONTACT), AGREE_AUDIT),
    ExtraWordForm(AUDIT, "Славнефть-МНГ — форма 16 Технический аудит.docx", "16",
                  ("Славнефть-МНГ",), (LOT_NUMBER, AUDIT_ADDRESS, AUDIT_RESOURCES, AUDIT_CONTACT), AGREE_AUDIT),
)
_FORMS_BY_FILENAME = {form.filename: form for form in EXTRA_FORMS}


def extra_word_form(document: str, spec: DocSpec) -> ExtraWordForm | None:
    form = _FORMS_BY_FILENAME.get(spec.template)
    return form if form and form.document == document else None


def extra_document_specs(customer: str) -> dict[str, DocSpec]:
    return {
        form.document: DocSpec(template=form.filename, prefix=f"Черновик форма {form.form_number}. ")
        for form in EXTRA_FORMS if customer in form.customers
    }
