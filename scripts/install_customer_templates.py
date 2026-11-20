"""Устанавливает формы первой партии в существующую базу программы.

Запуск из папки программы: python -m scripts.install_customer_templates
"""
import sys
from dataclasses import replace

from app import config, db
from app.models import ROLE_CUSTOMER, CustomerProfile, DocSpec, Organization
from app.render.xlsx_render import load_cellmap
from app.repo import Repo


BID = "Заявка на участие"
OFFER = "ТКП по форме заказчика"
PRICE_MAP = "gpn_offer_form2.json"


def _gpn_forms(label, price_label, assessment_number):
    return {
        BID: DocSpec(template=f"{label} — форма 1 Заявка.docx", prefix="Черновик 1. "),
        OFFER: DocSpec(template=f"{price_label} — форма 2 Ценовое предложение.xlsx",
                       cellmap=PRICE_MAP, prefix="Черновик 2. "),
        "Справка о кадровых ресурсах": DocSpec(template="ГПН — форма 5 Кадровые ресурсы.docx",
                                                prefix="Черновик форма 5. "),
        "Справка об опыте": DocSpec(template="ГПН — форма 6 Опыт.docx",
                                  prefix="Черновик форма 6. "),
        "Справка о МТР": DocSpec(template="ГПН — форма 8 МТР.docx", prefix="Черновик форма 8. "),
        "Оценка соответствия КТ-576": DocSpec(
            template=f"{price_label} — форма {assessment_number} КТ-576.docx",
            prefix=f"Черновик форма {assessment_number}. "),
    }


CUSTOMER_TEMPLATES = (
    (Organization(role=ROLE_CUSTOMER, short_name="АО «Газпромнефть – ННГ»",
                  full_name="Акционерное общество «Газпромнефть-Ноябрьскнефтегаз»"),
     _gpn_forms("ННГ", "ННГ и Хантос", 18)),
    (Organization(role=ROLE_CUSTOMER, short_name="ООО «Газпромнефть-Хантос»",
                  full_name="Общество с ограниченной ответственностью «Газпромнефть-Хантос»"),
     _gpn_forms("Хантос", "ННГ и Хантос", 18)),
    (Organization(role=ROLE_CUSTOMER, short_name="ПАО «СН-МНГ»",
                  full_name="Публичное акционерное общество «Славнефть-Мегионнефтегаз»"),
     _gpn_forms("Славнефть-МНГ", "Славнефть-МНГ", 15)),
    (Organization(role=ROLE_CUSTOMER, short_name="КБФ АО «Гознак»",
                  full_name="Краснокамская бумажная фабрика – филиал АО «Гознак»",
                  inn="7813252159", kpp="591643001",
                  postal_address="617060, г. Краснокамск, ул. Школьная, дом 13"),
     {
         BID: DocSpec(template="Гознак КБФ — форма 1 Информация об участнике.docx",
                      prefix="Черновик 1. "),
         OFFER: DocSpec(template="Гознак КБФ — форма 2 Техническое предложение.docx",
                        prefix="Черновик 2. "),
     }),
)


def _name(value: str) -> str:
    return "".join(ch for ch in value.casefold().replace("ё", "е") if ch.isalnum())


def _matching_customer(organizations, source):
    source_names = {_name(source.short_name), _name(source.full_name)}
    return [
        org for org in organizations
        if source_names.intersection({_name(org.short_name), _name(org.full_name)})
        or (source.inn and source.kpp and (org.inn, org.kpp) == (source.inn, source.kpp))
    ]


def install(repo: Repo) -> list[str]:
    """Добавляет недостающие формы. Текущие записи и назначенные бланки сохраняются."""
    for _, specs in CUSTOMER_TEMPLATES:
        for spec in specs.values():
            source = config.TEMPLATES_DOCS_DIR / spec.template
            if not source.is_file():
                raise FileNotFoundError(f"Не найден шаблон «{source.name}».")
            if spec.cellmap:
                load_cellmap(config.CELLMAPS_DIR / spec.cellmap)

    organizations = repo.list_organizations(ROLE_CUSTOMER)
    messages = []
    for source, specs in CUSTOMER_TEMPLATES:
        matches = _matching_customer(organizations, source)
        if len(matches) > 1:
            messages.append(f"{source.short_name} — несколько совпадений, профиль пропущен.")
            continue
        customer_id = matches[0].id if matches else repo.save_organization(source)
        profile = repo.find_customer_profile(customer_id) or CustomerProfile(customer_id=customer_id)
        docs = dict(profile.docs)
        added = []
        for document, spec in specs.items():
            if document in docs:
                continue
            docs[document] = spec
            added.append(document)
        if added:
            # У нефтяных заказчиков подбирают приложения и ценовую форму по лоту.
            default = [key for key in added
                       if key == BID or (key == OFFER and not specs[key].cellmap)]
            documents = list(dict.fromkeys([*profile.documents, *default]))
            repo.save_customer_profile(replace(profile, docs=docs, documents=documents))
            messages.append(f"{source.short_name} — добавлено: {', '.join(added)}.")
        else:
            messages.append(f"{source.short_name} — формы уже назначены.")
    return messages


def main() -> int:
    if not config.DB_PATH.is_file():
        print("Сначала один раз запустите программу, чтобы создать рабочую базу.", file=sys.stderr)
        return 1
    connection = db.connect(config.DB_PATH)
    try:
        db.migrate(connection)
        for message in install(Repo(connection)):
            print(message)
    finally:
        connection.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
