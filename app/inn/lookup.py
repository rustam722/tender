"""Черновик карточки организации: реквизиты ФНС плюс адрес и коды с rusprofile.

Банк, телефоны, почта и сайт не отдаёт ни один источник (в ЕГРЮЛ их нет, rusprofile
прячет контакты за платным доступом), поэтому они остаются пустыми и попадают
в предупреждения. Не БД: ничего не сохраняется, карточку сохраняет вызывающий.
"""
from app.inn import fns, rusprofile
from app.models import ROLE_CUSTOMER, Organization

_PAGE_FIELDS = "юридический адрес, ОКПО, ОКТМО, ОКВЭД и статус МСП"


def _pick_record(inn: str, records: list[fns.FnsRecord]) -> tuple[fns.FnsRecord, list[str]]:
    matching = [r for r in records if r.inn == inn]
    if not matching:
        raise fns.InnNotFound("ФНС вернула записи, но ИНН в них не совпадает с запрошенным. Проверьте ИНН.")
    warnings = []
    if len(matching) > 1:
        warnings.append(
            "В реестре ФНС по этому ИНН найдено несколько записей. "
            "Взята первая, проверьте реквизиты."
        )
    return matching[0], warnings


def _normalize_position(post: str) -> str:
    """«ГЕНЕРАЛЬНЫЙ ДИРЕКТОР» -> «Генеральный директор»: ФНС отдаёт должность capslock'ом."""
    post = post.strip()
    return post[:1].upper() + post[1:].lower()


def _msp_code(text: str) -> str:
    lowered = text.lower()
    for prefix, code in (
        ("микропредприятие", "micro"),
        ("малое предприятие", "small"),
        ("среднее предприятие", "medium"),
    ):
        if lowered.startswith(prefix):
            return code
    return ""


def _apply_rusprofile(org: Organization, rp: rusprofile.RusprofileRecord) -> list[str]:
    org.legal_address = rp.legal_address
    org.okpo = rp.okpo
    org.oktmo = rp.oktmo
    org.okved = rp.okved
    org.msp_status = _msp_code(rp.msp_status_text)
    warnings = []
    for value, name in (
        (org.legal_address, "Юридический адрес"),
        (org.okpo, "ОКПО"),
        (org.oktmo, "ОКТМО"),
        (org.okved, "ОКВЭД"),
    ):
        if not value:
            warnings.append(f"На странице rusprofile не найдено поле «{name}». Введите его вручную.")
    if not org.msp_status:
        if rp.msp_status_text:
            warnings.append(
                f"Статус МСП на rusprofile указан как «{rp.msp_status_text}» и не распознан. "
                "Выберите статус вручную."
            )
        else:
            warnings.append("На странице rusprofile не найден статус МСП. Укажите его вручную.")
    return warnings


def build_draft(
    inn: str, *, rusprofile_url: str | None = None, session=None
) -> tuple[Organization, list[str]]:
    """Собирает черновик карточки заказчика. Ошибки ФНС поднимаются, rusprofile даёт лишь предупреждения."""
    record, warnings = _pick_record(inn, fns.lookup(inn, session=session))
    org = Organization(
        role=ROLE_CUSTOMER,
        inn=record.inn,
        kpp=record.kpp,
        ogrn=record.ogrn,
        short_name=record.short_name,
        full_name=record.full_name,
        head_name=record.head_fio,
        head_position=_normalize_position(record.head_post),
    )
    if not rusprofile_url:
        warnings.append(
            f"Ссылка на страницу организации на rusprofile не указана, поэтому не заполнены: "
            f"{_PAGE_FIELDS}. Вставьте ссылку вида https://www.rusprofile.ru/id/… "
            "или введите эти данные вручную."
        )
    elif rusprofile.normalize_url(rusprofile_url) is None:
        warnings.append(
            f"Ссылка должна вести на rusprofile.ru (например, https://www.rusprofile.ru/id/…). "
            f"Страница не прочитана, поэтому не заполнены: {_PAGE_FIELDS}. "
            "Исправьте ссылку или введите эти данные вручную."
        )
    else:
        page = rusprofile.fetch(rusprofile_url, session=session)
        if page is None:
            warnings.append(
                f"Не удалось прочитать страницу rusprofile по указанной ссылке (сайт недоступен "
                f"или страница устроена иначе), поэтому не заполнены: {_PAGE_FIELDS}. "
                "Повторите позже или введите эти данные вручную."
            )
        else:
            warnings.extend(_apply_rusprofile(org, page))
    warnings.append(
        "Банковские реквизиты (банк, БИК, расчётный и корреспондентский счета) "
        "не найдены ни в одном источнике. Введите их вручную."
    )
    warnings.append(
        "Основание полномочий руководителя (например, «на основании Устава») "
        "не найдено ни в одном источнике. Введите его вручную."
    )
    warnings.append(
        "Почтовый адрес не найден ни в одном источнике. Введите его вручную."
    )
    warnings.append(
        "Телефоны, e-mail и сайт организации не найдены ни в одном источнике "
        "(на rusprofile контакты скрыты за платным доступом). Введите их вручную."
    )
    return org, warnings
