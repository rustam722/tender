"""Страница организации на rusprofile.ru: юридический адрес и коды статистики.

Страница находится по прямой ссылке, которую вставляет пользователь: поиск по ИНН
на сайте недоступен. Разбор идёт по видимым подписям блоков («Юридический адрес»,
«ОКПО», ...), а не по CSS-классам: классы меняются при каждой перевёрстке, подписи нет.
Любая неудача даёт None, а не исключение.
"""
import re
from dataclasses import dataclass
from urllib.parse import urlparse

import requests
from lxml import etree, html

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
ALLOWED_HOSTS = ("www.rusprofile.ru", "rusprofile.ru")


@dataclass
class RusprofileRecord:
    legal_address: str
    okpo: str
    oktmo: str
    okved: str
    msp_status_text: str


def normalize_url(url) -> str | None:
    """Ссылка на rusprofile.ru, достроенная до https://; None, если ссылка ведёт не туда.

    Допускает вставку без схемы (`www.rusprofile.ru/id/1`) и со схемой http://.
    """
    if not isinstance(url, str):
        return None
    url = url.strip()
    if not url:
        return None
    if "://" not in url:
        url = "https://" + url
    elif url.lower().startswith("http://"):
        url = "https://" + url[len("http://"):]
    try:
        parts = urlparse(url)
        ok = (
            parts.scheme == "https"
            and parts.hostname in ALLOWED_HOSTS
            and parts.port in (None, 443)
        )
    except ValueError:
        return None
    return url if ok else None


def _norm(text: str) -> str:
    return " ".join(text.split())


def _value_after(root, label: str) -> str:
    """Текст элемента, следующего за подписью `label`; пусто, если подписи нет."""
    for el in root.iter():
        if not isinstance(el.tag, str) or el.tag in ("script", "style"):
            continue
        if _norm(el.text or "") != label:
            continue
        for sibling in el.itersiblings():
            if not isinstance(sibling.tag, str):
                continue
            value = _norm(sibling.text_content())
            if value:
                return value
    return ""


def _digits(value: str) -> str:
    match = re.search(r"\d+", value)
    return match.group(0) if match else ""


def _okved(value: str) -> str:
    match = re.search(r"\b\d{2}\.\d{1,2}(?:\.\d{1,2})?\b", value)
    return match.group(0) if match else ""


def _parse(content: bytes) -> RusprofileRecord | None:
    try:
        try:
            document = content.decode("utf-8")  # без этого lxml читает байты без <meta> как latin-1
        except UnicodeDecodeError:
            document = content
        root = html.fromstring(document)
    except (etree.ParserError, etree.XMLSyntaxError, ValueError):
        return None
    record = RusprofileRecord(
        legal_address=_value_after(root, "Юридический адрес"),
        okpo=_digits(_value_after(root, "ОКПО")),
        oktmo=_digits(_value_after(root, "ОКТМО")),
        okved=_okved(_value_after(root, "Основной вид деятельности")),
        msp_status_text=re.sub(r"^Статус:\s*", "", _value_after(root, "Реестр МСП")),
    )
    if not (record.legal_address or record.okpo or record.oktmo):
        return None  # разметка не узнана: пустая карточка не должна выглядеть успехом
    return record


def fetch(url: str, *, timeout: float = 30.0, session=None) -> RusprofileRecord | None:
    """Читает страницу организации по ссылке (схему https:// можно не писать). None при любой неудаче."""
    url = normalize_url(url)
    if url is None:
        return None
    headers = {"User-Agent": USER_AGENT}
    try:
        if session is None:
            with requests.Session() as own:
                response = own.get(url, headers=headers, timeout=timeout)
        else:
            response = session.get(url, headers=headers, timeout=timeout)
    except requests.RequestException:
        return None
    if response.status_code != 200:
        return None
    return _parse(response.content)
