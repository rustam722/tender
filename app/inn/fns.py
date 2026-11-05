"""Клиент ЕГРЮЛ/ЕГРИП (egrul.nalog.ru): реквизиты организации по ИНН."""
import re
from dataclasses import dataclass
from datetime import date, datetime

import requests

BASE_URL = "https://egrul.nalog.ru/"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


class InnLookupError(Exception):
    """Базовая ошибка поиска по ИНН; сообщение на русском в args[0]."""


class InnNotFound(InnLookupError):
    pass


class CaptchaRequired(InnLookupError):
    pass


class SourceUnavailable(InnLookupError):
    pass


@dataclass
class FnsRecord:
    short_name: str
    full_name: str
    inn: str
    kpp: str
    ogrn: str
    head_fio: str
    head_post: str
    registered: date | None
    region: str


def _split_head(raw: str) -> tuple[str, str]:
    """«ДИРЕКТОР: Иванов Иван» -> («ДИРЕКТОР», «Иванов Иван»); без двоеточия весь текст — ФИО."""
    post, sep, fio = raw.partition(":")
    if not sep:
        return "", raw.strip()
    return post.strip(), fio.strip()


def _parse_date(raw: str) -> date | None:
    try:
        return datetime.strptime(raw, "%d.%m.%Y").date()
    except ValueError:
        return None


def _parse_row(row: dict) -> FnsRecord:
    head_post, head_fio = _split_head(row.get("g") or "")
    return FnsRecord(
        short_name=row.get("c") or "",
        full_name=row.get("n") or "",
        inn=row.get("i") or "",
        kpp=row.get("p") or "",
        ogrn=row.get("o") or "",
        head_fio=head_fio,
        head_post=head_post,
        registered=_parse_date(row.get("r") or ""),
        region=row.get("rn") or "",
    )


def _json(response) -> dict:
    if response.status_code != 200:
        raise SourceUnavailable(
            f"Сервис ФНС ответил ошибкой (код {response.status_code}). Попробуйте позже."
        )
    try:
        data = response.json()
    except ValueError:
        raise SourceUnavailable("Сервис ФНС вернул ответ в непонятном виде. Попробуйте позже.")
    if not isinstance(data, dict):
        raise SourceUnavailable("Сервис ФНС вернул ответ в непонятном виде. Попробуйте позже.")
    return data


def _search(session, inn: str, timeout: float) -> list[dict]:
    headers = {"User-Agent": USER_AGENT}
    try:
        first = _json(session.post(BASE_URL, data={"query": inn}, headers=headers, timeout=timeout))
        if first.get("captchaRequired"):
            raise CaptchaRequired(
                "Сервис ФНС требует ввести капчу. Заполните реквизиты вручную или повторите позже."
            )
        token = first.get("t")
        if not token:
            raise SourceUnavailable("Сервис ФНС не вернул номер запроса. Попробуйте позже.")
        second = _json(
            session.get(f"{BASE_URL}search-result/{token}", headers=headers, timeout=timeout)
        )
    except requests.Timeout:
        raise SourceUnavailable("Сервис ФНС не отвечает. Попробуйте позже.")
    except requests.RequestException:
        raise SourceUnavailable("Не удалось связаться с сервисом ФНС. Проверьте подключение к интернету.")
    rows = second.get("rows")
    if not isinstance(rows, list) or not all(isinstance(r, dict) for r in rows):
        raise SourceUnavailable("Сервис ФНС вернул ответ в непонятном виде. Попробуйте позже.")
    return rows


def lookup(inn: str, *, timeout: float = 30.0, session=None) -> list[FnsRecord]:
    """Ищет организации по ИНН. Несколько записей возвращаются все — выбор за вызывающим."""
    if not isinstance(inn, str) or not re.fullmatch(r"[0-9]{10}|[0-9]{12}", inn):
        raise InnLookupError("ИНН должен состоять из 10 или 12 цифр.")
    if session is None:
        with requests.Session() as own:
            rows = _search(own, inn, timeout)
    else:
        rows = _search(session, inn, timeout)
    if not rows:
        raise InnNotFound("Организация с таким ИНН не найдена")
    return [_parse_row(row) for row in rows]
