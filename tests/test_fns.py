import json
from datetime import date
from pathlib import Path

import pytest
import requests

from app.inn.fns import (
    CaptchaRequired,
    FnsRecord,
    InnLookupError,
    InnNotFound,
    SourceUnavailable,
    lookup,
)

FIXTURE = json.loads(
    (Path(__file__).parent / "fixtures" / "fns_response.json").read_text(encoding="utf-8")
)


class StubResponse:
    def __init__(self, payload=None, status_code=200, text=None):
        self.status_code = status_code
        self._payload = payload
        self._text = text

    def json(self):
        if self._text is not None:
            return json.loads(self._text)
        return self._payload


class StubSession:
    """Stands in for requests.Session: only .post() and .get(), and records every call."""

    def __init__(self, post_response=None, get_response=None, post_error=None, get_error=None):
        self.post_response = post_response
        self.get_response = get_response
        self.post_error = post_error
        self.get_error = get_error
        self.calls = []

    def post(self, url, **kwargs):
        self.calls.append(("POST", url, kwargs))
        if self.post_error:
            raise self.post_error
        return self.post_response

    def get(self, url, **kwargs):
        self.calls.append(("GET", url, kwargs))
        if self.get_error:
            raise self.get_error
        return self.get_response


def stub(search_result, token="TOKEN", captcha=False):
    return StubSession(
        post_response=StubResponse({"t": token, "captchaRequired": captcha}),
        get_response=StubResponse(search_result),
    )


def test_parses_fixture():
    session = stub(FIXTURE)
    r = lookup("5904188041", session=session)[0]
    assert isinstance(r, FnsRecord)
    assert r.short_name == 'ООО "ТЕХНОСФЕРА"'
    assert r.full_name == 'ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ "ТЕХНОСФЕРА"'
    assert (r.inn, r.kpp, r.ogrn) == ("5904188041", "590401001", "1085904010224")
    assert r.head_post == "ДИРЕКТОР"
    assert r.head_fio == "Рабичко Елена Михайловна"
    assert r.registered == date(2008, 5, 8)
    assert r.region == "Пермский край"


def test_protocol_two_requests_with_token_and_browser_agent():
    session = stub(FIXTURE, token="ABC123")
    lookup("5904188041", timeout=7.0, session=session)
    (m1, url1, kw1), (m2, url2, kw2) = session.calls
    assert (m1, url1) == ("POST", "https://egrul.nalog.ru/")
    assert kw1["data"] == {"query": "5904188041"}
    assert (m2, url2) == ("GET", "https://egrul.nalog.ru/search-result/ABC123")
    assert kw1["timeout"] == kw2["timeout"] == 7.0
    for kw in (kw1, kw2):
        assert "Mozilla" in kw["headers"]["User-Agent"]


def test_empty_rows_raise_inn_not_found():
    with pytest.raises(InnNotFound) as exc:
        lookup("5904188041", session=stub({"rows": []}))
    assert exc.value.args[0] == "Организация с таким ИНН не найдена"
    assert isinstance(exc.value, InnLookupError)


def test_captcha_required_raises_captcha():
    session = stub(FIXTURE, captcha=True)
    with pytest.raises(CaptchaRequired) as exc:
        lookup("5904188041", session=session)
    assert isinstance(exc.value, InnLookupError)
    assert exc.value.args[0]
    assert [c[0] for c in session.calls] == ["POST"]  # no second request


def test_timeout_raises_source_unavailable():
    session = StubSession(post_error=requests.Timeout("slow"))
    with pytest.raises(SourceUnavailable) as exc:
        lookup("5904188041", session=session)
    assert isinstance(exc.value, InnLookupError)
    assert "не отвечает" in exc.value.args[0]


def test_timeout_on_second_request_raises_source_unavailable():
    session = StubSession(
        post_response=StubResponse({"t": "X", "captchaRequired": False}),
        get_error=requests.Timeout("slow"),
    )
    with pytest.raises(SourceUnavailable):
        lookup("5904188041", session=session)


def test_connection_error_raises_source_unavailable():
    session = StubSession(post_error=requests.ConnectionError("down"))
    with pytest.raises(SourceUnavailable):
        lookup("5904188041", session=session)


def test_non_200_raises_source_unavailable():
    session = StubSession(post_response=StubResponse(status_code=503))
    with pytest.raises(SourceUnavailable) as exc:
        lookup("5904188041", session=session)
    assert "503" in exc.value.args[0]


def test_non_200_on_search_result_raises_source_unavailable():
    session = StubSession(
        post_response=StubResponse({"t": "X", "captchaRequired": False}),
        get_response=StubResponse(status_code=500),
    )
    with pytest.raises(SourceUnavailable):
        lookup("5904188041", session=session)


def test_unparseable_json_raises_source_unavailable():
    session = StubSession(post_response=StubResponse(text="<html>not json</html>"))
    with pytest.raises(SourceUnavailable):
        lookup("5904188041", session=session)


def test_unexpected_json_shape_raises_source_unavailable():
    with pytest.raises(SourceUnavailable):
        lookup("5904188041", session=stub({"status": "wait"}))


def test_several_rows_are_all_returned():
    second = dict(FIXTURE["rows"][0], c='ООО "ВТОРАЯ"', i="5904188042")
    result = lookup("5904188041", session=stub({"rows": [FIXTURE["rows"][0], second]}))
    assert [r.short_name for r in result] == ['ООО "ТЕХНОСФЕРА"', 'ООО "ВТОРАЯ"']
    assert result[1].inn == "5904188042"


def test_head_without_colon_does_not_crash():
    row = dict(FIXTURE["rows"][0], g="Рабичко Елена Михайловна")
    r = lookup("5904188041", session=stub({"rows": [row]}))[0]
    assert r.head_post == ""
    assert r.head_fio == "Рабичко Елена Михайловна"


def test_head_only_splits_on_first_colon():
    row = dict(FIXTURE["rows"][0], g="ДИРЕКТОР: Иванов: Иван")
    r = lookup("5904188041", session=stub({"rows": [row]}))[0]
    assert (r.head_post, r.head_fio) == ("ДИРЕКТОР", "Иванов: Иван")


def test_missing_optional_fields_give_empty_values():
    row = {"c": "ИП Иванов", "n": "ИП Иванов Иван Иванович", "i": "590400000000", "k": "fl"}
    r = lookup("590400000000", session=stub({"rows": [row]}))[0]
    assert (r.kpp, r.ogrn, r.head_post, r.head_fio, r.region) == ("", "", "", "", "")
    assert r.registered is None


def test_unparseable_registration_date_is_none():
    row = dict(FIXTURE["rows"][0], r="не дата")
    assert lookup("5904188041", session=stub({"rows": [row]}))[0].registered is None


@pytest.mark.parametrize("bad", ["", "123", "59041880411", "590418804a", "5904 188041", "5904188041 ", None, 5904188041])
def test_rejects_malformed_inn(bad):
    session = stub(FIXTURE)
    with pytest.raises(InnLookupError) as exc:
        lookup(bad, session=session)
    assert "10 или 12 цифр" in exc.value.args[0]
    assert session.calls == []  # rejected before any request


def test_accepts_twelve_digit_inn():
    session = stub(FIXTURE)
    lookup("590400000000", session=session)
    assert session.calls[0][2]["data"] == {"query": "590400000000"}


def test_non_object_row_raises_source_unavailable():
    with pytest.raises(SourceUnavailable):
        lookup("5904188041", session=stub({"rows": ["oops"]}))
