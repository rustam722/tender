import re
from pathlib import Path

import pytest
import requests

from app.inn import rusprofile
from app.inn.rusprofile import RusprofileRecord

FIXTURE = (Path(__file__).parent / "fixtures" / "rusprofile.html").read_bytes()
RP_URL = "https://www.rusprofile.ru/id/2854025"


class StubResponse:
    def __init__(self, content=b"", status_code=200):
        self.status_code = status_code
        self.content = content


class StubSession:
    """Stands in for requests.Session: only .get(), records every call."""

    def __init__(self, response=None, error=None):
        self.response = response
        self.error = error
        self.calls = []

    def get(self, url, **kwargs):
        self.calls.append((url, kwargs))
        if self.error:
            raise self.error
        return self.response


def stub(content=FIXTURE, status_code=200):
    return StubSession(response=StubResponse(content, status_code))


def test_parses_address_and_codes():
    r = rusprofile.fetch(RP_URL, session=stub())
    assert isinstance(r, RusprofileRecord)
    assert r.legal_address == "614016, Пермский край, город Пермь, ул. Куйбышева, д. 67/1, 90"
    assert r.okpo == "85081307"
    assert r.oktmo == "57701000001"
    assert r.okved == "71.12.13"
    assert r.msp_status_text.startswith("микропредприятие")


def test_parsing_does_not_depend_on_css_classes():
    page = re.sub(rb'\s(?:class|id)="[^"]*"', b"", FIXTURE)
    assert b"company-info" not in page
    r = rusprofile.fetch(RP_URL, session=stub(page))
    assert "Куйбышева" in r.legal_address
    assert (r.okpo, r.oktmo, r.okved) == ("85081307", "57701000001", "71.12.13")
    assert r.msp_status_text.startswith("микропредприятие")


def test_request_goes_to_given_url_with_timeout_and_browser_agent():
    session = stub()
    rusprofile.fetch(RP_URL, timeout=7.0, session=session)
    ((url, kwargs),) = session.calls
    assert url == RP_URL
    assert kwargs["timeout"] == 7.0
    assert kwargs["headers"]["User-Agent"].startswith("Mozilla/5.0")


@pytest.mark.parametrize(
    "content", [b"", "<html><body><p>Привет</p></body></html>".encode(), b"not html at all"]
)
def test_unknown_markup_returns_none(content):
    assert rusprofile.fetch(RP_URL, session=stub(content)) is None


def test_page_with_only_some_fields_is_partial_not_none():
    page = "<html><body><dl><dt>ОКПО</dt><dd>12345678</dd></dl></body></html>".encode()
    r = rusprofile.fetch(RP_URL, session=stub(page))
    assert r.okpo == "12345678"
    assert (r.legal_address, r.oktmo, r.okved, r.msp_status_text) == ("", "", "", "")


def test_network_error_returns_none():
    session = StubSession(error=requests.ConnectionError("boom"))
    assert rusprofile.fetch(RP_URL, session=session) is None


def test_timeout_returns_none():
    session = StubSession(error=requests.Timeout("slow"))
    assert rusprofile.fetch(RP_URL, session=session) is None


@pytest.mark.parametrize("status", [403, 404, 429, 500])
def test_non_200_returns_none(status):
    assert rusprofile.fetch(RP_URL, session=stub(status_code=status)) is None


@pytest.mark.parametrize(
    "bad",
    [
        "",
        "5904188041",
        "https://example.com/id/2854025",
        "https://www.rusprofile.ru.evil.com/id/1",
        "file:///etc/passwd",
    ],
)
def test_foreign_or_malformed_url_returns_none_without_any_request(bad):
    session = stub()
    assert rusprofile.fetch(bad, session=session) is None
    assert session.calls == []


def test_accepts_bare_rusprofile_host():
    session = stub()
    assert rusprofile.fetch("https://rusprofile.ru/id/2854025", session=session) is not None


@pytest.mark.parametrize(
    "url",
    [
        "www.rusprofile.ru/id/2854025",
        "rusprofile.ru/id/2854025",
        "http://www.rusprofile.ru/id/2854025",
        "  https://www.rusprofile.ru/id/2854025\n",
    ],
)
def test_url_without_scheme_or_with_http_is_completed_to_https(url):
    session = stub()
    assert rusprofile.fetch(url, session=session) is not None
    ((called, _),) = session.calls
    assert called == "https://" + url.strip().split("://")[-1]
    assert rusprofile.normalize_url(url) == called


@pytest.mark.parametrize(
    "bad", ["", "   ", "5904188041", "example.com/id/1", "rusprofile.ru.evil.com/id/1", "javascript:alert(1)"]
)
def test_normalize_url_rejects_foreign_hosts(bad):
    assert rusprofile.normalize_url(bad) is None
    session = stub()
    assert rusprofile.fetch(bad, session=session) is None
    assert session.calls == []
