import json
from pathlib import Path

import pytest
import requests

from app.inn import fns
from app.inn.fns import InnNotFound, SourceUnavailable
from app.inn.lookup import build_draft
from app.models import ROLE_CUSTOMER

FIXTURES = Path(__file__).parent / "fixtures"
FNS_ROW = json.loads((FIXTURES / "fns_response.json").read_text(encoding="utf-8"))["rows"][0]
RP_PAGE = (FIXTURES / "rusprofile.html").read_bytes()
RP_URL = "https://www.rusprofile.ru/id/2854025"


class Resp:
    def __init__(self, payload=None, content=b"", status_code=200):
        self.status_code = status_code
        self._payload = payload
        self.content = content

    def json(self):
        return self._payload


class RoutingSession:
    """One stub for both sources: egrul (POST + GET search-result) and rusprofile (GET page)."""

    def __init__(self, rows=None, rp_response=None, rp_error=None, fns_error=None):
        self.rows = [FNS_ROW] if rows is None else rows
        self.rp_response = rp_response if rp_response is not None else Resp(content=RP_PAGE)
        self.rp_error = rp_error
        self.fns_error = fns_error
        self.calls = []

    def post(self, url, **kwargs):
        self.calls.append(("POST", url))
        assert "egrul.nalog.ru" in url
        if self.fns_error:
            raise self.fns_error
        return Resp({"t": "TOKEN", "captchaRequired": False})

    def get(self, url, **kwargs):
        self.calls.append(("GET", url))
        if "egrul.nalog.ru" in url:
            return Resp({"rows": self.rows})
        assert url.startswith("https://www.rusprofile.ru/"), url
        if self.rp_error:
            raise self.rp_error
        return self.rp_response

    def hosts_called(self):
        return {url.split("/")[2] for _, url in self.calls}


def with_page(page):
    return RoutingSession(rp_response=Resp(content=page))


def test_draft_merges_both_sources():
    session = RoutingSession()
    org, warn = build_draft("5904188041", rusprofile_url=RP_URL, session=session)
    assert org.full_name.startswith("ОБЩЕСТВО")
    assert org.short_name == 'ООО "ТЕХНОСФЕРА"'
    assert (org.inn, org.kpp, org.ogrn) == ("5904188041", "590401001", "1085904010224")
    assert org.head_name == "Рабичко Елена Михайловна"
    assert org.head_position == "ДИРЕКТОР"
    assert "Куйбышева" in org.legal_address
    assert (org.okpo, org.oktmo, org.okved) == ("85081307", "57701000001", "71.12.13")
    assert org.msp_status == "micro"
    assert org.role == ROLE_CUSTOMER
    assert org.id is None
    assert not any("адрес" in w.lower() for w in warn)
    assert session.hosts_called() == {"egrul.nalog.ru", "www.rusprofile.ru"}


def test_draft_without_url_uses_fns_only():
    session = RoutingSession()
    org, warn = build_draft("5904188041", session=session)
    assert org.ogrn == "1085904010224"
    assert org.legal_address == ""
    assert org.okpo == org.oktmo == org.okved == org.msp_status == ""
    assert any("адрес" in w.lower() for w in warn)
    assert any("rusprofile" in w.lower() for w in warn)
    assert "www.rusprofile.ru" not in session.hosts_called()


@pytest.mark.parametrize(
    "rp_kwargs",
    [
        {"rp_response": Resp(content=b"", status_code=500)},
        {"rp_response": Resp(content=b"<html><body>nothing</body></html>")},
        {"rp_error": requests.ConnectionError("boom")},
    ],
)
def test_draft_survives_rusprofile_failure(rp_kwargs):
    org, warn = build_draft(
        "5904188041", rusprofile_url=RP_URL, session=RoutingSession(**rp_kwargs)
    )
    assert org.inn == "5904188041"
    assert org.full_name.startswith("ОБЩЕСТВО")
    assert org.legal_address == ""
    assert any("адрес" in w.lower() for w in warn)


def test_draft_with_foreign_url_makes_no_request_to_it_and_warns():
    session = RoutingSession()
    _, warn = build_draft("5904188041", rusprofile_url="https://example.com/x", session=session)
    assert session.hosts_called() == {"egrul.nalog.ru"}
    assert any("адрес" in w.lower() for w in warn)


def test_draft_partial_rusprofile_page_warns_about_missing_fields():
    page = "<html><body><dl><dt>ОКПО</dt><dd>12345678</dd></dl></body></html>".encode()
    org, warn = build_draft("5904188041", rusprofile_url=RP_URL, session=with_page(page))
    assert org.okpo == "12345678"
    assert org.legal_address == ""
    assert any("адрес" in w.lower() for w in warn)
    assert any("октмо" in w.lower() for w in warn)
    assert not any("окпо" in w.lower() for w in warn)


def test_draft_unrecognised_msp_text_is_left_empty_with_warning():
    page = RP_PAGE.replace("микропредприятие".encode(), "что-то новое".encode())
    org, warn = build_draft("5904188041", rusprofile_url=RP_URL, session=with_page(page))
    assert org.msp_status == ""
    assert any("мсп" in w.lower() for w in warn)


@pytest.mark.parametrize(
    "text, expected",
    [
        ("микропредприятие", "micro"),
        ("малое предприятие", "small"),
        ("среднее предприятие", "medium"),
    ],
)
def test_msp_text_maps_to_enum(text, expected):
    page = RP_PAGE.replace("микропредприятие".encode(), text.encode())
    org, _ = build_draft("5904188041", rusprofile_url=RP_URL, session=with_page(page))
    assert org.msp_status == expected


def test_draft_warns_about_bank_and_contacts():
    org, warn = build_draft("5904188041", session=RoutingSession())
    low = [w.lower() for w in warn]
    assert any("банк" in w for w in low)
    assert any("телефон" in w for w in low)
    assert any("e-mail" in w for w in low)
    assert any("сайт" in w for w in low)
    assert org.bank_name == org.bank_bik == org.bank_account == org.bank_corr_account == ""
    assert org.phones == org.email == org.website == ""


def test_warnings_are_russian_sentences():
    _, warn = build_draft("5904188041", session=RoutingSession())
    assert warn
    for w in warn:
        assert w.endswith(".")
        assert any("а" <= ch.lower() <= "я" for ch in w)


def test_draft_propagates_fns_not_found():
    with pytest.raises(InnNotFound):
        build_draft("5904188041", rusprofile_url=RP_URL, session=RoutingSession(rows=[]))


def test_draft_propagates_fns_unavailable():
    session = RoutingSession(fns_error=requests.ConnectionError("down"))
    with pytest.raises(SourceUnavailable):
        build_draft("5904188041", rusprofile_url=RP_URL, session=session)


def test_draft_propagates_malformed_inn_error():
    with pytest.raises(fns.InnLookupError):
        build_draft("123", session=RoutingSession())


def test_draft_does_not_touch_rusprofile_when_fns_fails():
    session = RoutingSession(rows=[])
    with pytest.raises(InnNotFound):
        build_draft("5904188041", rusprofile_url=RP_URL, session=session)
    assert "www.rusprofile.ru" not in session.hosts_called()


def test_draft_picks_row_matching_requested_inn():
    other = dict(FNS_ROW, i="5904000000", c='ООО "ДРУГАЯ"', n="ДРУГАЯ", o="1")
    org, warn = build_draft("5904188041", session=RoutingSession(rows=[other, FNS_ROW]))
    assert org.inn == "5904188041"
    assert org.short_name == 'ООО "ТЕХНОСФЕРА"'
    assert not any("нескольк" in w.lower() for w in warn)


def test_draft_with_several_matching_rows_takes_first_and_warns():
    twin = dict(FNS_ROW, c='ООО "ДВОЙНИК"', n="ДВОЙНИК")
    org, warn = build_draft("5904188041", session=RoutingSession(rows=[FNS_ROW, twin]))
    assert org.short_name == 'ООО "ТЕХНОСФЕРА"'
    assert any("нескольк" in w.lower() for w in warn)


def test_draft_with_no_row_matching_inn_raises_not_found():
    other = dict(FNS_ROW, i="5904000000")
    with pytest.raises(InnNotFound):
        build_draft("5904188041", session=RoutingSession(rows=[other]))


def test_unreadable_registration_date_does_not_break_draft():
    row = dict(FNS_ROW, r="")
    assert fns.lookup("5904188041", session=RoutingSession(rows=[row]))[0].registered is None
    org, _ = build_draft("5904188041", session=RoutingSession(rows=[row]))
    assert org.inn == "5904188041"
    assert not hasattr(org, "registered")
