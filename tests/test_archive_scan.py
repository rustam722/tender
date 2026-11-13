"""Сканер архива прошлых закупок: по ИМЕНАМ папок и файлов, содержимого не читает.

Архив фирмы -- сотни закупок в папках вида «Заказчик до 04.09.2025_исход», в
каждой документация заказчика и поданная заявка. Сканер раскладывает это в
записи; остальное (опыт, кадры, нормы часов) -- следующие шаги, и они читают
файлы уже сами и с проверкой человека.

Три правила, на которых держится доверие к такой записи:
  1. неизвестное остаётся неизвестным: исход, которого нет в названии, не
     «проигран» и не «выигран», а None -- иначе статистика врёт;
  2. сканер не открывает ни одного файла и не пишет в архив;
  3. всё, что он назвал догадкой, названо догадкой (заказчик из имени папки).
"""
import builtins
import io
import os
from pathlib import Path

import pytest

from app import archive_scan
from app.archive_scan import ArchiveError, classify_file, scan_archive


# ---- помощники ----------------------------------------------------------------

def make(root: Path, *relative: str) -> None:
    """Файлы с пустым содержимым: сканер на содержимое смотреть не должен."""
    for rel in relative:
        path = root / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"")


def entry_of(root: Path, folder: str, group: str | None = None):
    found = [e for e in scan_archive(root) if e.folder == folder]
    assert len(found) == 1, [e.folder for e in scan_archive(root)]
    return found[0]


@pytest.fixture
def archive(tmp_path):
    root = tmp_path / "архив"
    root.mkdir()
    return root


# ---- срок подачи из названия ---------------------------------------------------

@pytest.mark.parametrize("folder, expected", [
    ("Газпромнефть-Хантос до 04.09.2025", "2025-09-04"),
    ("Газпромнефть-ННГ_до 02.12.25", "2025-12-02"),
    ("Афипский НПЗ_КЭР_до 26.12.25", "2025-12-26"),
    ("Инфраструктурные проекты до 18.10.2024", "2024-10-18"),
    ("МУП УЭ_Черноголовка_аукцион12.08.26", "2026-08-12"),     # без «до»
])
def test_a_folder_name_gives_its_deadline_in_iso_form(archive, folder, expected):
    make(archive, f"{folder}/Заявка/1.docx")
    assert entry_of(archive, folder).deadline == expected


@pytest.mark.parametrize("folder", [
    "Гипротрубопровод_ноябрь  2025", "Махачкала", "СКБ"])
def test_a_folder_without_a_date_has_no_deadline_and_none_is_guessed(archive, folder):
    make(archive, f"{folder}/Заявка/1.docx")
    assert entry_of(archive, folder).deadline is None


def test_the_date_after_do_is_the_deadline_not_the_first_date_in_the_name(archive):
    """В названии бывают две даты: «Проект 12.01.24 до 05.03.25». Срок --
    та, что после «до»; первая может быть годом проекта или чем угодно."""
    make(archive, "Проект 12.01.24 до 05.03.25/Заявка/1.docx")
    assert entry_of(archive, "Проект 12.01.24 до 05.03.25").deadline == "2025-03-05"


def test_an_impossible_date_in_a_name_is_not_a_deadline(archive):
    """«до 31.02.26» -- не срок, а опечатка: календарь такой даты не знает."""
    make(archive, "Тест до 31.02.26/Заявка/1.docx")
    assert entry_of(archive, "Тест до 31.02.26").deadline is None


# ---- исход -----------------------------------------------------------------------

@pytest.mark.parametrize("folder, outcome", [
    ("Омскоблгаз до 07.10.26 проиграли", "lost"),
    ("Харанорская ГРЭС до 12.02.2025_не выйграли", "lost"),
    ("Югорская КЭК - Белоярский до 25.04.2025_не выйграли", "lost"),
    ("Госветинспекция ПК до 09.04.2025_выйграли", "won"),
    ("Скотомогильники в Кировской области_выйграли частично", "won_partly"),
    ("Омский каучук до 10.03.2025_не подавали", "not_applied"),
    ("Газпром нефть Шельф до 25.09.2025_не подходим", "not_applied"),
    ("Приазовнефть до 12.05.2025_отказали", "refused"),
    ("Форвард-Энерго до 25.04.2025_не успели с оплатой", "missed"),
    ("Администрация г.Губкинского_отменен", "cancelled"),
])
def test_the_outcome_is_read_from_the_name_when_the_name_says_it(
        archive, folder, outcome):
    make(archive, f"{folder}/Заявка/1.docx")
    found = entry_of(archive, folder)
    assert found.outcome == outcome
    assert found.outcome_source == "название папки"


def test_not_won_is_not_mistaken_for_won(archive):
    """«не выйграли» содержит «выйграли»: порядок проверки решает, выиграна
    закупка или проиграна, и ошибка тут переворачивает смысл."""
    make(archive, "А до 01.02.25_не выйграли/Заявка/1.docx")
    assert entry_of(archive, "А до 01.02.25_не выйграли").outcome == "lost"


def test_a_tender_whose_name_is_silent_has_an_unknown_outcome_not_a_lost_one(archive):
    """Правило 1. 55 закупок из 82 в реальном архиве не говорят об исходе в
    названии; записать их «проигранными» значило бы исказить всю статистику."""
    make(archive, "Конданефть до 02.06.2025/Заявка/1.docx")
    found = entry_of(archive, "Конданефть до 02.06.2025")
    assert found.outcome is None and found.outcome_source is None


def test_a_tender_in_the_won_folder_is_won_even_when_its_name_is_silent(archive):
    make(archive, "! ВЫЙГРАЛИ/ЗВКС до 22.10.26/Заявка/1.docx")
    found = entry_of(archive, "ЗВКС до 22.10.26")
    assert found.outcome == "won"
    assert found.outcome_source == "папка «! ВЫЙГРАЛИ»"
    assert found.group == "! ВЫЙГРАЛИ"


def test_the_name_wins_over_the_folder_it_lies_in(archive):
    """Название конкретной закупки точнее, чем общая папка: «выиграли
    частично» внутри «! ВЫЙГРАЛИ» остаётся частичной победой."""
    make(archive, "! ВЫЙГРАЛИ/Х до 01.02.25_выйграли частично/Заявка/1.docx")
    assert entry_of(archive, "Х до 01.02.25_выйграли частично").outcome == "won_partly"


# ---- услуга ----------------------------------------------------------------------

@pytest.mark.parametrize("folder, services", [
    ("Афипский НПЗ_КЭР_до 26.12.25", ("КЭР",)),
    ("ВЧНГ_Дулисьма_СЗЗ до 18.11.25", ("СЗЗ",)),
    ("Пермская ГРЭС_НДВ до 26.01.26", ("НДВ",)),
    ("Конданефть до 02.06.2025", ()),
])
def test_the_service_is_read_from_the_name_as_a_whole_word(archive, folder, services):
    make(archive, f"{folder}/Заявка/1.docx")
    assert entry_of(archive, folder).services == services


def test_a_service_code_inside_a_longer_word_is_not_a_service(archive):
    """«НДВ» внутри «Андвик», «КЭР» внутри «Пакэрия» -- не услуги."""
    make(archive, "Андвик до 01.02.25/Заявка/1.docx", "Пакэрия до 01.02.25/Заявка/1.docx")
    assert entry_of(archive, "Андвик до 01.02.25").services == ()
    assert entry_of(archive, "Пакэрия до 01.02.25").services == ()


# ---- заказчик: это ДОГАДКА -------------------------------------------------------

@pytest.mark.parametrize("folder, guess", [
    ("Афипский НПЗ_КЭР_до 26.12.25", "Афипский НПЗ"),
    ("Исполком Актанышского МР до 06.02.2025", "Исполком Актанышского МР"),
    ("Омскоблгаз до 07.10.26 проиграли", "Омскоблгаз"),
    ("Харанорская ГРЭС до 12.02.2025_не выйграли", "Харанорская ГРЭС"),
])
def test_the_customer_is_a_guess_from_the_name_without_dates_services_and_outcomes(
        archive, folder, guess):
    make(archive, f"{folder}/Заявка/1.docx")
    assert entry_of(archive, folder).customer_guess == guess


# ---- папки внутри закупки ---------------------------------------------------------

@pytest.mark.parametrize("name", [
    "Заявка", "Заявка ООО Техносфера", "3_заявка экоэ", "Заявка лот 2",
    "Заявка ЗапСибЦЭП"])
def test_a_folder_named_like_a_bid_is_a_bid_folder(archive, name):
    make(archive, f"Т до 01.02.25/{name}/1.docx")
    assert entry_of(archive, "Т до 01.02.25").bid_folders == (name,)


@pytest.mark.parametrize("name", [
    "Закупочная документация", "Документация", "!!новая тендерная документация",
    "КБФ_КЭР_закупочная документация"])
def test_a_folder_named_like_documentation_is_a_documentation_folder(archive, name):
    make(archive, f"Т до 01.02.25/{name}/1.docx")
    found = entry_of(archive, "Т до 01.02.25")
    assert found.documentation_folders == (name,)
    assert found.bid_folders == ()


@pytest.mark.parametrize("name", ["Пример", "Заявка_пример", "заявка пример"])
def test_an_example_folder_is_neither_a_bid_nor_documentation(archive, name):
    """«Пример» -- чужая подача как образец, а не заявка фирмы: считать её
    поданной значило бы приписать фирме чужой состав."""
    make(archive, f"Т до 01.02.25/{name}/Устав.pdf")
    found = entry_of(archive, "Т до 01.02.25")
    assert found.bid_folders == () and found.bid_papers == ()


def test_a_tender_without_subfolders_keeps_its_files_unclassified(archive):
    make(archive, "Т до 01.02.25/Устав.pdf", "Т до 01.02.25/ТЗ.docx")
    found = entry_of(archive, "Т до 01.02.25")
    assert found.bid_papers == () and found.unclassified_files == 2


def test_the_papers_of_a_bid_are_found_at_any_depth(archive):
    make(archive, "Т до 01.02.25/Заявка/Часть 1/Декларация.docx",
         "Т до 01.02.25/Заявка/Часть 2/1_Устав.pdf")
    found = entry_of(archive, "Т до 01.02.25")
    assert [p.path for p in found.bid_papers] == [
        "Заявка/Часть 1/Декларация.docx", "Заявка/Часть 2/1_Устав.pdf"]


# ---- мусор -----------------------------------------------------------------------

def test_temporary_files_thumbnails_and_shortcuts_are_not_papers(archive):
    make(archive, "Т до 01.02.25/Заявка/~$едложение.docx",
         "Т до 01.02.25/Заявка/Thumbs.db", "Т до 01.02.25/Заявка/x.tmp",
         "Т до 01.02.25/Заявка/сайт.url", "Т до 01.02.25/Заявка/Предложение.docx")
    assert [p.path for p in entry_of(archive, "Т до 01.02.25").bid_papers] == [
        "Заявка/Предложение.docx"]


def test_an_electronic_signature_file_is_not_a_paper_of_its_own(archive):
    """«Устав.pdf.sig» -- подпись к «Уставу.pdf», а не ещё один документ:
    в составе заявки она считалась бы лишним листом."""
    make(archive, "Т до 01.02.25/Заявка/Устав.pdf", "Т до 01.02.25/Заявка/Устав.pdf.sig")
    assert [p.path for p in entry_of(archive, "Т до 01.02.25").bid_papers] == [
        "Заявка/Устав.pdf"]


# ---- вид документа по имени файла -------------------------------------------------

@pytest.mark.parametrize("name, kinds", [
    ("2_Выписка из ЕГРЮЛ.pdf", ("egrul",)),
    ("4_Устав с изменениями от 2025 г..pdf", ("charter",)),
    ("5_Свидетельство о регистрации.pdf", ("reg_certificate",)),
    ("6_Свидетельство о постановке на учет.pdf", ("tax_registration",)),
    ("5. Свидетельство о постановке на учет в налоговом органе.pdf", ("tax_registration",)),
    ("3_Протокол № 1.pdf", ("authority",)),
    ("3. Решение №1, Приказ №1.pdf", ("authority",)),
    ("7_Справка о крупности.docx", ("large_deal",)),
    ("8_Справка о заинтересованности.docx", ("interest",)),
    ("6.Письмо о крупной сделке и заинтересованности.docx", ("large_deal", "interest")),
    ("12_Справка о МТР.docx", ("assets",)),
    ("13_Справка о кадрах.doc", ("staff",)),
    ("8_Приложение к справке о кадрах.pdf", ("staff",)),
    ("14_Справка об опыте.doc", ("experience",)),
    ("9_Сведения об опыте.doc", ("experience",)),
    ("15_Договор с ООО Лаборатория.pdf", ("contract",)),
    ("11_Уведомлеие о применении УСН.pdf", ("usn_notice",)),
    ("10_Копия бухгалтерского баланса и отчета о финансовых результатах за 2025.pdf",
     ("balance",)),
    ("10_Справка налоговой на 20.05.26.pdf", ("tax_clearance",)),
    ("17_Аттестат с областью аккредитации.pdf", ("accreditation",)),
    ("4_Выписка из реестра МСП_14.10.26.pdf", ("msp_register",)),
    ("Декларация.docx", ("declaration",)),
    ("01.Информация об участнике.docx", ("participant_info",)),
    ("1.Предложение о функциональных характеристиках.docx", ("proposal",)),
    ("Опись.docx", ("inventory",)),
    ("0_Опись.pdf", ("inventory",)),
    ("Надпись на конверт.docx", ("envelope",)),
    ("Заявка Гознак.zip", ("archive",)),
    ("doc04220920260609131312.pdf", ("other",)),
    # --- имена из настоящего архива, которые первая версия правил не узнала:
    # треть файлов поданных заявок оставалась «прочим», и среди них --
    # регулярные документы.
    ("Согласие на обработку персональных данных.docx", ("consent",)),
    ("5. Согласие на обработку ПД.docx", ("consent",)),
    ("16. Согласие на обр. перс. данных.doc", ("consent",)),
    ("7.3_Согласие на обработку данных сотрудников.docx", ("consent",)),
    ("Сведения о цепочке собственников.docx", ("ownership_chain",)),
    ("4_Информация о цепочке собственников.pdf", ("ownership_chain",)),
    ("Карточка предприятия.pdf", ("company_card",)),
    ("10_Справка о деловой репутации.docx.pdf", ("reputation",)),
    ("9.1_Сведения о субподрядчиках.docx", ("subcontract_info",)),
    ("9.2_Документы субподрядчика 1.pdf", ("subcontractor_docs",)),
    ("10.2. Отзывы.pdf", ("references",)),
    ("4_Документы специалистов.pdf", ("specialists_docs",)),
    ("9. ИСО 9001 Техносфера.pdf", ("certificate",)),
    ("22.Сертификаты и письма.pdf", ("certificate",)),
    ("19.Св-во собственности.pdf", ("property_docs",)),
    ("Форма 2_Письмо о подаче оферты.docx", ("cover_letter",)),
    ("7.Информационное письмо.docx", ("cover_letter",)),
    ("Форма 5_ Календарный график.docx", ("schedule",)),
    ("6. Основные сведения.docx", ("participant_info",)),
    ("Сведения из реестра СМСП от 14.10.2024.pdf", ("msp_register",)),
    ("Свидетельство о гос регистрации 2009.jpg", ("reg_certificate",)),
    ("Свидетельство о поставке на учет ТААН КОМБАТ.jpg", ("tax_registration",)),
    ("ul-1205900012230-20251212120137.pdf", ("egrul",)),
    ("ip-304590400012345-20251212120137.pdf", ("egrul",)),
    ("Техническая часть.zip.001", ("archive",)),
    ("добавить к документам.txt", ("note",)),
    ("!!!все документы нужны в ворде и пдф_убрать печати из ворда.txt", ("note",)),
])
def test_the_kind_of_a_paper_is_read_from_its_file_name(name, kinds):
    assert classify_file(name) == kinds


def test_a_note_in_a_text_file_is_not_counted_as_a_paper_of_the_bid():
    """Файл «убрать печати из ворда.txt» -- заметка фирмы себе, а не лист
    заявки. Он остаётся в списке файлов, но в состав заявки вид «note» не
    считается документом: его сорт «not_a_paper»."""
    assert archive_scan.KINDS["note"][1] == "not_a_paper"
    assert archive_scan.KINDS["customer_protocol"][1] == "not_a_paper"


def test_the_customer_protocol_in_the_bid_folder_is_not_a_paper_of_the_bid(archive):
    """Протокол закупки, случайно оказавшийся в папке «Заявка», -- документ
    заказчика, а не то, что фирма подала: в состав заявки он не идёт."""
    make(archive, "Т до 01.02.25/Заявка/Устав.pdf",
         "Т до 01.02.25/Заявка/Протокол_рассмотрения_заявок.pdf")
    found = entry_of(archive, "Т до 01.02.25")
    assert [p.path for p in found.bid_papers] == ["Заявка/Устав.pdf"]
    assert found.protocol_files == ("Заявка/Протокол_рассмотрения_заявок.pdf",)


def test_the_protocol_of_the_customer_is_not_the_directors_authority_paper():
    """Протокол рассмотрения заявок -- документ закупки, из него узнают исход;
    «Протокол № 1» -- решение учредителей о назначении директора. Одно слово,
    два документа с противоположным смыслом для состава заявки."""
    protocol = ("Протокол_рассмотрения_первых_частей_заявок_на_участие_в_конкурсе_"
                "(допущено_1_и_более)_(системный) (2).docx_с_данными_эл.подписей.pdf")
    assert classify_file(protocol) == ("customer_protocol",)
    assert classify_file("3_Протокол № 1.pdf") == ("authority",)


def test_a_pdf_and_a_docx_of_the_same_paper_count_once_in_the_composition(archive):
    make(archive, "Т до 01.02.25/Заявка/7_Справка о крупности.docx",
         "Т до 01.02.25/Заявка/7_Справка о крупности.pdf")
    found = entry_of(archive, "Т до 01.02.25")
    assert len(found.bid_papers) == 2                 # файлов два
    assert found.bid_kinds == ("large_deal",)         # документ один


def test_the_composition_of_a_bid_is_the_sorted_set_of_its_kinds(archive):
    make(archive, "Т до 01.02.25/Заявка/Устав.pdf", "Т до 01.02.25/Заявка/Опись.docx",
         "Т до 01.02.25/Заявка/Справка об опыте.doc")
    assert entry_of(archive, "Т до 01.02.25").bid_kinds == (
        "charter", "experience", "inventory")


def test_the_customer_protocol_is_kept_as_evidence_of_the_outcome(archive):
    """Исход, которого нет в названии, иногда есть в протоколе: сканер
    собирает такие файлы (из ЛЮБОЙ части закупки), но исхода из них не
    выводит -- читать содержимое ему нельзя."""
    make(archive, "Т до 01.02.25/Заявка/Устав.pdf",
         "Т до 01.02.25/Протокол_рассмотрения_заявок_итоговый.pdf")
    found = entry_of(archive, "Т до 01.02.25")
    assert found.protocol_files == ("Протокол_рассмотрения_заявок_итоговый.pdf",)
    assert found.outcome is None


# ---- каждый вид имеет русское имя и сорт ----------------------------------------

def test_every_kind_has_a_russian_label_and_a_sort():
    for code in archive_scan.KINDS:
        label, sort = archive_scan.KINDS[code]
        assert any("а" <= ch.lower() <= "я" for ch in label), code
        assert sort in archive_scan.SORTS, (code, sort)


def test_every_kind_the_classifier_can_return_is_described():
    probes = ["2_Выписка из ЕГРЮЛ.pdf", "Опись.docx", "Заявка Гознак.zip", "x.pdf",
              "Протокол_рассмотрения_заявок.pdf"]
    for name in probes:
        for code in classify_file(name):
            assert code in archive_scan.KINDS, (name, code)


# ---- группы верхнего уровня --------------------------------------------------------

def test_year_folders_are_groups_and_other_top_folders_are_tenders(archive):
    make(archive, "2025 год/А до 01.02.25/Заявка/1.docx",
         "2026 год/Б до 01.02.26/Заявка/1.docx",
         "Отдельная закупка до 01.02.26/Заявка/1.docx")
    found = {e.folder: e.group for e in scan_archive(archive)}
    assert found == {"А до 01.02.25": "2025 год", "Б до 01.02.26": "2026 год",
                     "Отдельная закупка до 01.02.26": "(отдельно)"}


def test_the_result_does_not_depend_on_the_order_the_disk_lists_things_in(
        archive, monkeypatch):
    """Порядок записей и файлов не должен зависеть от диска: NTFS отдаёт имена
    по алфавиту, другие файловые системы -- как придётся, и результат, который
    плавает от машины к машине, не сравнить и не проверить.

    Тест ломает порядок нарочно: `os.listdir` и `os.walk` возвращают всё
    в обратном порядке. На NTFS без этого проверка зелёна при ЛЮБОЙ
    реализации -- обычный диск и так отдаёт алфавит."""
    make(archive, "2025 год/Б до 01.02.25/Заявка/Б.docx",
         "2025 год/Б до 01.02.25/Заявка/А.docx",
         "2025 год/А до 01.02.25/Заявка/1.docx")
    real_listdir, real_walk = os.listdir, os.walk

    def backwards_listdir(path="."):
        return list(reversed(sorted(real_listdir(path))))

    def backwards_walk(top, *args, **kwargs):
        for dirpath, dirnames, filenames in real_walk(top, *args, **kwargs):
            dirnames[:] = list(reversed(sorted(dirnames)))
            yield dirpath, dirnames, list(reversed(sorted(filenames)))

    monkeypatch.setattr(archive_scan.os, "listdir", backwards_listdir)
    monkeypatch.setattr(archive_scan.os, "walk", backwards_walk)
    found = scan_archive(archive)
    assert [e.folder for e in found] == ["А до 01.02.25", "Б до 01.02.25"]
    assert [p.path for p in found[1].bid_papers] == ["Заявка/А.docx", "Заявка/Б.docx"]


# ---- правила доверия ---------------------------------------------------------------

def test_the_scan_never_opens_a_file(archive, monkeypatch):
    """Правило 2. Сканер читает имена. Открыть файл архива, да ещё чужой
    (документация заказчика, скачанная из интернета), -- отдельное решение с
    отдельной осторожностью, а не побочный эффект обхода."""
    make(archive, "Т до 01.02.25/Заявка/Устав.pdf", "Т до 01.02.25/Заявка/Опись.docx")
    root = str(archive)
    real_open, real_io_open = builtins.open, io.open

    def guard(real):
        def wrapper(file, *args, **kwargs):
            if str(file).startswith(root):
                raise AssertionError(f"сканер открыл файл архива: {file}")
            return real(file, *args, **kwargs)
        return wrapper

    monkeypatch.setattr(builtins, "open", guard(real_open))
    monkeypatch.setattr(io, "open", guard(real_io_open))
    monkeypatch.setattr(Path, "open", lambda self, *a, **k: guard(real_open)(self, *a, **k))
    monkeypatch.setattr(Path, "read_bytes", lambda self: guard(real_open)(self, "rb").read())
    monkeypatch.setattr(Path, "read_text", lambda self, *a, **k: guard(real_open)(self).read())
    assert len(scan_archive(archive)) == 1


def test_the_scan_never_writes_into_the_archive(archive):
    """Архив -- настоящие поданные заявки фирмы. Ни кэша, ни пометок, ни
    временных файлов рядом с ними."""
    make(archive, "Т до 01.02.25/Заявка/Устав.pdf", "2025 год/У до 01.02.25/Заявка/1.docx")

    def snapshot():
        return sorted(
            (str(p.relative_to(archive)), p.stat().st_size, p.stat().st_mtime_ns)
            for p in archive.rglob("*"))

    before = snapshot()
    scan_archive(archive)
    assert snapshot() == before


def test_a_missing_archive_folder_is_a_russian_refusal(tmp_path):
    with pytest.raises(ArchiveError) as info:
        scan_archive(tmp_path / "нет такой папки")
    assert "не найдена" in str(info.value) and "нет такой папки" in str(info.value)


def test_a_file_in_place_of_the_archive_folder_is_a_russian_refusal(tmp_path):
    path = tmp_path / "файл.txt"
    path.write_bytes(b"")
    with pytest.raises(ArchiveError) as info:
        scan_archive(path)
    assert "не папка" in str(info.value)


def test_an_empty_archive_is_an_empty_list_not_an_error(archive):
    assert scan_archive(archive) == []


def test_the_real_layout_of_a_won_tender_is_read_as_a_whole(archive):
    """Сквозной случай по образцу настоящей закупки «КБФ Гознак»: две заявки
    разных юрлиц и одна пачка «на диск»."""
    base = "! ВЫЙГРАЛИ/КБФ Гознак_КЭР до 18.06.26"
    make(archive,
         f"{base}/Заявка ЗапСибЦЭП/01.Информация об участнике.docx",
         f"{base}/Заявка ЗапСибЦЭП/4. Устав.pdf",
         f"{base}/Заявка ООО Техносфера/7_Справка о крупности.docx",
         f"{base}/Заявка ООО Техносфера/На диск/0_Опись.pdf",
         f"{base}/кбф_кэр_закупочная документация/ТЗ.docx")
    found = entry_of(archive, "КБФ Гознак_КЭР до 18.06.26")
    assert found.outcome == "won" and found.services == ("КЭР",)
    assert found.deadline == "2026-06-18"
    assert found.bid_folders == ("Заявка ЗапСибЦЭП", "Заявка ООО Техносфера")
    assert found.documentation_folders == ("кбф_кэр_закупочная документация",)
    assert found.bid_kinds == ("charter", "inventory", "large_deal", "participant_info")
