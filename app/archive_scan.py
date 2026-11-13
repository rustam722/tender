"""Сканер архива прошлых закупок: читает ИМЕНА папок и файлов, содержимое -- нет.

Архив фирмы -- сотни закупок в папках вида «Заказчик до 04.09.2025_исход»; в
каждой документация заказчика и поданная заявка. Сканер раскладывает это в
записи (`ArchiveEntry`): когда был срок, чем кончилось (если названо), какие
документы фирма подавала. Остальное -- опыт из справок, кадры, нормы часов --
следующие шаги: они читают файлы уже сами и с проверкой человека.

Три правила, на которых держится доверие к такой записи.

1. **Неизвестное остаётся неизвестным.** Исход, которого нет в названии папки,
   -- `None`, а не «проигран»: в реальном архиве 55 закупок из 82 об исходе
   молчат, и записать их проигранными значило бы исказить всю статистику.
2. **Файлы не открываются и в архив ничего не пишется.** Архив -- настоящие
   поданные заявки фирмы и документация заказчиков, скачанная из интернета.
   Читать её содержимое -- отдельное решение с отдельной осторожностью, а не
   побочный эффект обхода. Сканер зовёт только `os.walk`/`os.scandir`.
3. **Догадка названа догадкой.** Заказчик, вынутый из имени папки, лежит
   в поле `customer_guess`, а не `customer`: имя папки -- не справочник.

Что сканер НЕ делает намеренно: не выводит исход из «Протокола рассмотрения» --
он только собирает такие файлы в `protocol_files` (чтобы человек, или следующий
шаг, прочитали их глазами); не считает документ из папки «Пример» поданным --
это чужая подача как образец.
"""
import os
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path


class ArchiveError(Exception):
    """Архив просканировать нельзя; текст на русском, для человека."""


# ---- виды документов ---------------------------------------------------------------
#
# Сорт -- что с таким документом делает программа (те же четыре сорта, что в
# составе заявки): «ours» собирает она, «customer_form» -- по форме заказчика,
# «scan» -- ваш скан, программа его не делает. Остальное -- служебные.

SORTS = ("ours", "customer_form", "scan", "platform", "unknown", "not_a_paper")

KINDS: dict[str, tuple[str, str]] = {
    "inventory":         ("Опись документов", "ours"),
    "envelope":          ("Надпись на конверт", "ours"),
    "participant_info":  ("Информация об участнике (анкета)", "customer_form"),
    "application":       ("Заявка на участие", "customer_form"),
    "proposal":          ("Предложение / коммерческое предложение", "ours"),
    "calculation":       ("Калькуляция / расчёт цены", "customer_form"),
    "declaration":       ("Декларация соответствия", "ours"),
    "experience":        ("Справка об опыте", "ours"),
    "staff":             ("Справка о кадровых ресурсах", "ours"),
    "assets":            ("Справка о МТР / оборудовании", "ours"),
    "large_deal":        ("Письмо о крупности сделки", "ours"),
    "interest":          ("Справка о заинтересованности", "ours"),
    "charter":           ("Устав", "scan"),
    "egrul":             ("Выписка из ЕГРЮЛ", "scan"),
    "reg_certificate":   ("Свидетельство о регистрации", "scan"),
    "tax_registration":  ("Свидетельство о постановке на учёт", "scan"),
    "authority":         ("Документ о полномочиях директора (протокол, решение, приказ)", "scan"),
    "power_of_attorney": ("Доверенность", "scan"),
    "msp_register":      ("Выписка из реестра МСП", "scan"),
    "balance":           ("Бухгалтерская отчётность", "scan"),
    "usn_notice":        ("Уведомление о применении УСН", "scan"),
    "tax_clearance":     ("Справка налоговой об отсутствии задолженности", "scan"),
    "accreditation":     ("Аттестат аккредитации", "scan"),
    "license":           ("Лицензия", "scan"),
    "sro":               ("Допуск СРО", "scan"),
    "contract":          ("Договор (с соисполнителем и т. п.)", "scan"),
    "consent":           ("Согласие на обработку персональных данных", "ours"),
    "ownership_chain":   ("Информация о цепочке собственников", "ours"),
    "company_card":      ("Карточка предприятия", "ours"),
    "reputation":        ("Справка о деловой репутации", "ours"),
    "subcontract_info":  ("Сведения о субподрядчиках", "ours"),
    "subcontractor_docs": ("Документы субподрядчика", "scan"),
    "guarantee_letter":  ("Гарантийное письмо", "ours"),
    "offer":             ("Оферта", "ours"),
    "obligations":       ("Справка об обязательствах в текущем году", "ours"),
    "revenue_share":     ("Справка о доле выручки", "ours"),
    "confidentiality":   ("Соглашение о конфиденциальности", "customer_form"),
    "participant_consent": ("Согласие участника закупки", "customer_form"),
    "references":        ("Отзывы и рекомендательные письма", "scan"),
    "specialists_docs":  ("Документы специалистов (дипломы, удостоверения)", "scan"),
    "certificate":       ("Сертификат (ИСО и т. п.)", "scan"),
    "property_docs":     ("Документы о праве собственности", "scan"),
    "cover_letter":      ("Письмо о подаче заявки / информационное письмо", "customer_form"),
    "schedule":          ("Календарный график", "customer_form"),
    "note":              ("Заметка в текстовом файле (не документ заявки)", "not_a_paper"),
    "archive":           ("Архив (zip, rar, 7z) -- упакованная копия", "unknown"),
    "customer_protocol": ("Протокол закупки (документ заказчика)", "not_a_paper"),
    "other":             ("Прочее (по имени файла вид не определён)", "unknown"),
}

_CYR = "а-яё"
_B = rf"(?<![{_CYR}])"          # слева не буква: «опыт» в «Справка об опыте», но не в середине слова

# Порядок правил задаёт порядок видов в ответе; на файл может сработать несколько
# правил («Письмо о крупной сделке и заинтересованности»).
_RULES: list[tuple[str, re.Pattern]] = [(kind, re.compile(rx)) for kind, rx in [
    ("inventory",         rf"{_B}опись"),
    ("envelope",          r"конверт"),
    ("participant_info",  r"информаци\w* об участник|сведени\w* об участник|анкет|основные сведения"),
    ("declaration",       r"деклара"),
    ("experience",        rf"{_B}опыт"),
    ("staff",             rf"{_B}кадр|{_B}штат"),
    ("assets",            rf"{_B}мтр|оборудован|материально.?техни"),
    ("large_deal",        r"крупн"),
    ("interest",          r"заинтересован"),
    ("charter",           rf"{_B}устав"),
    # «ul-1205900012230-20251212120137.pdf» -- выписка с сайта налоговой:
    # ul -- юридическое лицо, ip -- предприниматель, дальше ОГРН и метка времени
    ("egrul",             r"егрюл|егрип|(?:^|[ -])(?:ul|ip)-\d{10,}"),
    ("reg_certificate",   rf"свидетельств\w* (о )?(гос\w*\.? )?регистрац|{_B}огрн"),
    # «поставке» -- опечатка в имени файла, которая в архиве встречается
    ("tax_registration",  r"свидетельств\w* о пост\w*ке|пост\w*ке на уч[её]т"),
    ("tax_clearance",     rf"справк\w* (из )?налогов|{_B}задолженност"),
    ("power_of_attorney", r"доверенност"),
    ("msp_register",      rf"{_B}с?мсп|малого и среднего"),
    ("balance",           r"бухгалтерск|баланс"),
    ("usn_notice",        rf"{_B}усн(?![{_CYR}])|упрощ[её]нн"),
    ("accreditation",     r"аттестат|аккредитац"),
    ("license",           r"лиценз"),
    ("sro",               rf"{_B}сро(?![{_CYR}])|{_B}допуск"),
    ("guarantee_letter",  r"гарантийн\w* письм|гарантии участника"),
    # «Оферта» -- сама оферта; «Письмо о подаче оферты» -- сопроводительное
    # письмо, и в вид «offer» оно не входит (отсюда запрет буквы после «оферт[а]»)
    ("offer",             rf"{_B}оферта?(?![{_CYR}])"),
    ("obligations",       r"справк\w* об обязательствах"),
    ("revenue_share",     r"справк\w* о доле выручки"),
    ("confidentiality",   r"соглашени\w* о конфиденциальности"),
    ("participant_consent", rf"{_B}согласие участника|{_B}согласие\.|^согласие$"),
    ("consent",           r"согласи\w* на обр\w*|согласи\w* субъекта персональных"),
    ("ownership_chain",   r"цепочк\w* собственник"),
    ("company_card",      r"карточк\w* предприяти"),
    ("reputation",        r"деловой репутации"),
    ("subcontract_info",  r"сведени\w* о субподр|перечень привлека\w* субподр|план привлечения субподр|согласие субподр|пояснение по субподр"),
    ("subcontractor_docs", r"документы субподр"),
    ("references",        r"отзыв|рекомендательн"),
    ("specialists_docs",  r"документы специалист|диплом|удостоверени"),
    ("certificate",       rf"сертификат|{_B}исо|{_B}iso"),
    ("property_docs",     r"св-во собственности|свидетельств\w* (о )?праве собственности|праве собственности"),
    ("cover_letter",      r"письм\w* о подаче|информационное письмо"),
    ("schedule",          r"календарн\w* (график|план)"),
    ("calculation",       rf"калькуляц|{_B}смет|нмцд|обоснован|сводная таблица стоимости|расч[её]т цп"),
    ("proposal",          rf"предложени|коммерческ|{_B}ткп|{_B}кп(?![{_CYR}])"),
    ("contract",          r"договор|контракт"),
    ("application",       rf"{_B}заявк"),
]]
# «Протокол № 1» -- решение учредителей о назначении директора; «Протокол
# рассмотрения заявок» -- документ закупки. Одно слово, два документа с
# противоположным смыслом для состава заявки, поэтому разбор отдельный.
_CUSTOMER_PROTOCOL = re.compile(
    r"протокол\w*.{0,100}?(рассмотр|итог|подвед|допущен|вскрыт|оценк)")
_AUTHORITY = re.compile(r"протокол|решени|приказ")

_ARCHIVE_EXT = {".zip", ".rar", ".7z"}
_SIGNATURE_EXT = {".sig", ".sgn", ".p7s", ".p7m"}
_JUNK_EXT = {".tmp", ".url"}
_JUNK_NAMES = {"thumbs.db", "desktop.ini"}


def _normal(name: str) -> str:
    low = name.lower().replace("ё", "е")
    return re.sub(r"[\s_]+", " ", low)


def classify_file(filename: str) -> tuple[str, ...]:
    """Вид документа по имени файла: коды из `KINDS`, в порядке правил.

    Содержимое файла не читается. Что по имени не определилось -- `("other",)`:
    честнее признать, чем записать наугад."""
    stem, ext = os.path.splitext(filename)
    # «Техническая часть.zip.001» -- том многотомного архива: расширение .001
    if ext.lower() in _ARCHIVE_EXT or re.search(r"\.(zip|rar|7z)\.\d{3}$", filename.lower()):
        return ("archive",)
    if ext.lower() == ".txt":
        return ("note",)
    text = _normal(stem)
    if _CUSTOMER_PROTOCOL.search(text):
        return ("customer_protocol",)
    found = [kind for kind, rx in _RULES if rx.search(text)]
    if _AUTHORITY.search(text):
        found.append("authority")
    return tuple(found) or ("other",)


# ---- название папки закупки -----------------------------------------------------------

SERVICE_CODES = ("КЭР", "СЗЗ", "НДВ", "НДС", "НООЛР", "ППЭЭ", "ПЭК")
_SERVICE = {code: re.compile(rf"(?<![А-Яа-яЁё]){code}(?![А-Яа-яЁё])") for code in SERVICE_CODES}

# Порядок важен: «не выйграли» содержит «выйграли», «выйграли частично» --
# тоже. Первое сработавшее правило решает, и ошибка тут переворачивает смысл.
_OUTCOMES: list[tuple[str, re.Pattern]] = [(code, re.compile(rx)) for code, rx in [
    ("won_partly",  r"вы[йи]грали\s+частично"),
    ("lost",        r"не\s*вы[йи]грали|проиграли"),
    ("won",         r"вы[йи]грали"),
    ("not_applied", r"не\s*подавали|не\s*подаем|не\s*подходим"),
    ("refused",     r"отказали"),
    ("missed",      r"не\s*успели(\s+с\s+оплатой)?"),
    ("cancelled",   r"отменен"),
]]

_DATE = re.compile(r"(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})")
_DEADLINE_AFTER_DO = re.compile(r"до[\s_]+(\d{1,2}\.\d{1,2}\.(?:\d{4}|\d{2}))")

_GROUP_YEAR = re.compile(r"^\d{4}\s*год$", re.I)
_GROUP_WON = re.compile(r"^!*\s*вы[йи]грали$", re.I)
LOOSE_GROUP = "(отдельно)"


def _as_date(text: str) -> str | None:
    found = _DATE.fullmatch(text) or _DATE.search(text)
    if not found:
        return None
    day, month, year = found.groups()
    try:
        return date(int(year) + (2000 if len(year) == 2 else 0), int(month), int(day)).isoformat()
    except ValueError:                      # «31.02.26» -- опечатка, а не срок
        return None


def _deadline(folder: str) -> str | None:
    after = _DEADLINE_AFTER_DO.search(folder)
    return _as_date(after.group(1)) if after else _as_date(folder)


def _outcome(folder: str) -> str | None:
    text = _normal(folder)
    for code, rx in _OUTCOMES:
        if rx.search(text):
            return code
    return None


def _services(folder: str) -> tuple[str, ...]:
    return tuple(code for code, rx in _SERVICE.items() if rx.search(folder))


def _customer_guess(folder: str) -> str:
    """Название папки без срока, исхода и кода услуги. ДОГАДКА, не справочник."""
    text = re.sub(r"[\s_]+до[\s_]+\d{1,2}\.\d{1,2}\.\d{2,4}.*$", "", folder)
    text = re.sub(r"\d{1,2}\.\d{1,2}\.\d{2,4}.*$", "", text)
    for _, rx in _OUTCOMES:
        text = re.sub(rx.pattern, " ", text, flags=re.I)
    for rx in _SERVICE.values():
        text = rx.sub(" ", text)
    text = re.sub(r"[()]", " ", text)
    text = re.sub(r"[\s_]+", " ", text)
    return text.strip(" -_")


# ---- папки внутри закупки ----------------------------------------------------------

def _role(subfolder: str) -> str:
    """«bid» -- наша подача, «documentation» -- документы заказчика, «other».

    «Пример» -- чужая подача как образец: считать её поданной значило бы
    приписать фирме чужой состав, поэтому она ни та, ни другая."""
    text = _normal(subfolder)
    if "пример" in text:
        return "other"
    if "заявк" in text:
        return "bid"
    if "документаци" in text:
        return "documentation"
    return "other"


@dataclass(frozen=True)
class BidPaper:
    path: str                        # от папки закупки, через «/»
    kinds: tuple[str, ...]


@dataclass
class ArchiveEntry:
    group: str
    folder: str
    path: Path
    customer_guess: str
    deadline: str | None
    services: tuple[str, ...]
    outcome: str | None
    outcome_source: str | None
    bid_folders: tuple[str, ...] = ()
    documentation_folders: tuple[str, ...] = ()
    bid_papers: tuple[BidPaper, ...] = ()
    protocol_files: tuple[str, ...] = ()
    unclassified_files: int = 0
    documentation_files: int = 0

    @property
    def bid_kinds(self) -> tuple[str, ...]:
        """Состав поданной заявки: отсортированный набор видов. Файл .pdf и .docx
        одного листа -- один документ, поэтому набор, а не список файлов."""
        return tuple(sorted({kind for paper in self.bid_papers for kind in paper.kinds}))


def _is_junk(name: str) -> bool:
    low = name.lower()
    return (name.startswith("~$") or low in _JUNK_NAMES
            or os.path.splitext(low)[1] in _JUNK_EXT | _SIGNATURE_EXT)


def _scan_tender(group: str, folder: Path, group_outcome: tuple[str, str] | None) -> ArchiveEntry:
    name = folder.name
    outcome = _outcome(name)
    source = "название папки" if outcome else None
    if outcome is None and group_outcome:
        outcome, source = group_outcome

    subfolders = sorted(d for d in os.listdir(folder) if (folder / d).is_dir())
    roles = {sub: _role(sub) for sub in subfolders}

    papers: list[BidPaper] = []
    protocols: list[str] = []
    unclassified = documentation = 0
    for dirpath, dirnames, filenames in os.walk(folder):
        dirnames.sort()
        relative_dir = Path(dirpath).relative_to(folder)
        top = relative_dir.parts[0] if relative_dir.parts else None
        role = roles.get(top, "other") if top else "root"
        for filename in sorted(filenames):
            if _is_junk(filename):
                continue
            relative = (relative_dir / filename).as_posix()
            kinds = classify_file(filename)
            if kinds == ("customer_protocol",):
                protocols.append(relative)
                continue              # документ заказчика, а не наша подача
            if role == "bid":
                papers.append(BidPaper(relative, kinds))
            elif role == "documentation":
                documentation += 1
            else:
                unclassified += 1

    return ArchiveEntry(
        group=group, folder=name, path=folder,
        customer_guess=_customer_guess(name), deadline=_deadline(name),
        services=_services(name), outcome=outcome, outcome_source=source,
        bid_folders=tuple(s for s in subfolders if roles[s] == "bid"),
        documentation_folders=tuple(s for s in subfolders if roles[s] == "documentation"),
        bid_papers=tuple(papers), protocol_files=tuple(protocols),
        unclassified_files=unclassified, documentation_files=documentation)


def scan_archive(root) -> list[ArchiveEntry]:
    """Записи по всем закупкам архива, в порядке (группа, название).

    Папки верхнего уровня вида «2025 год» и «! ВЫЙГРАЛИ» -- группы, закупки
    лежат в них; остальные папки верхнего уровня -- сами закупки, группа
    «(отдельно)». Закупка из папки «! ВЫЙГРАЛИ» выиграна, но название самой
    закупки точнее общей папки: «выиграли частично» внутри остаётся частичной."""
    root = Path(root)
    if not root.exists():
        raise ArchiveError(
            f"Папка архива не найдена: {root}. Проверьте путь и что диск "
            "подключён.")
    if not root.is_dir():
        raise ArchiveError(f"{root} -- не папка, а файл: укажите папку архива.")

    entries: list[ArchiveEntry] = []
    for top in sorted(d for d in os.listdir(root) if (root / d).is_dir()):
        path = root / top
        if _GROUP_YEAR.match(top) or _GROUP_WON.match(top):
            group_outcome = ("won", f"папка «{top}»") if _GROUP_WON.match(top) else None
            for sub in sorted(d for d in os.listdir(path) if (path / d).is_dir()):
                entries.append(_scan_tender(top, path / sub, group_outcome))
        else:
            entries.append(_scan_tender(LOOSE_GROUP, path, None))
    return entries
