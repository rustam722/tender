"""Сборка пакета документов тендера.

Берёт состав пакета, строит контекст на КАЖДЫЙ документ (раскладка таблиц у
«Справки о наличии оборудования» и «Справки о МТР» разная, а справочник один),
заполняет шаблон заказчика и кладёт файл в папку, которую выбрал пользователь.
Своей раскладки папки у программы нет: у существующих папок тендеров нет
общего устройства.

Что решено и почему:

  * Один неудавшийся документ не отменяет остальные. Он попадает в
    `PackageResult.skipped` парой «документ, причина по-русски», остальные
    записываются. Не собранный документ заметен, молча испорченный -- нет.
  * Имя файла = префикс из профиля заказчика + название документа + расширение
    шаблона. Нумерация -- соглашение заказчика («1.Декларация», «8.1_Справка
    о МТР»), а не наше, поэтому префикс берётся из `DocSpec.prefix`.
  * Имя очищается от символов, запрещённых в Windows, и укорачивается так,
    чтобы влезть и в 255 знаков имени, и в 259 знаков полного пути. Обрезается
    только имя файла: содержимое документа остаётся целым.
  * Любой документ пишется под временным именем (`~pkg-xxxxxxxx.ext`) рядом с
    целью и становится файлом одним rename/replace: остаток оборванной записи
    никогда не носит настоящего имени, а в начале сборки такие остатки
    подметаются.
  * Существующий файл по умолчанию не перезаписывается: документ уходит в
    `skipped`, чтобы не потерять правки в уже поданной или подписанной версии.
    `overwrite=True` разрешает замену (пересборка после правки данных); новый
    файл пишется рядом и подменяет старый только после успеха.
  * Ничего не пишется вне папки пакета: разделители пути в префиксе и названии
    документа заменяются при очистке имени.
  * Предупреждения контекста (`warnings`) собираются по документам в
    `PackageResult.warnings`, чтобы их показал интерфейс. Туда же ложатся
    графы документа Word, напечатанные пустыми (одно предупреждение на
    документ, с именами полей).
  * Цены, которые документы НАПЕЧАТАЛИ (`printed_price_wo_vat` контекста),
    сличаются между собой: разошлись -- предупреждение о пакете целиком.
    Место именно здесь: только сборка знает, какие документы в пакете.
    Цена документа попадает в сверку после того, как файл лёг в папку:
    назвать напечатанной цену бумаги, которой в папке нет, -- неправда.
  * Документ, снятый со сборки, мог остаться в папке файлом от прежнего
    раза -- и уйдёт в заявку вместе с остальными. Его цена тоже сверяется,
    но отдельно и другими словами: в файле стоит прежняя цена, а посчитать
    сейчас можно только то, что он напечатал БЫ (`_left_in_folder`).
"""
import os
import re
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from app import config
from app.render.context import (
    CUSTOMER_FORM_DOCUMENTS, ContextError, blank_fields_warning, build_context,
    money, printed_price,
)
from app.render.docx_render import TemplateError, render_docx
from app.render.layouts import LayoutError
from app.render.xlsx_render import CellMapError, fill_workbook, load_cellmap
from app.repo import Repo


MAX_NAME_LENGTH = 255       # предел имени файла в NTFS, в знаках UTF-16
MAX_PATH_LENGTH = 259       # MAX_PATH Windows без завершающего нуля: дальше
                            # Проводник, Word и Excel файл не открывают
_SUPPORTED = (".docx", ".xlsx")
# Временный файл заменяемого документа. Имя узнаваемо и не пересекается с
# замками Word («~$имя.docx») и чужими файлами: по нему сборка подметает остатки
# оборванных запусков, и ничего другого не трогает.
_TEMP_PREFIX = "~pkg-"
_TEMP_LENGTH = len(_TEMP_PREFIX) + 8     # «~pkg-» и 8 знаков, без расширения
_TEMP_NAME = re.compile(r"~pkg-[0-9a-f]{8}\.(docx|xlsx)")
_FALLBACK_NAME = "Документ"
_RESERVED = frozenset(
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{n}" for n in range(1, 10)} | {f"LPT{n}" for n in range(1, 10)})

# Разделители пути превращаются в дефис (как «СЗЗ/ТКП» -> «СЗЗ-ТКП»), прочие
# запрещённые знаки удаляются.
_SEPARATORS = "/\\"
_DROPPED = ':|?*"<>'
_TRANSLATION = {**{ord(c): "-" for c in _SEPARATORS},
                **{ord(c): None for c in _DROPPED}}
_WHITESPACE_CONTROLS = "\r\n\t\v\f"     # становятся пробелом; прочие управляющие -- удаляются
_DEL = "\x7f"


class PackageError(Exception):
    """Пакет собрать нельзя вовсе (в отличие от одного документа); по-русски."""


# Под этим именем в warnings стоит предупреждение не об одном документе, а о
# пакете целиком (две бумаги с разной ценой). Экран печатает «имя: текст»,
# и «Пакет: ...» читается так же, как «Калькуляция к КП: ...».
PACKAGE_WIDE = "Пакет"


@dataclass
class PackageResult:
    written: list[Path] = field(default_factory=list)
    skipped: list[tuple[str, str]] = field(default_factory=list)    # (документ, причина)
    # (документ, предупреждение контекста): что подставлено по умолчанию или
    # пропало из данных; собирается и для документов, которые не удалось собрать.
    # Под именем PACKAGE_WIDE стоит предупреждение обо всём пакете сразу.
    warnings: list[tuple[str, str]] = field(default_factory=list)


class _Refusal(Exception):
    """Документ не собран; текст -- причина для пользователя."""


def _units(text: str) -> int:
    """Длина в знаках UTF-16: так Windows считает имя и путь."""
    return sum(2 if ord(ch) > 0xFFFF else 1 for ch in text)


def _truncate(text: str, limit: int) -> str:
    out, used = [], 0
    for ch in text:
        width = 2 if ord(ch) > 0xFFFF else 1
        if used + width > limit:
            break
        out.append(ch)
        used += width
    return "".join(out)


def _clean(ch: str) -> str:
    if ch in _WHITESPACE_CONTROLS:
        return " "
    return "" if ch < " " or ch == _DEL else ch


def safe_filename(name: str, max_length: int = MAX_NAME_LENGTH) -> str:
    """Имя, допустимое в Windows: без `: / \\ | ? * " < >`, без пробелов по краям
    и точки на конце, не длиннее `max_length`. Кириллица, «» и № остаются.

    `/` и `\\` дают дефис, остальные запрещённые знаки пропадают. Переводы
    строк -- пробел, прочие управляющие знаки пропадают. Расширения здесь нет:
    его добавляет вызывающий, оставив для него место в `max_length`.
    """
    text = "".join(_clean(ch) for ch in name)
    text = text.translate(_TRANSLATION)
    text = re.sub(r"\s+", " ", text).strip().rstrip(". ")
    # CON, NUL, COM1... (и «CON.txt») Windows не даёт создать как файл
    if text.split(".")[0].strip().upper() in _RESERVED:
        text = "_" + text
    text = _truncate(text, max_length).rstrip(". ")
    return text or _truncate(_FALLBACK_NAME, max_length)


def _make_folder(out_dir: Path) -> None:
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise PackageError(
            f"Не удалось создать папку «{out_dir}»: проверьте, что в пути нет "
            "запрещённых знаков (? * \" < > |), что это не имя существующего "
            "файла и что в этом месте можно записывать.") from exc


def _sweep_temp_files(folder: Path) -> None:
    """Убирает остатки оборванных замен: питание отключили или задачу сняли
    посреди сборки, и уборка «по ходу» не сработала. Папка уходит заказчику,
    лишнего в ней быть не должно. Удаляются только файлы по шаблону имени."""
    try:
        entries = list(folder.iterdir())
    except OSError:
        return                  # уборка -- забота второго плана
    for entry in entries:
        if _TEMP_NAME.fullmatch(entry.name):
            try:
                if entry.is_file():
                    entry.unlink()
            except OSError:
                pass            # занят -- подметём в следующий раз


def _temp_path(target: Path) -> Path:
    """Свободное временное имя рядом с целью (в той же папке: замена -- один
    os.replace, без копирования между томами)."""
    while True:
        candidate = target.with_name(
            f"{_TEMP_PREFIX}{uuid.uuid4().hex[:8]}{target.suffix}")
        if not candidate.exists():
            return candidate


def _target(out_dir: Path, document: str, prefix: str, extension: str) -> Path:
    """Путь будущего файла. Имя -- префикс + документ, очищенные и обрезанные
    так, чтобы с расширением влезть в предел имени и в предел полного пути."""
    path_room = MAX_PATH_LENGTH - _units(str(out_dir)) - 1
    room = min(MAX_NAME_LENGTH, path_room) - len(extension)
    # документ пишется под временным именем, оно длиннее короткого целевого:
    # если не умещается оно, документ не записать, и отказ нужен сразу
    if room < 1 or path_room < _TEMP_LENGTH + len(extension):
        raise _Refusal(
            f"Путь к папке слишком длинный: для имени файла не остаётся места "
            f"(Windows открывает пути не длиннее {MAX_PATH_LENGTH} знаков). "
            "Выберите папку ближе к корню диска.")
    return out_dir / (safe_filename(prefix + document, room) + extension)


def _forms_of_other_customers(document: str, repo: Repo) -> str:
    """Подсказка к отказу: в чьих профилях такая форма уже назначена.

    Видно по самим профилям, и это прямо говорит, откуда взять файл. Ни у кого
    не назначена -- молчим: пустая подсказка хуже отсутствующей."""
    found = []
    for profile in repo.list_customer_profiles():
        spec = profile.docs.get(document)
        if not (spec and spec.template) or profile.customer_id is None:
            continue
        try:
            org = repo.get_organization(profile.customer_id)
        except LookupError:
            continue                    # профиль без заказчика -- не подсказка
        found.append(f"{org.short_name} -- «{spec.template}»")
    if not found:
        return ""
    return (" Такой же документ назначен в профилях: " + "; ".join(found) + ".")


def _template_path(document: str, template: str, repo: Repo) -> Path:
    # Документ по форме заказчика без назначенной формы: общего бланка у него
    # нет, и звать человека положить «<документ>.docx» в папку -- посылать его
    # делать бессмысленное.
    if not template and document in CUSTOMER_FORM_DOCUMENTS:
        raise _Refusal(
            f"«{document}» собирается по форме заказчика, а в профиле этого "
            "заказчика форма для него не назначена. Своего бланка у такого "
            "документа нет: форма у каждого заказчика своя. Укажите файл "
            "формы в описании документа в профиле заказчика или снимите "
            "галочку с документа в составе пакета."
            + _forms_of_other_customers(document, repo))
    name = template or f"{document}.docx"
    path = config.TEMPLATES_DOCS_DIR / name
    if path.suffix.lower() not in _SUPPORTED:
        raise _Refusal(
            f"Шаблон «{name}» имеет тип «{path.suffix or 'без расширения'}»: "
            "поддерживаются только Word (.docx) и Excel (.xlsx). Проверьте имя "
            "шаблона в профиле заказчика.")
    if not path.is_file():
        raise _Refusal(
            f"Шаблон «{name}» не найден в папке {config.TEMPLATES_DOCS_DIR.name}. "
            "Положите файл с этим именем в папку или укажите другой шаблон "
            "в профиле заказчика.")
    return path


def _plan_one(document: str, spec, out_dir: Path, taken: set[str],
              repo: Repo) -> tuple[Path, Path]:
    """Шаблон и итоговый путь документа, либо _Refusal, если документ заведомо
    не будет записан. Общий шаг для planned_files и build_package: путь, о
    котором спросил экран, и путь, который запишет сборка, считаются одним кодом.
    Занятое имя вносится в `taken`. `repo` -- только для подсказки в отказе
    (в чьих профилях назначена форма этого документа)."""
    template = _template_path(document, spec.template if spec else "", repo)
    target = _target(out_dir, document, spec.prefix if spec else "",
                     template.suffix.lower())
    if os.path.dirname(os.path.abspath(target)) != str(out_dir):
        raise _Refusal("Имя файла получилось с путём к другой папке; "
                       "документ не записан.")       # недостижимо после очистки
    if target.name.casefold() in taken:
        raise _Refusal(
            f"Имя файла «{target.name}» совпало с другим документом пакета: "
            "второй не записан. Задайте документам разные префиксы в профиле "
            "заказчика.")
    if target.is_dir():
        raise _Refusal(
            f"На месте файла «{target.name}» в папке лежит папка с таким же "
            "именем. Переименуйте или уберите её, либо выберите другую папку.")
    if template.suffix.lower() == ".xlsx":
        if not spec or not spec.cellmap:
            # Недостижимо: DocSpec отвергает .xlsx без карты при создании, а
            # пустой template всегда даёт .docx. Оставлено на случай профиля,
            # собранного в обход DocSpec; не полагайтесь на эту ветку.
            raise _Refusal(
                f"Для книги «{template.name}» в профиле заказчика не указана "
                "карта ячеек. Добавьте её в описание документа.")
        if not (config.CELLMAPS_DIR / spec.cellmap).is_file():
            raise _Refusal(
                f"Карта ячеек {spec.cellmap} не найдена в папке "
                f"{config.CELLMAPS_DIR.name}. Проверьте имя файла в профиле заказчика.")
    taken.add(target.name.casefold())
    return template, target


def _move_refusal(target: Path, replacing: bool) -> str:
    if replacing:
        return (f"Файл «{target.name}» не заменён: он открыт в Word или Excel "
                "либо защищён от записи. Закройте его и соберите заново; "
                "прежняя версия осталась целой.")
    return (f"Файл «{target.name}» не записан: такое имя уже занято или в "
            "папку нельзя записывать. Проверьте, что файл не открыт в другой "
            "программе, и соберите заново.")


def _listing(prices: dict[str, object]) -> str:
    return "; ".join(f"«{d}» -- {money(p)}" for d, p in prices.items())


def _price_warnings(printed: dict[str, object],
                    left: dict[str, object]) -> list[tuple[str, str]]:
    """Разошлись ли цены бумаг, которые лежат в папке пакета.

    Сравнивается не то, что посчитала одна подсистема, а то, что каждый
    документ выводит на бумагу: «Коммерческое предложение» печатает сумму
    напечатанных строк позиций лота, книга калькуляции -- итог расчёта. Эти
    два числа человек вводит в разных местах карточки, и до этой сверки их
    не сравнивал никто: пакет уходил в комиссию с двумя разными ценами.

    `printed` -- документы, записанные этой сборкой: их цена напечатана сейчас.
    `left` -- документы, которые сборка не пересобрала, а файл их в папке
    остался (чаще всего «файл уже есть, замена не разрешена»): их цена --
    то, что документ напечатал БЫ по нынешним данным, а в самом файле стоит
    прежняя. Эти два случая нельзя называть одним словом: сказать про старый
    файл «печатает X» -- соврать о бумаге, которой никто не печатал.

    Документы без цены (декларация, справки) не участвуют. Одна цена на всю
    папку -- молчание: сигнал, который горит всегда, через неделю перестают
    читать."""
    printed = {d: p for d, p in printed.items() if p is not None}
    left = {d: p for d, p in left.items() if p is not None}
    if len(set(printed.values()) | set(left.values())) < 2:
        return []
    parts = []
    if len(set(printed.values())) > 1:
        parts.append("Документы пакета печатают разную цену без НДС: "
                     f"{_listing(printed)}.")
    elif printed:
        parts.append("Цена в папке пакета не одна. Записано этой сборкой: "
                     f"{_listing(printed)}.")
    if left:
        parts.append(
            "В папке остались документы с ценой, которые сборка не "
            "пересобрала; в файле стоит прежняя цена, а по нынешним данным "
            f"напечатали бы: {_listing(left)}.")
    parts.append("Одна заявка с двумя ценами -- повод отклонить её. Сверьте "
                 "позиции лота и калькулятор на карточке тендера и соберите "
                 "пакет заново"
                 + (", разрешив замену существующих файлов." if left else "."))
    return [(PACKAGE_WIDE, " ".join(parts))]


def _left_in_folder(tender_id: int, document: str, spec, out_dir: Path,
                    repo: Repo, written: list[Path]) -> object:
    """Цена не пересобранного документа, файл которого лежит в папке; None --
    сверять нечего.

    Документ, снятый со сборки, уходит в `skipped`, но его файл в папке
    остаётся и в заявку попадёт вместе с остальными. Поэтому цена считается и
    для него -- из тендера, без контекста и без открытия книги.

    None, когда сверять нечего: пути файла мы не знаем (нет шаблона, слишком
    длинный путь), файла в папке нет, цены документ не печатает или его
    карта ячеек испорчена. Файл, записанный этой же сборкой, тоже не считается
    оставшимся: при совпадении имён в папке лежит бумага другого документа, и
    называть её этим именем было бы неправдой."""
    try:
        _, target = _plan_one(document, spec, out_dir, set(), repo)
    except _Refusal:
        return None
    if not target.is_file() or target in written:
        return None
    try:
        cell_map = (load_cellmap(config.CELLMAPS_DIR / spec.cellmap)
                    if spec and spec.cellmap else None)
        return printed_price(tender_id, repo, document, cell_map)
    except (CellMapError, LookupError):
        return None


def _build_one(tender_id: int, document: str, spec, out_dir: Path, repo: Repo,
               taken: set[str], warnings: list[tuple[str, str]],
               prices: dict[str, object],
               overwrite: bool, experience_ids) -> Path:
    """Собирает один документ; при отказе поднимает исключение. Нового файла
    после отказа не остаётся, а прежний (при overwrite) остаётся целым."""
    template, target = _plan_one(document, spec, out_dir, taken, repo)
    replacing = target.exists()
    if replacing and not overwrite:
        raise _Refusal(
            f"Файл «{target.name}» уже есть в папке: он не перезаписан, чтобы "
            "не потерять ваши правки. Разрешите замену существующих файлов, "
            "переименуйте или удалите его либо выберите другую папку.")

    context = build_context(tender_id, repo, document=document,
                           experience_ids=experience_ids)
    warnings.extend((document, text) for text in context["warnings"])
    # Единый путь записи: документ всегда пишется под временным именем рядом с
    # целью и становится настоящим файлом одним действием. Так оборванная или
    # неудавшаяся запись оставляет только `~pkg-…`, который уберёт подметание,
    # но не огрызок под настоящим именем, и заменяемый файл цел, пока новый не
    # готов. rename не затирает чужой файл, появившийся за это время;
    # replace -- только при разрешённой замене.
    work = _temp_path(target)
    blanks: list[str] = []
    try:
        try:
            if template.suffix.lower() == ".xlsx":
                fill_workbook(template, load_cellmap(config.CELLMAPS_DIR / spec.cellmap),
                              context, work)
            else:
                render_docx(template, context, work, blanks)
        except (TemplateError, CellMapError) as exc:
            # Модули рендера честно называют тот путь, который им дали, то есть
            # временный. Пользователь такого файла не видел: подменяется только
            # имя файла, а причина отказа остаётся как есть.
            raise _Refusal(str(exc).replace(work.name, target.name)) from exc
        try:
            (os.replace if replacing else os.rename)(work, target)
        except OSError as exc:
            raise _Refusal(_move_refusal(target, replacing)) from exc
    except BaseException:
        try:
            work.unlink(missing_ok=True)
        except OSError:
            pass        # настоящая причина важнее; остаток подметёт следующая сборка
        raise
    # Дальше -- только то, что верно про файл, уже лежащий в папке под своим
    # именем: его цена и его пустые графы. До переноса говорить об этом рано --
    # «напечатано» о бумаге, которой в папке нет, уже один раз было неправдой.
    prices[document] = context["printed_price_wo_vat"]
    blank_note = blank_fields_warning(document, blanks)
    if blank_note:
        warnings.append((document, blank_note))
    return target


def _prepare(tender_id: int, documents: list[str], out_dir: Path, repo: Repo):
    """Проверки, общие для планирования и сборки: тендер, папка шаблонов.
    Возвращает профиль заказчика (или None) и абсолютную папку пакета."""
    tender = repo.get_tender(tender_id)
    profile = (repo.find_customer_profile(tender.customer_id)
               if tender.customer_id is not None else None)
    if documents and not config.TEMPLATES_DOCS_DIR.is_dir():
        # одна причина на весь пакет -- одно сообщение, а не по строке на документ
        raise PackageError(
            f"Папка шаблонов «{config.TEMPLATES_DOCS_DIR}» не найдена: без неё не "
            "собрать ни один документ. Проверьте, что папка templates_docs "
            "лежит рядом с программой и в ней есть шаблоны.")
    return profile, Path(os.path.abspath(out_dir))


def planned_files(tender_id: int, documents: list[str], out_dir: Path,
                  repo: Repo) -> list[tuple[str, Path]]:
    """Пары «документ, итоговый путь» -- что `build_package` попытается записать.

    Ничего не пишет и папку не создаёт. Пути считаются тем же кодом, что и в
    сборке, поэтому совпадают с `PackageResult.written`. Файлы, которые уже
    лежат в папке, в список входят: по нему экран и узнаёт, какие документы
    будут заменены. Документы, которые заведомо не будут записаны (нет шаблона
    или карты, неподдерживаемый тип, повтор, совпавшее имя, слишком длинный
    путь), не входят. Ошибки те же, что у `build_package`: `LookupError` для
    неизвестного тендера, `PackageError` без папки шаблонов. Документ, который
    упадёт при самой сборке (ошибка в шаблоне, данных), заранее не угадать --
    он в списке остаётся.
    """
    profile, out_dir = _prepare(tender_id, documents, out_dir, repo)
    plan: list[tuple[str, Path]] = []
    seen: set[str] = set()
    taken: set[str] = set()
    for document in documents:
        if document in seen:
            continue
        seen.add(document)
        spec = profile.docs.get(document) if profile else None
        try:
            _, target = _plan_one(document, spec, out_dir, taken, repo)
        except _Refusal:
            continue
        plan.append((document, target))
    return plan


def build_package(tender_id: int, documents: list[str], out_dir: Path,
                  repo: Repo, overwrite: bool = False,
                  experience_ids: list[int] | tuple[int, ...] | None = None
                  ) -> PackageResult:
    """Собирает документы `documents` тендера в папку `out_dir`.

    Папка создаётся, если её нет. `overwrite=False` (по умолчанию): существующий
    файл не трогается, документ уходит в `skipped`. `overwrite=True`: файл
    заменяется, но только когда новый готов; не собрался -- прежний цел.
    Экран спрашивает пользователя один раз, если хоть один целевой файл есть.

    `experience_ids` -- записи опыта для справки об опыте, в нужном порядке;
    `None` -- отбор, сохранённый при тендере (как в `build_context`), пустой
    список -- без опыта.

    Перед сборкой из папки убираются остатки оборванных замен (`~pkg-xxxxxxxx`).

    Цены бумаг, которые лежат в папке, сличаются между собой: разошлись --
    в `warnings` ложится строка под именем `PACKAGE_WIDE` с именами документов
    и их числами. Записанные этой сборкой и оставшиеся от прежнего раза
    названы в ней отдельно и разными словами.

    Неизвестный тендер -- `LookupError`; нет папки шаблонов или нельзя создать
    папку пакета -- `PackageError`; в обоих случаях на диск ничего не пишется.
    Сбой одного документа в исключение не превращается: см. `PackageResult.skipped`.
    """
    profile, out_dir = _prepare(tender_id, documents, out_dir, repo)
    _make_folder(out_dir)
    _sweep_temp_files(out_dir)

    result = PackageResult()
    seen: set[str] = set()
    taken: set[str] = set()
    prices: dict[str, object] = {}      # документ -> цена, которую он НАПЕЧАТАЛ
    skipped: list[tuple[str, object]] = []      # документ -> его описание в профиле
    for document in documents:
        if document in seen:
            result.skipped.append((document, (
                "Документ указан в составе пакета дважды: обработан один раз.")))
            continue
        seen.add(document)
        spec = profile.docs.get(document) if profile else None
        try:
            path = _build_one(tender_id, document, spec, out_dir, repo,
                              taken, result.warnings, prices, overwrite,
                              experience_ids)
        except _Refusal as exc:
            result.skipped.append((document, str(exc)))
            skipped.append((document, spec))
        except (TemplateError, CellMapError, LayoutError, ContextError) as exc:
            result.skipped.append((document, str(exc)))
            skipped.append((document, spec))
        except Exception as exc:
            result.skipped.append((document, (
                f"Документ не собран из-за непредвиденной ошибки "
                f"({type(exc).__name__}). Остальные документы пакета не затронуты.")))
            skipped.append((document, spec))
        else:
            result.written.append(path)
    # Снятый со сборки документ мог остаться в папке файлом от прежнего раза --
    # и уйдёт в заявку вместе с остальными. Его цена считается здесь, когда
    # известно, что записано этой сборкой.
    left = {document: _left_in_folder(tender_id, document, spec, out_dir, repo,
                                      result.written)
            for document, spec in skipped}
    result.warnings.extend(_price_warnings(prices, left))
    return result
