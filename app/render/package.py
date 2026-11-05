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
  * Существующий файл по умолчанию не перезаписывается: документ уходит в
    `skipped`, чтобы не потерять правки в уже поданной или подписанной версии.
    `overwrite=True` разрешает замену (пересборка после правки данных); новый
    файл пишется рядом и подменяет старый только после успеха.
  * Ничего не пишется вне папки пакета: разделители пути в префиксе и названии
    документа заменяются при очистке имени.
  * Предупреждения контекста (`warnings`) собираются по документам в
    `PackageResult.warnings`, чтобы их показал интерфейс.
"""
import os
import re
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from app.config import BASE_DIR
from app.render.context import build_context
from app.render.docx_render import TemplateError, render_docx
from app.render.layouts import LayoutError
from app.render.xlsx_render import CellMapError, fill_workbook, load_cellmap
from app.repo import Repo

TEMPLATES_DOCS_DIR = BASE_DIR / "templates_docs"
CELLMAPS_DIR = BASE_DIR / "cellmaps"

MAX_NAME_LENGTH = 255       # предел имени файла в NTFS, в знаках UTF-16
MAX_PATH_LENGTH = 259       # MAX_PATH Windows без завершающего нуля: дальше
                            # Проводник, Word и Excel файл не открывают
_SUPPORTED = (".docx", ".xlsx")
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


@dataclass
class PackageResult:
    written: list[Path] = field(default_factory=list)
    skipped: list[tuple[str, str]] = field(default_factory=list)    # (документ, причина)
    # (документ, предупреждение контекста): что подставлено по умолчанию или
    # пропало из данных; собирается и для документов, которые не удалось собрать
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


def _target(out_dir: Path, document: str, prefix: str, extension: str) -> Path:
    """Путь будущего файла. Имя -- префикс + документ, очищенные и обрезанные
    так, чтобы с расширением влезть в предел имени и в предел полного пути."""
    room = min(MAX_NAME_LENGTH, MAX_PATH_LENGTH - _units(str(out_dir)) - 1)
    room -= len(extension)
    if room < 1:
        raise _Refusal(
            f"Путь к папке слишком длинный: для имени файла не остаётся места "
            f"(Windows открывает пути не длиннее {MAX_PATH_LENGTH} знаков). "
            "Выберите папку ближе к корню диска.")
    return out_dir / (safe_filename(prefix + document, room) + extension)


def _template_path(document: str, template: str) -> Path:
    name = template or f"{document}.docx"
    path = TEMPLATES_DOCS_DIR / name
    if path.suffix.lower() not in _SUPPORTED:
        raise _Refusal(
            f"Шаблон «{name}» имеет тип «{path.suffix or 'без расширения'}»: "
            "поддерживаются только Word (.docx) и Excel (.xlsx). Проверьте имя "
            "шаблона в профиле заказчика.")
    if not path.is_file():
        raise _Refusal(
            f"Шаблон «{name}» не найден в папке {TEMPLATES_DOCS_DIR.name}. "
            "Положите файл с этим именем в папку или укажите другой шаблон "
            "в профиле заказчика.")
    return path


def _build_one(tender_id: int, document: str, spec, out_dir: Path, repo: Repo,
               taken: set[str], warnings: list[tuple[str, str]],
               overwrite: bool) -> Path:
    """Собирает один документ; при отказе поднимает исключение. Нового файла
    после отказа не остаётся, а прежний (при overwrite) остаётся целым."""
    template = _template_path(document, spec.template if spec else "")
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
    replacing = target.exists()
    if replacing and not overwrite:
        raise _Refusal(
            f"Файл «{target.name}» уже есть в папке: он не перезаписан, чтобы "
            "не потерять ваши правки. Разрешите замену существующих файлов, "
            "переименуйте или удалите его либо выберите другую папку.")

    context = build_context(tender_id, repo, document=document)
    warnings.extend((document, text) for text in context["warnings"])
    # Заменяемый файл не трогается, пока новый не готов: пишется рядом под
    # временным именем и подменяет прежний одним действием. Сбой посреди записи
    # не должен стоить пользователю прежней версии.
    work = (target.with_name(f"~{uuid.uuid4().hex[:8]}{target.suffix}")
            if replacing else target)
    try:
        if template.suffix.lower() == ".xlsx":
            if not spec or not spec.cellmap:
                raise _Refusal(
                    f"Для книги «{template.name}» в профиле заказчика не указана "
                    "карта ячеек. Добавьте её в описание документа.")
            fill_workbook(template, load_cellmap(CELLMAPS_DIR / spec.cellmap),
                          context, work)
        else:
            render_docx(template, context, work)
        if replacing:
            try:
                os.replace(work, target)
            except OSError as exc:
                raise _Refusal(
                    f"Файл «{target.name}» не заменён: он открыт в Word или Excel "
                    "либо защищён от записи. Закройте его и соберите заново; "
                    "прежняя версия осталась целой.") from exc
    except BaseException:
        # work -- либо новый файл, либо временный: прежнего файла он не касается
        work.unlink(missing_ok=True)
        raise
    return target


def build_package(tender_id: int, documents: list[str], out_dir: Path,
                  repo: Repo, overwrite: bool = False) -> PackageResult:
    """Собирает документы `documents` тендера в папку `out_dir`.

    Папка создаётся, если её нет. `overwrite=False` (по умолчанию): существующий
    файл не трогается, документ уходит в `skipped`. `overwrite=True`: файл
    заменяется, но только когда новый готов; не собрался -- прежний цел.
    Экран спрашивает пользователя один раз, если хоть один целевой файл есть.

    Неизвестный тендер -- `LookupError`; нет папки шаблонов или нельзя создать
    папку пакета -- `PackageError`; в обоих случаях на диск ничего не пишется.
    Сбой одного документа в исключение не превращается: см. `PackageResult.skipped`.
    """
    tender = repo.get_tender(tender_id)
    profile = (repo.find_customer_profile(tender.customer_id)
               if tender.customer_id is not None else None)
    if documents and not TEMPLATES_DOCS_DIR.is_dir():
        # одна причина на весь пакет -- одно сообщение, а не по строке на документ
        raise PackageError(
            f"Папка шаблонов «{TEMPLATES_DOCS_DIR}» не найдена: без неё не "
            "собрать ни один документ. Проверьте, что папка templates_docs "
            "лежит рядом с программой и в ней есть шаблоны.")
    out_dir = Path(os.path.abspath(out_dir))
    _make_folder(out_dir)

    result = PackageResult()
    seen: set[str] = set()
    taken: set[str] = set()
    for document in documents:
        if document in seen:
            result.skipped.append((document, (
                "Документ указан в составе пакета дважды: собран один раз.")))
            continue
        seen.add(document)
        spec = profile.docs.get(document) if profile else None
        try:
            path = _build_one(tender_id, document, spec, out_dir, repo,
                              taken, result.warnings, overwrite)
        except _Refusal as exc:
            result.skipped.append((document, str(exc)))
        except (TemplateError, CellMapError, LayoutError) as exc:
            result.skipped.append((document, str(exc)))
        except Exception as exc:
            result.skipped.append((document, (
                f"Документ не собран из-за непредвиденной ошибки "
                f"({type(exc).__name__}). Остальные документы пакета не затронуты.")))
        else:
            result.written.append(path)
            taken.add(path.name.casefold())
    return result
