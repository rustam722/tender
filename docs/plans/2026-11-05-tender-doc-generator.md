# Генератор тендерных документов — план реализации

> **Для агентных исполнителей:** ОБЯЗАТЕЛЬНЫЙ ПОДНАВЫК — используйте
> superpowers:subagent-driven-development (рекомендуется) или
> superpowers:executing-plans для реализации задача за задачей.
> Шаги помечены чекбоксами (`- [ ]`) для отслеживания.

**Цель:** локальная программа в браузере, которая по карточке тендера из семи полей
и справочникам компании собирает готовый пакет тендерных документов в `.docx` и `.xlsx`.

**Архитектура:** Flask отдаёт одну HTML-страницу и REST-ручки; данные лежат в SQLite;
документы собираются подстановкой контекста в шаблоны Word (`docxtpl`) и заполнением
книг Excel заказчика по карте ячеек (`openpyxl`). Реквизиты организаций подтягиваются
по ИНН из ФНС и rusprofile. Ядро (база, калькулятор, рендер) не знает про веб
и тестируется отдельно.

**Технологии:** Python 3.14.6, Flask 3.1.3, SQLite (встроенный), `docxtpl`,
`openpyxl` 3.1.5, `python-docx`, `requests`, `lxml`, `pytest`.

**Описание:** [docs/specs/2026-11-05-tender-doc-generator-design.md](../specs/2026-11-05-tender-doc-generator-design.md)

## Канонические имена документов

Одна строка-ключ на документ. Эти же строки лежат в `CustomerProfile.documents`,
в ключах `CustomerProfile.docs` (задача 10а), в константе `TABLE_DOCUMENTS`
задачи 8, в именах шаблонов задачи 9 и в seed задачи 14.
Расхождение в один символ делает предупреждение о раскладке неработающим
и роняет сборку пакета, поэтому список здесь — единственный источник правды.

| Ключ | Таблица-справочник |
|---|---|
| `Коммерческое предложение` | нет |
| `Калькуляция к КП` | нет (свои подтаблицы) |
| `ТКП по форме заказчика` | нет |
| `Декларация соответствия` | нет |
| `Справка о кадровых ресурсах` | да, `employees` |
| `Справка о наличии оборудования` | да, `assets` |
| `Справка о МТР` | да, `assets` |
| `Письмо о крупности сделки` | нет |
| `Справка об опыте` | да, `experience` |
| `Анкета участника` | нет |

В разделе 8 описания документ опыта назван «Справка об опыте + приложение»,
потому что подаётся вместе с приложением. Ключ — `Справка об опыте`:
приложение — отдельный файл, который ждёт Word-версии.

## Общие ограничения

- Python 3.14.6, Windows 11, без прав администратора.
- Изолированный `venv` в папке программы; системный Python не используется напрямую.
- Работа на одном компьютере. Многопользовательский режим, блокировки и сетевая база не нужны.
- Все данные локально. Сеть — только `egrul.nalog.ru` и `www.rusprofile.ru`.
- Деньги и часы — только `decimal.Decimal`, никогда `float`.
- Проценты по умолчанию: страховые 30 %, травматизм 0,2 %, **накладные 5 % от
  расходов на оплату труда вместе со страховыми**, прибыль 10 % от
  (расходы на ОТ + субподряд + накладные).
- **Итог без НДС округляется до целых рублей** (`total_rounding_digits`,
  по умолчанию 0). Так устроена книга `КП.xlsx`: `ROUND(..., +D88)` при пустой
  `D88`. Именно это округление даёт ровные 4 000 000 из 3 999 999,83.
- **Единственное исключение из доктрины напечатанной строки:** итог без НДС
  может отличаться от суммы напечатанных над ним статей — на данных Арктикгаза
  на 17 копеек, в пределе до 50. Это конвенция самой книги заказчика, и мы ей
  следуем. Задачи 13 и 15 не должны считать итог равным сумме статей;
  задача 13 обязана брать итог из расчёта, а не складывать статьи на экране.
- **НДС по умолчанию 5 %** — ставка УСН, действующая для «Техносферы» как
  микропредприятия. В поданном КП Коряжмы НДС составил 285 000,00 на 5 700 000,00
  без НДС, что даёт ровно 5 %. Ставка переопределяется в карточке тендера:
  форма ТКП Роснефти прямо требует «указать размер ставки НДС или причину
  освобождения от уплаты НДС».
- Форма заказчика заполняется в его собственной книге: формулы, защита листа
  и оформление сохраняются. Своя копия не строится.
- Подпись и печать картинкой не вставляются. На выходе пустые подписные блоки.
- Интерфейс и тексты ошибок на русском, без трассировок.
- Выходные файлы открываются в Word и Excel без предупреждений о повреждении.

## На что смотреть ревьюеру

Пять условий, которые описание подразумевает, но ни одна задача не проверяла бы
сама по себе. Тест на каждое добавлен в задачу, владеющую кодом.

1. **ИНН не найден, найдено несколько организаций или ФНС требует капчу** —
   пользователь должен увидеть объяснение словами и возможность ввести вручную,
   а не пустую карточку. → Задача 3.
2. **Вложенные кавычки в наименовании.** `АО «НК «Роснефть»` — реальный заказчик
   из папки. Такое значение не должно ломать ни XML документа, ни разбор шаблона. → Задача 9.
3. **Rusprofile сменил разметку или отдал капчу.** Адрес и коды статистики не пришли —
   карточка должна заполниться данными ФНС и показать, чего не хватает,
   а не промолчать с пустыми полями. → Задача 4.
4. **Целевая цена ниже постоянных затрат.** Материалы и субподряд уже дают больше
   целевой суммы — подбор часов обязан отказаться с внятным сообщением,
   а не выдать отрицательные часы. → Задача 6.
5. **Длинные значения и недопустимые символы в имени файла.** Предмет закупки
   на 300 символов; `«»№:/\|?*` в наименовании заказчика. Имя файла должно
   остаться допустимым в Windows, текст — попасть в документ целиком. → Задача 11.

---

### Задача 1: Каркас проекта и база

**Файлы:**
- Создать: `requirements.txt`, `.gitignore`, `Тендеры.bat`, `pytest.ini`
- Создать: `app/__init__.py`, `app/config.py`, `app/models.py`, `app/db.py`
- Тесты: `tests/test_db.py`

**Интерфейсы:**
- Отдаёт: `config.DATA_DIR`, `config.TEMPLATES_DIR`, `config.DB_PATH`;
  датаклассы `Organization`, `Employee`, `Asset`, `ExperienceRecord`, `Service`,
  `OrgSettings`, `CustomerProfile`, `Tender`, `LotItem` в `app/models.py`;
  `db.connect(path: Path) -> sqlite3.Connection` и `db.migrate(conn) -> None`.

Все денежные поля и часы — `Decimal`. Поля справочников перечислены в разделе 7
описания; `Organization` несёт признак роли `role: str` со значениями
`'participant'` и `'customer'`.

- [ ] **Шаг 1: `git init`, создать `.gitignore`** (`venv/`, `__pycache__/`, `*.db`, `~$*`)

- [ ] **Шаг 2: Создать `venv` и `requirements.txt`, установить зависимости**

```bash
python -m venv venv
venv/Scripts/python -m pip install flask docxtpl openpyxl requests lxml pytest
venv/Scripts/python -m pip freeze > requirements.txt
```

- [ ] **Шаг 3: Написать падающий тест**

```python
def test_migrate_creates_tables(tmp_path):
    conn = db.connect(tmp_path / "t.db")
    db.migrate(conn)
    names = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"organization", "employee", "asset", "experience",
            "service", "org_settings", "customer_profile",
            "tender", "lot_item"} <= names

def test_migrate_is_idempotent(tmp_path):
    conn = db.connect(tmp_path / "t.db")
    db.migrate(conn); db.migrate(conn)          # второй раз не падает

def test_decimal_round_trips(tmp_path):
    # Decimal('5985000.00') читается обратно как Decimal, не как float
```

- [ ] **Шаг 4: Запустить тест, убедиться что падает**

Run: `venv/Scripts/python -m pytest tests/test_db.py -v`
Ожидаемо: FAIL, модуль `app.db` не найден.

- [ ] **Шаг 5: Написать `app/config.py`, `app/models.py`, `app/db.py`**

`db.connect` регистрирует адаптер и конвертер `Decimal` через
`sqlite3.register_adapter` / `register_converter` (хранить строкой, колонка типа
`DECIMAL`, соединение с `detect_types=sqlite3.PARSE_DECLTYPES`), включает
`foreign_keys=ON`. `db.migrate` исполняет `CREATE TABLE IF NOT EXISTS` и хранит
номер версии схемы в `PRAGMA user_version`.

- [ ] **Шаг 6: Запустить тесты, убедиться что проходят**

- [ ] **Шаг 7: Написать `Тендеры.bat`** — поднимает `venv/Scripts/python -m app.web`
  и открывает `http://127.0.0.1:8765`. Ручку сервера добавит задача 12;
  на этом шаге достаточно, чтобы файл запускался и печатал адрес.

- [ ] **Шаг 8: Коммит**

```bash
git add . && git commit -m "feat: каркас проекта, схема базы, модели"
```

---

### Задача 2: Репозиторий справочников

**Файлы:**
- Создать: `app/repo.py`
- Тесты: `tests/test_repo.py`

**Интерфейсы:**
- Использует: `app.db`, `app.models` из задачи 1.
- Отдаёт: класс `Repo(conn)` с методами на каждый справочник —
  `save_organization(org: Organization) -> int`,
  `get_organization(id: int) -> Organization`,
  `list_organizations(role: str | None = None) -> list[Organization]`,
  `find_organization_by_inn(inn: str) -> Organization | None`,
  и такие же четвёрки для `employee`, `asset`, `experience`, `service`,
  `customer_profile`, `tender`, `lot_item`;
  `get_settings() -> OrgSettings`, `save_settings(s: OrgSettings) -> None`;
  `delete_<entity>(id: int) -> None`.

- [ ] **Шаг 1: Написать падающие тесты**

```python
def test_save_then_get_organization(repo):
    oid = repo.save_organization(make_org(inn="5904188041"))
    assert repo.get_organization(oid).inn == "5904188041"

def test_save_with_id_updates_in_place(repo):
    # второй save того же id не создаёт вторую запись

def test_list_organizations_filters_by_role(repo):
    # участники и заказчики не смешиваются

def test_find_by_inn_returns_none_when_absent(repo)

def test_lot_items_keep_insertion_order(repo):
    # 11 позиций СЗЗ возвращаются в том же порядке

def test_settings_have_defaults_on_empty_db(repo):
    s = repo.get_settings()
    assert s.insurance_rate == Decimal("0.30")
    assert s.injury_rate    == Decimal("0.002")
    assert s.overhead_rate  == Decimal("0.05")
    assert s.profit_rate    == Decimal("0.10")
    assert s.vat_rate       == Decimal("0.05")
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Реализовать `Repo` в `app/repo.py`**

Один приватный помощник на `INSERT … ON CONFLICT(id) DO UPDATE`, остальные методы
через него. `get_settings` при пустой таблице возвращает значения из общих
ограничений этого плана.

- [ ] **Шаг 4: Запустить, убедиться что проходят**

- [ ] **Шаг 5: Коммит**

---

### Задача 3: Клиент ФНС по ИНН

**Файлы:**
- Создать: `app/inn/__init__.py`, `app/inn/fns.py`
- Тесты: `tests/test_fns.py`, `tests/fixtures/fns_response.json`

**Интерфейсы:**
- Отдаёт: `@dataclass FnsRecord` с полями `short_name, full_name, inn, kpp, ogrn,
  head_fio, head_post, registered: date, region`;
  `lookup(inn: str, *, timeout: float = 30.0, session=None) -> list[FnsRecord]`;
  исключения `InnNotFound`, `CaptchaRequired`, `SourceUnavailable`
  (все наследуют `InnLookupError`, у каждого русское сообщение в `args[0]`).

Протокол: `POST https://egrul.nalog.ru/` с полем `query=<ИНН>` возвращает
`{"t": "<токен>", "captchaRequired": false}`; затем
`GET https://egrul.nalog.ru/search-result/<токен>` возвращает `{"rows": [...]}`.
Ключи строки: `c` краткое наименование, `n` полное, `i` ИНН, `p` КПП, `o` ОГРН,
`g` руководитель в виде `"ДИРЕКТОР: Рабичко Елена Михайловна"`, `r` дата
регистрации `ДД.ММ.ГГГГ`, `rn` регион. Cookies между запросами сохраняются.

- [ ] **Шаг 1: Сохранить фикстуру**

Записать настоящий ответ по ИНН `5904188041` в `tests/fixtures/fns_response.json`.
Сеть в тестах не трогается: `session` подменяется заглушкой.

- [ ] **Шаг 2: Написать падающие тесты**

```python
def test_parses_fixture():
    r = lookup("5904188041", session=stub(FIXTURE))[0]
    assert r.short_name == 'ООО "ТЕХНОСФЕРА"'
    assert (r.inn, r.kpp, r.ogrn) == ("5904188041", "590401001", "1085904010224")
    assert r.head_post == "ДИРЕКТОР"
    assert r.head_fio  == "Рабичко Елена Михайловна"
    assert r.registered == date(2008, 5, 8)

def test_empty_rows_raise_inn_not_found()          # «Организация с таким ИНН не найдена»
def test_captcha_required_raises_captcha()         # captchaRequired: true
def test_timeout_raises_source_unavailable()
def test_several_rows_are_all_returned()           # вызывающий даёт выбрать
def test_rejects_malformed_inn()                   # не 10 и не 12 цифр → InnLookupError до запроса
```

- [ ] **Шаг 3: Запустить, убедиться что падает**

- [ ] **Шаг 4: Реализовать `lookup` в `app/inn/fns.py`**

`requests.Session` с браузерным `User-Agent`. Разбор `g` — split по первому `:`.

- [ ] **Шаг 5: Запустить, убедиться что проходят**

- [ ] **Шаг 6: Отдельно проверить на живой сети**

Run: `venv/Scripts/python -c "from app.inn.fns import lookup; print(lookup('5904188041'))"`
Ожидаемо: одна запись с ИНН 5904188041.

- [ ] **Шаг 7: Коммит**

---

### Задача 4: Rusprofile и сборка карточки организации

**Файлы:**
- Создать: `app/inn/rusprofile.py`, `app/inn/lookup.py`
- Тесты: `tests/test_rusprofile.py`, `tests/test_lookup.py`, `tests/fixtures/rusprofile.html`

**Интерфейсы:**
- Использует: `fns.lookup`, `FnsRecord`, `InnLookupError` из задачи 3;
  `Organization` из задачи 1.
- Отдаёт: `@dataclass RusprofileRecord` с полями `legal_address, okpo, oktmo,
  okved, msp_status_text`;
  `rusprofile.fetch(url: str, *, timeout: float = 30.0, session=None) -> RusprofileRecord | None`
  (возвращает `None` вместо исключения, если страница недоступна или разметка не узнана);
  `lookup.build_draft(inn: str, *, rusprofile_url: str | None = None, session=None)
  -> tuple[Organization, list[str]]` — черновик карточки и список предупреждений
  на русском о том, что не удалось получить.

Поиск на rusprofile по ИНН недоступен: `GET /search?query=<ИНН>` отдаёт 404 без
переадресации, форма на сайте — Vue-компонент без обычного `<form>`. Поэтому
страница адресуется прямой ссылкой, которую пользователь вставляет один раз
на организацию — именно так заказчик работы и описал этот шаг. Без ссылки
черновик собирается из данных ФНС, а недостающие поля перечисляются
в предупреждениях.

Поля `Organization` именно такие (проверено по `app/models.py`):
`bank_bik`, `bank_name`, `bank_account`, `bank_corr_account`, `head_name`,
`head_position`, `head_basis`, `website`, `phones`, `okpo`, `oktmo`, `okved`,
`msp_status`. Полей под ОКАТО, ОКФС, ОКОГУ, ОКОПФ и уставный капитал нет —
их не требует ни один документ пакета, и добавлять их не нужно. Поля под дату
регистрации тоже нет, поэтому `FnsRecord.registered` не отображается никуда.
`msp_status` — перечисление `'', 'micro', 'small', 'medium'`, так что текст
rusprofile «микропредприятие…» отображается в `'micro'` внутри `build_draft`.

`build_draft` обязателен к успеху, если ответил ФНС. Падение rusprofile даёт
предупреждение, а не исключение. Банк, телефоны, почта и сайт в черновике пустые —
их вводит пользователь, и каждое попадает в предупреждения.

- [ ] **Шаг 1: Сохранить фикстуру страницы**

Записать HTML `https://www.rusprofile.ru/id/2854025` в `tests/fixtures/rusprofile.html`.

- [ ] **Шаг 2: Написать падающие тесты**

```python
RP_URL = "https://www.rusprofile.ru/id/2854025"

def test_parses_address_and_codes():
    r = rusprofile.fetch(RP_URL, session=stub(FIXTURE))
    assert "Куйбышева" in r.legal_address
    assert r.okpo == "85081307"
    assert r.oktmo == "57701000001"
    assert r.msp_status_text.startswith("микропредприятие")

def test_unknown_markup_returns_none()             # пустой HTML → None, без исключения
def test_network_error_returns_none()
def test_non_200_returns_none()

def test_draft_merges_both_sources():
    org, warn = build_draft("5904188041", rusprofile_url=RP_URL, session=stub(...))
    assert org.full_name.startswith("ОБЩЕСТВО")    # из ФНС
    assert "Куйбышева" in org.legal_address        # из rusprofile
    assert org.msp_status == "micro"               # текст отображён в перечисление
    assert org.role == "customer"

def test_draft_without_url_uses_fns_only():
    org, warn = build_draft("5904188041", session=stub(...))
    assert org.ogrn == "1085904010224"
    assert org.legal_address == ""
    assert any("адрес" in w.lower() for w in warn)

def test_draft_survives_rusprofile_failure():
    # rusprofile отдаёт None при переданной ссылке
    org, warn = build_draft("5904188041", rusprofile_url=RP_URL, session=stub(...))
    assert org.inn == "5904188041"                 # данные ФНС на месте
    assert any("адрес" in w.lower() for w in warn) # Review Focus 3

def test_draft_warns_about_bank_and_contacts():
    org, warn = build_draft("5904188041", session=stub(...))
    assert any("банк" in w.lower() for w in warn)
    assert org.bank_bik == "" and org.email == ""

def test_draft_propagates_fns_failure():
    # InnNotFound из ФНС поднимается наружу — карточку строить не из чего

def test_draft_picks_row_matching_requested_inn():
    # ФНС вернула несколько строк: берётся совпавшая по ИНН,
    # при нескольких совпавших — первая, с предупреждением
```

- [ ] **Шаг 3: Запустить, убедиться что падает**

- [ ] **Шаг 4: Реализовать оба модуля**

Разбор страницы через `lxml.html`: искать по подписям блоков («Юридический адрес»,
«ОКПО», «ОКТМО»), а не по классам вёрстки — классы меняются чаще.

- [ ] **Шаг 5: Запустить, убедиться что проходят**

- [ ] **Шаг 6: Коммит**

---

### Задача 5: Калькулятор — прямой счёт

**Файлы:**
- Создать: `app/calc.py`
- Тесты: `tests/test_calc.py`

**Интерфейсы:**
- Использует: `OrgSettings` из задачи 1.
- Отдаёт: `@dataclass LaborLine(post: str, headcount: int, hours: Decimal, rate: Decimal)`;
  `@dataclass Estimate(materials, labor, insurance, subcontract, overhead, profit,
  total_wo_vat, vat, total_with_vat)` — все поля `Decimal`;
  `estimate(labor: list[LaborLine], materials: Decimal, subcontract: Decimal,
  settings: OrgSettings) -> Estimate`.

Формула — из раздела 8 описания. Округление каждой статьи до копеек,
`ROUND_HALF_UP`.

- [ ] **Шаг 1: Написать падающий тест на реальных числах КП Арктикгаза**

```python
ARKTIK = [LaborLine("Эколог-проектировщик",        4, Decimal("1200"), Decimal("270")),
          LaborLine("Начальник отдела проектирования", 1, Decimal("360"),  Decimal("420"))]

def test_arktikgaz_labor_matches_submitted():
    e = estimate(ARKTIK, Decimal("41686"), Decimal("1620000"), DEFAULTS)
    assert e.labor      == Decimal("1447200.00")   # 4*1200*270 + 1*360*420
    assert e.insurance  == Decimal("437054.40")    # 30 % + 0,2 %
    assert e.overhead   == Decimal("94212.72")     # 5 % от (ФОТ + страховые)
    assert e.profit     == Decimal("359846.71")
    assert e.total_wo_vat == Decimal("4000000")    # 3 999 999,83 -> целые рубли

def test_profit_is_ten_percent_of_base()
def test_total_rounds_to_whole_roubles():
    # сумма статей 3 999 999,83, итог 4 000 000 — как ROUND(..., +D88) в книге
def test_total_rounding_digits_is_configurable()

def test_vat_uses_settings_rate():
    # при vat_rate = 0.05 и итоге 5 700 000,00 без НДС
    # vat == 285 000,00 и total_with_vat == 5 985 000,00 — как в КП Коряжмы
    # при vat_rate = 0.20 на тех же данных vat == 1 140 000,00
def test_empty_labor_gives_zero_fot_but_keeps_materials()
def test_money_is_rounded_to_kopecks()             # ни одного float в Estimate
```

Модель взята дословно из формул книги `КП.xlsx`, а не восстановлена по виду
документа. Формулы и адреса ячеек выписаны в разделе 1 описания. Порядок:

```
ФОТ            = Σ (кол-во × часы × ставка)                 D16 = F54
страховые      = ФОТ × 30 % + ФОТ × 0,2 %                   D17, D18
расходы на ОТ  = ФОТ + страховые                            D15 = SUM(D16:D18)
субподряд      = Σ сторонних работ                          D21 = G54
накладные      = ROUND(расходы на ОТ × 5 %, 2)              D24, база D15
п.7            = расходы на ОТ + субподряд                  D23
прибыль        = ROUND((п.7 + накладные) × 10 %, 2)         D25
итого без НДС  = ROUND(материалы + п.7 + накладные + прибыль, 0)   D26
```

База накладных — `D15`, то есть расходы на оплату труда **вместе со
страховыми**: подпись «Накладные расходы 5 % от .п 2» указывает на п. 2,
а не на п. 2.1. Прежняя редакция плана читала её как ФОТ и объявляла файл
противоречивым — файл согласован, ошибка была в плане.

Контрольные значения поданного КП: ФОТ 1 447 200,00; страховые 437 054,40;
расходы на ОТ 1 884 254,40; накладные 94 212,72; п.7 3 504 254,40;
прибыль 359 846,71; материалы 41 686,00; сумма до округления 3 999 999,83;
итого без НДС 4 000 000.

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Реализовать `estimate`**

- [ ] **Шаг 4: Запустить, убедиться что проходят**

- [ ] **Шаг 5: Коммит**

---

### Задача 6: Калькулятор — обратный счёт

**Файлы:**
- Изменить: `app/calc.py`
- Тесты: `tests/test_calc_reverse.py`

**Интерфейсы:**
- Отдаёт: `solve_hours(target_wo_vat: Decimal, labor: list[LaborLine],
  materials: Decimal, subcontract: Decimal, settings: OrgSettings) -> list[LaborLine]`;
  исключение `TargetTooLow` с русским сообщением.

Часы масштабируются пропорционально: ищется один множитель `k` на все строки.
Замкнутая форма — пусть

```
A = ФОТ₀ × (1 + страховые + травматизм) × (1 + накладные)
k = ((target − материалы) / (1 + прибыль) − субподряд) / A
```

Множитель `(1 + накладные)` стоит отдельно, потому что накладные считаются
от расходов на оплату труда вместе со страховыми, а не от ФОТ. Для значений
по умолчанию `A = ФОТ₀ × 1,302 × 1,05 = ФОТ₀ × 1,3671`.

`k ≤ 0` или `A == 0` → `TargetTooLow`.

Часы округляются до целых, поэтому точное попадание в целевую сумму
недостижимо в принципе: один час работы стоит денег, и шаг сетки равен
стоимости одного часа самой дешёвой строки с её наценками.

На данных Арктикгаза обратный счёт даёт `k ≈ 1,0000001`: собственные часы
фирмы — 1200 и 360 — и есть решение для цели 4 000 000. Это и подтверждает
вывод описания, что ровная цифра получилась снизу вверх с финальным
округлением, а не подгонкой. Прежняя редакция плана приводила здесь
`k ≈ 1,0112` и часы 1213/364 — артефакт неверной базы накладных.

После пропорционального масштабирования выполняется жадная доводка: часы
отдельных строк меняются на ±1, пока расхождение с целью уменьшается.
Исполнитель обязан пересчитать и выписать в отчёт все контрольные числа
доводки заново: прежние значения в этом разделе получены по неверной модели.

Инвариант, который закрепляет тест: остаточное расхождение не превышает
стоимости одного часа самой дешёвой строки с наценками. Это настоящая граница
достижимого, а не произвольное число.

Пропорции между должностями сохраняются приблизительно, а не с точностью до
одного часа: жёсткий коридор ±1 час вокруг `k × часы` несовместим с инвариантом
расхождения. При реализации нашлась цель, для которой внутри коридора лучшее
достижимое расхождение превышало шаг сетки, а вне коридора инвариант
выполнялся. Конкретные числа того контрпримера получены по прежней, неверной
модели накладных и здесь не приводятся: если коридор когда-нибудь вернут,
контрпример нужно искать заново. Инвариант расхождения важнее коридора:
итог виден заказчику, а часы — внутреннее обоснование трудозатрат.
Тест закрепляет снос не более двух часов на строку.

Остаточное расхождение показывает интерфейс (задача 13) всегда, даже когда оно
в копейках.

- [ ] **Шаг 1: Написать падающие тесты**

```python
def test_solved_hours_reproduce_target():
    solved = solve_hours(Decimal("4000000"), ARKTIK,
                         Decimal("41686"), Decimal("1620000"), DEFAULTS)
    got = estimate(solved, Decimal("41686"), Decimal("1620000"), DEFAULTS)
    assert abs(got.total_wo_vat - Decimal("4000000")) < Decimal("1000")

def test_residual_is_within_one_hour_of_the_cheapest_line():
    # настоящая граница достижимого: расхождение не больше стоимости
    # одного часа самой дешёвой строки с наценками
    solved = solve_hours(Decimal("4000000"), ARKTIK,
                         Decimal("41686"), Decimal("1620000"), DEFAULTS)
    got = estimate(solved, Decimal("41686"), Decimal("1620000"), DEFAULTS)
    # наценка выводится из DEFAULTS, а не прописывается числом:
    # (1 + страховые + травматизм) x (1 + накладные) x (1 + прибыль)
    step = min(l.headcount * l.rate for l in ARKTIK) * LABOR_MARKUP * PROFIT_MARKUP
    assert abs(got.total_wo_vat - Decimal("4000000")) <= step

def test_greedy_refinement_beats_plain_scaling():
    # жадная доводка обязана быть лучше простого масштабирования:
    # 3 999 340,62 без неё против расхождения в десятки рублей с ней

def test_proportions_between_posts_are_preserved():
    # отношение часов 1200:360 сохраняется с точностью до одного часа на строку

def test_target_below_fixed_costs_raises():        # Review Focus 4
    with pytest.raises(TargetTooLow):
        solve_hours(Decimal("1000000"), ARKTIK,
                    Decimal("41686"), Decimal("1620000"), DEFAULTS)

def test_target_exactly_at_fixed_costs_raises()
def test_no_labor_lines_raises_rather_than_dividing_by_zero()
def test_hours_are_whole_numbers()
def test_solve_does_not_mutate_input_lines()
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Реализовать `solve_hours`**

- [ ] **Шаг 4: Запустить, убедиться что проходят**

- [ ] **Шаг 5: Коммит**

---

### Задача 7: Раскладки колонок табличных справок

**Файлы:**
- Создать: `app/render/__init__.py`, `app/render/layouts.py`
- Создать: `layouts/employees_koryazhma.json`, `layouts/assets_koryazhma.json`,
  `layouts/assets_rnpk.json`
- Тесты: `tests/test_layouts.py`

**Интерфейсы:**
- Отдаёт: `@dataclass Column(header: str, parts: list[str], join: str = " ")`;
  `@dataclass TableLayout(source: str, columns: list[Column], numbered: bool)`
  где `source` — одно из `'employees' | 'assets' | 'experience'`;
  `load_layout(path: Path) -> TableLayout`;
  `build_rows(layout: TableLayout, records: Sequence[object]) -> list[list[str]]`;
  исключение `LayoutError`.

`build_rows` возвращает только строки данных — шапка берётся из `layout.columns`.
При `numbered=True` первой колонкой подставляется порядковый номер.
Неизвестное имя поля — `LayoutError` с указанием имени и имени колонки, не `KeyError`.

Колонка описывается не одним полем, а списком полей: разбор поданных документов
показал, что одной клетке формы соответствует несколько полей справочника.
У Коряжмы колонка «Источник» содержит «Право собственности: Свидетельство
№ 59-БД 081725 от 31.10.2013 г.» — это `ownership_basis` и `ownership_doc`
через `": "`. У РНПК колонка «Наименование» содержит «Компьютер Intel Core
i3 2130» — это `name` и `maker_model` через пробел, тогда как у Коряжмы те же
данные разнесены по двум колонкам. Пустые части отбрасываются до соединения,
поэтому отсутствующий документ не оставляет висящего разделителя.

Колонка с `parts: []` рендерится пустой. Это нужно для колонок формы, под
которые в справочнике нет данных: «Мощность (производительность)» у Коряжмы
и «Примечание» у РНПК в поданных документах пусты во всех строках.

Фактические поля моделей (проверено по `app/models.py`):
`Asset` — `name, kind, maker_model, quantity, year, location, ownership_basis,
ownership_doc, purpose, condition`; `Employee` — `full_name, position,
education, graduation_year, specialty, qualification, experience_years,
hourly_rate`; `ExperienceRecord` — `customer_name, contract_number,
contract_date, subject, amount, term, status, service_id, contract_scan,
act_scan`.

Раскладки, которые нужно описать:

| Файл | Колонки |
|---|---|
| `assets_koryazhma.json` | Наименование `[name]` · Изготовитель, марка `[maker_model]` · Количество `[quantity]` · Год выпуска `[year]` · Мощность `[]` · Нынешнее состояние `[location]` · Источник `[ownership_basis, ownership_doc]` через `": "` |
| `assets_rnpk.json` | № п/п (нумерация) · Наименование `[name, maker_model]` · Место нахождения `[location]` · Право собственности `[ownership_basis, ownership_doc]` через `": "` · Назначение `[purpose]` · Характеристика, состояние `[condition]` · Примечание `[]` |
| `employees_koryazhma.json` | № п/п (нумерация) · Должность `[position]` · ФИО `[full_name]` · Образование, квалификация `[education, qualification]` через `", "` · Стаж, лет `[experience_years]` |

Колонка образования берёт `education` целиком, а не собирает фразу из
`graduation_year` и `specialty`: в поданной справке стоит «Окончила Томский
политехнический институт в 1974 г., Химик-технолог» — живая фраза, которую
человек пишет сам. Сборка такой фразы из частей была бы форматированием,
а оно в этой задаче запрещено. `graduation_year` и `specialty` остаются
структурированными полями для форм, которые просят их отдельно.

- [ ] **Шаг 1: Написать падающие тесты**

```python
ASSET = make_asset(name="Компьютер", maker_model="Intel Core i3 2130",
                   quantity=3, location="г.Пермь, Куйбышева, 57а",
                   ownership_basis="Право собственности",
                   ownership_doc="Свидетельство № 59-БД 081725 от 31.10.2013 г.")

def test_same_asset_renders_differently_under_both_layouts():
    kor  = build_rows(load_layout("layouts/assets_koryazhma.json"), [ASSET])[0]
    rnpk = build_rows(load_layout("layouts/assets_rnpk.json"), [ASSET])[0]
    assert len(kor) == 7 and len(rnpk) == 7      # у РНПК первая — нумерация
    # у Коряжмы наименование и марка разнесены, у РНПК — в одной клетке
    assert kor[0] == "Компьютер" and kor[1] == "Intel Core i3 2130"
    assert rnpk[1] == "Компьютер Intel Core i3 2130"

def test_parts_are_joined_with_the_declared_separator():
    kor = build_rows(load_layout("layouts/assets_koryazhma.json"), [ASSET])[0]
    assert kor[6] == ("Право собственности: "
                      "Свидетельство № 59-БД 081725 от 31.10.2013 г.")

def test_empty_part_leaves_no_dangling_separator():
    a = replace(ASSET, ownership_doc="")
    kor = build_rows(load_layout("layouts/assets_koryazhma.json"), [a])[0]
    assert kor[6] == "Право собственности"       # без ": " на конце

def test_column_without_parts_renders_empty():
    kor = build_rows(load_layout("layouts/assets_koryazhma.json"), [ASSET])[0]
    assert kor[4] == ""                          # «Мощность» — нет данных

def test_numbered_layout_adds_sequence():
    rows = build_rows(load_layout("layouts/assets_rnpk.json"), [ASSET, ASSET])
    assert rows[0][0] == "1" and rows[1][0] == "2"

def test_unknown_field_raises_layout_error_naming_field_and_column()
def test_none_renders_as_empty_string_not_the_word_none()
def test_headers_come_from_the_layout_not_from_the_rows()
def test_employees_layout_has_five_columns()
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Реализовать и описать три файла раскладок**

- [ ] **Шаг 4: Запустить, убедиться что проходят**

- [ ] **Шаг 5: Коммит**

---

### Задача 8: Сборка контекста шаблона

**Файлы:**
- Создать: `app/render/context.py`
- Изменить: `app/models.py` (добавить `TenderLaborRow`), `app/db.py` (миграция 2),
  `app/repo.py` (методы новой таблицы)
- Тесты: `tests/test_context.py`, дополнить `tests/test_repo.py`

**Интерфейсы:**
- Использует: `Repo` (задача 2), `estimate` и `LaborLine` (задача 5),
  `load_layout` / `build_rows` (задача 7).
- Отдаёт: `build_context(tender_id: int, repo: Repo) -> dict`;
  `@dataclass TenderLaborRow(tender_id, position: int, post: str,
  headcount: int, hours: Decimal, rate: Decimal, id)`;
  `Repo.list_tender_labor(tender_id) -> list[TenderLaborRow]`,
  `Repo.save_tender_labor(row) -> int`, `Repo.delete_tender_labor(id) -> None`.

### Недостающий кусок: строки ФОТ тендера

В схеме задачи 1 нет ни таблицы, ни связи, которая держала бы трудозатраты
тендера: `service_id` есть только у записи опыта, для отбора. Значит
`build_context` не может построить `Estimate` — калькулятору нужны строки
`LaborLine`, а взять их негде. Добавляется таблица `tender_labor`
(`tender_id` с `ON DELETE CASCADE`, `position` для порядка, `post`,
`headcount`, `hours`, `rate`) миграцией 2; `hours` и `rate` объявляются
`DECIMAL TEXT`, как все денежные колонки.

Почему строки хранятся, а не выводятся из норм каталога: в КП Арктикгаза
стоит «Эколог-проектировщик, 4 чел.», тогда как в фирме один эколог-проектировщик.
Численность в КП — плановая величина, которую подбирают под конкретный тендер,
а не состав штата. Нормы каталога служат заготовкой: задача 13 заполняет ими
строки при создании тендера, дальше пользователь правит. Обратный счёт задачи 6
детерминирован, поэтому часы можно не хранить отдельно — при том же целевом
итоге, тех же нормах и ставках он даст те же часы.

### Ставка НДС: цепочка подстановки

`Tender.vat_rate` → `Organization.vat_rate` участника → `OrgSettings.vat_rate`.
Первое ненулевое значение в этой цепочке и есть ставка тендера. `None` означает
«брать из следующего звена», что и записано в комментариях модели задачи 1.
Форма ТКП Роснефти требует указывать ставку или причину освобождения, поэтому
переопределение на уровне тендера обязательно работает.

### Ключи контекста

Тот словарь, на который опираются все шаблоны: `org` и `customer` (словари полей
организации), `tender`, `items` (позиции лота), `estimate`, `labor` (строки ФОТ),
`employees`, `assets` и `experience` (`{"headers": [...], "rows": [[...]]}` под
раскладку профиля заказчика), `declaration_text`, `conditions` (ответы
«подтверждаю / не подтверждаю»), `today` и `money`.

`money` форматирует `Decimal` в `5 985 000,00`: запятая как десятичный
разделитель, разряды через **неразрывный пробел** `U+00A0`. Неразрывный — чтобы
Word не разорвал число переносом строки посередине. Денежные значения в таблицах
(`experience` — сумма контракта, `items` — цена позиции) проходят через `money`,
потому что задача 7 форматирование запрещает и отдаёт в клетке сырое
`str(Decimal)` вида `5700000.00`.

### Четыре дополнения после реализации

Реализация вскрыла четыре места, где данных негде взять. Все закрываются
миграцией 3 и двумя функциями контекста.

**Раскладка привязывается к профилю заказчика явно.** У `CustomerProfile`
не было поля, называющего раскладку, и привязка по имени документа сделала бы
обещание «новый заказчик — только описание колонок» ложным: заказчик со своей
раскладкой оборудования потребовал бы нового имени документа. Добавляется
`layouts: dict[str, str]` — имя документа в имя файла раскладки, JSON-колонкой,
как `file_prefixes`. Плюс `layouts/experience_default.json`, которого не было.

**Контекст собирается под документ, а не один на весь пакет.** Раскладка
привязана к документу, а ключ `assets` в контексте один — поэтому «первая
раскладка на источник побеждает» ломается именно на реальном случае: в списке
раздела 8 есть и «Справка о наличии оборудования» с раскладкой Коряжмы,
и «Справка о МТР» с раскладкой РНПК, обе по одному и тому же справочнику МТР.
Один контекст на пакет физически не может отдать обе таблицы.
Поэтому `build_context(tender_id, repo, *, document: str | None = None,
experience_ids=None)`: при переданном имени документа табличные ключи
собираются по раскладке этого документа из профиля, без имени — по раскладкам
по умолчанию. Задача 11 вызывает сборку на каждый документ пакета.

**Пропавшие записи опыта не ломают сборку.** Если в сохранённом отборе остался
идентификатор удалённой записи, сборка не падает: такие идентификаторы
пропускаются, а предупреждение на русском попадает в ключ `warnings:
list[str]`. Молча терять опыт нельзя — он влияет на оценку заявки. Ключ
`warnings` повторяет приём `build_draft` из задачи 4, где предупреждения уже
возвращаются рядом с данными.

**Отбор опыта хранится при тендере.** Добавляется `Tender.experience_ids:
list[int]` JSON-колонкой, как `deal_answers`. `build_context(..., experience_ids=None)`
при `None` читает отбор из тендера, иначе берёт переданный — повторная сборка
пакета даёт тот же результат, а задача 13 может показать предпросмотр
с другим отбором, не сохраняя его.

**Материалы и субподряд хранятся разбивкой, а не суммой.** Без них калькуляция
не воспроизводит поданное КП: строки ФОТ, страховых и накладных верны, а прибыль
и итог ниже. Сумма не годится: Excel-калькуляция Арктикгаза печатает обе
разбивки подтаблицами — материалы «инвентарь, оборудование 24 000» и
«расходные материалы 17 686», субподряд «проведение исследований 1 500 000»,
«фоновые концентрации 70 000», «экспертиза проектной документации 50 000».
Добавляется таблица `tender_cost_item` (`tender_id` с `ON DELETE CASCADE`,
`kind` из `'materials' | 'subcontract'`, `position`, `name`, `unit`,
`quantity`, `price`, `note`; `quantity` и `price` — `DECIMAL TEXT`).
`build_context` суммирует её по `kind` и передаёт итоги в `estimate`, а сами
строки кладёт в контекст под ключами `materials` и `subcontract` для задачи 10.

**Даты получают русский формат.** Два ключа рядом с `money`, потому что
документы требуют оба вида: `date_ru` даёт `24.10.2025`, а `date_long` —
`«24» октября 2025 г.`, как в декларации Коряжмы («№ 0524600000225000002
от «24» октября 2025 г.»). Вход — ISO-строка, как хранит задача 1;
пустая строка на входе даёт пустую строку на выходе, а не падение.

### Доктрина напечатанной строки распространяется на подтаблицы

Правило задачи 5 — каждая напечатанная строка воспроизводится калькулятором
от напечатанных соседей — относится и к разбивкам материалов и субподряда.
Поэтому произведение `quantity × price` каждой строки округляется до копеек
один раз, и в `estimate` передаётся сумма **этих округлённых** значений,
а не округление сырой суммы. Иначе подтаблица, которую печатает задача 10,
не сойдётся со строкой итога под ней: сумма округлённых не равна округлению
суммы. На данных Арктикгаза расхождение скрыто (везде `quantity = 1` и целые
цены), но дробные количества — ожидаемый ввод.

### Единый контракт денежных полей в списках

Списки `items`, `materials`, `subcontract` и `labor` типизируют деньги
одинаково: отформатированная строка под естественным именем (`price`, `amount`,
`rate`) и `Decimal` под именем с суффиксом `_value`. У `labor` дополнительно
есть `amount` — `headcount × hours × rate`, которое печатает калькуляция
Арктикгаза. `hours` остаётся `Decimal` без форматирования: это количество
часов, а не деньги, и это оговорено в документации модуля.

Без единого контракта автор шаблона, выучивший приём на трёх списках,
напишет ставку в четвёртом так же и получит в документе `270.00` — точку
вместо запятой и неразрывный пробел мимо. Три следующие задачи пишутся
под этот набор ключей, поэтому расхождение даёт неверный документ,
а не падение.

### Предупреждение о раскладке — только для документов с таблицей

Предупреждение «раскладка не назначена, колонки по умолчанию» имеет смысл
лишь там, где таблица есть. Документы вроде декларации или письма о крупности
сделки таблиц-справочников не содержат, и предупреждение для них — шум.
А шум опаснее молчания: пользователь, привыкший пролистывать предупреждения,
пропустит и настоящее.

Поэтому в `context.py` объявляется константа с именами документов, несущих
табличную справку: справка о кадровых ресурсах, справка о наличии
оборудования, справка о МТР, справка об опыте. Предупреждение выдаётся только
для них. Константу придётся держать в согласии с шаблонами задач 9 и 14 —
это цена того, что контекст не знает, какой шаблон его отрендерит.

- [ ] **Шаг 1: Написать падающие тесты**

```python
NBSP = " "

def test_context_has_all_documented_keys(repo_with_koryazhma)

def test_money_formats_russian():
    assert ctx["money"](Decimal("5985000")) == f"5{NBSP}985{NBSP}000,00"
    assert ctx["money"](Decimal("285000.00")) == f"285{NBSP}000,00"
    assert ctx["money"](Decimal("0")) == "0,00"

def test_money_uses_non_breaking_space_not_plain_space():
    assert " " not in ctx["money"](Decimal("5985000"))

def test_tender_labor_round_trips(repo)          # в tests/test_repo.py
def test_tender_labor_cascades_with_tender(repo) # в tests/test_repo.py
def test_tender_labor_keeps_position_order(repo) # в tests/test_repo.py

def test_estimate_is_built_from_stored_labor_rows():
    # строки tender_labor Арктикгаза дают ФОТ 1 447 200,00

def test_vat_rate_falls_back_tender_then_org_then_settings():
    # 0.07 у тендера побеждает; при None у тендера берётся 0.06 организации;
    # при None у обоих — 0.05 из настроек

def test_assets_table_follows_customer_profile_layout()
def test_declaration_text_comes_from_customer_profile()
def test_experience_includes_only_selected_records()
def test_experience_amount_is_formatted_by_money()
def test_missing_customer_profile_falls_back_to_default_layouts()
def test_unknown_tender_id_raises_with_russian_message()
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Реализовать `build_context`**

- [ ] **Шаг 4: Запустить, убедиться что проходят**

- [ ] **Шаг 5: Коммит**

---

### Задача 9: Рендер документов Word

**Файлы:**
- Создать: `app/render/docx_render.py`
- Тесты: `tests/test_docx_render.py`, `tests/fixtures/minimal.docx`

**Интерфейсы:**
- Использует: контекст из задачи 8.
- Отдаёт: `render_docx(template: Path, context: dict, out: Path) -> Path`;
  исключение `TemplateError` с именем шаблона и причиной на русском.

`docxtpl.DocxTemplate` + `render(context)`. Циклы по таблицам —
через `{%tr for … %}` в строке таблицы шаблона.

- [ ] **Шаг 1: Сделать `tests/fixtures/minimal.docx`** — один абзац `{{ org.short_name }}`
  и таблица с `{%tr for r in assets.rows %}`.

- [ ] **Шаг 2: Написать падающие тесты**

```python
def test_substitutes_scalar(tmp_path)
def test_expands_table_rows(tmp_path):
    # 23 позиции МТР дают 23 строки таблицы

def test_nested_quotes_survive(tmp_path):          # Review Focus 2
    out = render_docx(FIXTURE, {"org": {"short_name": 'АО «НК «Роснефть»'}}, tmp_path/"o.docx")
    assert 'АО «НК «Роснефть»' in docx.Document(out).paragraphs[0].text

def test_xml_special_chars_are_escaped(tmp_path):
    # значение 'Иванов & Партнёры <ООО>' попадает в документ как есть
    # и файл открывается python-docx без ошибки разбора

def test_braces_in_value_are_not_treated_as_syntax(tmp_path):
    # значение '{{ ловушка }}' остаётся текстом

def test_output_opens_in_python_docx(tmp_path)     # файл не повреждён
def test_broken_template_raises_template_error(tmp_path)
```

- [ ] **Шаг 3: Запустить, убедиться что падает**

### Таблица с переменным числом колонок

Шаблон с `{%tr %}` имеет фиксированное число клеток в строке, а раскладки
задачи 7 задают колонки данными. Прямое следствие: заказчик с шестью колонками
и заказчик с семью потребовали бы двух шаблонов, и обещание «новый заказчик —
только описание колонок» стало бы ложным.

`docxtpl` 0.20.2 поддерживает `{%tc %}` — цикл по клеткам строки, наравне
с `{%tr %}` (проверено по исходникам пакета: `for y in ["tr", "tc", "p", "r"]`).

**Теги удаляют клетку или строку, в которой стоят,** поэтому открытие и
закрытие цикла не могут находиться в одной клетке: `for` и `endfor` рядом
оставляют цикл незакрытым, и Jinja сообщает «unknown tag endfor». Рабочая
форма — таблица из четырёх строк, где пара тегов разнесена по отдельным
клеткам, а между ними стоит клетка-прототип:

```
строка 1 (шапка):   [{%tc for h in assets.headers %}] [{{ h }}] [{%tc endfor %}]
строка 2:           [{%tr for row in assets.rows %}]
строка 3 (данные):  [{%tc for c in row %}]            [{{ c }}] [{%tc endfor %}]
строка 4:           [{%tr endfor %}]
```

Средняя клетка каждой пары несёт оформление — границы, шрифт, заливку — и
именно она копируется на все N колонок, поэтому форматирование поданного
документа сохраняется. Строки 2 и 4 исчезают при рендере.

**Ширину клетки в шаблоне ставить «авто».** Иначе ширина прототипа копируется
на все N колонок, и таблица уезжает за край страницы. `tblGrid` docxtpl
перестраивает сам под число клеток.

Один шаблон обслуживает любое число колонок. Это закрепляется тестом: тот же
шаблон рендерится с раскладками на 6 и на 7 колонок, и в собранном файле
проверяются число колонок по строкам, текст клеток, `tblGrid` и оформление
шапки на каждой сгенерированной клетке.

**Объединённые клетки шапки (`gridSpan`, двухъярусная шапка) не проверены** —
в разобранных формах их нет. Проверить до того, как писать шаблон для
заказчика, чья шапка объединена.

### Настройки рендера, без которых документ портится

- **`autoescape=True`** в вызове `render`. По умолчанию `docxtpl` его
  не включает, и значение с `&` или `<` ломает XML документа. Утверждение
  «docxtpl экранирует сам» верно только с этим флагом.
- **`StrictUndefined`**: отсутствующий ключ роняет рендер `TemplateError`
  с русским текстом, а не оставляет пустую клетку. Молча пустое поле в заявке
  хуже отказа: его не заметят до подачи.
- **`finalize`**, превращающий `None` в пустую строку. Иначе Jinja печатает
  в документе слово «None».

- [ ] **Шаг 4: Реализовать `render_docx`**

Значения подаются строками, не `RichText` — тогда `docxtpl` экранирует XML сам
(при `autoescape=True`).

- [ ] **Шаг 5: Запустить, убедиться что проходят**

- [ ] **Шаг 6: Коммит**

---

### Задача 10: Заполнение формы заказчика в Excel

**Файлы:**
- Создать: `app/render/xlsx_render.py`
- Создать: `cellmaps/tkp_rosneft_murmansk.json`
- Тесты: `tests/test_xlsx_render.py`

**Интерфейсы:**
- Использует: контекст из задачи 8.
- Отдаёт: `@dataclass CellMap(sheet: str, cells: dict[str, str],
  rows: RowBlock | None)`;
  `@dataclass RowBlock(first_row: int, source: str, columns: dict[str, str])`
  где ключ — буква колонки, значение — выражение контекста;
  `load_cellmap(path: Path) -> CellMap`;
  `fill_workbook(template: Path, cell_map: CellMap, context: dict, out: Path) -> Path`.

Книга заказчика открывается как есть и сохраняется с сохранением формул
(`load_workbook(data_only=False)`) и защиты листа.

### Заполняются только поля ввода, итоги книга считает сама

Главный принцип: в книгу заказчика записываются **только те клетки, которые
заказчик отвёл участнику**, а всё производное считают её собственные формулы.
Записать вычисленный итог поверх формулы — значит сломать документ, который
проверяющий откроет и пересчитает.

Выгрузка формы ТКП Роснефти (`Форма_ТКП_СЗЗ_29.10.2025[1].xlsx`, лист `Лист1`,
защита листа не включена, 12 объединённых диапазонов) даёт такую карту:

| Клетки | Что там | Кто пишет |
|---|---|---|
| `A6:E6` | шапка: № поз., наименование, цена без НДС, НДС, цена с НДС | заказчик |
| `B7` | `Лот №2000726410` | заказчик |
| `B8` | предмет лота | заказчик |
| `D8` | `Ставка НДС: ____*%` | **мы** — ставка или причина освобождения |
| `A9:A19` | номера позиций 1…11 | заказчик |
| `B9:B19` | наименования: СЗЗ для АЗК № 4, 6, 33, 22, 10, 16, 23, 27, 32, 18, Автоколонна | заказчик |
| `C9:C19` | цена без НДС | **мы** |
| `D9:D19` | сумма НДС | **мы** |
| `E9:E19` | `=C9+D9` | формула книги |
| `C20`, `E20` | `=SUM(C9:C19)`, `=SUM(E9:E19)` | формулы книги |
| `B23:B29`, `C23:C29` | семь условий сделки и условия заказчика | заказчик |
| `E23:E29` | `подтверждаю/не подтверждаю` | **мы** |
| `B31` | `[Руководитель/Должность/ФИО]` | **мы** |

То есть карта пишет `D8`, `C9:C19`, `D9:D19`, `E23:E29` и `B31` — и больше
ничего. Тест обязан проверить, что `E9:E19`, `C20` и `E20` после заполнения
по-прежнему содержат формулы, а не числа.

Позиций в форме ровно одиннадцать, и строки удалять запрещено («при заполнении
не разрешается удалять строки», `A3`). Поэтому если позиций лота в тендере
больше одиннадцати — отказ с русским сообщением, а если меньше — лишние
строки остаются пустыми.

### Решения, принятые по ходу реализации

**`rows` — список блоков, а не один блок.** Карте Арктикгаза нужны три
подтаблицы (материалы, трудозатраты, субподряд) в одной книге. Плюс у блока
есть ёмкость: `last_row`, потому что строки удалять и добавлять запрещено.

**Ключи условий сделки — дословный текст колонки заказчика.** В форме Роснефти
это семь строк `B23:B29`: «Цена и условия:», «Форма, срок и порядок оплаты:»,
«Условия и место оказания услуг:», «Срок оказания услуг:», «Срок действия
коммерческого предложения:», «Срок действия договора:», «Возможность
привлечения субподрядчиков/субисполнителей». Под этими же ключами задача 14
засеивает `Tender.deal_answers`.

Чтобы расхождение не прошло молча, `fill_workbook` **отказывается с русским
сообщением**, если карта ссылается на условие, которого нет в `conditions`.
Иначе в графу ответа попала бы пустота, и заявку отклонили бы за незаполненное
поле, а программа об этом не сказала бы. Отказ срабатывает и на значении
из одних пробелов: графа, выглядящая пустой, ничем не лучше пустой.

Отказ по умолчанию распространяется на **все** одиночные клетки, а не только
на ответы. Поэтому у клетки карты есть признак `optional` (по умолчанию ложь):
им помечаются клетки, пустота которых законна. Это нужно задаче 14 — в книге
`КП.xlsx` часть одиночных полей ввода пуста законно (например разрядность
округления `D88`). Признак записан в карте явно, а не держится в голове того,
кто её правил: умолчание остаётся строгим, а исключения видны в файле.

Подписант `B31` требует фамилии: должность без ФИО — не подписант, и запись
одного слова «Директор» в графу подписи недопустима.

**Клетка `A2` не заполняется.** Она жёлтая, но содержит собственный текст
заказчика с пробелами внутри: «Письмо о подаче оферты от «____»_______ г.
№______ / Наименование претендента…». Записать её целиком означает перебить
текст заказчика, что запрещает `A3` («не вносить иные изменения»). Пробелы
в ней заполняются от руки или в сопроводительном письме. Задача 15 проверяет
это глазами.

**Причина освобождения от НДС пока негде не хранится.** `D8` получает ставку
из `vat_percent`. «0 %» причину освобождения не заменяет, но «Техносфера»
на УСН со ставкой 5 % и от НДС не освобождена, поэтому поле под причину
не добавляется до первого тендера, где оно понадобится.

**Позиции сопоставляются с формой по порядку, а не по названию** — названия
в форме принадлежат заказчику, и сверять по ним нельзя.

### Предел точности openpyxl — проверять глазами в задаче 15

Круговой проход книги через `openpyxl` теряет то, чего библиотека не знает:
в форме Роснефти отброшены 1920 пустых невидимых текстовых полей,
`printerSettings` и `calcChain`. Потеря `calcChain` безвредна — в файл ставится
`fullCalcOnLoad`, и Excel пересчитает итоги при открытии. Но **потеря
`printerSettings` меняет настройки печати**, а заявку печатают и подписывают.
Расширения другого заказчика (например списки x14) библиотека может потерять
молча.

Поэтому задача 15 обязана открыть заполненную форму заказчика в Excel и
сверить разметку печати, а не только числа. Если потери окажутся значимыми,
альтернатива — правка XML внутри книги без кругового прохода; это отдельная
работа, и до доказанной нужды её не делаем.

### Заметка для задачи 14: наша калькуляция устроена так же

Книга `КП.xlsx` Арктикгаза связана формулами внутри себя: `D16 = F54`,
`D21 = G54`, `D13 = F34`. Значит её карта ячеек тоже заполняет **подтаблицы**
(материалы `B35:F38`, трудозатраты `B48:E53`, субподряд `B60:F62`), а главную
таблицу книга пересчитает сама. Это ещё одна причина хранить разбивки
материалов и субподряда строками, а не суммами.

- [ ] **Шаг 1: Написать падающие тесты**

```python
def test_writes_scalar_cell(tmp_path)
def test_fills_eleven_lot_rows(tmp_path)
def test_preserves_existing_formulas(tmp_path):
    # ячейка с '=C9+D9' после заполнения всё ещё содержит формулу
def test_preserves_sheet_protection(tmp_path)
def test_does_not_touch_cells_outside_the_map(tmp_path)
def test_more_rows_than_template_raises_with_counts(tmp_path)
def test_output_opens_in_openpyxl(tmp_path)
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Реализовать и описать карту ячеек**

- [ ] **Шаг 4: Запустить, убедиться что проходят**

- [ ] **Шаг 5: Коммит**

---

### Задача 10а: Профиль заказчика — одно описание на документ

**Файлы:**
- Изменить: `app/models.py`, `app/db.py` (миграция 5), `app/repo.py`,
  `app/render/context.py`
- Тесты: дополнить `tests/test_repo.py`, `tests/test_db.py`, `tests/test_context.py`

Сборке пакета нужно знать, каким шаблоном рендерится документ и какой картой
ячеек заполняется книга. Сейчас этого нет нигде, а описание документа у
заказчика уже размазано по двум колонкам — `layouts` и `file_prefixes`.
Добавить ещё две (`templates`, `cellmaps`) значит получить четыре словаря
с одними ключами, которые разъедутся при первой правке руками.

Поэтому колонки `layouts` и `file_prefixes` заменяются одной:

```
CustomerProfile.docs: dict[str, DocSpec]
DocSpec = {template: str, cellmap: str, layout: str, prefix: str}
```

Ключ — каноническое имя документа. Любое поле `DocSpec` может быть пустым:
пустой `template` означает умолчание `<имя документа>.docx` в `templates_docs/`,
пустой `layout` — раскладку по умолчанию, пустой `prefix` — имя файла без
префикса. Непустой `cellmap` обязателен, если шаблон с расширением `.xlsx`:
без карты книгу заполнять нечем, и это отказ с русским сообщением.

Один объект на документ — это и есть то, что описывает нового заказчика:
какой шаблон, какая карта, какая раскладка, какой префикс имени файла.
Миграция 5 переносит данные из старых колонок, если они есть, и удаляет их.

Это чистка перед задачей 11, а не новая возможность: набор ключей контекста
и поведение раскладок не меняются, меняется только место, где они описаны.

- [ ] **Шаг 1: Написать падающие тесты**

```python
def test_docspec_round_trips_with_every_field(repo)
def test_docspec_tolerates_empty_fields(repo)
def test_migration_5_moves_layouts_and_prefixes_into_docs(repo)
def test_xlsx_template_without_cellmap_is_refused()
def test_context_takes_layout_from_docs(repo)   # вместо profile.layouts
def test_missing_document_in_docs_falls_back_to_defaults(repo)
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Реализовать модель, миграцию 5, методы репозитория и правку контекста**

- [ ] **Шаг 4: Запустить всю сборку под `-W error`**

- [ ] **Шаг 5: Коммит**

---

### Задача 11: Сборка пакета

**Файлы:**
- Создать: `app/render/package.py`
- Тесты: `tests/test_package.py`

**Интерфейсы:**
- Использует: `build_context` (8), `render_docx` (9), `fill_workbook` (10),
  `CustomerProfile` (1).
- Отдаёт: `safe_filename(name: str) -> str`;
  `@dataclass PackageResult(written: list[Path], skipped: list[tuple[str, str]])`
  где `skipped` — пары «документ, причина на русском»;
  `build_package(tender_id: int, documents: list[str], out_dir: Path, repo: Repo)
  -> PackageResult`.

Имена файлов складываются из префикса профиля заказчика и названия документа.
Один неудавшийся документ не отменяет остальные: он попадает в `skipped`.

### Повторная сборка в ту же папку

Отказ при существующем файле — верное умолчание: два почти одинаковых файла
в папке, которая уходит заказчику, хуже отказа. Но пользователь будет
пересобирать пакет постоянно: собрал, открыл, увидел не ту цену, поправил,
собрал снова. С одним лишь отказом ему пришлось бы удалять десять файлов
руками каждый раз.

Поэтому у `build_package` есть параметр `overwrite: bool = False`. Умолчание
остаётся безопасным, а экран задачи 13 спрашивает один раз, когда хотя бы один
целевой файл уже есть: «папка не пуста, перезаписать?» Так пользователь не
теряет ни правленый вручную файл, ни время на уборку.

Чтобы экран мог спросить, ему нужно знать заранее, какие файлы будут записаны.
Поэтому рядом есть `planned_files(tender_id, documents, out_dir, repo)
-> list[tuple[str, Path]]` — пары «документ, итоговый путь», без записи на диск.
Задача 13 сверяет их с содержимым папки и показывает, что именно будет заменено:
список имён убедительнее вопроса «папка не пуста».

**Замена файла — только целиком.** Новый файл пишется под временным именем
и подменяет старый через `os.replace` лишь когда готов полностью. Иначе
неудачный рендер, обрыв записи или файл, открытый в Word, оставили бы
пользователя без прежней версии и без новой. При любом отказе прежний файл
остаётся цел, временный не остаётся.

### Отсутствующая папка шаблонов — одна ошибка, а не десять

Если нет самой папки `templates_docs/`, собирать нечем ни один документ.
Это `PackageError` с русским сообщением, а не десять одинаковых записей
в `skipped`: десять строк об одной причине пользователь читать не станет.
Отсутствие **одного** шаблона при существующей папке остаётся записью
в `skipped` — там причина действительно своя у каждого документа.

- [ ] **Шаг 1: Написать падающие тесты**

```python
def test_safe_filename_strips_forbidden_chars():   # Review Focus 5
    assert safe_filename('АО «НК «Роснефть»: СЗЗ/ТКП №2000726410?') \
        == 'АО «НК «Роснефть» СЗЗ-ТКП №2000726410'
    # ни одного из : / \ | ? * " < > и без точки в конце

def test_safe_filename_truncates_to_windows_limit()
def test_long_subject_reaches_document_in_full(tmp_path):
    # предмет закупки на 300 символов не обрезается внутри документа

def test_prefixes_come_from_customer_profile(tmp_path):
    # Коряжма даёт '1.Декларация.docx', РНПК — '8.1_Справка о МТР.docx'

def test_failing_document_is_skipped_not_fatal(tmp_path)
def test_creates_output_dir_if_absent(tmp_path)
def test_existing_file_is_not_silently_overwritten(tmp_path)
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Реализовать**

- [ ] **Шаг 4: Запустить, убедиться что проходят**

- [ ] **Шаг 5: Коммит**

---

### Задача 12: Веб-ручки

**Файлы:**
- Создать: `app/web/__init__.py`, `app/web/routes.py`, `app/web/__main__.py`
- Тесты: `tests/test_web.py`

**Интерфейсы:**
- Использует: `Repo` (2), `build_draft` (4), `estimate` / `solve_hours` (5, 6),
  `build_package` (11).
- Отдаёт: `create_app(db_path: Path) -> flask.Flask`.

Ручки: `GET/POST/DELETE /api/organizations`, `/api/employees`, `/api/assets`,
`/api/experience`, `/api/services`, `/api/profiles`, `/api/settings`,
`/api/tenders`, `/api/tenders/<id>/items`;
`GET /api/inn/<inn>` — черновик карточки и предупреждения;
`POST /api/tenders/<id>/estimate` — прямой счёт;
`POST /api/tenders/<id>/solve` — подбор часов;
`POST /api/tenders/<id>/build` — сборка пакета;
`GET /` — страница.

Ошибки домена отдаются как `{"error": "<русский текст>"}` со статусом 400,
отказы внешних источников — 502. Трассировки наружу не попадают.

- [ ] **Шаг 1: Написать падающие тесты** (через `app.test_client()`)

```python
def test_crud_round_trip_for_organizations(client)
def test_inn_endpoint_returns_draft_and_warnings(client, monkeypatch)
def test_inn_not_found_gives_400_with_russian_message(client, monkeypatch)
def test_source_unavailable_gives_502(client, monkeypatch)
def test_solve_below_target_gives_400_not_500(client)
def test_build_returns_written_and_skipped(client, tmp_path)
def test_index_serves_html(client)
def test_no_traceback_in_any_error_body(client, monkeypatch)
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Реализовать `create_app` и `__main__`** (порт 8765, `127.0.0.1`)

- [ ] **Шаг 4: Запустить, убедиться что проходят**

- [ ] **Шаг 5: Проверить `Тендеры.bat`** — страница открывается в браузере

- [ ] **Шаг 6: Коммит**

---

### Задача 13: Интерфейс

**Файлы:**
- Создать: `app/web/templates/index.html`, `app/web/static/app.js`,
  `app/web/static/app.css`
- Тесты: `tests/test_ui_smoke.py`

**Интерфейсы:**
- Использует: ручки задачи 12.

Три экрана по разделу 2 описания: справочники, карточка тендера, сборка.
Экран тендера содержит семь полей шапки, редактор позиций лота со вставкой
списком из буфера, калькулятор с переключателем «целевая цена с НДС / без НДС»
и постоянным показом расхождения итога с целевой суммой — даже когда оно в копейках.
На экране справочников кнопка «Загрузить по ИНН» заполняет карточку и показывает
предупреждения о том, чего не хватает. Экран сборки содержит отбор опыта
галочками с предварительным фильтром по услуге закупки, галочки документов
(предзаполненные из профиля заказчика), выбор папки назначения и список
несобранных документов с причинами.

Без сборки и без `npm`: один файл скриптов, запросы через `fetch`.

- [ ] **Шаг 1: Написать падающий тест**

```python
def test_page_references_all_three_screens(client):
    html = client.get("/").text
    for anchor in ("Справочники", "Карточка тендера", "Сборка"):
        assert anchor in html

def test_page_loads_its_own_assets(client):
    # /static/app.js и /static/app.css отдаются со статусом 200
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Сверстать три экрана**

- [ ] **Шаг 4: Запустить тест, затем открыть программу и пройти путь руками:**
  создать заказчика по ИНН, завести тендер, подобрать часы, собрать пакет

- [ ] **Шаг 5: Коммит**

---

### Задача 14: Наполнение справочников и разметка шаблонов

**Файлы:**
- Создать: `scripts/seed.py`
- Создать: `templates_docs/` — восемь файлов по таблице раздела 9 описания
- Тесты: `tests/test_seed.py`

**Интерфейсы:**
- Использует: `Repo` (2).
- Отдаёт: `seed(repo: Repo) -> None` — идемпотентно, повторный запуск не дублирует.

Данные берутся из разобранных заявок в `C:\Users\Рустам\Desktop\Тендеры`:
реквизиты и подписант ООО «Техносфера»; пять сотрудников (Рабичко, Мельникова,
Ивонина, Четина, Шилоносова) с вузами, годами и стажем; около 23 позиций МТР
(офис на Куйбышева 57а по свидетельству 59-БД 081725, девять компьютеров,
четыре печатающих устройства, брошюратор DSB CB-120, Unibind, десять лицензий
«Эколог» от «Интеграла»); каталог услуг КЭР, СЗЗ, НДВ, НДС, НООЛР, ППЭЭ, ПЭК
с формулировками этапов; профили трёх заказчиков — МУП «Полигон»,
АО «АРКТИКГАЗ», АО «НК «Роснефть».

Шаблоны делаются из готовых файлов: значения заменяются на метки, форматирование
не трогается. В этот этап не входят три документа из списка раздела 8 описания:
приложение к справке об опыте и приложение к справке о кадрах ждут Word-версий
от заказчика работы, анкета участника — формы из ближайшего тендера, где её
затребуют. Их шаблоны добавляются позже, без правок ядра.

- [ ] **Шаг 1: Написать падающие тесты**

```python
def test_seed_creates_five_employees(repo)
def test_seed_creates_twenty_plus_assets(repo)
def test_seed_is_idempotent(repo)                  # дважды — те же количества
def test_technosfera_requisites_match_registry(repo):
    org = repo.find_organization_by_inn("5904188041")
    assert (org.kpp, org.ogrn) == ("590401001", "1085904010224")
    assert org.head_name == "Рабичко Елена Михайловна"
def test_every_seeded_profile_has_existing_template_files(repo)
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Написать `seed.py`, разметить восемь шаблонов**

- [ ] **Шаг 4: Запустить, убедиться что проходят**

- [ ] **Шаг 5: Коммит**

---

### Задача 15: Приёмка на тендере Коряжмы

**Файлы:**
- Создать: `tests/test_acceptance_koryazhma.py`

**Интерфейсы:**
- Использует: всё предыдущее.

Собрать пакет по закупке № 0524600000225000002 от 24.10.2025 и сверить
с тем, что было подано вручную.

- [ ] **Шаг 1: Написать приёмочный тест**

```python
def test_koryazhma_package_matches_submitted(tmp_path):
    # 1. seed, 2. завести тендер с реальными данными закупки,
    # 3. build_package, 4. сверить с эталонами:
    #    - декларация содержит номер и дату закупки и полный предмет
    #    - справка о кадрах содержит все пять ФИО
    #    - справка об оборудовании содержит 23 позиции и 7 колонок
    #    - письмо о крупности называет МУП «Полигон» и предмет
    #    - КП даёт итог 5 985 000,00 с НДС при 5 700 000,00 без НДС
    assert result.skipped == []
```

- [ ] **Шаг 2: Запустить, убедиться что падает**

- [ ] **Шаг 3: Довести до прохождения** — правки в шаблонах и раскладках, не в ядре

- [ ] **Шаг 4: Запустить весь набор тестов**

Run: `venv/Scripts/python -m pytest -v`
Ожидаемо: всё зелёное.

- [ ] **Шаг 5: Открыть собранные файлы в Word и Excel** и сравнить с поданными
  глазами: шрифты, границы таблиц, подписные блоки

- [ ] **Шаг 6: Коммит**
