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

## Общие ограничения

- Python 3.14.6, Windows 11, без прав администратора.
- Изолированный `venv` в папке программы; системный Python не используется напрямую.
- Работа на одном компьютере. Многопользовательский режим, блокировки и сетевая база не нужны.
- Все данные локально. Сеть — только `egrul.nalog.ru` и `www.rusprofile.ru`.
- Деньги и часы — только `decimal.Decimal`, никогда `float`.
- Проценты по умолчанию: страховые 30 %, травматизм 0,2 %, накладные 5 % от ФОТ,
  прибыль 10 % от (ФОТ + страховые + субподряд + накладные).
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
    assert e.overhead   == Decimal("72360.00")     # 5 % от ФОТ

def test_profit_is_ten_percent_of_base()
def test_total_sums_all_articles()

def test_vat_uses_settings_rate():
    # при vat_rate = 0.05 и итоге 5 700 000,00 без НДС
    # vat == 285 000,00 и total_with_vat == 5 985 000,00 — как в КП Коряжмы
    # при vat_rate = 0.20 на тех же данных vat == 1 140 000,00
def test_empty_labor_gives_zero_fot_but_keeps_materials()
def test_money_is_rounded_to_kopecks()             # ни одного float в Estimate
```

Числа страховых и накладных в поданном КП посчитаны от разных баз
(страховые от 1 447 200, накладные подписаны «5 % от п. 2»). План фиксирует
единую базу — ФОТ — и проверяет её. Расхождение с поданным файлом в накладных
ожидаемо и допустимо: проценты настраиваются.

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
Замкнутая форма — пусть `A = ФОТ₀ × (1 + страховые + травматизм + накладные)`, тогда

```
k = ((target − материалы) / (1 + прибыль) − субподряд) / A
```

`k ≤ 0` или `A == 0` → `TargetTooLow`. Часы округляются до целых,
поэтому точное попадание в копейку не гарантируется — остаточное расхождение
показывает интерфейс (задача 13).

- [ ] **Шаг 1: Написать падающие тесты**

```python
def test_solved_hours_reproduce_target():
    solved = solve_hours(Decimal("4000000"), ARKTIK,
                         Decimal("41686"), Decimal("1620000"), DEFAULTS)
    got = estimate(solved, Decimal("41686"), Decimal("1620000"), DEFAULTS)
    assert abs(got.total_wo_vat - Decimal("4000000")) < Decimal("100")

def test_proportions_between_posts_are_preserved():
    # отношение часов 1200:360 сохраняется после подбора

def test_target_below_fixed_costs_raises():        # Review Focus 4
    with pytest.raises(TargetTooLow):
        solve_hours(Decimal("1000000"), ARKTIK,
                    Decimal("41686"), Decimal("1620000"), DEFAULTS)

def test_target_exactly_at_fixed_costs_raises()
def test_no_labor_lines_raises_rather_than_dividing_by_zero()
def test_hours_are_whole_numbers()
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
- Отдаёт: `@dataclass Column(header: str, field: str)`;
  `@dataclass TableLayout(source: str, columns: list[Column], numbered: bool)`
  где `source` — одно из `'employees' | 'assets' | 'experience'`;
  `load_layout(path: Path) -> TableLayout`;
  `build_rows(layout: TableLayout, records: Sequence[object]) -> list[list[str]]`;
  исключение `LayoutError`.

`build_rows` возвращает только строки данных — шапка берётся из `layout.columns`.
При `numbered=True` первой колонкой подставляется порядковый номер.
Неизвестное имя поля — `LayoutError` с указанием имени, не `KeyError`.

Это и есть то, что пользователь настраивает сам: колонки заказчика описываются
в JSON, данные берутся из того же справочника. Колонки из описания —
Коряжма: `Наименование · Изготовитель, марка · Количество · Год выпуска ·
Мощность · Нынешнее состояние · Источник`; РНПК: `Наименование · Место нахождения ·
Право собственности · Назначение в отношении предмета закупки ·
Характеристика, состояние · Примечание`.

- [ ] **Шаг 1: Написать падающие тесты**

```python
def test_same_assets_render_under_both_layouts():
    assets = [make_asset(name="Компьютер", maker_model="Intel Core i3 2130", qty=3)]
    kor = build_rows(load_layout("layouts/assets_koryazhma.json"), assets)
    rnpk = build_rows(load_layout("layouts/assets_rnpk.json"), assets)
    assert len(kor[0]) == 7 and len(rnpk[0]) == 6   # разные колонки
    assert "Intel Core i3 2130" in " ".join(kor[0])
    assert "Intel Core i3 2130" in " ".join(rnpk[0])

def test_numbered_layout_adds_sequence()
def test_unknown_field_raises_layout_error()
def test_missing_value_renders_empty_string_not_none()
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
- Тесты: `tests/test_context.py`

**Интерфейсы:**
- Использует: `Repo` (задача 2), `estimate` (задача 5),
  `load_layout` / `build_rows` (задача 7).
- Отдаёт: `build_context(tender_id: int, repo: Repo) -> dict`.

Ключи контекста — тот словарь, на который опираются все шаблоны:
`org` и `customer` (словари полей организации), `tender`, `items` (позиции лота),
`estimate`, `labor` (строки ФОТ), `employees` и `assets` и `experience`
(`{"headers": [...], "rows": [[...]]}` под раскладку профиля заказчика),
`declaration_text`, `conditions` (ответы «подтверждаю / не подтверждаю»),
`today` и `money` — функция форматирования `Decimal` в `5 985 000,00`.

- [ ] **Шаг 1: Написать падающие тесты**

```python
def test_context_has_all_documented_keys(repo_with_koryazhma)
def test_money_formats_russian():
    assert ctx["money"](Decimal("5985000")) == "5 985 000,00"
def test_assets_table_follows_customer_profile_layout()
def test_declaration_text_comes_from_customer_profile()
def test_experience_includes_only_selected_records()
def test_missing_customer_profile_falls_back_to_default_layouts()
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

- [ ] **Шаг 4: Реализовать `render_docx`**

Значения подаются строками, не `RichText` — тогда `docxtpl` экранирует XML сам.

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
(`load_workbook(data_only=False)`) и защиты листа. Карта ячеек для ТКП Роснефти
покрывает 11 позиций лота начиная со строки с первым АЗК, колонку цены без НДС
и колонку ответов «подтверждаю / не подтверждаю».

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
    assert org.head_fio == "Рабичко Елена Михайловна"
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
