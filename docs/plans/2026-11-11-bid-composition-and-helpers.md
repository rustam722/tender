# Состав заявки, помощник по формам и живой предпросмотр — план работ

> **Исполнителю:** работа идёт через superpowers:subagent-driven-development.
> Шаги отмечены галочками (`- [ ]`) для отслеживания.

**Цель:** закрыть три вещи, которые заказчик назвал сам: он не знает, какие
документы нужны для заявки; карту ячеек для чужой книги Excel пишу я; он не
видит, что стало в документе, пока не соберёт пакет. Плюс резервная копия
базы — дефект, найденный мной.

**Устройство:** всё внутри нынешней программы. Flask + одностраничный экран
на `app.js` без сборки, SQLite с нумерованными миграциями, `Decimal` для
денег, отрисовка через `docxtpl`/`openpyxl`. Новых зависимостей не
добавляется ни одной.

**Шаблон технологий:** Python 3.14, Flask, python-docx, docxtpl, openpyxl,
sqlite3, node:test для браузерной логики.

**Описание:** `docs/specs/2026-11-11-bid-composition-and-helpers-design.md`

## Общие ограничения

Действуют на каждую задачу, повторять в брифах не нужно — читать здесь.

- **Деньги не считаются в браузере.** В `app.js` ни одной арифметики над
  суммами. `grep -nE "parseFloat|parseInt|Number\(|toFixed|Math\.|toLocale|Intl|BigInt"`
  по `app.js` должен давать только счётчики и индексы, как сейчас.
- **Все суммы — `Decimal`**, `float` запрещён.
- **Любой текст для человека — по-русски**, отказ говорит, что сделать.
- **Имена полей для человека** берутся из словарей `app/models.py`, свои не
  пишутся.
- **Пути — только из `app/config.py`**, читаются как `config.X` в момент
  вызова (один шов для подмены в тестах).
- **Имя теста — утверждение о поведении**, и тест обязан падать ровно тогда,
  когда утверждение перестаёт быть правдой. За эту работу десять раз ловили
  проверку, которая лжёт.
- **Мутанты обязательны** на каждую новую проверку: таблица «поломка →
  упавший тест» в отчёте, мутанты ставятся в копии дерева в scratchpad.
- **Запрещено:** писать в `C:\Users\Рустам\Desktop\Тендеры` (настоящие
  поданные заявки, только чтение); писать за пределы папки проекта и
  scratchpad; запускать `app.web` и всё, что создаёт `data/tender.db`
  в папке проекта (это будет рабочая база фирмы; на наборе тестов её держит
  сторож `tests/conftest.py`); трогать папки `demo-пакет-*`; звать
  субагентов; коммитить мутантов.
- **Коммиты по-русски**, последней строкой
  `Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>`.

## На что смотреть при ревью

Пять мест, где эта работа сломается тише всего:

1. **Предпросмотр, который показывает не то, что напечатается.** Если он
   когда-нибудь начнёт подставлять значения сам, а не гнать документ через
   настоящую отрисовку, он станет хуже своего отсутствия.
2. **Временный файл предпросмотра**, оставшийся после ошибки или оборванного
   запроса, и предпросмотр, пишущий в папку вывода.
3. **Строка состава заявки без ссылки на первоисточник** — перечень,
   которому нельзя посмотреть в основание.
4. **Разбор документации, применяющий найденное сам**, без нажатия человека.
5. **Помощник, предлагающий заполнить клетку с формулой** или молчащий
   о клетке, которая ни в одну сумму не входит.

---

## Задача 1: резервная копия базы

**Файлы:**
- Изменить: `app/db.py` (функция копии), `app/web/__main__.py` (вызов до
  миграций), `app/config.py` (папка копий)
- Изменить: `app/web/routes.py` (ручка «сохранить копию»),
  `app/web/static/app.js` (кнопка на экране справочников)
- Тест: `tests/test_db.py`, `tests/test_web.py`

**Интерфейсы:**
- Отдаёт: `db.backup(path: Path, keep: int = 20) -> Path | None` — кладёт
  копию в `config.BACKUPS_DIR`, удаляет лишние, возвращает путь копии или
  `None`, если базы ещё нет.

- [ ] **Шаг 1: падающий тест**

```python
def test_a_backup_is_made_before_migrations_and_the_oldest_are_dropped(tmp_path):
    # 21 копия подряд -> в папке ровно 20, самая старая удалена
def test_the_first_launch_on_an_empty_machine_has_nothing_to_copy(tmp_path):
    # базы нет -> None, папка копий не создаётся, запуск не падает
def test_a_backup_is_a_readable_database_with_the_same_rows(tmp_path):
    # копия открывается и содержит те же организации
def test_a_failed_backup_does_not_stop_the_program(tmp_path, monkeypatch):
    # копирование бросает OSError -> запуск продолжается, сказано по-русски
```

- [ ] **Шаг 2: прогнать, убедиться, что падают**
- [ ] **Шаг 3: `db.backup`** через `sqlite3.Connection.backup` (не
      `shutil.copy`: база может быть открыта), имя
      `tender-ГГГГ-ММ-ДД-ЧЧММ.db`, папка `data/копии`
- [ ] **Шаг 4: вызов в `app/web/__main__.py`** до `migrate`, с сообщением
      в консоль; сбой копии не мешает запуску
- [ ] **Шаг 5: ручка `POST /api/backup`** и кнопка «Сохранить копию
      справочников» на экране справочников; ответ называет полный путь файла
- [ ] **Шаг 6: прогон всего набора, мутанты, коммит**

---

## Задача 2: предпросмотр документа — сервер

**Файлы:**
- Создать: `app/render/preview.py`
- Изменить: `app/web/routes.py` (ручка), `app/render/package.py` (вынести
  общее, если потребуется — но путь записи не менять)
- Тест: `tests/test_preview.py`

**Интерфейсы:**
- Отдаёт: `preview_document(tender_id, document, repo, experience_ids=None) -> Preview`,
  где `Preview` — датакласс: `kind` («word» / «excel» / «refusal»),
  `blocks` (список абзацев и таблиц), `warnings` (список строк),
  `refusal` (строка или None).
- Блок абзаца: `{"kind": "p", "text": str}`; блок таблицы:
  `{"kind": "table", "rows": [[str, ...], ...]}`.

- [ ] **Шаг 1: падающий тест**

```python
def test_a_preview_shows_the_paragraphs_and_tables_the_document_prints():
    # значения из тендера видны в блоках; порядок как в документе
def test_a_preview_names_the_blank_mandatory_fields_like_the_package_does():
    # те же имена граф, что у blank_fields_warning
def test_a_document_that_cannot_be_built_is_a_refusal_not_an_empty_preview():
    # форма заказчика без формы в профиле -> refusal с русским текстом
def test_the_temporary_file_of_a_preview_is_always_removed():
    # после успеха, после отказа и после исключения в отрисовке
def test_a_preview_never_writes_into_the_output_folder_of_the_tender():
    # папка вывода тендера пуста после предпросмотра
def test_an_excel_document_is_not_rendered_but_names_the_cells_it_would_fill():
    # kind == "excel", перечислены клетки из карты
```

- [ ] **Шаг 2: прогнать, убедиться, что падают**
- [ ] **Шаг 3: `preview_document`** — `build_context` → `render_docx`
      во временный файл рядом с временной папкой (`tempfile`), чтение
      абзацев и таблиц через `python-docx`, `finally` с удалением
- [ ] **Шаг 4: ручка `POST /api/tenders/<id>/preview`** с телом
      `{"document": "<имя>"}`; отказ — 200 с `kind: "refusal"`, а не 400:
      отказ по данным не ошибка запроса
- [ ] **Шаг 5: прогон, мутанты, коммит**

---

## Задача 3: предпросмотр документа — экран

**Файлы:**
- Изменить: `app/web/static/app.js` (область предпросмотра в карточке
  тендера), `app/web/static/app.css`
- Тест: `tests/ui_logic.test.js`, `tests/ui_render.test.js`

- [ ] **Шаг 1: падающий тест** (node:test)

```js
// предпросмотр запрашивается с задержкой: три правки подряд -> один запрос
// ответ старого запроса не затирает новый (маркер отрисовки)
// отказ показывается как отказ, предупреждения -- как предупреждения
// таблица документа рисуется таблицей, абзац -- абзацем
// выбор другого документа запрашивает его, а не показывает прежний
```

- [ ] **Шаг 2: прогнать, убедиться, что падают**
- [ ] **Шаг 3: область предпросмотра** — выбор документа, блоки, пауза в
      вводе 700 мс, маркер отрисовки как `isCurrentRender`
- [ ] **Шаг 4: предупреждения и отказ** в той же области, видами
      `notice()`
- [ ] **Шаг 5: `grep` по арифметике в `app.js`** — убедиться, что новых
      вычислений над суммами нет
- [ ] **Шаг 6: прогон, мутанты, коммит**

---

## Задача 4: состав заявки — данные и репозиторий

**Файлы:**
- Изменить: `app/db.py` (миграция 7), `app/models.py` (датакласс и словарь
  имён), `app/repo.py` (сущность и методы)
- Тест: `tests/test_db.py`, `tests/test_repo.py`

**Интерфейсы:**
- Отдаёт: `BidItem` — `tender_id`, `position`, `name`, `kind`, `document`,
  `source`, `source_quote`, `done`, `note`, `id`; константы сортов
  `BID_KIND_OURS`, `BID_KIND_CUSTOMER_FORM`, `BID_KIND_SCAN`,
  `BID_KIND_PLATFORM` со словарём русских названий.
- Отдаёт: `Repo.list_bid_items(tender_id)`, `save_bid_item`, `delete_bid_item`,
  `replace_bid_items(tender_id, items)` — одной транзакцией, как
  `replace_lot_items`.

- [ ] **Шаг 1: падающий тест**

```python
def test_the_bid_items_of_a_tender_keep_their_order_and_their_source():
def test_replacing_the_bid_items_is_all_or_nothing():
    # плохая строка -> ни одна не записана, прежние целы
def test_an_unknown_sort_of_bid_item_is_refused_in_russian():
def test_deleting_a_tender_deletes_its_bid_items():
    # ON DELETE CASCADE, как у позиций лота
def test_migration_seven_keeps_the_tenders_that_already_exist():
    # база версии 6 с тендерами -> миграция не теряет их
```

- [ ] **Шаг 2: прогнать, убедиться, что падают**
- [ ] **Шаг 3: миграция 7** — таблица `bid_item`, `tender_id` с
      `ON DELETE CASCADE`, `PRAGMA user_version = 7`
- [ ] **Шаг 4: датакласс и сущность репозитория** по образцу `LotItem` /
      `_LOT_ITEM`
- [ ] **Шаг 5: прогон, мутанты, коммит**

---

## Задача 5: состав заявки — экран и сверка на сборке

**Файлы:**
- Изменить: `app/web/schema.py` (проверка строк), `app/web/routes.py` (ручки),
  `app/web/static/app.js` (раздел «Состав заявки» в карточке тендера)
- Изменить: `app/render/package.py` (сверка состава в итоге сборки)
- Тест: `tests/test_web.py`, `tests/test_package.py`, `tests/ui_logic.test.js`

**Интерфейсы:**
- Потребляет: `BidItem` и методы репозитория из задачи 4.
- Отдаёт: `package.bid_composition_warning(tender_id, repo, written, skipped) -> str`
  — предупреждение о составе или пустая строка.

- [ ] **Шаг 1: падающий тест**

```python
def test_the_package_counts_the_bid_items_it_cannot_build_itself():
    # «ваших сканов не отмечено 2: устав, выписка ЕГРЮЛ»
def test_a_bid_item_whose_document_was_built_is_not_called_missing():
def test_a_complete_bid_composition_says_nothing():
    # всё отмечено и собрано -> предупреждения нет
def test_a_tender_without_a_bid_composition_is_not_nagged():
    # состава нет вовсе -> молчание: человек мог его не заводить
def test_the_sort_of_a_bid_item_decides_what_the_person_is_told_to_do():
    # четыре сорта -> четыре разных совета
```

- [ ] **Шаг 2: прогнать, убедиться, что падают**
- [ ] **Шаг 3: ручки** `GET/POST /api/tenders/<id>/bid-items` по образцу
      позиций лота (замена списком, всё-или-ничего)
- [ ] **Шаг 4: раздел «Состав заявки»** в карточке тендера: строки с сортом,
      отметкой, источником и цитатой (цитата — раскрываемая, чтобы не
      загромождать); кнопка «Взять из профиля заказчика»
- [ ] **Шаг 5: сверка в сборке** — `bid_composition_warning` в итоге
- [ ] **Шаг 6: прогон, мутанты, коммит**

---

## Задача 6: базовые наборы по закону

**Файлы:**
- Создать: `bid_sets/44fz-auction.json`, `bid_sets/44fz-tender.json`,
  `bid_sets/223fz-general.json`
- Создать: `app/bid_sets.py` (загрузка и отказы)
- Изменить: `app/config.py` (`BID_SETS_DIR`), `app/web/routes.py`,
  `app/web/static/app.js` (кнопка «Базовый набор по закону»)
- Тест: `tests/test_bid_sets.py`

**Интерфейсы:**
- Отдаёт: `load_bid_set(law: str, method: str) -> list[BidItem]` —
  строки с `source` вида «базовый набор 44-ФЗ, ст. 43 ч. 1 п. 2».

**ВАЖНО для исполнителя:** содержимое файлов наборов пишет контролёр, а не
исполнитель. Исполнитель делает загрузку, проверку формата, отказы и экран;
в файлы кладёт **заготовки с пометкой «проверяется по тексту закона»**.
Выдумывать пункты закона запрещено: неверная ссылка в юридическом документе
хуже её отсутствия.

- [ ] **Шаг 1: падающий тест**

```python
def test_a_bid_set_file_that_names_an_unknown_sort_is_refused_in_russian():
def test_a_bid_set_row_without_a_source_reference_is_refused():
    # ссылка на пункт обязательна: перечень без основания -- то, что мы ругали
def test_every_bid_set_file_in_the_folder_loads():
    # данными, как сходимость профилей с папкой шаблонов
def test_an_unknown_pair_of_law_and_method_says_which_pairs_exist():
```

- [ ] **Шаг 2: прогнать, убедиться, что падают**
- [ ] **Шаг 3: загрузка и отказы** по образцу `load_cellmap` / `load_layout`
- [ ] **Шаг 4: заготовки файлов** с пометками; законные тексты впишет
      контролёр
- [ ] **Шаг 5: кнопка «Базовый набор по закону»** — предложение, а не
      применение: показать список и ждать «Применить»
- [ ] **Шаг 6: прогон, мутанты, коммит**

---

## Задача 7: разбор документации заказчика

**Файлы:**
- Создать: `app/docs_reader.py`
- Изменить: `app/web/routes.py` (загрузка файла и разбор),
  `app/web/static/app.js` (кнопка «Разобрать документацию»)
- Тест: `tests/test_docs_reader.py`, `tests/fixtures/` (учебные .docx)

**Интерфейсы:**
- Отдаёт: `find_bid_composition(path: Path) -> list[Found]`, где `Found` —
  `heading` (где нашлось), `quote` (дословный текст), `items` (разбитые
  строки с предполагаемым сортом).

- [ ] **Шаг 1: падающий тест**

```python
def test_the_clause_about_the_bid_composition_is_found_by_its_own_words():
    # «заявка должна содержать», «в состав заявки входят», «состав заявки»
def test_the_quote_is_the_text_of_the_document_word_for_word():
    # ни одного слова от программы: человек сверяет с файлом
def test_a_numbered_list_inside_the_clause_becomes_separate_items():
def test_a_document_without_such_a_clause_says_so_and_does_not_guess():
def test_a_file_that_is_not_a_docx_is_refused_in_russian():
def test_the_sort_of_an_item_is_a_guess_and_is_marked_as_one():
    # «устав», «выписка» -> ваш скан; «декларация» -> наш документ; остальное
    # -- «не определён», и это честнее, чем угадать
def test_nothing_is_written_into_the_bid_composition_by_the_reader_itself():
```

- [ ] **Шаг 2: прогнать, убедиться, что падают**
- [ ] **Шаг 3: поиск раздела** по устойчивым оборотам, с запасом на
      разрыв абзаца и на таблицу (у Автоспецтранса состав лежит в пункте
      3.4 текстом, а требования — в таблице информационной карты)
- [ ] **Шаг 4: разбор на строки** и предположение сорта по словарю слов,
      с пометкой «предположение»
- [ ] **Шаг 5: ручка и экран** — показать найденное, ждать «Применить»
- [ ] **Шаг 6: прогон, мутанты, коммит**

**Проба на настоящих файлах (обязательна):** документация Автоспецтранса
(223-ФЗ) и документация Коряжмы (44-ФЗ). Обе лежат у контролёра; если
исполнителю они нужны — спросить, не искать самому в папках фирмы.

---

## Задача 8: помощник по формам — чтение книги

**Файлы:**
- Создать: `app/cellmap_helper.py`
- Тест: `tests/test_cellmap_helper.py`

**Интерфейсы:**
- Отдаёт: `read_workbook(path) -> list[SheetView]` — листы, клетки
  с текстом/формулой/числом, адреса.
- Отдаёт: `unreachable_cells(path, addresses) -> list[str]` — те из
  `addresses`, что ни в одну формулу книги не входят.

- [ ] **Шаг 1: падающий тест**

```python
def test_the_view_of_a_workbook_shows_the_text_formula_and_address_of_a_cell():
def test_a_cell_that_holds_a_formula_is_marked_as_one():
def test_a_cell_no_formula_ever_reads_is_named_unreachable():
    # форма Арктикгаза: строки субподряда 4-6 ни в одну сумму не входят --
    # это нашли руками, читая формулы, и это должно находиться само
def test_a_cell_inside_a_summed_range_is_not_called_unreachable():
    # СУММ(G48:G50) -- G49 достижима
def test_an_indirect_sum_through_another_cell_is_followed():
    # D21 = G54, G54 = СУММ(G48:G50): G49 достижима через два шага
def test_a_workbook_that_cannot_be_opened_is_refused_in_russian():
```

- [ ] **Шаг 2: прогнать, убедиться, что падают**
- [ ] **Шаг 3: чтение книги** через `openpyxl` в двух проходах
      (`data_only=False` для формул, `data_only=True` для значений)
- [ ] **Шаг 4: достижимость** — разбор диапазонов и ссылок в формулах,
      обход по ссылкам с защитой от циклов
- [ ] **Шаг 5: прогон, мутанты, коммит**

---

## Задача 9: помощник по формам — разметка и проверка

**Файлы:**
- Изменить: `app/cellmap_helper.py` (сборка карты), `app/web/routes.py`,
  `app/web/static/app.js` (новый экран «Формы заказчиков»),
  `app/web/static/app.css`
- Тест: `tests/test_cellmap_helper.py`, `tests/test_web.py`,
  `tests/ui_logic.test.js`

**Интерфейсы:**
- Потребляет: `read_workbook`, `unreachable_cells` из задачи 8.
- Отдаёт: `build_cellmap(sheet, cells, rows, rates) -> dict` — карта того же
  формата, что написанные руками, и `save_cellmap(name, data)`.

- [ ] **Шаг 1: падающий тест**

```python
def test_a_cellmap_built_by_the_helper_loads_by_the_same_loader():
    # формат один: load_cellmap принимает её без оговорок
def test_the_helper_refuses_to_fill_a_cell_that_holds_a_formula():
    # и называет формулу, которую человек затёр бы
def test_the_helper_warns_about_a_cell_no_formula_reads():
def test_a_row_block_is_written_with_its_first_row_last_row_and_columns():
def test_a_field_is_offered_by_its_russian_name_not_by_its_path():
def test_a_map_is_never_saved_over_another_without_being_asked():
```

- [ ] **Шаг 2: прогнать, убедиться, что падают**
- [ ] **Шаг 3: сборка и сохранение карты**
- [ ] **Шаг 4: экран «Формы заказчиков»** — книга, листы, клетки, выбор
      поля по русскому имени, блоки строк
- [ ] **Шаг 5: кнопка «Проверить на тендере»** — заполнить копию и
      показать, что куда легло; исходный шаблон не трогать
- [ ] **Шаг 6: прогон, мутанты, коммит**

---

## Задача 10: живой прогон руками и приёмка этапа

**Файлы:**
- Создать: `docs/ПРИЁМКА-этап-2.md`
- Изменить: `docs/СОСТОЯНИЕ.md`, `docs/РЕШЕНИЯ.md`, `docs/КАК-РАБОТАТЬ.md`

**Это делает контролёр сам, не исполнитель.** За первый этап живая работа
руками нашла шесть дефектов, мимо которых прошли восемь кругов ревью, 813
тестов и сотня мутантов. Это самая дешёвая проверка из всех, и подменять её
тестами нельзя.

- [ ] **Шаг 1:** пройти руками состав заявки на закупке Автоспецтранса:
      базовый набор → разбор документации → правка → сборка со сверкой
- [ ] **Шаг 2:** пройти помощником по формам книгу Арктикгаза с нуля и
      сверить полученную карту с написанной руками (`cellmaps/kp_arktikgaz.json`)
- [ ] **Шаг 3:** предпросмотр: поменять ИНН, увидеть изменение; стереть
      обязательную графу, увидеть предупреждение
- [ ] **Шаг 4:** записать находки в журнал и закрыть их кругами правок
- [ ] **Шаг 5:** обновить документы заказчика и отправить ему
