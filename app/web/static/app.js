/* Тендеры: экраны программы. Один файл, без сборки и без внешних библиотек.
 *
 * Главное правило: ЭКРАН НИЧЕГО НЕ СЧИТАЕТ. В JavaScript нет Decimal, а любая
 * арифметика над деньгами здесь -- float. Деньги, часы, количества, ставки в
 * процентах, суммы строк, итоги и расхождение с целевой ценой приходят с сервера
 * готовыми строками («5 985 000,00»); экран их показывает и отправляет обратно
 * то, что ввёл человек. Конвенция сервера: голое имя поля -- готовая строка,
 * <поле>_value -- сырое значение, <поле>_percent -- ставка в процентах. Запись
 * отправляется на сервер ЦЕЛИКОМ, как получена, с правками поверх: что
 * победит -- голое имя или _value -- решает сервер (см. Задачу 12).
 *
 * Файл в двух частях:
 *   1. логика без страницы (запросы, разбор вставки, замок сборки). Её же
 *      гоняют тесты в Node (tests/ui_logic.test.js);
 *   2. страница: DOM строится функцией h(), в разметку попадает только textContent.
 */
'use strict';

/* ======================================================================== */
/* 1. Логика без страницы                                                     */
/* ======================================================================== */

const NO_CONNECTION = 'Не удалось связаться с программой. Проверьте, что чёрное окно программы '
  + 'не закрыто, и повторите действие.';

class ApiFailure extends Error {
  constructor(message, status, body) {
    super(message);
    this.name = 'ApiFailure';
    this.status = status;
    this.body = body || null;
    this.field = body && body.field ? body.field : undefined;
    this.path = body && body.path ? body.path : undefined;
  }
}

function sentenceFor(status, data) {
  if (data && typeof data.error === 'string' && data.error.trim()) return data.error;
  return `Программа ответила ошибкой (код ${status}), но не объяснила причину. Повторите действие; `
    + 'если ошибка повторяется, покажите чёрное окно программы разработчику.';
}

/* Единственное место, где страница ходит на сервер. Каждый POST и PUT идёт с
 * Content-Type: application/json (иначе сервер отвечает 415), даже с пустым телом. */
async function api(method, path, body) {
  const init = { method, headers: { Accept: 'application/json' } };
  if (method === 'POST' || method === 'PUT') {
    init.headers['Content-Type'] = 'application/json';
    init.body = JSON.stringify(body === undefined ? {} : body);
  }
  let response;
  try {
    response = await fetch(path, init);
  } catch (error) {
    throw new ApiFailure(NO_CONNECTION, 0);
  }
  let data = null;
  try {
    const text = await response.text();
    data = text ? JSON.parse(text) : null;
  } catch (error) {
    data = null;
  }
  if (!response.ok) throw new ApiFailure(sentenceFor(response.status, data), response.status, data);
  return data;
}

/* Ответ на запись может нести warning: запись сохранилась, а вид после неё
 * построить не удалось. Суммы в таком ответе могут отсутствовать; экран их не
 * достраивает и запись не отправляет повторно. */
function splitWarning(body) {
  const { warning, ...record } = body || {};
  return { record, warning };
}

function recordForWrite(record, edits) {
  const { warning, ...rest } = record;
  return { ...rest, ...edits };
}

async function writeRecord(path, record) {
  return splitWarning(await api('POST', path, record));
}

/* Вставка списком (из Excel или Word): таблица копируется с табуляцией между
 * колонками. Цена остаётся ТЕКСТОМ, как вставлена: разбирает её сервер. Строку без
 * чёткого разделителя не угадываем -- «АЗК № 4 350 000,00» не отличить от
 * «АЗК № 4 350» и «000,00», а неверная цена в заявке хуже лишнего вопроса. */
function parsePasted(text) {
  const lines = String(text).split(/\r\n|\r|\n/).filter((line) => line.trim() !== '');
  const delimiter = lines.some((line) => line.includes('\t')) ? '\t'
    : lines.some((line) => line.includes(';')) ? ';' : '';
  const rows = lines.map((line) => {
    let cells = delimiter ? line.split(delimiter).map((cell) => cell.trim()) : [line.trim()];
    if (cells.length >= 3 && /^\d+[.)]?$/.test(cells[0])) cells = cells.slice(1);
    while (cells.length > 1 && cells.at(-1) === '') cells = cells.slice(0, -1);
    const name = cells[0];
    const price = cells.length > 1 ? cells.at(-1) : '';
    const problems = [];
    if (!name) problems.push('Нет наименования.');
    if (!price) {
      problems.push(delimiter
        ? 'Не указана цена.'
        : 'Не найдена цена: отделите её от наименования знаком табуляции (так копирует Excel) или точкой с запятой.');
    }
    if (cells.length > 2) problems.push('Больше двух колонок: проверьте, какая из них цена, и поправьте строку.');
    return { name, price, problem: problems.join(' ') };
  });
  return { rows };
}

function lotPayload(rows) {
  return { items: rows.map((row, position) => ({ position, name: row.name, price: row.price })) };
}

/* Знак расхождения читается по строке сервера, без вычитания: «0.17» -- итог
 * выше цели, «-0.17» -- ниже, «0.00» -- совпал. */
function differenceKind(text) {
  const raw = String(text === undefined || text === null ? '' : text).trim();
  if (!raw) return 'unknown';
  if (/^-?0+([.,]0+)?$/.test(raw)) return 'equal';
  return raw.startsWith('-') ? 'below' : 'above';
}

/* Галочки документов: сначала состав пакета из профиля заказчика (отмечены),
 * затем остальное, что где-либо описано в профилях (не отмечено). Никакого
 * списка «всех документов» экран не придумывает: его берёт только из профилей. */
function documentChoices(profiles, tender) {
  const mine = profiles.find((profile) => String(profile.customer_id) === String(tender.customer_id));
  const seen = new Set();
  const choices = [];
  const add = (name, checked) => {
    if (name && !seen.has(name)) {
      seen.add(name);
      choices.push({ name, checked });
    }
  };
  if (mine) {
    (mine.documents || []).forEach((name) => add(name, true));
    Object.keys(mine.docs || {}).forEach((name) => add(name, false));
  }
  profiles.filter((profile) => profile !== mine).forEach((profile) => {
    (profile.documents || []).forEach((name) => add(name, false));
    Object.keys(profile.docs || {}).forEach((name) => add(name, false));
  });
  return choices;
}

/* В тендере нет поля «услуга», поэтому услугу для предварительного фильтра опыта
 * угадываем по предмету закупки -- и только когда совпадение одно. */
function inferServiceId(subject, services) {
  const escape = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const hits = services.filter((service) => [service.code, service.name].some((word) => {
    if (!word || word.trim().length < 2) return false;
    const pattern = new RegExp(`(^|[^\\p{L}\\p{N}])${escape(word.trim())}($|[^\\p{L}\\p{N}])`, 'iu');
    return pattern.test(subject || '');
  }));
  return hits.length === 1 ? String(hits[0].id) : '';
}

function toggleId(list, id, on) {
  const key = String(id);
  const rest = list.map(String);
  if (on) return rest.includes(key) ? rest : [...rest, key];
  return rest.filter((item) => item !== key);
}

/* Сборка пакета. Замок: пока одна сборка идёт, вторая не стартует -- её
 * подметание удалило бы временный файл первой. Кнопка блокируется до первого
 * ответа сервера (setBusy вызывается синхронно). Перед сборкой -- plan: если в
 * папке уже лежат файлы, вопрос называет документы, которые будут заменены. */
function createBuilder({ api: call, confirmReplace, setBusy }) {
  let running = false;

  async function attempt(request) {
    const { tenderId, documents, experienceIds } = request;
    const outputDir = String(request.outputDir || '').trim();
    if (!documents.length) {
      return { status: 'failed', error: 'Отметьте хотя бы один документ пакета.' };
    }
    if (!outputDir) {
      return { status: 'failed', error: 'Укажите папку, в которую собирать пакет.' };
    }
    const base = `/api/tenders/${tenderId}`;
    const body = { documents, output_dir: outputDir, experience_ids: experienceIds };
    const plan = await call('POST', `${base}/plan`, body);
    const existing = plan.files.filter((file) => file.exists);
    let overwrite = false;
    if (existing.length) {
      if (!(await confirmReplace(existing))) return { status: 'cancelled' };
      overwrite = true;
    }
    const result = await call('POST', `${base}/build`, { ...body, overwrite });
    let note = '';
    try {
      await call('POST', '/api/tenders', { id: tenderId, output_dir: outputDir, experience_ids: experienceIds });
    } catch (error) {
      note = `Пакет собран, но папку и отбор опыта запомнить в тендере не удалось: ${error.message}`;
    }
    return { status: 'built', result, note };
  }

  async function run(request) {
    if (running) return { status: 'busy' };
    running = true;
    setBusy(true);
    try {
      return await attempt(request);
    } catch (error) {
      if (error instanceof ApiFailure) return { status: 'failed', error: error.message };
      console.error(error);
      return { status: 'failed', error: 'Сборка прервана из-за ошибки в самой странице. Обновите страницу и повторите.' };
    } finally {
      running = false;
      setBusy(false);
    }
  }

  return { run, get busy() { return running; } };
}

if (typeof module === 'object' && module.exports) {
  module.exports = {
    ApiFailure, api, splitWarning, recordForWrite, writeRecord, parsePasted, lotPayload,
    differenceKind, documentChoices, inferServiceId, toggleId, createBuilder,
  };
}

/* ======================================================================== */
/* 2. Страница                                                                */
/* ======================================================================== */

if (typeof document !== 'undefined' && document.getElementById('screen')) {
  (function page() {
    /* ---- мелочи ---- */
    const $ = (selector, root = document) => root.querySelector(selector);
    const str = (value) => (value === null || value === undefined ? '' : String(value));
    const DASH = '—';

    function h(tag, attrs, ...children) {
      const el = document.createElement(tag);
      for (const [key, value] of Object.entries(attrs || {})) {
        if (value === undefined || value === null || value === false) continue;
        if (key === 'class') el.className = value;
        else if (key === 'text') el.textContent = value;
        else if (key === 'value') el.value = value;
        else if (key === 'checked') el.checked = Boolean(value);
        else if (key === 'disabled') el.disabled = Boolean(value);
        else if (key.startsWith('on')) el.addEventListener(key.slice(2), value);
        else el.setAttribute(key, value === true ? '' : value);
      }
      append(el, children);
      return el;
    }
    function append(el, children) {
      for (const child of children.flat(Infinity)) {
        if (child === null || child === undefined || child === false) continue;
        el.append(child instanceof Node ? child : document.createTextNode(String(child)));
      }
    }

    /* Сообщения. Ошибка -- всегда русская фраза сервера, не код. */
    function notice() {
      const el = h('div');
      return {
        el,
        clear() { el.replaceChildren(); },
        show(kind, title, lines) {
          const box = h('div', { class: `notice ${kind}`, role: kind === 'error' ? 'alert' : 'status' },
            h('p', { class: 'title', text: title }));
          if (lines && lines.length) box.append(h('ul', {}, lines.map((line) => h('li', { text: line }))));
          el.replaceChildren(box);
          if (kind === 'error') box.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        },
        error(error) { this.show('error', error instanceof ApiFailure ? error.message : failureText(error)); },
      };
    }
    function failureText(error) {
      console.error(error);
      return 'Что-то пошло не так в самой странице. Обновите её (F5) и повторите; если повторится, '
        + 'покажите разработчику чёрное окно программы.';
    }

    function ask({ title, message, items, ok = 'Да', cancel = 'Отмена' }) {
      return new Promise((resolve) => {
        const dialog = $('#dialog');
        $('#dialog-title').textContent = title;
        $('#dialog-message').textContent = message || '';
        $('#dialog-items').replaceChildren(...(items || []).map((item) => h('li', { text: item })));
        $('#dialog-items').hidden = !(items && items.length);
        $('#dialog-ok').textContent = ok;
        $('#dialog-cancel').textContent = cancel;
        dialog.returnValue = '';
        dialog.addEventListener('close', () => resolve(dialog.returnValue === 'ok'), { once: true });
        dialog.showModal();
      });
    }

    /* ---- текущий тендер и переходы ---- */
    const REMEMBER_KEY = 'tender-ui.current';
    function rememberedTender() {
      try { return localStorage.getItem(REMEMBER_KEY) || ''; } catch (error) { return ''; }
    }
    function rememberTender(id) {
      try { localStorage.setItem(REMEMBER_KEY, String(id)); } catch (error) { /* без запоминания тоже работает */ }
      $('#nav-tender').setAttribute('href', `#/tender/${id}`);
      $('#nav-build').setAttribute('href', `#/build/${id}`);
    }
    function forgetTender() {
      try { localStorage.removeItem(REMEMBER_KEY); } catch (error) { /* не страшно */ }
      $('#nav-tender').setAttribute('href', '#/tender');
      $('#nav-build').setAttribute('href', '#/build');
      $('#current-tender').textContent = '';
    }
    function describeTender(tender, orgs) {
      const customer = orgs.find((org) => org.id === tender.customer_id);
      const head = tender.purchase_number ? `№ ${tender.purchase_number}` : 'без номера';
      return [head, tender.subject || 'без предмета', customer ? customer.short_name || customer.full_name : '']
        .filter(Boolean).join(' — ');
    }

    let active = { token: 0, dirty: [] };
    let renderToken = 0;
    let shownHash = location.hash;
    let revertingHash = false;

    function anyUnsaved() { return active.dirty.some((check) => check()); }

    function go(hash) { location.hash = hash; }

    window.addEventListener('beforeunload', (event) => {
      if (anyUnsaved()) { event.preventDefault(); event.returnValue = ''; }
    });
    window.addEventListener('hashchange', () => {
      if (revertingHash) { revertingHash = false; return; }
      if (anyUnsaved() && !window.confirm('На этом экране есть несохранённые изменения. Уйти без сохранения?')) {
        revertingHash = true;
        location.hash = shownHash;
        return;
      }
      shownHash = location.hash;
      render();
    });

    async function render() {
      renderToken += 1;
      const mine = { token: renderToken, dirty: [] };
      active = mine;
      const parts = location.hash.replace(/^#\/?/, '').split('/');
      const screen = parts[0] || 'refs';
      for (const [name, id] of [['refs', 'nav-refs'], ['tender', 'nav-tender'], ['build', 'nav-build']]) {
        const link = $(`#${id}`);
        if (name === screen) link.setAttribute('aria-current', 'page'); else link.removeAttribute('aria-current');
      }
      const target = $('#screen');
      target.replaceChildren(h('p', { class: 'loading', text: 'Загрузка…' }));
      const root = h('div');
      try {
        const handler = { refs: refsScreen, tender: tenderScreen, build: buildScreen }[screen];
        if (!handler) { go('#/refs'); return; }
        await handler(root, parts.slice(1), mine);
      } catch (error) {
        const box = notice();
        box.show('error', error instanceof ApiFailure ? error.message : failureText(error));
        root.replaceChildren(box.el);
      }
      if (mine.token === renderToken) target.replaceChildren(root);
    }

    /* ---- формы: поля, значения, сбор правок ---- */
    const KINDS = {
      text: { inputmode: undefined }, textarea: {}, date: {}, select: {},
      money: { inputmode: 'decimal', num: true }, number: { inputmode: 'decimal', num: true },
      int: { inputmode: 'numeric', num: true }, percent: { inputmode: 'decimal', num: true },
    };
    /* Ключ значения в записи: у ставки это <поле>_percent, остальное -- само поле. */
    const keyOf = (spec) => (spec.kind === 'percent' ? `${spec.key}_percent` : spec.key);

    function makeControl(spec, record, lookups) {
      const kind = KINDS[spec.kind] || KINDS.text;
      const value = str(record[keyOf(spec)]);
      if (spec.kind === 'textarea') return h('textarea', { value, rows: spec.rows || 3 });
      if (spec.kind === 'select' || spec.ref) {
        const options = spec.ref ? [['', DASH + ' не выбрано']].concat(lookups[spec.ref] || []) : spec.options;
        const select = h('select', {}, options.map(([v, label]) => h('option', { value: str(v), text: label })));
        select.value = value;
        if (select.value !== value && value) {         // значения нет в списке -- не теряем его молча
          select.append(h('option', { value, text: `${value} (нет в списке)` }));
          select.value = value;
        }
        return select;
      }
      return h('input', { type: spec.kind === 'date' ? 'date' : 'text', value,
        inputmode: kind.inputmode, class: kind.num ? 'num' : '', autocomplete: 'off',
        list: spec.list, placeholder: spec.placeholder });
    }

    /* fields: [{group}, {key, label, kind, wide, hint, ref, options}, {custom(ctx)}]
     * Возвращает {el, inputs, collect()}; collect даёт правки, а не всю запись --
     * запись целиком собирает recordForWrite. */
    function buildForm(fields, record, lookups, onDirty) {
      const inputs = {};
      const collectors = [];
      const root = h('div');
      let grid = null;
      const ctx = { record, inputs, lookups, onDirty };
      for (const spec of fields) {
        if (spec.group) {
          grid = h('div', { class: 'form-grid' });
          root.append(h('fieldset', {}, h('legend', { text: spec.group }), grid));
          continue;
        }
        if (!grid) { grid = h('div', { class: 'form-grid' }); root.append(grid); }
        if (spec.custom) {
          const built = spec.custom(ctx);
          grid.append(built.el);
          if (built.collect) collectors.push(built.collect);
          continue;
        }
        const control = makeControl(spec, record, lookups);
        inputs[keyOf(spec)] = control;
        grid.append(h('label', { class: `field${spec.wide ? ' wide' : ''}` },
          h('span', { class: 'lbl', text: spec.label }), control,
          spec.hint ? h('span', { class: 'hint', text: spec.hint }) : null));
      }
      root.addEventListener('input', () => onDirty && onDirty());
      root.addEventListener('change', () => onDirty && onDirty());
      return {
        el: root, inputs,
        collect() {
          const edits = {};
          for (const [key, control] of Object.entries(inputs)) edits[key] = control.value;
          collectors.forEach((collector) => collector(edits));
          return edits;
        },
        mark(error) {
          Object.values(inputs).forEach((control) => control.classList.remove('invalid'));
          const control = error && error.field ? inputs[error.field] : null;
          if (control) { control.classList.add('invalid'); control.focus(); }
        },
      };
    }

    /* Редактор пар «название -- значение» (нормы часов, ответы по условиям сделки). */
    function pairsEditor({ key, valueKey, label, left, right, record, pairs, hint, addText }) {
      const body = h('tbody');
      const el = h('div', { class: 'wide' });
      const dirtyEvent = () => el.dispatchEvent(new Event('input', { bubbles: true }));
      function addRow(a, b) {
        const first = h('input', { type: 'text', value: str(a), 'aria-label': left, autocomplete: 'off' });
        const second = h('input', { type: 'text', value: str(b), 'aria-label': right, class: valueKey ? 'num' : '',
          inputmode: valueKey ? 'decimal' : undefined, autocomplete: 'off' });
        const tr = h('tr', {}, h('td', {}, first), h('td', {}, second),
          h('td', {}, h('button', { type: 'button', class: 'small', text: 'Убрать',
            onclick: () => { tr.remove(); dirtyEvent(); } })));
        tr.pairInputs = [first, second];
        body.append(tr);
      }
      (pairs || Object.entries(record[key] || {})).forEach(([a, b]) => addRow(a, b));
      el.append(h('div', { class: 'lbl', style: 'font-weight:600;font-size:.9rem;margin-bottom:4px', text: label }),
        h('div', { class: 'table-wrap' }, h('table', {}, h('thead', {}, h('tr', {},
          h('th', { text: left }), h('th', { text: right }), h('th'))), body)),
        h('div', { class: 'actions' }, h('button', { type: 'button', class: 'small', text: addText || 'Добавить строку',
          onclick: () => { addRow('', ''); dirtyEvent(); } })),
        hint ? h('p', { class: 'hint', text: hint }) : null);
      return {
        el,
        collect(edits) {
          const dict = {};
          for (const tr of body.children) {
            const [a, b] = tr.pairInputs.map((input) => input.value.trim());
            if (a === '' && b === '') continue;
            dict[a] = b;
          }
          edits[key] = dict;
        },
      };
    }

    /* ---- справочные списки для выпадающих полей ---- */
    const orgLabel = (org) => org.short_name || org.full_name || org.inn || `организация ${org.id}`;
    async function loadLookups(names) {
      const out = {};
      const jobs = [];
      if (names.includes('customers') || names.includes('participants')) {
        jobs.push(api('GET', '/api/organizations').then((orgs) => {
          out.orgs = orgs;
          out.customers = orgs.filter((o) => o.role === 'customer').map((o) => [o.id, orgLabel(o)]);
          out.participants = orgs.filter((o) => o.role === 'participant').map((o) => [o.id, orgLabel(o)]);
        }));
      }
      if (names.includes('services')) {
        jobs.push(api('GET', '/api/services').then((services) => {
          out.serviceRecords = services;
          out.services = services.map((s) => [s.id, [s.code, s.name].filter(Boolean).join(' — ')]);
        }));
      }
      await Promise.all(jobs);
      return out;
    }

    /* ================================================================== */
    /* Экран «Справочники»                                                  */
    /* ================================================================== */

    const MSP = [['', 'не указан'], ['micro', 'микропредприятие'], ['small', 'малое предприятие'],
      ['medium', 'среднее предприятие']];
    const ROLES = [['customer', 'Заказчик'], ['participant', 'Наше юрлицо (участник)']];
    const ASSET_KINDS = [['premises', 'Помещение'], ['equipment', 'Оборудование'],
      ['software', 'Программное обеспечение'], ['', 'Не указан']];

    /* Блок «Загрузить по ИНН»: заполняет поля карточки из реестра и показывает,
     * чего не хватило. Ничего не сохраняет -- карточку сохраняет человек. */
    function innLoader(ctx) {
      const urlInput = h('input', { type: 'text', autocomplete: 'off',
        placeholder: 'https://www.rusprofile.ru/id/…' });
      const status = notice();
      const button = h('button', { type: 'button', class: 'primary', text: 'Загрузить по ИНН' });
      button.addEventListener('click', async () => {
        const inn = (ctx.inputs.inn.value || '').replace(/\s+/g, '');
        if (!/^(\d{10}|\d{12})$/.test(inn)) {
          status.show('error', 'Введите ИНН: 10 цифр у организации или 12 у предпринимателя.');
          ctx.inputs.inn.focus();
          return;
        }
        const link = urlInput.value.trim();
        button.disabled = true;
        button.textContent = 'Запрашиваю реестр…';
        status.show('warning', 'Запрашиваю данные в реестре ФНС (и на rusprofile, если указана ссылка). Это может занять несколько секунд.');
        try {
          const query = link ? `?rusprofile_url=${encodeURIComponent(link)}` : '';
          const data = await api('GET', `/api/inn/${inn}${query}`);
          for (const [key, control] of Object.entries(ctx.inputs)) {
            const fetched = data.organization[key];
            if (key === 'role' || key === 'vat_rate_percent') continue;
            if (typeof fetched === 'string' && fetched.trim() !== '') control.value = fetched;
          }
          ctx.onDirty();
          status.show(data.warnings.length ? 'warning' : 'ok',
            'Поля заполнены из реестра. Проверьте их и нажмите «Сохранить» — пока карточка не сохранена, в справочнике ничего не изменилось.',
            data.warnings);
        } catch (error) {
          status.error(error);
        } finally {
          button.disabled = false;
          button.textContent = 'Загрузить по ИНН';
        }
      });
      return {
        el: h('div', { class: 'wide' },
          h('div', { class: 'inline' },
            h('div', {}, h('label', { class: 'field' },
              h('span', { class: 'lbl', text: 'Ссылка на страницу rusprofile (необязательно)' }), urlInput,
              h('span', { class: 'hint', text: 'У rusprofile нет поиска по ИНН: найдите компанию на сайте и вставьте адрес её страницы — оттуда возьмутся адрес, коды и статус МСП.' }))),
            button),
          status.el),
      };
    }

    const orgFields = [
      { group: 'Основное' },
      { custom: innLoader },
      { key: 'role', label: 'Роль', kind: 'select', options: ROLES },
      { key: 'inn', label: 'ИНН', kind: 'text' },
      { key: 'kpp', label: 'КПП', kind: 'text' },
      { key: 'ogrn', label: 'ОГРН', kind: 'text' },
      { key: 'short_name', label: 'Краткое наименование', kind: 'text', wide: true },
      { key: 'full_name', label: 'Полное наименование', kind: 'text', wide: true },
      { key: 'legal_address', label: 'Юридический адрес', kind: 'text', wide: true },
      { key: 'postal_address', label: 'Почтовый адрес', kind: 'text', wide: true },
      { group: 'Контакты (в реестрах их нет — вводятся один раз вручную)' },
      { key: 'phones', label: 'Телефоны', kind: 'text' },
      { key: 'email', label: 'E-mail', kind: 'text' },
      { key: 'website', label: 'Сайт', kind: 'text' },
      { group: 'Банк (в реестрах его нет — вводится один раз вручную)' },
      { key: 'bank_name', label: 'Банк', kind: 'text', wide: true },
      { key: 'bank_bik', label: 'БИК', kind: 'text' },
      { key: 'bank_account', label: 'Расчётный счёт', kind: 'text' },
      { key: 'bank_corr_account', label: 'Корреспондентский счёт', kind: 'text' },
      { group: 'Руководитель' },
      { key: 'head_name', label: 'ФИО руководителя', kind: 'text' },
      { key: 'head_position', label: 'Должность руководителя', kind: 'text' },
      { key: 'head_basis', label: 'Основание полномочий', kind: 'text', hint: 'Например: Устава' },
      { group: 'Коды и налоги' },
      { key: 'okpo', label: 'ОКПО', kind: 'text' },
      { key: 'oktmo', label: 'ОКТМО', kind: 'text' },
      { key: 'okved', label: 'ОКВЭД', kind: 'text' },
      { key: 'vat_rate', label: 'Ставка НДС, %', kind: 'percent', hint: 'Пусто — берётся из настроек организации.' },
      { key: 'msp_status', label: 'Статус МСП', kind: 'select', options: MSP },
    ];

    const REF_TABS = [
      { key: 'organizations', title: 'Организации', path: '/api/organizations', one: 'организацию',
        fields: orgFields, filter: ROLES,
        label: (r) => orgLabel(r),
        sub: (r) => [r.role === 'customer' ? 'Заказчик' : 'Наше юрлицо', r.inn && `ИНН ${r.inn}`].filter(Boolean).join(' · '),
        blank: (filter) => ({ role: filter || 'customer' }) },
      { key: 'employees', title: 'Сотрудники', path: '/api/employees', one: 'сотрудника',
        fields: [
          { key: 'full_name', label: 'ФИО', kind: 'text', wide: true },
          { key: 'position', label: 'Должность', kind: 'text', hint: 'По ней ставка подставляется в трудозатраты.' },
          { key: 'hourly_rate', label: 'Ставка, руб/ч', kind: 'money' },
          { key: 'education', label: 'Образование', kind: 'text', wide: true },
          { key: 'graduation_year', label: 'Год окончания', kind: 'int' },
          { key: 'specialty', label: 'Специальность', kind: 'text' },
          { key: 'qualification', label: 'Квалификация и сертификаты', kind: 'textarea', wide: true },
          { key: 'experience_years', label: 'Стаж, лет', kind: 'int' },
        ],
        label: (r) => r.full_name || 'без имени',
        sub: (r) => [r.position, r.hourly_rate && `${r.hourly_rate} руб/ч`].filter(Boolean).join(' · ') },
      { key: 'assets', title: 'Оборудование и ПО', path: '/api/assets', one: 'ресурс',
        fields: [
          { key: 'name', label: 'Наименование', kind: 'text', wide: true },
          { key: 'kind', label: 'Вид', kind: 'select', options: ASSET_KINDS },
          { key: 'maker_model', label: 'Изготовитель и марка', kind: 'text' },
          { key: 'quantity', label: 'Количество', kind: 'int' },
          { key: 'year', label: 'Год', kind: 'int' },
          { key: 'location', label: 'Местонахождение', kind: 'text' },
          { key: 'ownership_basis', label: 'Основание владения', kind: 'text' },
          { key: 'ownership_doc', label: 'Документ о владении', kind: 'text' },
          { key: 'purpose', label: 'Назначение', kind: 'text' },
          { key: 'condition', label: 'Состояние', kind: 'text' },
        ],
        label: (r) => r.name || 'без названия',
        sub: (r) => [(ASSET_KINDS.find(([v]) => v === r.kind) || [])[1], r.maker_model].filter(Boolean).join(' · ') },
      { key: 'experience', title: 'Опыт', path: '/api/experience', one: 'запись об опыте', lookups: ['services'],
        fields: [
          { key: 'customer_name', label: 'Заказчик', kind: 'text', wide: true },
          { key: 'contract_number', label: 'Номер контракта', kind: 'text' },
          { key: 'contract_date', label: 'Дата контракта', kind: 'text', hint: 'Как печатать в справке, например 24.10.2025.' },
          { key: 'subject', label: 'Предмет', kind: 'textarea', wide: true },
          { key: 'amount', label: 'Сумма, руб.', kind: 'money' },
          { key: 'term', label: 'Срок', kind: 'text' },
          { key: 'status', label: 'Статус исполнения', kind: 'text' },
          { key: 'service_id', label: 'Услуга', kind: 'select', ref: 'services', hint: 'По ней отбирается опыт при сборке.' },
          { key: 'contract_scan', label: 'Скан контракта (путь)', kind: 'text' },
          { key: 'act_scan', label: 'Скан акта (путь)', kind: 'text' },
        ],
        label: (r) => r.customer_name || 'без заказчика',
        sub: (r) => [r.contract_number && `№ ${r.contract_number}`, r.contract_date, r.amount && `${r.amount} руб.`].filter(Boolean).join(' · ') },
      { key: 'services', title: 'Услуги', path: '/api/services', one: 'услугу',
        fields: [
          { key: 'code', label: 'Код', kind: 'text', hint: 'КЭР, СЗЗ, НДВ…' },
          { key: 'name', label: 'Название', kind: 'text' },
          { key: 'stage_wording', label: 'Формулировки этапов', kind: 'textarea', wide: true, rows: 5 },
          { custom: (ctx) => pairsEditor({ key: 'norms', valueKey: true, label: 'Нормы часов по должностям',
            left: 'Должность', right: 'Часы', record: ctx.record,
            hint: 'Часы на одного специалиста. Должность должна совпадать с должностью в справочнике сотрудников — тогда ставка подставится сама.',
            addText: 'Добавить должность' }) },
        ],
        label: (r) => [r.code, r.name].filter(Boolean).join(' — ') || 'без названия',
        sub: () => '' },
      { key: 'profiles', title: 'Профили заказчиков', path: '/api/profiles', one: 'профиль', lookups: ['customers'],
        fields: [
          { key: 'customer_id', label: 'Заказчик', kind: 'select', ref: 'customers', wide: true,
            hint: 'У заказчика один профиль.' },
          { key: 'documents_text', label: 'Состав пакета (по одному документу в строке)', kind: 'textarea', wide: true, rows: 6,
            hint: 'Названия документов — как в разделе «Описание документов» ниже.' },
          { custom: (ctx) => docsEditor(ctx) },
          { key: 'declaration_text', label: 'Текст декларации', kind: 'textarea', wide: true, rows: 8 },
        ],
        prepare: (record) => ({ ...record, documents_text: (record.documents || []).join('\n') }),
        finish: (edits) => {
          edits.documents = String(edits.documents_text || '').split(/\r\n|\r|\n/).map((s) => s.trim()).filter(Boolean);
          delete edits.documents_text;
        },
        label: (r, look) => (look.customers.find(([id]) => id === r.customer_id) || [0, 'заказчик не выбран'])[1],
        sub: (r) => (r.documents || []).join(', ') },
    ];

    /* Описание документов профиля: шаблон, карта ячеек, раскладка, префикс. */
    function docsEditor(ctx) {
      const body = h('tbody');
      const el = h('div', { class: 'wide' });
      const dirtyEvent = () => el.dispatchEvent(new Event('input', { bubbles: true }));
      const FIELDS = [['template', 'Шаблон'], ['cellmap', 'Карта ячеек'], ['layout', 'Раскладка'], ['prefix', 'Префикс']];
      function addRow(name, spec) {
        const nameInput = h('input', { type: 'text', value: str(name), 'aria-label': 'Документ', autocomplete: 'off' });
        const inputs = FIELDS.map(([key, label]) => h('input', { type: 'text', value: str((spec || {})[key]),
          'aria-label': label, autocomplete: 'off' }));
        const tr = h('tr', {}, h('td', {}, nameInput), inputs.map((input) => h('td', {}, input)),
          h('td', {}, h('button', { type: 'button', class: 'small', text: 'Убрать', onclick: () => { tr.remove(); dirtyEvent(); } })));
        tr.docInputs = { name: nameInput, fields: inputs };
        body.append(tr);
      }
      Object.entries(ctx.record.docs || {}).forEach(([name, spec]) => addRow(name, spec));
      el.append(h('div', { class: 'lbl', style: 'font-weight:600;font-size:.9rem;margin-bottom:4px', text: 'Описание документов' }),
        h('div', { class: 'table-wrap' }, h('table', {}, h('thead', {}, h('tr', {},
          h('th', { text: 'Документ' }), FIELDS.map(([, label]) => h('th', { text: label })), h('th'))), body)),
        h('div', { class: 'actions' }, h('button', { type: 'button', class: 'small', text: 'Добавить документ',
          onclick: () => { addRow('', {}); dirtyEvent(); } })),
        h('p', { class: 'hint', text: 'Любое поле можно оставить пустым: пустой шаблон — файл «<документ>.docx» в папке шаблонов; для книги Excel нужна карта ячеек.' }));
      return {
        el,
        collect(edits) {
          const docs = {};
          for (const tr of body.children) {
            const name = tr.docInputs.name.value.trim();
            if (!name) continue;
            docs[name] = Object.fromEntries(FIELDS.map(([key], i) => [key, tr.docInputs.fields[i].value]));
          }
          edits.docs = docs;
        },
      };
    }

    const SETTINGS_FIELDS = [
      { key: 'insurance_rate', label: 'Страховые взносы, %', kind: 'percent' },
      { key: 'injury_rate', label: 'Травматизм, %', kind: 'percent' },
      { key: 'overhead_rate', label: 'Накладные, % от оплаты труда со страховыми', kind: 'percent' },
      { key: 'profit_rate', label: 'Прибыль, %', kind: 'percent', hint: 'От оплаты труда, субподряда и накладных.' },
      { key: 'vat_rate', label: 'НДС по умолчанию, %', kind: 'percent', hint: 'Можно переопределить в организации и в тендере.' },
      { key: 'total_rounding_digits', label: 'Округление итога без НДС, знаков после запятой', kind: 'int',
        hint: '0 — до целых рублей (как в книге заказчика), 2 — до копеек, −3 — до тысяч. Из-за округления итог может отличаться от суммы статей — до 50 копеек.' },
    ];

    async function refsScreen(root, args, ctx) {
      const tabKey = args[0] || 'organizations';
      const tab = REF_TABS.find((t) => t.key === tabKey);
      const tabs = h('nav', { class: 'tabs', 'aria-label': 'Справочники' },
        REF_TABS.map((t) => h('a', { href: `#/refs/${t.key}`, text: t.title, 'aria-current': t.key === tabKey ? 'page' : undefined })),
        h('a', { href: '#/refs/settings', text: 'Настройки', 'aria-current': tabKey === 'settings' ? 'page' : undefined }));
      root.append(h('h1', { text: 'Справочники' }), tabs);
      if (tabKey === 'settings') { await settingsView(root, ctx); return; }
      if (!tab) { root.append(h('p', { class: 'empty', text: 'Такого справочника нет.' })); return; }
      await crudView(root, tab, ctx);
    }

    async function crudView(root, cfg, ctx) {
      const [items, lookups] = await Promise.all([api('GET', cfg.path), loadLookups(cfg.lookups || [])]);
      let list = items;
      let selected = null;          // { record, isNew }
      let formDirty = false;
      let filterRole = '';
      let search = '';
      ctx.dirty.push(() => formDirty);

      const msg = notice();
      const listBox = h('ul', { class: 'pick-list' });
      const formBox = h('div');
      const searchInput = h('input', { type: 'search', placeholder: 'Поиск по списку', 'aria-label': 'Поиск по списку', autocomplete: 'off' });
      const roleSelect = cfg.filter ? h('select', { 'aria-label': 'Показать' },
        [['', 'Все организации']].concat(cfg.filter.map(([v, l]) => [v, v === 'customer' ? 'Только заказчики' : 'Только наши юрлица']))
          .map(([v, l]) => h('option', { value: v, text: l }))) : null;

      function drawList() {
        const needle = search.trim().toLowerCase();
        const shown = list.filter((r) => (!filterRole || r.role === filterRole)
          && (!needle || `${cfg.label(r, lookups)} ${cfg.sub(r, lookups)}`.toLowerCase().includes(needle)));
        listBox.replaceChildren(...shown.map((r) => h('li', {}, h('button', { type: 'button',
          'aria-current': selected && !selected.isNew && selected.record.id === r.id ? 'true' : undefined,
          onclick: () => choose(r, false) },
        cfg.label(r, lookups), h('span', { class: 'sub', text: cfg.sub(r, lookups) })))));
        if (!shown.length) listBox.replaceChildren(h('li', { class: 'empty', style: 'padding:10px 12px', text: list.length ? 'Ничего не найдено.' : 'Справочник пока пуст.' }));
      }

      async function choose(record, isNew) {
        if (formDirty && !(await ask({ title: 'Несохранённые изменения', message: 'В открытой карточке есть правки, которые не сохранены. Закрыть её без сохранения?', ok: 'Закрыть без сохранения' }))) return;
        selected = { record, isNew };
        formDirty = false;
        msg.clear();
        drawList();
        drawForm();
      }

      function drawForm() {
        if (!selected) {
          formBox.replaceChildren(h('div', { class: 'card' }, h('p', { class: 'empty', text: `Выберите запись слева или нажмите «Добавить», чтобы завести ${cfg.one}.` })));
          return;
        }
        const record = cfg.prepare ? cfg.prepare(selected.record) : selected.record;
        const form = buildForm(cfg.fields, record, lookups, () => { formDirty = true; dirtyTag.hidden = false; });
        const dirtyTag = h('span', { class: 'state-tag', text: 'есть несохранённые изменения', hidden: true });
        const save = h('button', { type: 'button', class: 'primary', text: 'Сохранить' });
        const remove = h('button', { type: 'button', class: 'danger', text: 'Удалить', hidden: selected.isNew });
        save.addEventListener('click', async () => {
          save.disabled = true;
          msg.clear();
          try {
            const edits = form.collect();
            if (cfg.finish) cfg.finish(edits);
            const { record: saved, warning } = await writeRecord(cfg.path, recordForWrite(selected.record, edits));
            formDirty = false;
            list = await api('GET', cfg.path);
            selected = { record: list.find((r) => r.id === saved.id) || saved, isNew: false };
            drawList();
            drawForm();
            if (warning) msg.show('warning', warning); else msg.show('ok', 'Сохранено.');
          } catch (error) {
            form.mark(error);
            msg.error(error);
          } finally {
            save.disabled = false;
          }
        });
        remove.addEventListener('click', async () => {
          const name = cfg.label(selected.record, lookups);
          if (!(await ask({ title: 'Удалить запись?', message: `«${name}» будет удалена из справочника. Это нельзя отменить.`, ok: 'Удалить' }))) return;
          try {
            await api('DELETE', `${cfg.path}/${selected.record.id}`);
            list = await api('GET', cfg.path);
            selected = null;
            formDirty = false;
            drawList();
            drawForm();
            msg.show('ok', 'Запись удалена.');
          } catch (error) {
            msg.error(error);
          }
        });
        const title = selected.isNew ? `Новая запись: ${cfg.one}` : cfg.label(selected.record, lookups);
        formBox.replaceChildren(h('div', { class: 'card' }, h('h2', {}, title, ' ', dirtyTag), form.el,
          h('div', { class: 'actions' }, save, remove)));
      }

      searchInput.addEventListener('input', () => { search = searchInput.value; drawList(); });
      if (roleSelect) roleSelect.addEventListener('change', () => { filterRole = roleSelect.value; drawList(); });
      const addButton = h('button', { type: 'button', class: 'primary', text: `Добавить ${cfg.one}`,
        onclick: () => choose(cfg.blank ? cfg.blank(filterRole) : {}, true) });
      drawList();
      drawForm();
      root.append(msg.el, h('div', { class: 'refs-layout' },
        h('div', { class: 'card' }, h('h2', { text: cfg.title }), roleSelect, searchInput, listBox, h('div', { class: 'actions' }, addButton)),
        formBox));
    }

    async function settingsView(root, ctx) {
      const record = await api('GET', '/api/settings');
      const msg = notice();
      let dirty = false;
      ctx.dirty.push(() => dirty);
      const tag = h('span', { class: 'state-tag', text: 'есть несохранённые изменения', hidden: true });
      const form = buildForm(SETTINGS_FIELDS, record, {}, () => { dirty = true; tag.hidden = false; });
      const save = h('button', { type: 'button', class: 'primary', text: 'Сохранить настройки' });
      save.addEventListener('click', async () => {
        save.disabled = true;
        try {
          const { record: saved } = await writeRecord('/api/settings', recordForWrite(record, form.collect()));
          dirty = false;
          tag.hidden = true;
          Object.assign(record, saved);
          msg.show('ok', 'Настройки сохранены. Они применяются к новым расчётам.');
        } catch (error) {
          form.mark(error);
          msg.error(error);
        } finally {
          save.disabled = false;
        }
      });
      root.append(msg.el, h('div', { class: 'card' }, h('h2', {}, 'Настройки калькуляции ', tag),
        h('p', { class: 'hint', text: 'Проценты по умолчанию для всех тендеров. В организации и в самом тендере ставку НДС можно переопределить.' }),
        form.el, h('div', { class: 'actions' }, save)));
    }

    /* ================================================================== */
    /* Общее для «Карточки тендера» и «Сборки»: выбор тендера                */
    /* ================================================================== */

    function tenderPicker({ tenders, orgs, currentId, screen, withNew }) {
      const select = h('select', { 'aria-label': 'Тендер' },
        h('option', { value: '', text: tenders.length ? 'Выберите тендер…' : 'Тендеров пока нет' }),
        tenders.map((t) => h('option', { value: String(t.id), text: describeTender(t, orgs) })));
      select.value = currentId ? String(currentId) : '';
      select.addEventListener('change', () => { if (select.value) go(`#/${screen}/${select.value}`); });
      return h('div', { class: 'card inline' },
        h('div', {}, h('label', { class: 'field' }, h('span', { class: 'lbl', text: 'Тендер' }), select)),
        withNew ? h('button', { type: 'button', text: 'Новый тендер', onclick: () => go('#/tender/new') }) : null);
    }

    async function chooseTender(root, args, ctx, screen) {
      const [tenders, orgs] = await Promise.all([api('GET', '/api/tenders'), api('GET', '/api/organizations')]);
      let idArg = args[0] || '';
      if (!idArg) {
        const remembered = rememberedTender();
        if (remembered && tenders.some((t) => String(t.id) === remembered)) { go(`#/${screen}/${remembered}`); return { redirected: true }; }
      }
      const tender = idArg && idArg !== 'new' ? tenders.find((t) => String(t.id) === idArg) : null;
      if (idArg && idArg !== 'new' && !tender) {
        root.append(h('h1', { text: screen === 'tender' ? 'Карточка тендера' : 'Сборка' }),
          tenderPicker({ tenders, orgs, currentId: '', screen, withNew: screen === 'tender' }),
          h('div', { class: 'notice error' }, h('p', { class: 'title', text: 'Такого тендера в программе нет. Выберите другой из списка.' })));
        return { redirected: true };
      }
      if (tender) {
        rememberTender(tender.id);
        $('#current-tender').textContent = describeTender(tender, orgs);
      } else if (idArg !== 'new') {
        $('#current-tender').textContent = '';
      }
      return { tenders, orgs, tender, isNew: idArg === 'new' };
    }

    /* ================================================================== */
    /* Экран «Карточка тендера»                                             */
    /* ================================================================== */

    const TENDER_FIELDS = (lookups) => [
      { group: 'Заказчик и участник' },
      { key: 'customer_id', label: 'Заказчик', kind: 'select', ref: 'customers',
        hint: 'Нет в справочнике? Заведите его по ИНН на экране «Справочники».' },
      { key: 'participant_id', label: 'Участник (наше юрлицо)', kind: 'select', ref: 'participants' },
      { group: 'Закупка' },
      { key: 'purchase_number', label: 'Номер закупки', kind: 'text' },
      { key: 'purchase_date', label: 'Дата закупки', kind: 'date' },
      { key: 'method', label: 'Способ закупки', kind: 'text' },
      { key: 'platform', label: 'Площадка', kind: 'text' },
      { key: 'subject', label: 'Предмет закупки', kind: 'textarea', wide: true },
      { group: 'Объект' },
      { key: 'object_name', label: 'Объект', kind: 'text', wide: true },
      { key: 'object_address', label: 'Адрес объекта', kind: 'text', wide: true },
      { group: 'Цена' },
      { key: 'price_wo_vat', label: 'Цена без НДС, руб.', kind: 'money',
        hint: 'Пусто — 0,00. Итог калькулятора можно перенести сюда кнопкой в калькуляторе.' },
      { key: 'vat_rate', label: 'Ставка НДС, %', kind: 'percent',
        hint: 'Пусто — берётся из организации, затем из настроек. Действующая ставка видна в калькуляторе.' },
      { group: 'Сроки и документы' },
      { key: 'deadline', label: 'Сроки исполнения', kind: 'textarea', rows: 2 },
      { key: 'payment_terms', label: 'Условия оплаты', kind: 'textarea', rows: 2 },
      { key: 'doc_date', label: 'Дата документов', kind: 'date' },
      { custom: (ctx) => pairsEditor({ key: 'deal_answers', label: 'Ответы по условиям сделки (формы вроде Роснефти)',
        left: 'Условие', right: 'Ответ', record: ctx.record,
        hint: 'Для условий, которых здесь нет, в форме остаётся «подтверждаю».', addText: 'Добавить условие' }) },
    ];

    /* Таблица строк тендера (позиции лота, трудозатраты, материалы, субподряд).
     * Экран суммы не считает: сумма строки и итог приходят с сервера. После
     * записи нет суммы -- показывается прочерк, а не придуманное число. */
    function makeRowTable(cfg) {
      const rows = [];
      const tbody = h('tbody');
      const totals = h('p', { class: 'totals' });
      const empty = h('p', { class: 'empty', text: cfg.emptyText });
      const msg = notice();
      const wrap = h('div', { class: 'table-wrap' }, h('table', { class: 'rows' },
        h('thead', {}, h('tr', {}, h('th', { class: 'cell-num', text: '№' }),
          cfg.columns.map((c) => h('th', { class: c.num ? 'cell-num' : '', text: c.label })),
          cfg.computed.map((c) => h('th', { class: 'money', text: c.label })), h('th'))),
        tbody));

      const rowName = (row) => str(row.inputs[cfg.nameKey].value).trim() || 'новая строка';
      function paint(row) {
        row.tr.classList.toggle('dirty', row.dirty);
        for (const c of cfg.computed) {
          const value = row.rec[c.key];
          row.cells[c.key].textContent = value === undefined ? DASH : value;
          row.cells[c.key].classList.toggle('missing', value === undefined);
        }
        row.tag.textContent = row.dirty ? (row.rec.id === undefined ? 'не сохранено' : 'изменено') : '';
        row.tag.hidden = !row.dirty;
      }
      function syncInputs(row) {
        for (const c of cfg.columns) row.inputs[c.key].value = str(row.rec[c.key]);
      }
      function syncEmpty() {
        empty.hidden = rows.length > 0;
        wrap.hidden = rows.length === 0;
      }
      function addRow(rec, { draft = false, needs = [] } = {}) {
        const row = { rec, dirty: draft, inputs: {}, cells: {}, tr: h('tr'), tag: h('span', { class: 'state-tag' }) };
        row.tr.append(h('td', { class: 'ordinal cell-num' }));
        for (const c of cfg.columns) {
          const input = h('input', { type: 'text', value: str(rec[c.key]), class: c.num ? 'num' : '',
            inputmode: c.num ? (c.int ? 'numeric' : 'decimal') : undefined, 'aria-label': c.label,
            autocomplete: 'off', list: c.list });
          if (needs.includes(c.key)) input.classList.add('needed');
          input.addEventListener('input', () => {
            row.dirty = true;
            input.classList.remove('needed');
            row.tr.classList.remove('invalid');
            paint(row);
          });
          row.inputs[c.key] = input;
          row.tr.append(h('td', { class: c.num ? 'cell-num' : '' }, input));
        }
        for (const c of cfg.computed) {
          row.cells[c.key] = h('td', { class: 'computed money' });
          row.tr.append(row.cells[c.key]);
        }
        row.tr.append(h('td', {}, row.tag, ' ', h('button', { type: 'button', class: 'small danger', text: 'Удалить',
          onclick: () => removeRow(row) })));
        if (needs.length) row.tr.classList.add('needs-input');
        rows.push(row);
        tbody.append(row.tr);
        paint(row);
        syncEmpty();
        return row;
      }
      function dropRow(row) {
        rows.splice(rows.indexOf(row), 1);
        row.tr.remove();
        syncEmpty();
      }
      function setTotals(view) {
        totals.replaceChildren(...(view ? cfg.totals(view) : cfg.totals({})).flatMap(([label, value]) => [
          h('span', { class: 'muted', text: `${label}: ` }),
          h('strong', { text: value === undefined ? DASH : value }), '   ']));
      }
      async function removeRow(row) {
        msg.clear();
        if (row.rec.id === undefined) { dropRow(row); return; }
        const sure = await ask({ title: 'Удалить строку?', message: `«${rowName(row)}» будет удалена из тендера. Это нельзя отменить.`, ok: 'Удалить' });
        if (!sure) return;
        try {
          await api('DELETE', `${cfg.path}/${row.rec.id}`);
          dropRow(row);
          await refresh();
          cfg.changed();
        } catch (error) {
          msg.error(error);
        }
      }
      /* Свежий вид с сервера. Правки в несохранённых строках не трогаем. */
      function applyView(view) {
        const incoming = cfg.rowsOf(view);
        const byId = new Map(incoming.map((r) => [String(r.id), r]));
        for (const row of [...rows]) {
          if (row.rec.id === undefined) continue;
          const fresh = byId.get(String(row.rec.id));
          if (!fresh) { dropRow(row); continue; }
          if (!row.dirty) { row.rec = fresh; syncInputs(row); }
          else row.rec = { ...row.rec, ...Object.fromEntries(cfg.computed.map((c) => [c.key, fresh[c.key]])) };
          paint(row);
        }
        for (const fresh of incoming) {
          if (!rows.some((r) => String(r.rec.id) === String(fresh.id))) addRow(fresh);
        }
        setTotals(view);
      }
      function setView(view) {
        rows.splice(0).forEach((row) => row.tr.remove());
        cfg.rowsOf(view).forEach((rec) => addRow(rec));
        syncEmpty();
        setTotals(view);
      }
      async function refresh() {
        try { await cfg.reload(); } catch (error) { setTotals(null); msg.error(error); }
      }
      async function saveAll() {
        msg.clear();
        rows.forEach((row) => row.tr.classList.remove('invalid'));
        const pending = rows.filter((row) => row.dirty);
        if (!pending.length) { msg.show('ok', 'Нет изменений для сохранения.'); return; }
        const warnings = [];
        for (const row of pending) {
          const edits = Object.fromEntries(cfg.columns.map((c) => [c.key, row.inputs[c.key].value]));
          try {
            const { record, warning } = await writeRecord(cfg.path, recordForWrite(row.rec, edits));
            row.rec = record;
            row.dirty = false;
            syncInputs(row);
            paint(row);
            if (warning && !warnings.includes(warning)) warnings.push(warning);
          } catch (error) {
            row.tr.classList.add('invalid');
            row.tr.scrollIntoView({ block: 'nearest' });
            msg.show('error', `Строка «${rowName(row)}» не сохранена: ${error.message}`);
            await refresh().catch(() => {});
            return;
          }
        }
        if (warnings.length) {
          setTotals(null);
          msg.show('warning', 'Строки сохранены, но суммы не пересчитались.', warnings);
          return;
        }
        await refresh();
        cfg.changed();
        if (!msg.el.firstChild) msg.show('ok', 'Сохранено.');
      }
      const saveButton = h('button', { type: 'button', class: 'primary', text: 'Сохранить изменения', onclick: saveAll });
      const addButton = h('button', { type: 'button', text: cfg.addText,
        onclick: () => { const row = addRow({ ...cfg.defaults(), position: rows.length }, { draft: true }); row.inputs[cfg.nameKey].focus(); } });
      const reloadButton = h('button', { type: 'button', text: 'Обновить с сервера', onclick: async () => { msg.clear(); await refresh(); } });

      const el = h('div', {}, msg.el, empty, wrap, totals,
        h('div', { class: 'actions' }, addButton, saveButton, reloadButton, ...(cfg.extraButtons || [])));
      setTotals(null);
      syncEmpty();
      return { el, addRow, applyView, setView, msg,
        hasDirty: () => rows.some((r) => r.dirty),
        rows: () => rows,
        findById: (id) => rows.find((r) => String(r.rec.id) === String(id)) };
    }

    async function tenderScreen(root, args, ctx) {
      const picked = await chooseTender(root, args, ctx, 'tender');
      if (picked.redirected) return;
      const { tenders, orgs, tender, isNew } = picked;
      root.append(h('h1', { text: 'Карточка тендера' }),
        tenderPicker({ tenders, orgs, currentId: tender ? tender.id : '', screen: 'tender', withNew: true }));
      if (!tender && !isNew) {
        root.append(h('p', { class: 'empty', text: tenders.length ? 'Выберите тендер из списка или создайте новый.' : 'Тендеров пока нет. Нажмите «Новый тендер».' }));
        return;
      }
      const lookups = await loadLookups(['customers', 'participants', 'services']);
      const [employees] = await Promise.all([api('GET', '/api/employees')]);
      const base = tender ? `/api/tenders/${tender.id}` : '';
      const record = tender || {};

      /* -- шапка -- */
      const headMsg = notice();
      let headDirty = false;
      const dirtyTag = h('span', { class: 'state-tag', text: 'есть несохранённые изменения', hidden: true });
      const form = buildForm(TENDER_FIELDS(lookups), record, lookups, () => { headDirty = true; dirtyTag.hidden = false; });
      ctx.dirty.push(() => headDirty);
      const saveHead = h('button', { type: 'button', class: 'primary', text: isNew ? 'Создать тендер' : 'Сохранить тендер' });
      saveHead.addEventListener('click', async () => {
        saveHead.disabled = true;
        headMsg.clear();
        try {
          const edits = form.collect();
          if (edits.price_wo_vat.trim() === '') edits.price_wo_vat = '0';      // цену ещё не задали: 0,00 на виду, а не отказ
          const { record: saved } = await writeRecord('/api/tenders', recordForWrite(record, edits));
          headDirty = false;
          if (isNew) { go(`#/tender/${saved.id}`); return; }
          Object.assign(record, saved);
          dirtyTag.hidden = true;
          headMsg.show('ok', 'Тендер сохранён.');
          form.mark(null);
          if (calc) calc.refreshEstimate();
        } catch (error) {
          form.mark(error);
          headMsg.error(error);
        } finally {
          saveHead.disabled = false;
        }
      });
      const removeHead = h('button', { type: 'button', class: 'danger', text: 'Удалить тендер', hidden: isNew });
      removeHead.addEventListener('click', async () => {
        if (!(await ask({ title: 'Удалить тендер?', message: 'Тендер вместе с его позициями лота, трудозатратами и разбивкой расходов будет удалён. Это нельзя отменить.', ok: 'Удалить тендер' }))) return;
        try {
          await api('DELETE', base);
          headDirty = false;
          forgetTender();
          go('#/tender');
        } catch (error) {
          headMsg.error(error);
        }
      });
      const toBuild = h('a', { class: 'button', href: tender ? `#/build/${tender.id}` : '#/build', text: 'Перейти к сборке →', hidden: isNew });
      const headCard = h('div', { class: 'card' }, h('h2', {}, isNew ? 'Новый тендер' : 'Шапка тендера', ' ', dirtyTag), headMsg.el, form.el,
        h('div', { class: 'actions' }, saveHead, removeHead, h('span', { class: 'spacer' }), toBuild));

      if (isNew) {
        root.append(headCard, h('p', { class: 'hint', text: 'Позиции лота, трудозатраты, расходы и калькулятор появятся после того, как тендер будет создан.' }));
        return;
      }

      /* -- строки тендера -- */
      let calc = null;
      const changed = () => { if (calc) calc.refreshEstimate(); };

      const lotTable = makeRowTable({
        path: `${base}/items`, nameKey: 'name', addText: 'Добавить позицию', emptyText: 'Позиций лота пока нет. Добавьте их или вставьте списком.',
        columns: [{ key: 'name', label: 'Наименование' }, { key: 'price', label: 'Цена без НДС, руб.', num: true }],
        computed: [{ key: 'vat', label: 'НДС, руб.' }],
        defaults: () => ({}), rowsOf: (view) => view.rows || [],
        totals: (view) => [['Итого без НДС', view.total], [`НДС${view.vat_percent ? ` по ставке ${view.vat_percent} %` : ''}`, view.vat_total]],
        reload: async () => lotTable.applyView(await api('GET', `${base}/items`)), changed,
      });
      const laborTable = makeRowTable({
        path: `${base}/labor`, nameKey: 'post', addText: 'Добавить строку', emptyText: 'Трудозатрат пока нет. Добавьте строки или заполните их по нормам услуги.',
        columns: [{ key: 'post', label: 'Должность', list: 'post-list' }, { key: 'headcount', label: 'Человек', num: true, int: true },
          { key: 'hours', label: 'Часов на человека', num: true }, { key: 'rate', label: 'Ставка, руб/ч', num: true }],
        computed: [{ key: 'amount', label: 'Сумма, руб.' }],
        defaults: () => ({ headcount: '1' }), rowsOf: (view) => view.rows || [],
        totals: (view) => [['Оплата труда (ФОТ)', view.total]],
        reload: async () => laborTable.applyView(await api('GET', `${base}/labor`)), changed,
      });
      const costColumns = [{ key: 'name', label: 'Наименование' }, { key: 'unit', label: 'Единица' },
        { key: 'quantity', label: 'Количество', num: true }, { key: 'price', label: 'Цена за единицу, руб.', num: true },
        { key: 'note', label: 'Примечание' }];
      const reloadCosts = async () => {
        const view = await api('GET', `${base}/costs`);
        materialsTable.applyView(view);
        subcontractTable.applyView(view);
      };
      const costTable = (kind, noun, empty) => makeRowTable({
        path: `${base}/costs`, nameKey: 'name', addText: 'Добавить строку', emptyText: empty,
        columns: costColumns, computed: [{ key: 'amount', label: 'Сумма, руб.' }],
        defaults: () => ({ kind, unit: 'усл.ед.', quantity: '1' }),
        rowsOf: (view) => (view.rows || []).filter((r) => r.kind === kind),
        totals: (view) => [[`Итого ${noun}`, view.totals ? view.totals[kind] : undefined]],
        reload: reloadCosts, changed,
      });
      const materialsTable = costTable('materials', 'материалы', 'Материалов пока нет.');
      const subcontractTable = costTable('subcontract', 'субподряд', 'Субподряда пока нет.');
      ctx.dirty.push(() => [lotTable, laborTable, materialsTable, subcontractTable].some((t) => t.hasDirty()));

      const [lotView, laborView, costsView] = await Promise.all([
        api('GET', `${base}/items`), api('GET', `${base}/labor`), api('GET', `${base}/costs`)]);
      lotTable.setView(lotView);
      laborTable.setView(laborView);
      materialsTable.setView(costsView);
      subcontractTable.setView(costsView);

      /* -- вставка списком в лот -- */
      const lotCard = h('div', { class: 'card' }, h('h2', {}, 'Позиции лота ', h('span', { class: 'sub', text: 'то, что заказчик просит расценить: объекты, этапы, виды работ' })),
        lotTable.el, pasteBox());
      function pasteBox() {
        const pasteMsg = notice();
        const area = h('textarea', { rows: 6, placeholder: 'Вставьте сюда таблицу из Excel или Word: наименование и цена в двух колонках', 'aria-label': 'Список позиций' });
        const previewBox = h('div');
        let preview = [];       // { name, price, problem, tr, nameInput, priceInput, problemCell }
        const applyBtn = h('button', { type: 'button', class: 'primary', text: 'Заменить список лота этим списком', disabled: true });

        function drawPreview(parsed) {
          preview = [];
          previewBox.replaceChildren();
          if (!parsed.length) { applyBtn.disabled = true; return; }
          const body = h('tbody');
          for (const item of parsed) {
            const nameInput = h('input', { type: 'text', value: item.name, 'aria-label': 'Наименование', autocomplete: 'off' });
            const priceInput = h('input', { type: 'text', value: item.price, class: 'num', 'aria-label': 'Цена', autocomplete: 'off', inputmode: 'decimal' });
            const problemCell = h('td', { class: 'row-problem', text: item.problem });
            const entry = { nameInput, priceInput, problemCell };
            entry.tr = h('tr', { class: item.problem ? 'invalid' : '' }, h('td', { class: 'ordinal cell-num' }), h('td', {}, nameInput), h('td', { class: 'cell-num' }, priceInput), problemCell,
              h('td', {}, h('button', { type: 'button', class: 'small', text: 'Убрать', onclick: () => { entry.tr.remove(); preview.splice(preview.indexOf(entry), 1); recheck(); } })));
            const onEdit = () => { entry.tr.classList.remove('invalid'); problemCell.textContent = ''; recheck(); };
            nameInput.addEventListener('input', onEdit);
            priceInput.addEventListener('input', onEdit);
            preview.push(entry);
            body.append(entry.tr);
          }
          previewBox.append(h('p', { class: 'hint', text: 'Проверьте разбор: цены остаются такими, как вставлены. Строки можно поправить или убрать. Позиции лота заменятся целиком.' }),
            h('div', { class: 'table-wrap' }, h('table', { class: 'rows' }, h('thead', {}, h('tr', {}, h('th', { class: 'cell-num', text: '№' }), h('th', { text: 'Наименование' }), h('th', { class: 'cell-num', text: 'Цена, руб.' }), h('th', { text: 'Замечание' }), h('th'))), body)));
          recheck();
        }
        function recheck() {
          const blank = preview.some((e) => !e.nameInput.value.trim() || !e.priceInput.value.trim());
          const flagged = preview.some((e) => e.problemCell.textContent);
          applyBtn.disabled = !preview.length || blank || flagged;
        }
        const parseBtn = h('button', { type: 'button', text: 'Разобрать список', onclick: () => {
          pasteMsg.clear();
          const { rows } = parsePasted(area.value);
          if (!rows.length) { pasteMsg.show('warning', 'В поле пусто: вставьте список (Ctrl+V) и нажмите «Разобрать список».'); drawPreview([]); return; }
          drawPreview(rows);
          if (rows.some((r) => r.problem)) pasteMsg.show('warning', 'В некоторых строках есть замечания — они подсвечены. Поправьте или уберите такие строки.');
        } });
        const clipBtn = h('button', { type: 'button', text: 'Вставить из буфера обмена', onclick: async () => {
          pasteMsg.clear();
          try {
            area.value = await navigator.clipboard.readText();
            parseBtn.click();
          } catch (error) {
            pasteMsg.show('warning', 'Браузер не дал прочитать буфер обмена. Щёлкните в поле ниже и нажмите Ctrl+V.');
            area.focus();
          }
        } });
        applyBtn.addEventListener('click', async () => {
          pasteMsg.clear();
          const existing = lotTable.rows().filter((r) => r.rec.id !== undefined);
          if (existing.length) {
            const sure = await ask({ title: 'Заменить список лота?', message: 'Эти позиции будут удалены, на их месте встанет вставленный список:',
              items: existing.map((r) => r.rec.name), ok: 'Заменить список' });
            if (!sure) return;
          }
          applyBtn.disabled = true;
          try {
            const view = await api('PUT', `${base}/items`, lotPayload(preview.map((e) => ({ name: e.nameInput.value.trim(), price: e.priceInput.value.trim() }))));
            lotTable.setView(view);
            area.value = '';
            drawPreview([]);
            if (view.warning) pasteMsg.show('warning', view.warning, ['Список заменён. Суммы по строкам и итоги не пришли — не пересчитывайте их вручную: обновите страницу.']);
            else { pasteMsg.show('ok', 'Список лота заменён.'); changed(); }
          } catch (error) {
            const index = error.path && error.path[0] === 'items' ? error.path[1] : undefined;
            if (index !== undefined && preview[index]) {
              preview[index].tr.classList.add('invalid');
              preview[index].problemCell.textContent = error.message;
              preview[index].tr.scrollIntoView({ block: 'nearest' });
            }
            pasteMsg.error(error);
            recheck();
          }
        });
        return h('details', { class: 'paste' }, h('summary', { text: 'Вставить список из буфера обмена', style: 'cursor:pointer;font-weight:600;margin-top:12px' }),
          pasteMsg.el, h('div', { class: 'actions' }, clipBtn), h('div', { style: 'margin-top:8px' }, area),
          h('div', { class: 'actions' }, parseBtn), previewBox, h('div', { class: 'actions' }, applyBtn));
      }

      /* -- трудозатраты: список должностей и заполнение по нормам -- */
      const postList = h('datalist', { id: 'post-list' },
        [...new Set(employees.map((e) => e.position).filter(Boolean))].map((p) => h('option', { value: p })));
      const normSelect = h('select', { 'aria-label': 'Услуга для норм часов' },
        h('option', { value: '', text: 'Услуга…' }), lookups.services.map(([id, label]) => h('option', { value: String(id), text: label })));
      normSelect.value = inferServiceId(record.subject, lookups.serviceRecords);
      const normBtn = h('button', { type: 'button', text: 'Заполнить по нормам услуги', onclick: () => {
        const service = lookups.serviceRecords.find((s) => String(s.id) === normSelect.value);
        if (!service) { laborTable.msg.show('warning', 'Выберите услугу, по нормам которой заполнить строки.'); return; }
        laborTable.msg.clear();
        const have = new Set(laborTable.rows().map((r) => r.inputs.post.value.trim().toLowerCase()));
        const missingRate = [];
        for (const post of Object.keys(service.norms)) {
          if (have.has(post.trim().toLowerCase())) continue;
          const person = employees.find((e) => (e.position || '').trim().toLowerCase() === post.trim().toLowerCase());
          const rec = { post, headcount: '1', hours: service.norms[post], hours_value: (service.norms_value || {})[post],
            rate: person ? person.hourly_rate : '', rate_value: person ? person.hourly_rate_value : undefined };
          laborTable.addRow(rec, { draft: true, needs: person ? [] : ['rate'] });
          if (!person) missingRate.push(post);
        }
        laborTable.msg.show(missingRate.length ? 'warning' : 'ok',
          'Строки добавлены, но ещё не сохранены: проверьте численность и нажмите «Сохранить изменения».',
          missingRate.length ? [`В справочнике сотрудников нет должности «${missingRate.join('», «')}» — ставку введите вручную (поле подсвечено).`] : []);
      } });
      const laborCard = h('div', { class: 'card' }, h('h2', {}, 'Трудозатраты (ФОТ)'), postList,
        h('div', { class: 'inline', style: 'margin-bottom:10px' }, h('div', {}, normSelect), normBtn), laborTable.el);
      const materialsCard = h('div', { class: 'card' }, h('h2', { text: 'Материалы' }), materialsTable.el);
      const subcontractCard = h('div', { class: 'card' }, h('h2', { text: 'Субподряд' }), subcontractTable.el);

      /* -- калькулятор -- */
      calc = calculator();
      function calculator() {
        const targetInput = h('input', { type: 'text', class: 'num', inputmode: 'decimal', autocomplete: 'off', placeholder: 'Например, 5 700 000,00', 'aria-label': 'Целевая цена' });
        const noVat = h('input', { type: 'radio', name: 'vat-mode', value: 'wo', checked: true });
        const withVat = h('input', { type: 'radio', name: 'vat-mode', value: 'with' });
        const msg = notice();
        const resultBox = h('div');
        const proposalBox = h('div');
        const stale = h('p', { class: 'stale', text: 'Цель или режим изменились — нажмите «Посчитать».', hidden: true });
        let proposal = null;     // последнее предложение подбора (ещё не записано)
        let shown = null;        // последний показанный расчёт
        let vatPercent;          // ставка НДС из последнего ответа «Посчитать»: в ответе подбора её нет
        const includesVat = () => withVat.checked;
        const targetBody = () => {
          const text = targetInput.value.trim();
          return text ? { target: text, target_includes_vat: includesVat() } : {};
        };
        [targetInput, noVat, withVat].forEach((c) => c.addEventListener('input', () => { stale.hidden = false; }));
        [noVat, withVat].forEach((c) => c.addEventListener('change', () => { stale.hidden = false; }));

        const ROWS = [['materials', 'Материалы'], ['labor', 'Оплата труда (ФОТ)'], ['insurance', 'Страховые взносы'],
          ['subcontract', 'Субподряд'], ['overhead', 'Накладные расходы'], ['profit', 'Прибыль'],
          ['total_wo_vat', 'Итого без НДС'], ['vat', 'НДС'], ['total_with_vat', 'Итого с НДС']];

        function drawEstimate(data, { pending }) {
          shown = { estimate: data.estimate, pending };
          if (data.vat_percent !== undefined) vatPercent = data.vat_percent;
          const table = h('table', { class: 'estimate-table' }, h('tbody', {}, ROWS.map(([key, label]) => h('tr', { class: key.startsWith('total') ? 'total' : '' },
            h('td', { text: key === 'vat' ? (vatPercent === undefined ? 'НДС' : `НДС по ставке ${vatPercent} %`) : label }), h('td', { text: data.estimate[key] })))));
          const pieces = [pending ? h('p', { class: 'notice warning', text: 'Это расчёт по подобранным часам — они ещё не записаны в тендер.' }) : null, table,
            h('p', { class: 'hint', text: 'Итог без НДС округляется по настройке округления (обычно до целых рублей), поэтому может отличаться от суммы напечатанных выше статей — до 50 копеек. Так устроена книга заказчика.' })];
          if (data.difference !== undefined) {
            const kind = differenceKind(data.difference_value);
            const which = data.target_includes_vat ? 'с НДС' : 'без НДС';
            const own = data.target_includes_vat ? data.estimate.total_with_vat : data.estimate.total_wo_vat;
            const verdict = { above: 'итог выше цели', below: 'итог ниже цели', equal: 'итог равен цели', unknown: '' }[kind];
            pieces.push(h('div', { class: `gap ${kind}`, role: 'status' },
              h('div', { class: 'line', text: `Цель ${which}: ${data.target} руб.` }),
              h('div', { class: 'line', text: `Итог ${which} по расчёту: ${own} руб.` }),
              h('div', { class: 'line' }, 'Расхождение (итог минус цель): ', h('span', { class: 'figure', text: `${data.difference} руб.` }), verdict ? ` — ${verdict}` : '')));
          } else {
            pieces.push(h('p', { class: 'hint', text: 'Целевая цена не задана — расхождение показывается, когда она введена.' }));
          }
          resultBox.replaceChildren(...pieces.filter(Boolean));
          priceBtn.disabled = pending;
          stale.hidden = true;
        }

        async function refreshEstimate() {
          msg.clear();
          proposal = null;
          proposalBox.replaceChildren();
          try {
            drawEstimate(await api('POST', `${base}/estimate`, targetBody()), { pending: false });
          } catch (error) {
            resultBox.replaceChildren();
            msg.error(error);
          }
        }

        async function solve() {
          msg.clear();
          const body = targetBody();
          if (!body.target) { msg.show('warning', 'Введите целевую цену — под неё подбираются часы.'); targetInput.focus(); return; }
          solveBtn.disabled = true;
          try {
            const data = await api('POST', `${base}/solve`, body);
            proposal = data;
            drawEstimate(data, { pending: true });
            drawProposal(data);
          } catch (error) {
            msg.error(error);
          } finally {
            solveBtn.disabled = false;
          }
        }

        function drawProposal(data) {
          const apply = h('button', { type: 'button', class: 'primary', text: 'Применить подобранные часы', onclick: applyProposal });
          const discard = h('button', { type: 'button', text: 'Отменить предложение', onclick: refreshEstimate });
          proposalBox.replaceChildren(h('h3', { text: 'Предложение: подобранные часы' }),
            h('p', { class: 'hint', text: 'Часы специалистов пропорционально масштабируются под цель. Ставки, проценты, материалы и субподряд не меняются. В тендере ничего не изменено, пока вы не нажмёте «Применить».' }),
            h('div', { class: 'table-wrap' }, h('table', { class: 'estimate-table' },
              h('thead', {}, h('tr', {}, h('th', { text: 'Должность' }), h('th', { text: 'Человек' }), h('th', { text: 'Было, ч' }), h('th', { text: 'Станет, ч' }), h('th', { text: 'Сумма, руб.' }))),
              h('tbody', {}, data.labor.map((row) => {
                const current = laborTable.findById(row.id);
                return h('tr', {}, h('td', { text: row.post }), h('td', { text: str(row.headcount) }),
                  h('td', { text: current ? current.rec.hours : DASH }), h('td', { text: row.hours }), h('td', { text: row.amount }));
              })))),
            h('div', { class: 'actions' }, apply, discard));
        }

        async function applyProposal() {
          if (!proposal) return;
          if (laborTable.hasDirty()) {
            msg.show('warning', 'В трудозатратах есть несохранённые правки. Сохраните их или обновите строки, затем подберите часы заново.');
            return;
          }
          msg.clear();
          const warnings = [];
          try {
            for (const row of proposal.labor) {
              const { warning } = await writeRecord(`${base}/labor`, recordForWrite(row, {}));
              if (warning && !warnings.includes(warning)) warnings.push(warning);
            }
          } catch (error) {
            msg.error(error);
            return;
          }
          const keep = targetBody();
          if (warnings.length) {
            msg.show('warning', 'Часы записаны, но суммы не пересчитались.', warnings);
            proposalBox.replaceChildren();
            proposal = null;
            return;
          }
          laborTable.setView(await api('GET', `${base}/labor`));
          await refreshEstimate();
          if (keep.target) msg.show('ok', 'Подобранные часы записаны в тендер.');
        }

        const solveBtn = h('button', { type: 'button', text: 'Подобрать часы под цель', onclick: solve });
        const calcBtn = h('button', { type: 'button', class: 'primary', text: 'Посчитать', onclick: refreshEstimate });
        const fromHead = h('button', { type: 'button', class: 'small', text: 'Взять цену из шапки', onclick: () => {
          targetInput.value = form.inputs.price_wo_vat.value;
          noVat.checked = true;
          stale.hidden = false;
        } });
        const priceBtn = h('button', { type: 'button', text: 'Записать итог без НДС в цену тендера', onclick: () => {
          if (!shown || shown.pending) return;
          form.inputs.price_wo_vat.value = shown.estimate.total_wo_vat;
          form.inputs.price_wo_vat.dispatchEvent(new Event('input', { bubbles: true }));
          headMsg.show('warning', 'Цена без НДС в шапке заменена на итог расчёта. Нажмите «Сохранить тендер», чтобы запомнить.');
          form.inputs.price_wo_vat.scrollIntoView({ block: 'center', behavior: 'smooth' });
        } });
        priceBtn.disabled = true;
        const el = h('div', { class: 'card sticky' }, h('h2', { text: 'Калькулятор цены' }), msg.el,
          h('label', { class: 'field' }, h('span', { class: 'lbl', text: 'Целевая цена, руб.' }), targetInput),
          h('div', { class: 'radio-row', role: 'radiogroup', 'aria-label': 'Цель с НДС или без НДС', style: 'margin:8px 0' },
            h('label', {}, noVat, 'Цена без НДС'), h('label', {}, withVat, 'Цена с НДС')),
          h('div', { class: 'actions', style: 'margin-top:0' }, calcBtn, solveBtn, fromHead), stale,
          resultBox, h('div', { class: 'actions' }, priceBtn), proposalBox);
        return { el, refreshEstimate };
      }

      root.append(h('div', { class: 'two-col' }, h('div', {}, headCard, lotCard, laborCard, materialsCard, subcontractCard), calc.el));
      calc.refreshEstimate();
    }

    /* ================================================================== */
    /* Экран «Сборка»                                                       */
    /* ================================================================== */

    async function buildScreen(root, args, ctx) {
      const picked = await chooseTender(root, args, ctx, 'build');
      if (picked.redirected) return;
      const { tenders, orgs, tender } = picked;
      root.append(h('h1', { text: 'Сборка пакета' }), tenderPicker({ tenders, orgs, currentId: tender ? tender.id : '', screen: 'build', withNew: false }));
      if (!tender) {
        root.append(h('p', { class: 'empty', text: tenders.length ? 'Выберите тендер, пакет документов для которого нужно собрать.' : 'Тендеров пока нет. Заведите тендер на экране «Карточка тендера».' }));
        return;
      }
      const [profiles, experience, services] = await Promise.all([
        api('GET', '/api/profiles'), api('GET', '/api/experience'), api('GET', '/api/services')]);

      /* документы */
      const choices = documentChoices(profiles, tender);
      const chosen = new Set(choices.filter((c) => c.checked).map((c) => c.name));
      const docList = h('div', { class: 'check-list' });
      const hasProfile = profiles.some((p) => String(p.customer_id) === String(tender.customer_id));
      function drawDocs() {
        docList.replaceChildren(...choices.map((c) => h('label', {},
          h('input', { type: 'checkbox', checked: chosen.has(c.name), onchange: (e) => { if (e.target.checked) chosen.add(c.name); else chosen.delete(c.name); } }),
          h('span', { text: c.name }))));
        if (!choices.length) docList.replaceChildren(h('p', { class: 'empty', text: 'Профилей заказчиков пока нет, поэтому список документов пуст. Заведите профиль на экране «Справочники» или впишите название документа ниже.' }));
      }
      const extraName = h('input', { type: 'text', placeholder: 'Название документа, как в профиле', 'aria-label': 'Название документа', autocomplete: 'off' });
      const addDoc = () => {
        const name = extraName.value.trim();
        if (!name) return;
        if (!choices.some((c) => c.name === name)) choices.push({ name, checked: true });
        chosen.add(name);
        extraName.value = '';
        drawDocs();
      };
      extraName.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addDoc(); } });
      drawDocs();

      /* опыт */
      let selected = (tender.experience_ids || []).map(String);
      let serviceFilter = inferServiceId(tender.subject, services);
      const serviceSelect = h('select', { 'aria-label': 'Фильтр опыта по услуге' },
        h('option', { value: '', text: 'Все услуги' }), services.map((s) => h('option', { value: String(s.id), text: [s.code, s.name].filter(Boolean).join(' — ') })));
      serviceSelect.value = serviceFilter;
      const expList = h('div', { class: 'check-list' });
      const expSummary = h('p', { class: 'hint' });
      const describeExp = (r) => [r.customer_name || 'без заказчика', r.subject].filter(Boolean).join(' — ');
      function drawExperience() {
        const visible = experience.filter((r) => !serviceFilter || String(r.service_id) === serviceFilter);
        expList.replaceChildren(...visible.map((r) => h('label', {},
          h('input', { type: 'checkbox', checked: selected.includes(String(r.id)),
            onchange: (e) => { selected = toggleId(selected, r.id, e.target.checked); drawExperience(); } }),
          h('span', {}, describeExp(r), h('span', { class: 'meta', text: [r.contract_number && `контракт № ${r.contract_number}`, r.contract_date, r.amount && `${r.amount} руб.`, r.term].filter(Boolean).join(' · ') })))));
        if (!visible.length) expList.replaceChildren(h('p', { class: 'empty', text: experience.length ? 'По выбранной услуге опыта нет. Выберите «Все услуги».' : 'Справочник опыта пуст.' }));
        const names = selected.map((id) => experience.find((r) => String(r.id) === id)).filter(Boolean).map(describeExp);
        expSummary.textContent = names.length ? `В справку пойдёт, в порядке отметки: ${names.join('; ')}.` : 'Опыт не выбран: справка об опыте будет без записей.';
      }
      serviceSelect.addEventListener('change', () => { serviceFilter = serviceSelect.value; drawExperience(); });
      drawExperience();

      /* папка и сборка */
      const folder = h('input', { type: 'text', value: str(tender.output_dir), autocomplete: 'off', placeholder: 'C:\\Тендеры\\Название заявки', 'aria-label': 'Папка для пакета' });
      const msg = notice();
      const resultBox = h('div');
      const buildButton = h('button', { type: 'button', class: 'primary big', text: 'Собрать пакет' });
      const builder = createBuilder({
        api,
        setBusy: (busy) => {
          buildButton.disabled = busy;
          buildButton.textContent = busy ? 'Идёт сборка… не закрывайте страницу' : 'Собрать пакет';
          buildButton.setAttribute('aria-busy', String(busy));
        },
        confirmReplace: (files) => ask({
          title: 'В папке уже есть документы',
          message: 'Эти документы будут заменены новыми. Правки, внесённые в них вручную, пропадут:',
          items: files.map((f) => `${f.document} — файл «${f.name}»`),
          ok: 'Заменить и собрать', cancel: 'Отмена' }),
      });
      buildButton.addEventListener('click', async () => {
        msg.clear();
        resultBox.replaceChildren();
        const outcome = await builder.run({ tenderId: tender.id, documents: choices.filter((c) => chosen.has(c.name)).map((c) => c.name),
          outputDir: folder.value, experienceIds: selected });
        if (outcome.status === 'failed') msg.show('error', outcome.error);
        else if (outcome.status === 'cancelled') msg.show('warning', 'Сборка отменена: файлы в папке не тронуты.');
        else if (outcome.status === 'built') drawResult(outcome);
      });
      function drawResult({ result, note }) {
        const parts = [h('h2', { text: 'Результат сборки' }),
          h('p', {}, 'Папка: ', h('code', { class: 'path', text: result.output_dir }))];
        parts.push(h('h3', { text: 'Записано' }),
          result.written.length ? h('ul', { class: 'result-list' }, result.written.map((w) => h('li', {}, w.name, ' ', h('code', { class: 'path', text: w.path }))))
            : h('p', { class: 'empty', text: 'Ни один документ не записан.' }));
        if (result.skipped.length) {
          parts.push(h('div', { class: 'notice error' }, h('p', { class: 'title', text: 'Не собрано — и почему' }),
            h('ul', {}, result.skipped.map((s) => h('li', { text: `${s.document}: ${s.reason}` })))));
        }
        if (result.warnings.length) {
          parts.push(h('div', { class: 'notice warning' }, h('p', { class: 'title', text: 'Предупреждения — проверьте перед подачей' }),
            h('ul', {}, result.warnings.map((w) => h('li', { text: `${w.document}: ${w.text}` })))));
        }
        parts.push(note ? h('div', { class: 'notice warning' }, h('p', { text: note }))
          : h('p', { class: 'hint', text: 'Папка и отбор опыта запомнены в тендере — в следующий раз они подставятся сами.' }));
        resultBox.replaceChildren(...parts);
        resultBox.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      }

      root.append(h('div', { class: 'two-col' },
        h('div', {},
          h('div', { class: 'card' }, h('h2', { text: 'Документы пакета' }),
            hasProfile ? h('p', { class: 'hint', text: 'Отмечены документы из профиля заказчика. Остальные описаны в других профилях — их можно добавить галочкой.' })
              : h('p', { class: 'notice warning', text: 'У заказчика этого тендера нет профиля: ничего не отмечено заранее. Отметьте документы или заведите профиль в справочниках.' }),
            docList, h('div', { class: 'inline', style: 'margin-top:8px' }, h('div', {}, extraName), h('button', { type: 'button', text: 'Добавить документ', onclick: addDoc }))),
          h('div', { class: 'card' }, h('h2', { text: 'Опыт для справки' }),
            h('div', { class: 'inline' }, h('div', {}, h('label', { class: 'field' }, h('span', { class: 'lbl', text: 'Показать опыт по услуге' }), serviceSelect))),
            h('p', { class: 'hint', text: tender.subject && serviceFilter ? 'Услуга подобрана по предмету закупки; выберите «Все услуги», чтобы увидеть остальное.' : 'В тендере нет поля «услуга»: выберите её сами или оставьте «Все услуги».' }),
            expList, expSummary)),
        h('div', { class: 'card sticky' }, h('h2', { text: 'Сборка' }),
          h('label', { class: 'field' }, h('span', { class: 'lbl', text: 'Папка, в которую собрать пакет' }), folder,
            h('span', { class: 'hint', text: 'Полный путь на этом компьютере, например C:\\Тендеры\\Коряжма. Браузер не умеет выбирать папку, поэтому путь вводится текстом (его можно скопировать из проводника). Папка запоминается в тендере после первой сборки.' })),
          h('div', { class: 'actions' }, buildButton), msg.el, resultBox)));
    }

    /* ---- запуск ---- */
    if (!location.hash) location.hash = '#/refs';
    const remembered = rememberedTender();
    if (remembered) rememberTender(remembered);
    render();
  }());
}
