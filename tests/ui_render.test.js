// Отрисовка экранов: код страницы из app/web/static/app.js ИСПОЛНЯЕТСЯ в Node поверх
// заглушки DOM (tests/fake_dom.js) и разговаривает с настоящим сервером.
// Запуск: из pytest (tests/test_ui_smoke.py), с TENDER_BASE и TENDER_OUT_DIR.
//
// Что это проверяет: что построение таблиц, отрисовка расчёта, результата сборки,
// переходы между экранами и замок сборки доходят до конца и не падают на ответах
// сервера -- полных, без итогов и с предупреждением -- и что на экране нет
// «undefined», «null», «NaN». Чего это НЕ проверяет: как выглядит страница,
// удобно ли ей пользоваться, настоящее поведение элементов браузера (фокус,
// прокрутка, раскрытие details, диалог, буфер обмена). Это проверяет человек.
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createWorld } = require('./fake_dom.js');

const BASE = process.env.TENDER_BASE;
const OUT_DIR = process.env.TENDER_OUT_DIR;
const live = BASE ? test : test.skip;
const realFetch = globalThis.fetch;
const NBSP = ' ';
const DASH = '—';
const APP_PATH = require.resolve('../app/web/static/app.js');
let requestsSent = 0;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// ---- обращение к серверу для заготовки данных (не через код страницы) -------------

async function call(method, url, body) {
  requestsSent += 1;
  const response = await realFetch(BASE + url, { method, headers: { 'Content-Type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body) });
  const data = await response.json();
  if (!response.ok) throw new Error(`${method} ${url}: ${data.error}`);
  return data;
}

async function seed(tag) {
  const me = await call('POST', '/api/organizations', { role: 'participant', short_name: `ООО «Техносфера ${tag}»`, inn: '5904188041' });
  const customer = await call('POST', '/api/organizations', { role: 'customer', short_name: `ПАО «Заказчик ${tag}»`, inn: '7708514824' });
  const service = await call('POST', '/api/services', { code: `КЭР${tag}`, name: 'Комплексный экологический отчёт', stage_wording: '',
    norms: { 'Инженер-эколог': '40,5', 'Руководитель проекта': '20' } });
  const employees = [
    await call('POST', '/api/employees', { full_name: `Петров ${tag}`, position: 'Инженер-эколог', hourly_rate: '1250,50' }),
    await call('POST', '/api/employees', { full_name: `Сидоров ${tag}`, position: 'Руководитель проекта', hourly_rate: '2000' }),
  ];
  const asset = await call('POST', '/api/assets', { name: `Газоанализатор ГАНК-4 ${tag}`, kind: 'equipment',
    maker_model: 'НПО «Прибор», ГАНК-4', quantity: '2', year: '2021', location: 'Пермь, ул. Ласьвинская, 98' });
  // настройки -- одна строка на всю базу; ставка травматизма засевается, чтобы на
  // вкладке «Настройки» было что утверждать, кроме заголовка
  const settings = await call('POST', '/api/settings', { injury_rate_percent: '0,3' });
  const tender = await call('POST', '/api/tenders', { customer_id: String(customer.id), participant_id: String(me.id),
    service_id: String(service.id), purchase_number: `N-${tag}`, purchase_date: '2025-10-24',
    subject: 'Разработка КЭР', price_wo_vat: '5700000' });
  const base = `/api/tenders/${tender.id}`;
  await call('PUT', `${base}/items`, { items: [{ position: 0, name: 'Объект 1', price: '350 000,00' },
    { position: 1, name: 'Объект 2', price: '420 000,50' }] });
  await call('POST', `${base}/labor`, { post: 'Инженер-эколог', headcount: '2', hours: '40,5', rate: '1250,50' });
  await call('POST', `${base}/labor`, { post: 'Руководитель проекта', headcount: '1', hours: '20', rate: '2000' });
  await call('POST', `${base}/costs`, { kind: 'materials', name: 'Бумага', unit: 'пач.', quantity: '10', price: '350' });
  await call('POST', `${base}/costs`, { kind: 'subcontract', name: 'Лаборатория', unit: 'усл.', quantity: '1', price: '12 000' });
  await call('POST', '/api/profiles', { customer_id: String(customer.id),
    documents: ['Справка о МТР', 'Декларация соответствия'], docs: {}, declaration_text: '' });
  const experience = [];
  for (const n of ['1', '2']) {
    experience.push(await call('POST', '/api/experience', { customer_name: `Заказчик опыта ${n}`, contract_number: n,
      contract_date: '01.02.2024', subject: `Работа ${n}`, amount: '1000000', term: '2024', status: 'исполнен', service_id: String(service.id) }));
  }
  return { tender, customer, me, service, employees, asset, settings, experience, base };
}

/* Что на какой вкладке справочников обязано быть видно: ЗАСЕЯННОЕ значение,
 * а не только отсутствие «undefined». Строки берутся из ответов сервера, а не
 * сочиняются здесь, поэтому утверждение не зависит от того, как сервер печатает
 * деньги. Опечатка в имени поля даёт пустые клетки, и assertClean её пропускает:
 * текста на экране хватает за счёт заголовков, вкладок и кнопок -- ровно тот
 * отказ, ради которого заглушка DOM и вводилась. */
function seededOnTabs(s, tag) {
  return {
    organizations: { text: [`ПАО «Заказчик ${tag}»`, `ООО «Техносфера ${tag}»`, 'ИНН 7708514824'] },
    employees: { text: [s.employees[0].full_name, s.employees[1].full_name, `${s.employees[0].hourly_rate} руб/ч`] },
    assets: { text: [s.asset.name, s.asset.maker_model] },
    experience: { text: [s.experience[0].customer_name, `№ ${s.experience[1].contract_number}`,
      `${s.experience[0].amount} руб.`] },
    services: { text: [`КЭР${tag}`, 'Комплексный экологический отчёт'] },
    profiles: { text: [`ПАО «Заказчик ${tag}»`, 'Справка о МТР', 'Декларация соответствия'] },
    // настройки живут в полях ввода: textContent у input пуст, поэтому значения
    // сверяются с ответом сервера по каждому полю формы
    settings: { fields: [['Страховые взносы, %', 'insurance_rate_percent'], ['Травматизм, %', 'injury_rate_percent'],
      ['Накладные, % от оплаты труда со страховыми', 'overhead_rate_percent'], ['Прибыль, %', 'profit_rate_percent'],
      ['НДС по умолчанию, %', 'vat_rate_percent'],
      ['Округление итога без НДС, знаков после запятой', 'total_rounding_digits']] },
  };
}

// ---- страница в заглушке DOM ---------------------------------------------------------

function loadPage() {
  const page = { log: [], tweaks: [], gates: [], inFlight: 0, held: 0 };
  const fetchImpl = async (url, init = {}) => {
    const method = init.method || 'GET';
    const body = init.body ? JSON.parse(init.body) : undefined;
    page.log.push({ method, url, body });
    page.inFlight += 1;
    try {
      for (const gate of page.gates) {
        if (gate.match(url, method)) { page.held += 1; await gate.promise; page.held -= 1; }
      }
      requestsSent += 1;
      const real = await realFetch(BASE + url, init);
      const text = await real.text();
      let data = null;
      try { data = text ? JSON.parse(text) : null; } catch (error) { data = null; }
      for (const tweak of page.tweaks) {
        if (data && tweak.match(url, method)) data = tweak.edit(data, body) ?? data;
      }
      return new Response(JSON.stringify(data), { status: real.status, headers: { 'Content-Type': 'application/json' } });
    } finally {
      page.inFlight -= 1;
    }
  };

  const world = createWorld({ fetch: fetchImpl });
  const saved = {};
  const globals = { document: world.document, window: world.window, location: world.location,
    localStorage: world.localStorage, Node: world.Node, Event: world.Event, fetch: fetchImpl };
  for (const [name, value] of Object.entries(globals)) {
    saved[name] = Object.getOwnPropertyDescriptor(globalThis, name);
    Object.defineProperty(globalThis, name, { value, configurable: true, writable: true });
  }
  delete require.cache[APP_PATH];
  require(APP_PATH);                       // страница стартует сама: render() на #/refs

  page.unload = () => {
    for (const [name, descriptor] of Object.entries(saved)) {
      if (descriptor) Object.defineProperty(globalThis, name, descriptor); else delete globalThis[name];
    }
    delete require.cache[APP_PATH];
  };
  page.world = world;
  page.idle = async () => {
    let calm = 0;
    for (let i = 0; i < 500 && calm < 4; i += 1) {
      await sleep(10);
      calm = page.inFlight === page.held ? calm + 1 : 0;
    }
  };
  page.goto = async (hash) => { world.location.hash = hash; await page.idle(); };

  page.text = () => world.screen.textContent;
  page.all = (predicate) => world.screen.all(predicate);
  const hasClass = (el, name) => el.className.split(/\s+/).includes(name);
  page.card = (title) => {
    const found = page.all((e) => hasClass(e, 'card') && e.children.some((c) => c.tagName === 'H2' && c.textContent.startsWith(title)));
    assert.ok(found.length, `нет карточки «${title}»`);
    return found[0];
  };
  page.button = (label, within = world.screen) => {
    const found = within.all((e) => e.tagName === 'BUTTON' && e.textContent.trim() === label);
    assert.ok(found.length, `нет кнопки «${label}»`);
    return found[0];
  };
  page.input = (label) => {
    const found = page.all((e) => e.tagName === 'INPUT' && e.attributes['aria-label'] === label);
    assert.ok(found.length, `нет поля «${label}»`);
    return found[0];
  };
  page.field = (label) => {
    const found = page.all((e) => e.tagName === 'LABEL' && hasClass(e, 'field') && e.children[0] && e.children[0].textContent === label);
    assert.ok(found.length, `нет поля формы «${label}»`);
    return found[0].children[1];
  };
  page.rows = (within) => within.all((e) => e.tagName === 'TR' && e.parentNode.tagName === 'TBODY');
  page.inputs = (within) => within.all((e) => e.tagName === 'INPUT');
  page.type = (el, value) => { el.value = value; el.dispatchEvent({ type: 'input', bubbles: true }); };
  page.press = async (el) => { assert.equal(el.disabled, false, 'кнопка заблокирована'); el.click(); await page.idle(); };
  page.answerDialog = async (answer) => {
    assert.equal(world.dialog.open, true, 'диалог не открыт');
    world.dialog.returnValue = answer;
    world.dialog.close();
    await page.idle();
  };
  page.dialogItems = () => world.root.querySelector('#dialog-items').textContent;
  return page;
}

/* Прочерк утверждается в ЗНАЧЕНИИ, а не в строке целиком: подпись и значение --
 * разные узлы, и `includes('—')` на всей строке зелен и тогда, когда прочерк
 * попал в подпись, а значение просто пусто. Ровно это и нужно различать. */
function totalsOf(within) {
  return within.all((e) => e.className === 'totals').map((line) => ({
    text: line.textContent,
    values: line.all((e) => e.tagName === 'STRONG').map((v) => v.textContent),
  }));
}

function assertDashes(lines, where) {
  assert.ok(lines.length, `нет итоговых строк (${where})`);
  for (const line of lines) {
    assert.ok(line.values.length, `в итоговой строке нет значения (${where}): ${line.text}`);
    assert.deepEqual([...new Set(line.values)], [DASH], `значение итога -- не прочерк (${where}): ${line.text}`);
    assert.match(line.text, /[А-Яа-я].*:/, `итог без подписи -- не отличить от пустой клетки (${where}): ${line.text}`);
  }
}

function assertClean(text, where) {
  for (const bad of ['undefined', 'null', 'NaN', '[object']) {
    assert.ok(!text.includes(bad), `«${bad}» на экране (${where}): ${text.slice(Math.max(0, text.indexOf(bad) - 80), text.indexOf(bad) + 80)}`);
  }
  assert.ok(text.trim().length > 20, `пустой экран (${where})`);
}

let current = null;
async function withPage(fn) {
  const page = loadPage();
  current = page;
  try { await page.idle(); await fn(page); } finally { page.unload(); current = null; }
}
test.afterEach(() => { if (current) { current.unload(); current = null; } });

const isGet = (re) => (url, method) => method === 'GET' && re.test(url);
const isPost = (re) => (url, method) => method === 'POST' && re.test(url);
const isPut = (re) => (url, method) => method === 'PUT' && re.test(url);
const WARNING = 'Запись сохранена, но пересчитать суммы не удалось. Обновите страницу, чтобы увидеть их.';

// ---- все экраны на настоящих данных ---------------------------------------------------

live('render: every reference tab, the tender and the build screen draw real data with no placeholders', async () => {
  const s = await seed('R1');
  await withPage(async (page) => {
    const expected = seededOnTabs(s, 'R1');
    assert.deepEqual(Object.keys(expected),
      ['organizations', 'employees', 'assets', 'experience', 'services', 'profiles', 'settings'],
      'утверждение нужно на каждую вкладку справочников, а не на одну');
    for (const [tab, what] of Object.entries(expected)) {
      await page.goto(`#/refs/${tab}`);
      assertClean(page.text(), `справочник ${tab}`);
      for (const value of what.text || []) {
        assert.ok(page.text().includes(value), `на вкладке «${tab}» не видно засеянного значения «${value}»`);
      }
      for (const [label, key] of what.fields || []) {
        assert.equal(page.field(label).value, String(s.settings[key]),
          `на вкладке «${tab}» поле «${label}» показывает не то, что прислал сервер`);
      }
    }
    assert.equal(s.settings.injury_rate_percent, '0,3', 'засеянная ставка травматизма дошла до сервера');

    await page.goto(`#/tender/${s.tender.id}`);
    const text = page.text();
    assertClean(text, 'карточка тендера');
    const items = await call('GET', `${s.base}/items`);
    const labor = await call('GET', `${s.base}/labor`);
    const costs = await call('GET', `${s.base}/costs`);
    const estimate = await call('POST', `${s.base}/estimate`, {});
    for (const shown of [items.total, items.vat_total, labor.total, costs.totals.materials, costs.totals.subcontract,
      labor.rows[0].amount, items.rows[1].vat, estimate.estimate.total_wo_vat, estimate.estimate.total_with_vat,
      estimate.estimate.insurance, `по ставке ${estimate.vat_percent} %`]) {
      assert.ok(text.includes(shown), `на экране нет значения от сервера: ${shown}`);
    }
    assert.ok(text.includes('Целевая цена не задана'), 'без цели расхождение не придумывается');

    await page.goto(`#/build/${s.tender.id}`);
    assertClean(page.text(), 'сборка');
    const ticked = page.all((e) => e.tagName === 'INPUT' && e.attributes.type === 'checkbox' && e.checked);
    assert.equal(ticked.length, 2, 'отмечены документы профиля');
  });
});

live('render: the tender screen survives answers WITHOUT totals, vat percent and sums', async () => {
  const s = await seed('R2');
  await withPage(async (page) => {
    const strip = (keys) => (data) => { for (const k of keys) delete data[k]; if (data.rows) data.rows = data.rows.map((r) => { const c = { ...r }; delete c.amount; delete c.amount_value; delete c.vat; delete c.vat_value; return c; }); return data; };
    page.tweaks.push({ match: isGet(/\/items$/), edit: strip(['total', 'total_value', 'vat_total', 'vat_total_value', 'vat_percent']) });
    page.tweaks.push({ match: isGet(/\/labor$/), edit: strip(['total', 'total_value']) });
    page.tweaks.push({ match: isGet(/\/costs$/), edit: strip(['totals', 'totals_value']) });
    page.tweaks.push({ match: isPost(/\/estimate$/), edit: (d) => { delete d.vat_percent; delete d.estimate.vat; delete d.estimate.total_with_vat; return d; } });
    await page.goto(`#/tender/${s.tender.id}`);
    assertClean(page.text(), 'тендер без итогов');
    const totals = totalsOf(page.world.screen);
    assert.equal(totals.length, 4);
    assertDashes(totals, 'тендер без итогов');
    const amounts = page.all((e) => e.className.includes('computed'));
    assert.ok(amounts.length >= 6, `клеток сумм строк мало: ${amounts.length}`);
    for (const cell of amounts) {
      assert.equal(cell.textContent, DASH, 'сумма строки без данных -- прочерк, не выдумка');
      assert.ok(cell.className.includes('missing'), 'клетка с прочерком помечена как «значения нет»');
    }
    // расчёт: чего сервер не прислал -- прочерк, что прислал -- напечатано как есть
    const estimateRows = page.all((e) => e.tagName === 'TR' && e.parentNode.tagName === 'TBODY' && e.parentNode.parentNode.className === 'estimate-table');
    assert.equal(estimateRows.length, 9);
    const cells = new Map(estimateRows.map((tr) => [tr.children[0].textContent, tr.children[1].textContent]));
    const full = await call('POST', `${s.base}/estimate`, {});
    assert.equal(cells.get('НДС'), DASH, 'НДС, которого в ответе нет, -- прочерк, а не пустая клетка');
    assert.equal(cells.get('Итого с НДС'), DASH, 'итог с НДС без данных -- прочерк');
    assert.equal(cells.get('Итого без НДС'), full.estimate.total_wo_vat, 'присланный итог напечатан как есть');
    assert.equal(cells.get('Оплата труда (ФОТ)'), full.estimate.labor);
    assert.equal(cells.get('Страховые взносы'), full.estimate.insurance);
    for (const [label, value] of cells) assert.ok(value.trim() !== '', `пустая клетка расчёта «${label}» вместо прочерка`);
  });
});

live('render: saving a labor row answered with a warning keeps what arrived, shows the warning and sends nothing twice', async () => {
  const s = await seed('R3');
  await withPage(async (page) => {
    await page.goto(`#/tender/${s.tender.id}`);
    page.tweaks.push({ match: isPost(/\/labor$/), edit: (d) => { delete d.amount; delete d.amount_value; d.warning = WARNING; return d; } });
    const labor = page.card('Трудозатраты');
    const first = page.inputs(page.rows(labor)[0]);
    page.type(first[2], '41');
    const before = page.log.length;
    await page.press(page.button('Сохранить изменения', labor));
    const sent = page.log.slice(before);
    assert.deepEqual(sent.map((c) => `${c.method} ${c.url}`), [`POST ${s.base}/labor`], 'ровно одна отправка, без перезапроса и без повторной записи');
    assert.ok(page.text().includes(WARNING));
    assertDashes(totalsOf(labor), 'трудозатраты после предупреждения');
    const computed = page.rows(labor)[0].all((e) => e.className.includes('computed'))[0];
    assert.equal(computed.textContent, DASH);
    assert.ok(computed.className.includes('missing'), 'клетка с прочерком помечена как «значения нет»');
    assertClean(page.text(), 'после предупреждения');
    // в базе строка одна, а не две
    assert.equal((await call('GET', `${s.base}/labor`)).rows.length, 2);
  });
});

live('render: pasted list answered with a warning (no totals) is shown, and sent once', async () => {
  const s = await seed('R4');
  await withPage(async (page) => {
    await page.goto(`#/tender/${s.tender.id}`);
    page.tweaks.push({ match: isPut(/\/items$/), edit: (d) => ({ rows: d.rows.map((r) => { const c = { ...r }; delete c.vat; delete c.vat_value; return c; }), warning: WARNING }) });
    const lot = page.card('Позиции лота');
    const area = lot.all((e) => e.tagName === 'TEXTAREA')[0];
    area.value = 'АЗК № 4\t350 000,00\nАЗК № 6\t420 000,50\nАвтоколонна\t99 500,5';
    await page.press(page.button('Разобрать список', lot));
    const apply = page.button('Заменить список лота этим списком', lot);
    assert.equal(apply.disabled, false);
    await page.press(apply);
    assert.match(page.dialogItems(), /Объект 1/, 'вопрос называет позиции, которые будут заменены');
    const before = page.log.length;
    await page.answerDialog('ok');
    const puts = page.log.slice(before).filter((c) => c.method === 'PUT');
    assert.equal(puts.length, 1);
    assert.ok(page.text().includes(WARNING));
    assertClean(page.text(), 'после вставки с предупреждением');
    assert.equal(page.rows(lot).filter((tr) => tr.all((e) => e.tagName === 'INPUT').length === 2).length >= 3, true);
    assertDashes(totalsOf(lot), 'лот после вставки с предупреждением');
  });
});

// ---- калькулятор ------------------------------------------------------------------------

live('render: calculator always shows the gap, even zero; a missing target reads as a dash; the proposal applies and the state follows', async () => {
  const s = await seed('R5');
  await withPage(async (page) => {
    await page.goto(`#/tender/${s.tender.id}`);
    const calc = page.card('Калькулятор цены');
    const target = page.input('Целевая цена');
    page.type(target, '5 700 000');
    const reference = await call('POST', `${s.base}/estimate`, { target: '5 700 000', target_includes_vat: false });
    await page.press(page.button('Посчитать', calc));
    assert.ok(calc.textContent.includes('Расхождение'), 'расхождение показано');
    assert.ok(calc.textContent.includes(`${reference.difference} руб.`));
    assert.ok(calc.textContent.includes(reference.target));
    assertClean(calc.textContent, 'калькулятор с целью');

    // расхождение ровно ноль -- всё равно показано
    page.tweaks.push({ match: isPost(/\/estimate$/), edit: (d) => ({ ...d, difference: '0,00', difference_value: '0.00' }) });
    await page.press(page.button('Посчитать', calc));
    assert.ok(calc.textContent.includes(`0,00 руб.`) && calc.textContent.includes('итог равен цели'));
    page.tweaks.length = 0;

    // сервер не прислал цель -- прочерк, а не «undefined»
    page.tweaks.push({ match: isPost(/\/estimate$/), edit: (d) => { delete d.target; return d; } });
    await page.press(page.button('Посчитать', calc));
    assertClean(calc.textContent, 'ответ без цели');
    assert.ok(calc.textContent.includes(`Цель без НДС: ${DASH} руб.`));
    page.tweaks.length = 0;

    // подбор часов и предложение
    await page.press(page.button('Подобрать часы под цель', calc));
    assert.ok(calc.textContent.includes('Предложение: подобранные часы'));
    assert.ok(calc.textContent.includes('ещё не записаны'));
    assertClean(calc.textContent, 'предложение');
    const solved = await call('POST', `${s.base}/solve`, { target: '5 700 000', target_includes_vat: false });
    for (const row of solved.labor) assert.ok(calc.textContent.includes(row.hours));
    assert.ok(calc.textContent.includes(`по ставке ${solved.vat_percent} %`), 'ставка НДС -- из самого ответа подбора');

    // применить: ответ с предупреждением -- состояние обновляется, надпись не лжёт
    page.tweaks.push({ match: isPut(/\/labor$/), edit: (d) => ({ rows: d.rows.map((r) => { const c = { ...r }; delete c.amount; delete c.amount_value; return c; }), warning: WARNING }) });
    const before = page.log.length;
    await page.press(page.button('Применить подобранные часы', calc));
    assert.equal(page.log.slice(before).filter((c) => c.method === 'PUT').length, 1);
    assert.ok(!calc.textContent.includes('ещё не записаны'), 'часы записаны -- «ещё не записаны» больше нельзя');
    assert.ok(!calc.textContent.includes('Применить подобранные часы'), 'предложения больше нет');
    assert.ok(calc.textContent.includes(WARNING));
    assert.equal(page.button('Записать итог без НДС в цену тендера', calc).disabled, true, 'расчёта на экране нет, переносить в цену нечего');
    assertClean(page.text(), 'после применения с предупреждением');
    page.tweaks.length = 0;

    // без предупреждения: расчёт перестраивается, расхождение то же, что обещал подбор
    await page.press(page.button('Подобрать часы под цель', calc));
    await page.press(page.button('Применить подобранные часы', calc));
    assert.ok(!calc.textContent.includes('ещё не записаны'));
    assert.ok(calc.textContent.includes(`${solved.difference} руб.`));
    assert.equal(page.button('Записать итог без НДС в цену тендера', calc).disabled, false);
  });
});

// ---- сборка ---------------------------------------------------------------------------------

live('render: build screen -- real ids of the selected experience, the replace question by names, the result with skipped and warnings', async () => {
  const s = await seed('R6');
  const out = path.join(OUT_DIR, 'render-R6');
  await withPage(async (page) => {
    await page.goto(`#/build/${s.tender.id}`);
    const exp = page.card('Опыт для справки');
    const boxes = exp.all((e) => e.tagName === 'INPUT' && e.attributes.type === 'checkbox');
    assert.equal(boxes.length, 2);
    // отмечаем в порядке «второй, первый»
    for (const box of [boxes[1], boxes[0]]) { box.checked = true; box.dispatchEvent({ type: 'change', bubbles: true }); }
    page.type(page.input('Папка для пакета'), out);
    const build = page.button('Собрать пакет');
    await page.press(build);
    const result = page.text();
    assert.ok(result.includes('Справка о МТР.docx') && result.includes('Декларация соответствия.docx'));
    assertClean(result, 'результат сборки');
    const sent = page.log.filter((c) => c.method === 'POST' && c.url.endsWith('/build'))[0];
    assert.deepEqual(sent.body.experience_ids, [String(s.experience[1].id), String(s.experience[0].id)], 'идентификаторы опыта уходят строками, в порядке отметки');
    const remembered = await call('GET', `${s.base}`);
    assert.deepEqual(remembered.experience_ids, [s.experience[1].id, s.experience[0].id], 'сервер принял строки и запомнил числа');
    assert.equal(remembered.output_dir, path.normalize(out));

    // файлы уже есть: вопрос с названиями, потом замена
    await page.press(build);
    assert.match(page.dialogItems(), /Справка о МТР/);
    assert.match(page.dialogItems(), /Декларация соответствия/);
    await page.answerDialog('ok');
    assert.ok(page.text().includes('Справка о МТР.docx'));

    // результат с причинами и предупреждениями: рисуется целиком
    page.tweaks.push({ match: isPost(/\/build$/), edit: (d) => ({ ...d,
      skipped: [{ document: 'Анкета участника', reason: 'Нет шаблона «Анкета участника».' }],
      warnings: [{ document: 'Справка о МТР', text: 'Не назначена раскладка таблиц.' }] }) });
    await page.press(build);
    await page.answerDialog('ok');
    const drawn = page.text();
    assert.ok(drawn.includes('Не собрано') && drawn.includes('Нет шаблона «Анкета участника».'));
    assert.ok(drawn.includes('Предупреждения') && drawn.includes('Не назначена раскладка таблиц.'));
    assertClean(drawn, 'результат с причинами');

    // сервер не прислал ни skipped, ни warnings -- экран не падает
    page.tweaks.length = 0;
    page.tweaks.push({ match: isPost(/\/build$/), edit: (d) => ({ output_dir: d.output_dir, written: [] }) });
    await page.press(build);
    await page.answerDialog('ok');
    assertClean(page.text(), 'результат без списков');
    assert.ok(page.text().includes('Ни один документ не записан'));
    assert.equal(build.disabled, false, 'кнопка снова доступна');
  });
});

live('render: leaving the build screen mid-build and coming back keeps the button locked and still shows the result', async () => {
  const s = await seed('R7');
  const out = path.join(OUT_DIR, 'render-R7');
  await withPage(async (page) => {
    await page.goto(`#/build/${s.tender.id}`);
    page.type(page.input('Папка для пакета'), out);
    // снятие замка -- в finally: регресс замка роняет не это утверждение, а всю
    // проверку по таймауту (около 120 секунд), и о причине падения не говорит
    let release = () => {};
    try {
      page.gates.push({ match: (url) => url.endsWith('/plan'), promise: new Promise((resolve) => { release = resolve; }) });
      page.button('Собрать пакет').click();                 // сборка идёт: ждёт ответа plan
      assert.equal(page.button('Идёт сборка… не закрывайте страницу').disabled, true);

      await page.goto(`#/tender/${s.tender.id}`);           // ушли на «Карточку тендера»
      await page.goto(`#/build/${s.tender.id}`);            // и вернулись: экран собран заново
      const again = page.button('Идёт сборка… не закрывайте страницу');
      assert.equal(again.disabled, true, 'после возврата сборка всё ещё идёт, кнопка заблокирована');
      assert.equal(page.input('Папка для пакета').value, out, 'поле папки восстановлено из идущей сборки, а не пусто');
      const plansBefore = page.log.filter((c) => c.url.endsWith('/plan')).length;
      again.click();                                        // даже «силой» вторую не запустить
      await sleep(30);
      assert.equal(page.log.filter((c) => c.url.endsWith('/plan')).length, plansBefore);

      release();
      await page.idle();
      assert.ok(page.text().includes('Справка о МТР.docx'), 'результат первой сборки виден на новом экране');
      assert.equal(page.button('Собрать пакет').disabled, false);
      assert.equal(page.log.filter((c) => c.url.endsWith('/build')).length, 1, 'сборка была одна');
    } finally {
      release();
    }
  });
});

live('render: a finished build is shown when the user returns to the screen later', async () => {
  const s = await seed('R8');
  const out = path.join(OUT_DIR, 'render-R8');
  await withPage(async (page) => {
    await page.goto(`#/build/${s.tender.id}`);
    page.type(page.input('Папка для пакета'), out);
    let release = () => {};
    try {
      page.gates.push({ match: (url) => url.endsWith('/plan'), promise: new Promise((resolve) => { release = resolve; }) });
      page.button('Собрать пакет').click();
      await page.goto(`#/refs/organizations`);              // ушли совсем
      release();
      await page.idle();
      await page.goto(`#/build/${s.tender.id}`);
      assert.ok(page.text().includes('Справка о МТР.docx'), 'результат дождался возвращения пользователя');
      assert.equal(page.button('Собрать пакет').disabled, false);
    } finally {
      release();
    }
  });
});

live('render: two overlapping renders of one build screen -- the live screen keeps the button and gets the result', async () => {
  const s = await seed('R9');
  const out = path.join(OUT_DIR, 'render-R9');
  await withPage(async (page) => {
    let release = () => {};
    try {
      // первая отрисовка экрана сборки застревает на списке документов и доходит
      // до конца ПОЗЖЕ третьей: подключиться к реестру сборок она не должна
      let seen = 0;
      page.gates.push({ match: (url, method) => method === 'GET' && url.endsWith('/api/documents') && (seen += 1) === 1,
        promise: new Promise((resolve) => { release = resolve; }) });
      page.world.location.hash = `#/build/${s.tender.id}`;
      await sleep(40);
      page.world.location.hash = `#/tender/${s.tender.id}`;
      await sleep(40);
      await page.goto(`#/build/${s.tender.id}`);            // живой экран -- этот
      release();
      await page.idle();                                    // и теперь дошла первая

      page.type(page.input('Папка для пакета'), out);
      const build = page.button('Собрать пакет');
      await page.press(build);
      assert.ok(page.text().includes('Справка о МТР.docx'),
        'результат нарисован на живом экране, а не в отсоединённом от страницы');
      assert.equal(page.button('Собрать пакет').disabled, false,
        'кнопка живого экрана разблокирована, а не ждёт перехода туда и обратно');
    } finally {
      release();
    }
  });
});

live('render: build screen -- no stale warning next to a fresh result, and a validation refusal is not replayed', async () => {
  const s = await seed('R10');
  const out = path.join(OUT_DIR, 'render-R10');
  await withPage(async (page) => {
    await page.goto(`#/build/${s.tender.id}`);
    await page.press(page.button('Собрать пакет'));        // папка не заполнена
    assert.ok(page.text().includes('Укажите папку'), 'отказ назван словами');

    await page.goto(`#/tender/${s.tender.id}`);
    await page.goto(`#/build/${s.tender.id}`);
    assert.ok(!page.text().includes('Укажите папку'),
      'отказ валидации при возврате не переигрывается: человек видел бы жалобу на то, что уже исправил');

    page.type(page.input('Папка для пакета'), out);
    let release = () => {};
    try {
      page.gates.push({ match: (url) => url.endsWith('/plan'), promise: new Promise((resolve) => { release = resolve; }) });
      page.button('Собрать пакет').click();
      await sleep(30);
      page.button('Идёт сборка… не закрывайте страницу').click();   // второй щелчок «силой»
      await sleep(30);
      assert.ok(page.text().includes('уже идёт'), 'второй щелчок объяснён предупреждением');
      release();
      await page.idle();
      assert.ok(page.text().includes('Справка о МТР.docx'), 'результат нарисован');
      assert.ok(!page.text().includes('уже идёт'),
        'предупреждение прошлого действия не висит рядом со свежим результатом');
    } finally {
      release();
    }
  });
});

process.on('exit', () => {
  if (process.env.TENDER_REPORT_REQUESTS) require('node:fs').writeSync(1, `REQUESTS_SENT=${requestsSent}\n`);
});

live('render: the tender card shows the conditions the customer form requires, not a promise', async () => {
  const s = await seed('R9');
  await withPage(async (page) => {
    // список условий приходит с сервера; здесь он подменён, чтобы проверялась
    // отрисовка, а не наполнение профиля заказчика
    page.tweaks.push({ match: isGet(/\/conditions$/), edit: () => ({
      rows: [
        { condition: 'Цена и условия:', answer: 'по прилагаемой калькуляции', documents: ['ТКП по форме заказчика'] },
        { condition: 'Возможность привлечения субподрядчиков/субисполнителей', answer: '', documents: ['ТКП по форме заказчика'] },
      ],
      warnings: ['Условия документа «Справка о МТР» прочитать не удалось: карта ячеек испорчена.'],
    }) });
    await page.goto(`#/tender/${s.tender.id}`);
    const text = page.text();
    assertClean(text, 'условия сделки');
    assert.ok(text.includes('ТКП по форме заказчика'), 'не названа форма, которая требует условий');
    assert.ok(text.includes('прочитать не удалось'), 'карта, которую не разобрать, спрятана от человека');
    assert.ok(!text.includes('остаётся «подтверждаю»'), 'на экране осталось обещание умолчания');
    // ключи стоят в полях ввода: человек их ВИДИТ и правит, а не угадывает
    // дословно -- включая седьмое условие, у которого нет двоеточия на конце
    const keys = page.all((e) => e.tagName === 'INPUT' && e.attributes['aria-label'] === 'Условие');
    assert.deepEqual(keys.map((e) => e.value),
      ['Цена и условия:', 'Возможность привлечения субподрядчиков/субисполнителей'],
      'условия формы на экране не те, что прислал сервер');
    const answers = page.all((e) => e.tagName === 'INPUT' && e.attributes['aria-label'] === 'Ответ');
    assert.equal(answers[0].value, 'по прилагаемой калькуляции', 'нынешний ответ не подставлен');
    assert.equal(answers[1].value, '', 'ответ за человека не придуман');
  });
});

live('render: creating a tender whose number is taken says so on the new tender, and the two rows differ in the list', async () => {
  const s = await seed('D1');
  await withPage(async (page) => {
    // первый тендер этого номера уже лежит в базе (его завёл seed): заводим второй через экран
    await page.goto('#/tender/new');
    page.type(page.field('Номер закупки'), s.tender.purchase_number);
    page.type(page.field('Предмет закупки'), s.tender.subject);
    const customer = page.field('Заказчик');
    customer.value = String(s.customer.id);
    customer.dispatchEvent({ type: 'change', bubbles: true });
    await page.press(page.button('Создать тендер'));

    const all = await call('GET', '/api/tenders');
    const mine = all.filter((t) => t.purchase_number === s.tender.purchase_number);
    assert.equal(mine.length, 2, 'второй тендер с этим номером создан (это предупреждение, а не отказ)');
    const second = mine.find((t) => t.id !== s.tender.id);

    // экран сменился на новый тендер, и предупреждение переехало вместе с ним
    assert.equal(page.world.location.hash, `#/tender/${second.id}`);
    const text = page.text();
    assertClean(text, 'новый тендер-двойник');
    assert.ok(text.includes(`запись № ${s.tender.id}`), 'не названа запись, с которой совпал номер');
    assert.ok(text.includes(s.tender.purchase_number) && text.includes(s.tender.subject),
      'предупреждение не называет номер и предмет');
    const boxes = page.all((e) => e.className.includes('notice') && e.className.includes('warning'));
    assert.ok(boxes.length, 'предупреждение выведено не как предупреждение');

    // повторное сохранение того же тендера: двойник по-прежнему назван
    await page.press(page.button('Сохранить тендер'));
    assert.ok(page.text().includes(`запись № ${s.tender.id}`), 'при сохранении существующего тендера двойник не назван');

    // в списке строки различимы, и это видно без щелчка по каждой
    const options = page.all((e) => e.tagName === 'OPTION' && e.value && e.textContent.includes(s.tender.purchase_number));
    assert.equal(options.length, 2, options.map((o) => o.textContent).join(' | '));
    assert.notEqual(options[0].textContent, options[1].textContent, 'две строки списка совпадают буква в букву');
    assert.ok(options.some((o) => o.textContent.includes(`запись № ${s.tender.id}`)));
    assert.ok(options.some((o) => o.textContent.includes(`запись № ${second.id}`)));
  });
});

live('render: a tender whose number is free is created with no warning', async () => {
  const s = await seed('D2');
  await withPage(async (page) => {
    await page.goto('#/tender/new');
    page.type(page.field('Номер закупки'), 'N-D2-свободный');
    await page.press(page.button('Создать тендер'));
    const text = page.text();
    assertClean(text, 'новый тендер без двойника');
    assert.ok(!text.includes('уже записан'), 'предупреждение о двойнике без двойника');
    assert.equal(page.all((e) => e.className.includes('notice') && e.className.includes('warning')).length, 0);
  });
});
