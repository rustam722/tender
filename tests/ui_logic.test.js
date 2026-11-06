// Логика экрана без страницы. Запуск: node tests/ui_logic.test.js
// (из pytest -- tests/test_ui_smoke.py). Только встроенные модули Node: ни npm, ни сборки.
//
// Часть 1 -- подставной fetch: что уходит на сервер и как разбирается ответ.
// Часть 2 -- только при TENDER_BASE: те же функции против настоящего сервера.
// Здесь нет ничего про вид и удобство: страница в браузере сюда не входит.
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

const app = require(path.join(__dirname, '..', 'app', 'web', 'static', 'app.js'));
const realFetch = globalThis.fetch;
const NBSP = '\u00a0';
let requestsSent = 0;

function fakeFetch(handler) {
  const calls = [];
  globalThis.fetch = async (url, init) => {
    calls.push({ url, init, body: init && init.body ? JSON.parse(init.body) : undefined });
    const answer = await handler(url, init, calls);
    return {
      ok: answer.status === undefined || (answer.status >= 200 && answer.status < 300),
      status: answer.status || 200,
      text: async () => (answer.raw !== undefined ? answer.raw : JSON.stringify(answer.json ?? {})),
    };
  };
  return calls;
}
test.afterEach(() => { globalThis.fetch = realFetch; });

// ---- api ------------------------------------------------------------------------

test('POST and PUT carry Content-Type: application/json, even with an empty body', async () => {
  const calls = fakeFetch(() => ({ json: { ok: true } }));
  await app.api('POST', '/api/x', { a: 1 });
  await app.api('POST', '/api/x');
  await app.api('PUT', '/api/x', { items: [] });
  assert.equal(calls.length, 3);
  for (const call of calls) {
    assert.equal(call.init.headers['Content-Type'], 'application/json');
    assert.ok(call.init.body !== undefined, 'у POST/PUT всегда есть тело');
  }
  assert.deepEqual(calls[1].body, {});
});

test('GET and DELETE send no body', async () => {
  const calls = fakeFetch(() => ({ json: {} }));
  await app.api('GET', '/api/x');
  await app.api('DELETE', '/api/x/1');
  for (const call of calls) assert.equal(call.init.body, undefined);
});

test('a refusal is shown as the server\'s Russian sentence, not as a status code', async () => {
  fakeFetch(() => ({ status: 400, json: { error: 'Поле «Цена»: введите число.', field: 'price', path: ['price'] } }));
  await assert.rejects(app.api('POST', '/api/x', {}), (error) => {
    assert.equal(error.message, 'Поле «Цена»: введите число.');
    assert.equal(error.status, 400);
    assert.deepEqual(error.path, ['price']);
    return true;
  });
});

test('a reply without a sentence still reads as Russian prose, never a bare code', async () => {
  for (const answer of [{ status: 500, raw: '<html>boom</html>' }, { status: 502, raw: '' },
                        { status: 404, raw: 'null' }]) {
    fakeFetch(() => answer);
    await assert.rejects(app.api('GET', '/api/x'), (error) => {
      assert.match(error.message, /[А-Яа-я]{4,}/);
      assert.doesNotMatch(error.message, /boom/);
      return true;
    });
  }
});

test('no connection reads as Russian prose', async () => {
  globalThis.fetch = async () => { throw new TypeError('Failed to fetch'); };
  await assert.rejects(app.api('GET', '/api/x'), (error) => {
    assert.match(error.message, /программ/);
    assert.doesNotMatch(error.message, /Failed/);
    return true;
  });
});

// ---- запись ---------------------------------------------------------------------

test('writeRecord sends the whole record once and strips the warning', async () => {
  const calls = fakeFetch(() => ({ status: 201, json: { id: 5, name: 'АЗК № 4', price: `1${NBSP}000,00`,
    price_value: '1000.00', warning: 'Запись сохранена, но пересчитать суммы не удалось.' } }));
  const sent = { name: 'АЗК № 4', price: `1${NBSP}000,00`, price_value: '1000.00' };
  const result = await app.writeRecord('/api/tenders/3/items', sent);
  assert.equal(calls.length, 1, 'запись сохранена -- повторно не отправляется');
  assert.deepEqual(calls[0].body, sent);
  assert.equal(result.warning, 'Запись сохранена, но пересчитать суммы не удалось.');
  assert.equal('warning' in result.record, false);
  assert.equal(result.record.id, 5);
  assert.equal('amount' in result.record, false, 'недостающую сумму экран не придумывает');
});

test('recordForWrite keeps everything that came back and overlays only the edits', () => {
  const server = { id: 2, price: `10,01`, price_value: '10.005', warning: 'x', name: 'А' };
  const body = app.recordForWrite(server, { price: '20' });
  assert.deepEqual(body, { id: 2, price: '20', price_value: '10.005', name: 'А' });
  assert.equal('warning' in body, false);
  assert.equal(server.price, '10,01', 'исходная запись не портится');
});

// ---- вставка списком ---------------------------------------------------------------

test('paste from Excel: tab-separated name and price are kept as text, untouched', () => {
  const { rows } = app.parsePasted(`АЗК № 4\t350${NBSP}000,00\r\nАЗК № 6\t420000\r\n\r\nАвтоколонна\t99 500,50\r\n`);
  assert.deepEqual(rows.map((r) => [r.name, r.price, r.problem]), [
    ['АЗК № 4', `350${NBSP}000,00`, ''],
    ['АЗК № 6', '420000', ''],
    ['Автоколонна', '99 500,50', ''],
  ]);
});

test('paste: a leading ordinal column is dropped', () => {
  const { rows } = app.parsePasted('1\tАЗК № 4\t350 000,00\n2.\tАЗК № 6\t420000');
  assert.deepEqual(rows.map((r) => [r.name, r.price]), [['АЗК № 4', '350 000,00'], ['АЗК № 6', '420000']]);
});

test('paste: semicolons work when there are no tabs', () => {
  const { rows } = app.parsePasted('Разработка проекта; 5 700 000,00');
  assert.deepEqual(rows.map((r) => [r.name, r.price, r.problem]), [['Разработка проекта', '5 700 000,00', '']]);
});

test('paste: a line without a clear separator is NOT guessed -- a wrong price is worse than a question', () => {
  const { rows } = app.parsePasted('Разработка проекта 5 700 000,00');
  assert.equal(rows.length, 1);
  assert.equal(rows[0].price, '');
  assert.match(rows[0].problem, /цен/);
  assert.equal(rows[0].name, 'Разработка проекта 5 700 000,00');
});

test('paste: a name with no price is flagged, a price with no name is flagged', () => {
  const { rows } = app.parsePasted('Только имя\n\t5000');
  assert.match(rows[0].problem, /цен/);
  assert.match(rows[1].problem, /наименован/);
});

test('paste: more than two columns is flagged, not silently dropped', () => {
  const { rows } = app.parsePasted('АЗК\tшт.\t5\t350000');
  assert.equal(rows.length, 1);
  assert.match(rows[0].problem, /колон/);
});

test('empty paste gives no rows', () => {
  assert.deepEqual(app.parsePasted('  \n\r\n').rows, []);
});

test('lotPayload sends rows in order without computing numbers', () => {
  const body = app.lotPayload([{ name: 'А', price: '1' }, { name: 'Б', price: '2,50' }]);
  assert.deepEqual(body, { items: [{ position: 0, name: 'А', price: '1' }, { position: 1, name: 'Б', price: '2,50' }] });
});

// ---- расхождение -------------------------------------------------------------------

test('differenceKind reads the sign of the server string, kopecks included', () => {
  assert.equal(app.differenceKind('0.17'), 'above');
  assert.equal(app.differenceKind('-0.17'), 'below');
  assert.equal(app.differenceKind('0.00'), 'equal');
  assert.equal(app.differenceKind('-0.00'), 'equal');
  assert.equal(app.differenceKind('0'), 'equal');
  assert.equal(app.differenceKind('12.00'), 'above');
  assert.equal(app.differenceKind('0.01'), 'above');
});

// ---- документы и опыт -----------------------------------------------------------------

test('document checkboxes: the customer\'s profile is ticked, other profiles\' documents are offered unticked', () => {
  const profiles = [
    { customer_id: 2, documents: ['Справка о МТР', 'Декларация соответствия'],
      docs: { 'Справка о МТР': {}, 'Письмо о крупности сделки': {} } },
    { customer_id: 9, documents: ['Анкета участника', 'Справка о МТР'], docs: {} },
  ];
  const choices = app.documentChoices(profiles, { customer_id: 2 });
  assert.deepEqual(choices.map((c) => [c.name, c.checked]), [
    ['Справка о МТР', true], ['Декларация соответствия', true],
    ['Письмо о крупности сделки', false], ['Анкета участника', false],
  ]);
});

test('document checkboxes: no profile -- nothing ticked, nothing invented', () => {
  assert.deepEqual(app.documentChoices([], { customer_id: 2 }), []);
  const other = [{ customer_id: 9, documents: ['Анкета участника'], docs: {} }];
  assert.deepEqual(app.documentChoices(other, { customer_id: 2 }).map((c) => c.checked), [false]);
});

test('the tender\'s service is guessed from the subject only when it is unambiguous', () => {
  const services = [{ id: 1, code: 'КЭР', name: 'Комплексный экологический отчёт' },
                    { id: 2, code: 'СЗЗ', name: 'Санитарно-защитная зона' }];
  assert.equal(app.inferServiceId('Разработка КЭР для АЗК', services), '1');
  assert.equal(app.inferServiceId('Проект СЗЗ и КЭР', services), '');
  assert.equal(app.inferServiceId('Услуги', services), '');
  assert.equal(app.inferServiceId('Разработка проекта санитарно-защитная зона АЗК', services), '2');
  assert.equal(app.inferServiceId('ПОКЭРНЫЙ', services), '', 'код внутри слова -- не совпадение');
});

test('toggleId keeps the order the user ticked and never duplicates', () => {
  let ids = app.toggleId(['3', '1'], '7', true);
  assert.deepEqual(ids, ['3', '1', '7']);
  assert.deepEqual(app.toggleId(ids, '7', true), ['3', '1', '7']);
  assert.deepEqual(app.toggleId(ids, '1', false), ['3', '7']);
});

// ---- сборка: замок, вопрос о замене, итог --------------------------------------------

function builderWith(handler, overrides = {}) {
  const calls = fakeFetch(handler);
  const asked = [];
  const busy = [];
  const builder = app.createBuilder({
    api: app.api,
    confirmReplace: async (files) => { asked.push(files); return overrides.answer ?? true; },
    setBusy: (value) => busy.push(value),
  });
  return { builder, calls, asked, busy };
}
const REQUEST = { tenderId: 7, documents: ['Справка о МТР', 'Декларация соответствия'],
                  outputDir: 'C:\\Тендеры\\Заявка', experienceIds: ['3', '4'] };
const planOf = (existing) => ({ output_dir: 'C:\\Тендеры\\Заявка', files: [
  { document: 'Справка о МТР', name: 'Справка о МТР.docx', path: 'x', exists: existing.includes('Справка о МТР') },
  { document: 'Декларация соответствия', name: 'Декларация соответствия.docx', path: 'y',
    exists: existing.includes('Декларация соответствия') }] });
const builtOf = { output_dir: 'C:\\Тендеры\\Заявка', written: [{ name: 'a.docx', path: 'p' }],
                  skipped: [{ document: 'Справка о МТР', reason: 'Нет шаблона.' }], warnings: [{ document: 'Декларация соответствия', text: 'Пустое поле.' }] };

function buildRoutes(existing, record = []) {
  return (url, init) => {
    if (url.endsWith('/plan')) return { json: planOf(existing) };
    if (url.endsWith('/build')) return { json: builtOf };
    if (url === '/api/tenders') { record.push(1); return { json: { id: 7 } }; }
    return { status: 404, json: { error: 'Нет такого адреса.' } };
  };
}

test('build: plans first, then builds; nothing exists -- no question, overwrite is false', async () => {
  const { builder, calls, asked } = builderWith(buildRoutes([]));
  const outcome = await builder.run(REQUEST);
  assert.equal(outcome.status, 'built');
  assert.equal(asked.length, 0);
  assert.deepEqual(calls.map((c) => c.url), ['/api/tenders/7/plan', '/api/tenders/7/build', '/api/tenders']);
  assert.deepEqual(calls[1].body, { documents: REQUEST.documents, output_dir: REQUEST.outputDir,
    overwrite: false, experience_ids: ['3', '4'] });
  for (const call of calls) assert.equal(call.init.headers['Content-Type'], 'application/json');
});

test('build: the replace question names the documents, and only the ones that exist', async () => {
  const { builder, calls, asked } = builderWith(buildRoutes(['Справка о МТР']));
  const outcome = await builder.run(REQUEST);
  assert.equal(outcome.status, 'built');
  assert.equal(asked.length, 1);
  assert.deepEqual(asked[0].map((f) => [f.document, f.name]), [['Справка о МТР', 'Справка о МТР.docx']]);
  assert.equal(calls[1].body.overwrite, true);
});

test('build: declining the question builds nothing and unlocks the button', async () => {
  const { builder, calls, busy } = builderWith(buildRoutes(['Справка о МТР']), { answer: false });
  const outcome = await builder.run(REQUEST);
  assert.equal(outcome.status, 'cancelled');
  assert.deepEqual(calls.map((c) => c.url), ['/api/tenders/7/plan']);
  assert.deepEqual(busy, [true, false]);
});

test('build: a second click while the first build runs starts nothing', async () => {
  let release;
  const gate = new Promise((resolve) => { release = resolve; });
  const { builder, calls, busy } = builderWith(async (url) => {
    if (url.endsWith('/plan')) { await gate; return { json: planOf([]) }; }
    return buildRoutes([])(url);
  });
  const first = builder.run(REQUEST);
  assert.deepEqual(busy, [true], 'кнопка блокируется сразу, до первого ответа сервера');
  const second = await builder.run(REQUEST);
  assert.equal(second.status, 'busy');
  assert.equal(calls.length, 1, 'вторая сборка не отправила ни одного запроса');
  release();
  assert.equal((await first).status, 'built');
  assert.deepEqual(busy, [true, false]);
  assert.equal((await builder.run(REQUEST)).status, 'built', 'после окончания -- снова можно');
});

test('build: the button unlocks after a refusal, and the server sentence is kept', async () => {
  const { builder, busy } = builderWith(() => ({ status: 400, json: { error: 'Папка шаблонов не найдена.' } }));
  const outcome = await builder.run(REQUEST);
  assert.equal(outcome.status, 'failed');
  assert.equal(outcome.error, 'Папка шаблонов не найдена.');
  assert.deepEqual(busy, [true, false]);
});

test('build: skipped reasons and warnings reach the result untouched', async () => {
  const { builder } = builderWith(buildRoutes([]));
  const outcome = await builder.run(REQUEST);
  assert.deepEqual(outcome.result.skipped, builtOf.skipped);
  assert.deepEqual(outcome.result.warnings, builtOf.warnings);
  assert.deepEqual(outcome.result.written, builtOf.written);
});

test('build: nothing ticked or no folder -- Russian refusal, no request at all', async () => {
  const { builder, calls, busy } = builderWith(buildRoutes([]));
  const none = await builder.run({ ...REQUEST, documents: [] });
  assert.equal(none.status, 'failed');
  assert.match(none.error, /документ/);
  const noFolder = await builder.run({ ...REQUEST, outputDir: '  ' });
  assert.match(noFolder.error, /папк/);
  assert.equal(calls.length, 0);
  assert.deepEqual(busy, [true, false, true, false]);
});

test('build: a built package whose folder could not be remembered still reports success, with a note', async () => {
  const { builder } = builderWith((url) => url === '/api/tenders'
    ? { status: 400, json: { error: 'Не сохранено.' } } : buildRoutes([])(url));
  const outcome = await builder.run(REQUEST);
  assert.equal(outcome.status, 'built');
  assert.match(outcome.note, /Не сохранено/);
});

// ---- настоящий сервер ---------------------------------------------------------------------

const BASE = process.env.TENDER_BASE;
const live = BASE ? test : test.skip;

function useRealServer() {
  globalThis.fetch = (url, init) => { requestsSent += 1; return realFetch(BASE + url, init); };
}

live('live: whole-record round trip keeps precision, a typed edit wins, a percent goes as percent', async () => {
  useRealServer();
  const made = await app.api('POST', '/api/employees', { full_name: 'Иванов И. И.', position: 'Инженер',
    hourly_rate: '10.005', education: '', specialty: '', qualification: '' });
  assert.equal(made.hourly_rate, '10,01');
  assert.equal(made.hourly_rate_value, '10.005');

  // ничего не трогали -- ушла вся запись, и копейки не округлились
  const same = await app.writeRecord('/api/employees', app.recordForWrite(made, { position: 'Главный инженер' }));
  assert.equal(same.record.hourly_rate_value, '10.005');
  assert.equal(same.record.position, 'Главный инженер');

  // правка не тонет в _value
  const edited = await app.writeRecord('/api/employees', app.recordForWrite(same.record, { hourly_rate: '20' }));
  assert.equal(edited.record.hourly_rate_value, '20');

  // ставка отправляется процентами, делить на сто экрану не нужно
  const org = await app.api('POST', '/api/organizations', { role: 'customer', short_name: 'МУП «Полигон»',
    inn: '2904016130', vat_rate_percent: '0,7' });
  assert.equal(org.vat_rate, '0.007');
  assert.equal(org.vat_rate_percent, '0,7');
  const back = await app.writeRecord('/api/organizations', app.recordForWrite(org, { short_name: 'МУП «Полигон-2»' }));
  assert.equal(back.record.vat_rate, '0.007', 'эхо ставки не теряет знаков');
  const cleared = await app.writeRecord('/api/organizations', app.recordForWrite(back.record, { vat_rate_percent: '' }));
  assert.equal(cleared.record.vat_rate, null);
});

live('live: a refusal arrives as the server\'s Russian sentence', async () => {
  useRealServer();
  await assert.rejects(app.api('POST', '/api/employees', { hourly_rate: 'много' }), (error) => {
    assert.match(error.message, /число/);
    assert.equal(error.status, 400);
    return true;
  });
  await assert.rejects(app.api('GET', '/api/organizations/9999'), (error) => {
    assert.match(error.message, /не найдена/);
    assert.equal(error.status, 404);
    return true;
  });
});

live('live: paste replaces the lot in one go; a bad row names itself and nothing is written', async () => {
  useRealServer();
  const me = await app.api('POST', '/api/organizations', { role: 'participant', short_name: 'ООО «Техносфера»', inn: '5904188041' });
  const customer = await app.api('POST', '/api/organizations', { role: 'customer', short_name: 'ПАО «Заказчик»', inn: '7708514824' });
  const tender = await app.api('POST', '/api/tenders', { customer_id: String(customer.id), participant_id: String(me.id),
    subject: 'Разработка КЭР', purchase_date: '2025-10-24', price_wo_vat: '0' });

  const pasted = app.parsePasted(`Объект 1\t350${NBSP}000,00\nОбъект 2\t420 000,50\n`);
  const first = await app.api('PUT', `/api/tenders/${tender.id}/items`, app.lotPayload(pasted.rows));
  assert.equal(first.rows.length, 2);
  assert.equal(first.rows[0].price, `350${NBSP}000,00`);
  assert.equal(first.total, `770${NBSP}000,50`, 'итог пришёл с сервера');
  assert.equal(first.vat_percent, '5');
  assert.match(first.vat_total, /^\d/);

  const bad = app.parsePasted('Раз\t100\nДва\tдорого\nТри\t300');
  await assert.rejects(app.api('PUT', `/api/tenders/${tender.id}/items`, app.lotPayload(bad.rows)), (error) => {
    assert.match(error.message, /Позиция №2/);
    assert.deepEqual(error.path.slice(0, 2), ['items', 1], 'экран знает, какую строку подсветить');
    return true;
  });
  const kept = await app.api('GET', `/api/tenders/${tender.id}/items`);
  assert.equal(kept.rows.length, 2, 'прежний список цел');
});

live('live: labor, estimate with a target, solve, apply, and the gap is always shown', async () => {
  useRealServer();
  const tenders = await app.api('GET', '/api/tenders');
  const tender = tenders[tenders.length - 1];
  const base = `/api/tenders/${tender.id}`;

  const row = await app.writeRecord(`${base}/labor`, { post: 'Инженер', headcount: '2', hours: '100,5', rate: '1 000' });
  assert.equal(row.record.hours, '100,5');
  assert.equal(row.record.amount, `201${NBSP}000,00`);
  assert.equal(row.warning, undefined);
  await app.writeRecord(`${base}/costs`, { kind: 'materials', name: 'Бумага', unit: 'пач.', quantity: '10', price: '350' });
  await app.writeRecord(`${base}/costs`, { kind: 'subcontract', name: 'Лаборатория', unit: 'усл.', quantity: '1', price: '12 000' });
  const costs = await app.api('GET', `${base}/costs`);
  assert.equal(costs.rows.length, 2);
  assert.match(costs.totals.materials, /^3[ \u00a0]?500,00$/);

  const direct = await app.api('POST', `${base}/estimate`, {});
  assert.equal(direct.target, undefined, 'цели нет -- сравнивать не с чем');
  assert.match(direct.estimate.total_wo_vat, /\d,\d\d$/);

  const solved = await app.api('POST', `${base}/solve`, { target: '1 000 000', target_includes_vat: false });
  assert.equal(typeof solved.difference, 'string');
  assert.ok(['above', 'below', 'equal'].includes(app.differenceKind(solved.difference_value)));
  assert.equal(solved.labor.length, 1);

  // применить предложение: запись с id возвращается как получена
  const applied = await app.writeRecord(`${base}/labor`, app.recordForWrite(solved.labor[0], {}));
  assert.equal(applied.warning, undefined);
  assert.equal(applied.record.hours_value, solved.labor[0].hours_value, 'часы не потеряли знаков');

  const gap = await app.api('POST', `${base}/estimate`, { target: '1 000 000', target_includes_vat: false });
  assert.equal(gap.difference, solved.difference);
  assert.equal(gap.estimate.total_wo_vat, solved.estimate.total_wo_vat);

  const withVat = await app.api('POST', `${base}/solve`, { target: '1 050 000', target_includes_vat: true });
  assert.equal(withVat.target_includes_vat, true);
  assert.equal(typeof withVat.difference, 'string');
});

live('live: build asks about replacement by document names, locks, and remembers the folder', async () => {
  useRealServer();
  const tenders = await app.api('GET', '/api/tenders');
  const tender = tenders[tenders.length - 1];
  const outDir = process.env.TENDER_OUT_DIR;
  const profile = await app.writeRecord('/api/profiles', { customer_id: String(tender.customer_id),
    documents: ['Справка о МТР', 'Декларация соответствия'], docs: {}, declaration_text: '' });
  assert.equal(profile.warning, undefined);

  const profiles = await app.api('GET', '/api/profiles');
  const choices = app.documentChoices(profiles, tender);
  assert.deepEqual(choices.filter((c) => c.checked).map((c) => c.name), ['Справка о МТР', 'Декларация соответствия']);

  const asked = [];
  const busy = [];
  const builder = app.createBuilder({ api: app.api, setBusy: (v) => busy.push(v),
    confirmReplace: async (files) => { asked.push(files.map((f) => f.document)); return true; } });
  const request = { tenderId: tender.id, documents: choices.filter((c) => c.checked).map((c) => c.name),
                    outputDir: outDir, experienceIds: [] };

  const [one, two] = await Promise.all([builder.run(request), builder.run(request)]);
  assert.deepEqual([one.status, two.status].sort(), ['built', 'busy'], 'двойной щелчок -- одна сборка');
  assert.equal(one.status === 'built' ? one.result.written.length : two.result.written.length, 2);
  assert.equal(asked.length, 0, 'папка была пуста -- вопроса нет');

  const again = await builder.run(request);
  assert.equal(again.status, 'built');
  assert.deepEqual(asked, [['Справка о МТР', 'Декларация соответствия']], 'вопрос назвал документы');
  assert.equal(again.result.written.length, 2);
  assert.deepEqual(again.result.skipped, []);
  assert.deepEqual(busy.slice(-2), [true, false]);

  const remembered = await app.api('GET', `/api/tenders/${tender.id}`);
  assert.equal(remembered.output_dir, path.normalize(outDir), 'папка запомнена в тендере');

  // без разрешения на замену файл остаётся нетронутым и причина названа
  const refusal = await app.api('POST', `/api/tenders/${tender.id}/build`, {
    documents: request.documents, output_dir: outDir, overwrite: false });
  assert.equal(refusal.written.length, 0);
  assert.equal(refusal.skipped.length, 2);
  assert.match(refusal.skipped[0].reason, /уже есть/);
});

process.on('exit', () => {
  if (process.env.TENDER_REPORT_REQUESTS) require('node:fs').writeSync(1, `REQUESTS_SENT=${requestsSent}\n`);
});
