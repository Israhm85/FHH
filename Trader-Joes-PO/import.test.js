'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const now = new Date('2026-11-23T20:00:00Z');

function body(count = 1, po = '0160295112') {
  const items = Array.from({ length: count }, (_, i) =>
    `PRODUCT NAME : ${i ? 'PRODUCT ' + i : 'HOL SIDE CORN PUDDING'}\nSKU : ${151493 + i} 8 16 Wgt Oz BOX\nCase QTY : ${i + 1}\nCOST : $ 31.20`).join('\n');
  return `PURCHASE AGREEMENT\nP.O. : ${po}\nDELIVER TO : WCD PALMDALE 5028 / 5028-PALMDALE XDOCK\nARRIVE : 11/28/26\n${items}\nTOTALS : $ 32.37`;
}

function load() {
  const context = vm.createContext({ Date, Number, JSON, Error,
    SpreadsheetApp: { BorderStyle: { SOLID: 'SOLID' } },
    Utilities: {
      formatDate(date, timezone) {
        return new Intl.DateTimeFormat('en-CA', { timeZone: timezone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(date);
      },
      parseDate(value) { return new Date(value.replace(' ', 'T') + ':00-08:00'); }
    } });
  vm.runInContext(fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8'), context);
  return context;
}

function setup(context, messages, existing = [
  ['', '', '', '', 'Regular', '', '', '', ''],
  ['', '', '', '', 'Regular', '', '', '', '']
]) {
  const header = ['PROD_DATE', 'ARRIVE_DATE', 'LOCATION', 'PO_NUMBER', 'ORDER_TYPE', 'PRODUCT', 'QTY', 'UOM', 'NOTES'];
  const cells = [header, ...existing];
  const writes = [];
  const formatting = [];
  const queries = [];
  let maxRows = 100;
  const sheet = {
    getLastRow: () => cells.length,
    getMaxRows: () => maxRows,
    getRowHeight: row => row === 2 ? 26 : 27,
    setRowHeight: (row, value) => formatting.push({ kind: 'height', row, value }),
    insertRowsAfter: (_, n) => { maxRows += n; },
    getRange(row, col, numRows, numCols) {
      return {
        row, col, numRows, numCols,
        getValues: () => cells.slice(row - 1, row - 1 + numRows).map(r => r.slice(col - 1, col - 1 + numCols)),
        getDisplayValues: () => cells.slice(row - 1, row - 1 + numRows).map(r => r.slice(col - 1, col - 1 + numCols).map(v =>
          v instanceof Date ? context.Utilities.formatDate(v, 'America/Los_Angeles') : String(v))),
        getNumberFormats: () => Array.from({ length: numRows }, () => Array(numCols).fill('original')),
        setNumberFormats: value => { formatting.push({ kind: 'numberFormats', row, value }); },
        copyTo: (target, options) => { formatting.push({ kind: 'copy', source: row, row: target.row, options }); },
        setBackgrounds: value => { formatting.push({ kind: 'backgrounds', row, value }); },
        setHorizontalAlignments: value => { formatting.push({ kind: 'alignments', row, value }); },
        setVerticalAlignment: value => { formatting.push({ kind: 'vertical', row, value }); },
        setBorder: (...args) => { formatting.push({ kind: 'border', row, args }); },
        setFontWeight: weight => { formatting.push({ kind: 'weight', row, weight }); },
        setNumberFormat: () => {},
        setValues: rows => { writes.push(...rows); cells.push(...rows); }
      };
    }
  };
  const ss = { getSheetByName: name => name === 'Datos' ? sheet : null, getSpreadsheetTimeZone: () => 'America/Los_Angeles' };
  context.GmailApp = { search: (query, offset, count) => {
    queries.push(query);
    return messages.slice(offset, offset + count).map((message, i) => ({
      getMessages: () => [{
        getDate: () => message.date ? new Date(message.date) : now,
        getPlainBody: () => message.text || message,
        getId: () => String(offset + i)
      }]
    }));
  } };
  return { ss, sheet, cells, writes, formatting, queries };
}

test('example is a one-case demo and dates cross month boundaries correctly', () => {
  const c = load();
  const order = c.tjParseOrder_(body());
  assert.equal(order.po, '0160295112');
  assert.equal(order.type, 'Demo');
  assert.equal(order.location, 'PALMDALE');
  assert.equal(order.items[0].product, 'Holiday Side Corn Pudding');
  assert.equal(order.items[0].qty, 1);
  assert.equal(order.items[0].uom, 'CASES');
  assert.equal(c.tjShiftDate_('2026-03-01', -1), '2026-02-28');
  assert.equal(c.tjShiftDate_('2024-03-01', -1), '2024-02-29');
});

test('seven products are demo and eight are regular; each quantities are units', () => {
  const c = load();
  assert.equal(c.tjParseOrder_(body(7)).type, 'Demo');
  const order = c.tjParseOrder_(body(8).replace('Case QTY : 1', 'Each QTY : 1,225'));
  assert.equal(order.type, 'Regular');
  assert.equal(order.items[0].qty, 1225);
  assert.equal(order.items[0].uom, 'UNITS');
});

test('rejects invalid or incomplete orders rather than misclassifying them', () => {
  const c = load();
  for (const invalid of [
    body().replace('11/28/26', '02/30/26'),
    body().replace('Case QTY : 1', 'Case QTY : 0'),
    body().replace('Case QTY : 1', 'Case QTY : 1.5'),
    body().replace('Case QTY : 1', 'Case QTY : 1,2'),
    body().replace('SKU : 151493', 'MISSING : 151493'),
    body(2).replace('PRODUCT NAME : PRODUCT 1', 'PRODUCT NAME PRODUCT 1'),
    body(2).replace('SKU : 151494', 'SKU : 151493'),
    body().replace('PALMDALE 5028 / 5028-PALMDALE', 'UNKNOWN'),
    body().replace('HOL SIDE CORN PUDDING', '=1+1'),
    body() + '\nP.O. : 0160295113'
  ]) assert.throws(() => c.tjParseOrder_(invalid));
});

test('imports once, preserves PO leading zero and warns on changed existing PO', () => {
  const c = load();
  const { ss, cells, writes } = setup(c, [body()]);
  assert.match(c.tjImport_(ss, '2026-11-28', '2026-11-28', now), /Filas nuevas: 1/);
  assert.equal(writes[0][3], '0160295112');
  assert.equal(c.Utilities.formatDate(writes[0][0], 'America/Los_Angeles'), '2026-11-27');
  assert.equal(c.Utilities.formatDate(writes[0][1], 'America/Los_Angeles'), '2026-11-28');
  assert.match(c.tjImport_(ss, '2026-11-28', '2026-11-28', now), /sin cambios: 1/);
  assert.equal(writes.length, 1);
  cells[3][6] = 9;
  assert.match(c.tjImport_(ss, '2026-11-28', '2026-11-28', now), /diferencias \(no modificados\): 0160295112/);
  assert.equal(cells[3][6], 9);
  assert.equal(writes.length, 1);
});

test('formats new regular and demo rows using alternating templates without altering values', () => {
  const c = load();
  const { ss, writes, formatting } = setup(c, [body(8, '0160295111'), body()]);
  c.tjImport_(ss, '2026-11-28', '2026-11-28', now);
  assert.equal(writes.length, 9);
  assert.deepEqual(formatting.filter(f => f.kind === 'copy').map(f => [f.source, f.row]),
    [[2, 4], [3, 5], [2, 6], [3, 7], [2, 8], [3, 9], [2, 10], [3, 11], [2, 12]]);
  const colors = formatting.find(f => f.kind === 'backgrounds').value;
  assert.deepEqual(Array.from(colors, row => row[0]), [
    '#f8fafc', '#ffffff', '#f8fafc', '#ffffff', '#f8fafc', '#ffffff', '#f8fafc', '#ffffff', '#fef3c9'
  ]);
  assert.deepEqual(Array.from(formatting.find(f => f.kind === 'alignments').value[0]),
    ['center', 'center', 'center', 'center', 'center', 'left', 'right', 'left', 'left']);
  assert.equal(formatting.find(f => f.kind === 'border').args[6], '#cbd5e0');
  assert.equal(formatting.find(f => f.kind === 'vertical').value, 'middle');
  assert.deepEqual(formatting.filter(f => f.kind === 'weight').map(f => [f.row, f.weight]), [[12, 'bold']]);
  assert.deepEqual(formatting.filter(f => f.kind === 'height').slice(0, 2).map(f => [f.row, f.value]), [[4, 26], [5, 27]]);
});

test('existing rows retain values and number formats while restoring alternating colors', () => {
  const c = load();
  const original = [
    ['', '', '', '', 'Regular', '', '', '', ''],
    ['', '', '', '', 'Regular', '', '', '', ''],
    ['date', 'date', 'PALMDALE', '000123', 'Demo', 'Corn', 1, 'CASES', 'nota'],
    ['date', 'date', 'LACEY', '000124', 'Regular', 'Corn', 9, 'CASES', '']
  ];
  const { sheet, cells, formatting } = setup(c, [], original);
  const before = JSON.stringify(cells);
  assert.equal(c.tjFormatRows_(sheet, 4, ['Demo', 'Regular'], true), true);
  assert.equal(JSON.stringify(cells), before);
  assert.deepEqual(Array.from(formatting.find(f => f.kind === 'backgrounds').value, r => r[0]), ['#fef3c9', '#ffffff']);
  assert.equal(formatting.find(f => f.kind === 'weight').row, 4);
  assert.equal(formatting.find(f => f.kind === 'numberFormats').value[0][3], 'original');
});

test('missing regular templates never copies demo formatting to other orders', () => {
  const c = load();
  const { ss, writes, formatting } = setup(c, [body()], []);
  const result = c.tjImport_(ss, '2026-11-28', '2026-11-28', now);
  assert.match(result, /No se aplicó el formato/);
  assert.equal(writes.length, 1);
  assert.equal(formatting.length, 0);
});

test('conflicting or malformed messages block all writes', () => {
  for (const messages of [[body(), body().replace('Case QTY : 1', 'Case QTY : 2')],
    [body(), body(1, '0160295113').replace('Case QTY : 1', 'Case QTY : bad')]]) {
    const c = load();
    const { ss, writes } = setup(c, messages);
    assert.throws(() => c.tjImport_(ss, '2026-11-28', '2026-11-28', now));
    assert.equal(writes.length, 0);
  }
});

test('scans all pages, deduplicates messages and ignores non-order mail', () => {
  const c = load();
  const messages = Array(101).fill(body()).concat(['No purchase order here']);
  const { ss, writes } = setup(c, messages);
  assert.match(c.tjImport_(ss, '2026-11-28', '2026-11-28', now), /Filas nuevas: 1/);
  assert.equal(writes.length, 1);
});

test('filters by inclusive ARRIVE dates while examining only the last ten calendar days', () => {
  const c = load();
  const dated = (po, arrive, date) => ({ text: body(1, po).replace('11/28/26', arrive), date });
  const messages = [
    dated('001', '11/28/26', '2026-11-14T08:00:00Z'),
    dated('002', '11/30/26', '2026-11-23T23:59:59Z'),
    dated('003', '11/27/26', '2026-11-23T20:00:00Z'),
    dated('004', '12/01/26', '2026-11-23T20:00:00Z'),
    dated('005', '11/29/26', '2026-11-14T07:59:59Z'),
    dated('006', '11/29/26', '2026-11-24T08:00:00Z')
  ];
  const { ss, writes, queries } = setup(c, messages);
  const result = c.tjImport_(ss, '2026-11-28', '2026-11-30', now);
  assert.deepEqual(writes.map(row => row[3]), ['001', '002']);
  assert.match(result, /Correos revisados del 2026-11-14 al 2026-11-23/);
  assert.match(result, /ARRIVE del 2026-11-28 al 2026-11-30/);
  assert.equal(queries[0], 'label:[TRJO] after:2026/11/13 before:2026/11/25');
});

test('does not duplicate existing POs or identical emails even when ARRIVE and receipt dates differ', () => {
  const c = load();
  const dated = date => ({ text: body(), date });
  const { ss, writes } = setup(c, [dated('2026-11-15T20:00:00Z'), dated('2026-11-23T20:00:00Z')]);
  assert.match(c.tjImport_(ss, '2026-11-28', '2026-11-28', now), /Filas nuevas: 1/);
  assert.match(c.tjImport_(ss, '2026-11-28', '2026-11-30', now), /sin cambios: 1/);
  assert.equal(writes.length, 1);
});

test('existing numeric PO without leading zero prevents duplication and preserves the existing row', () => {
  const c = load();
  const existing = [
    ['', '', '', '', 'Regular', '', '', '', ''],
    ['', '', '', '', 'Regular', '', '', '', ''],
    [new Date('2026-11-27T20:00:00Z'), new Date('2026-11-28T20:00:00Z'), 'PALMDALE',
      160295112, 'Demo', 'Holiday Side Corn Pudding', 1, 'CASES', 'nota']
  ];
  const { ss, cells, writes } = setup(c, [body()], existing);
  const before = JSON.stringify(cells);
  assert.match(c.tjImport_(ss, '2026-11-28', '2026-11-28', now), /sin cambios: 1/);
  assert.equal(writes.length, 0);
  assert.equal(JSON.stringify(cells), before);
});

test('out-of-range orders do not block import, but missing or ambiguous ARRIVE blocks it', () => {
  const c = load();
  const outOfRange = body(1, '001').replace('11/28/26', '12/01/26').replace('Case QTY : 1', 'Case QTY : bad');
  const selected = body(1, '002');
  const valid = setup(c, [outOfRange, selected]);
  c.tjImport_(valid.ss, '2026-11-28', '2026-11-28', now);
  assert.equal(valid.writes.length, 1);
  for (const invalid of [
    selected.replace('ARRIVE : 11/28/26', 'ARRIVE : 02/30/26'),
    selected.replace('ARRIVE : 11/28/26', 'ARRIVE : '),
    selected + '\nARRIVE : 11/30/26'
  ]) {
    const batch = setup(c, [selected, invalid]);
    assert.throws(() => c.tjImport_(batch.ss, '2026-11-28', '2026-11-28', now));
    assert.equal(batch.writes.length, 0);
  }
});
