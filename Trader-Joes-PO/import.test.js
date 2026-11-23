'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function body(count = 1, po = '0160295112') {
  const items = Array.from({ length: count }, (_, i) =>
    `PRODUCT NAME : ${i ? 'PRODUCT ' + i : 'HOL SIDE CORN PUDDING'}\nSKU : ${151493 + i} 8 16 Wgt Oz BOX\nCase QTY : ${i + 1}\nCOST : $ 31.20`).join('\n');
  return `PURCHASE AGREEMENT\nP.O. : ${po}\nDELIVER TO : WCD PALMDALE 5028 / 5028-PALMDALE XDOCK\nARRIVE : 11/28/26\n${items}\nTOTALS : $ 32.37`;
}

function load() {
  const context = vm.createContext({ Date, Number, JSON, Error,
    Utilities: {
      formatDate(date, timezone) {
        return new Intl.DateTimeFormat('en-CA', { timeZone: timezone, year: 'numeric', month: '2-digit', day: '2-digit' }).format(date);
      },
      parseDate(value) { return new Date(value.replace(' ', 'T') + ':00-08:00'); }
    } });
  vm.runInContext(fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8'), context);
  return context;
}

function setup(context, messages, existing = []) {
  const header = ['PROD_DATE', 'ARRIVE_DATE', 'LOCATION', 'PO_NUMBER', 'ORDER_TYPE', 'PRODUCT', 'QTY', 'UOM', 'NOTES'];
  const cells = [header, ...existing];
  const writes = [];
  let maxRows = 100;
  const sheet = {
    getLastRow: () => cells.length,
    getMaxRows: () => maxRows,
    insertRowsAfter: (_, n) => { maxRows += n; },
    getRange(row, col, numRows, numCols) {
      return {
        getValues: () => cells.slice(row - 1, row - 1 + numRows).map(r => r.slice(col - 1, col - 1 + numCols)),
        getDisplayValues: () => cells.slice(row - 1, row - 1 + numRows).map(r => r.slice(col - 1, col - 1 + numCols).map(v =>
          v instanceof Date ? context.Utilities.formatDate(v, 'America/Los_Angeles') : String(v))),
        setNumberFormat: () => {},
        setValues: rows => { writes.push(...rows); cells.push(...rows); }
      };
    }
  };
  const ss = { getSheetByName: name => name === 'Datos' ? sheet : null, getSpreadsheetTimeZone: () => 'America/Los_Angeles' };
  context.GmailApp = { search: (_, offset, count) => messages.slice(offset, offset + count).map((text, i) => ({
    getMessages: () => [{ getDate: () => new Date('2026-11-23T20:00:00Z'), getPlainBody: () => text, getId: () => String(offset + i) }]
  })) };
  return { ss, cells, writes };
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
  assert.match(c.tjImport_(ss, '2026-11-23', '2026-11-23'), /Filas nuevas: 1/);
  assert.equal(writes[0][3], '0160295112');
  assert.equal(c.Utilities.formatDate(writes[0][0], 'America/Los_Angeles'), '2026-11-27');
  assert.equal(c.Utilities.formatDate(writes[0][1], 'America/Los_Angeles'), '2026-11-28');
  assert.match(c.tjImport_(ss, '2026-11-23', '2026-11-23'), /sin cambios: 1/);
  assert.equal(writes.length, 1);
  cells[1][6] = 9;
  assert.match(c.tjImport_(ss, '2026-11-23', '2026-11-23'), /diferencias \(no modificados\): 0160295112/);
  assert.equal(cells[1][6], 9);
  assert.equal(writes.length, 1);
});

test('conflicting or malformed messages block all writes', () => {
  for (const messages of [[body(), body().replace('Case QTY : 1', 'Case QTY : 2')],
    [body(), body(1, '0160295113').replace('Case QTY : 1', 'Case QTY : bad')]]) {
    const c = load();
    const { ss, writes } = setup(c, messages);
    assert.throws(() => c.tjImport_(ss, '2026-11-23', '2026-11-23'));
    assert.equal(writes.length, 0);
  }
});

test('scans all pages, deduplicates messages and ignores non-order mail', () => {
  const c = load();
  const messages = Array(101).fill(body()).concat(['No purchase order here']);
  const { ss, writes } = setup(c, messages);
  assert.match(c.tjImport_(ss, '2026-11-23', '2026-11-23'), /Filas nuevas: 1/);
  assert.equal(writes.length, 1);
});
