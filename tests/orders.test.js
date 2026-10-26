'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'Orders.gs'), 'utf8');

function load(options = {}) {
  const calls = [];
  const tabs = options.tabs || [{ name: 'Weekly', displayed: block() }];
  const properties = { EAST_COAST_SPREADSHEET_ID: 'test-sheet',
    EAST_COAST_SHEET_NAMES: JSON.stringify(tabs.map(tab => tab.name)), ...options.properties };
  const sheets = tabs.map(tab => ({
    getName: () => tab.name,
    getDataRange: () => {
      if (tab.error) throw new Error(tab.error);
      return {
        getValues: () => tab.typed || tab.displayed,
        getDisplayValues: () => tab.displayed
      };
    }
  }));
  const context = vm.createContext({
    PropertiesService: { getScriptProperties: () => ({ getProperty: key => properties[key] || null }) },
    SpreadsheetApp: {
      openById: id => {
        calls.push(id);
        if (options.accessError) throw new Error(options.accessError);
        return {
          getSheets: () => sheets,
          getSheetByName: name => sheets.find((sheet, index) => tabs[index].name === name) || null,
          getSpreadsheetTimeZone: () => options.timezone || 'America/New_York'
        };
      }
    },
    Utilities: {
      formatDate: (date, timezone, format) => {
        calls.push({ timezone, format });
        const parts = new Intl.DateTimeFormat('en-US', {
          timeZone: timezone, year: 'numeric', month: '2-digit', day: '2-digit'
        }).formatToParts(date);
        const part = type => parts.find(value => value.type === type).value;
        return `${part('year')}-${part('month')}-${part('day')}`;
      }
    }
  });
  vm.runInContext(source, context, { filename: 'Orders.gs' });
  return { context, calls, get: date => JSON.parse(JSON.stringify(context.getEastCoastOrders(date))),
    getWeek: week => JSON.parse(JSON.stringify(context.getEastCoastOrdersByWeek(week))) };
}

function block(options = {}) {
  const stores = options.stores || ['Nazareth'];
  const pos = options.pos || stores.map(() => '0159543664');
  const arrivals = options.arrivals || stores.map(() => '11/04/2026');
  const rows = options.rows || [['Kung Pao', ...stores.map(() => 143)]];
  return [
    ['Week 44'],
    ['', ...stores, 'TOTAL CASES', 'TOTAL PALLETS', 'V-BOARDS'],
    ['PO#', ...pos, '', '', ''],
    ['PICK UP DATE', ...stores.map(() => '11/02/2026'), '', '', ''],
    ['ARRIVE (RDD)', ...arrivals, '', '', ''],
    ...rows,
    ['total pallets', ...stores.map(() => 99), 999, 999, 999],
    ['V-BOARDS', ...stores.map(() => 10), 999, 999, 999]
  ];
}

const holidayRows = [
  ['Holiday Creamed Corn Pudding', 572],
  ['Roasted vegetables w/ garlic', 286],
  ['Holiday Apple Cranberry Stuffing', 143],
  ['Four cheese scalloped potatoes', 286],
  ['Kung Pao', 143],
  ['Garlic Asiago Dip', 143],
  ['Roasted green vegetables', 143],
  ['Loaded mashed potatoes', 429],
  ['Holiday Vanilla Bean Whipped Sw…', 429],
  ['Stuffed peppers beef', 143],
  ['Stuffed poblanos chicken', 0],
  ['Holiday Entree Turkey Pot Pie', 180]
];

test('verified full sweet potatoes name matches the existing product in date and week batches', () => {
  const rows = holidayRows.map(([name, cases]) => [
    name === 'Holiday Vanilla Bean Whipped Sw…'
      ? '  HOLIDAY Vanilla Bean Whipped   Sweet Potatoes  ' : name, cases
  ]);
  const app = load({ tabs: [{ name: 'EAST COAST', displayed: block({ rows }) }] });
  for (const result of [app.get('2026-11-04'), app.getWeek('WK 44')]) {
    assert.equal(result.totalPallets, 21);
    const products = result.stores[0].products;
    const sweetPotatoes = products[8];
    assert.equal(sweetPotatoes.product.name, 'Holiday Vanilla Bean Whipped Sw…');
    assert.equal(sweetPotatoes.product.casesPerPallet, 143);
    assert.equal(sweetPotatoes.product.size, '');
    assert.equal(sweetPotatoes.totalCases, 429);
    assert.equal(sweetPotatoes.totalPallets, 3);
    assert.equal(sweetPotatoes.poNumber, '0159543664');
    assert.equal(products.length, 11);
  }
  const mixed = load({ tabs: [{ name: 'EAST COAST', displayed: block({
    rows: [['Holiday Vanilla Bean Whipped Sw…', 143],
      ['Holiday Vanilla Bean Whipped Sweet Potatoes', 286]]
  }) }] });
  assert.equal(mixed.getWeek('WK 44').stores[0].products.length, 1);
  assert.equal(mixed.getWeek('WK 44').totalPallets, 3);
});

test('full vanilla product generates every store quantity with its own PO and pallet count', () => {
  const stores = ['MINOOKA', 'DAYTONA', 'HARTFORD', 'CAROLINE', 'KENTUCKY', 'NAZARETH', 'NEW YORK'];
  const quantities = [143, 286, 429, 572, 143, 429, 1];
  const pos = stores.map((_, i) => '000' + i);
  const displayed = block({ stores, pos,
    rows: [['Holiday Vanilla Bean Whipped Sweet Potatoes', ...quantities, 2003]] });
  displayed[1][8] = '';
  const app = load({ tabs: [{ name: 'EAST COAST', displayed }] });
  for (const result of [app.get('2026-11-04'), app.getWeek('WK 44')]) {
    assert.deepEqual(result.stores.map(store => store.destination), stores);
    assert.deepEqual(result.stores.map(store => store.products[0].totalCases), quantities);
    assert.deepEqual(result.stores.map(store => store.products[0].poNumber), pos);
    assert.deepEqual(result.stores.map(store => store.totalPallets), [1, 2, 3, 4, 1, 3, 1]);
    assert.equal(result.totalPallets, 15);
  }
});

test('zero items never suppress other positive items or stores', () => {
  const displayed = block({ stores: ['NAZARETH', 'DAYTONA'], pos: ['0001', '0002'],
    rows: [
      ['Stuffed poblanos chicken', 0, 0],
      ['Holiday Vanilla Bean Whipped Sweet Potatoes', 429, 0],
      ['Kung Pao', 0, 286],
      ['Garlic Asiago Dip', 143, 143]
    ] });
  const app = load({ tabs: [{ name: 'EAST COAST', displayed }] });
  for (const result of [app.get('2026-11-04'), app.getWeek('WK 44')]) {
    assert.equal(result.stores.length, 2);
    assert.deepEqual(result.stores.map(store => store.totalPallets), [4, 3]);
    assert.deepEqual(result.stores.map(store => store.products.length), [2, 2]);
    assert.ok(result.stores.every(store => store.products.every(order => order.totalCases > 0)));
    assert.equal(result.totalPallets, 7);
  }
});

test('EAST COAST excludes unlabeled column I item totals in date and week batches', () => {
  const stores = ['MINOOKA', 'DAYTONA', 'HARTFORD', 'CAROLINE', 'KENTUCKY', 'Nazareth', 'NEW YORK'];
  const rows = holidayRows.map(([name, cases]) => [name, 0, 143, 0, 0, 0, cases, 143, cases + 286]);
  const displayed = block({ stores, pos: stores.map((_, i) => i === 5 ? '0159543664' : '000' + i),
    arrivals: stores.map(() => '11/04/2026'), rows });
  displayed[1][8] = '';
  const before = JSON.stringify(displayed);
  const app = load({ tabs: [{ name: 'EAST COAST', displayed }] });
  const date = app.get('2026-11-04');
  const week = app.getWeek('WK 44');
  [date, week].forEach(result => {
    assert.deepEqual(result.stores.map(store => store.destination), ['DAYTONA', 'Nazareth', 'NEW YORK']);
    const nazareth = result.stores.find(store => store.destination === 'Nazareth');
    assert.equal(nazareth.totalPallets, 21);
    assert.deepEqual(nazareth.products.map(order => order.totalPallets), [4, 2, 1, 2, 1, 1, 1, 3, 3, 1, 2]);
    assert.ok(nazareth.products.every(order => order.poNumber === '0159543664'));
  });
  assert.equal(JSON.stringify(displayed), before, 'read-only preprocessing must not mutate sheet arrays');
});

test('column I exclusions never hide order metadata or affect other tabs', () => {
  const stores = ['MINOOKA', 'DAYTONA', 'HARTFORD', 'CAROLINE', 'KENTUCKY', 'Nazareth', 'NEW YORK'];
  for (const row of [2, 3, 4]) {
    const displayed = block({ stores });
    displayed[row][8] = row === 2 ? '0009' : '11/04/2026';
    const app = load({ tabs: [{ name: 'EAST COAST', displayed }] });
    assert.throws(() => app.get('2026-11-04'), /Column I.*order metadata/);
    assert.throws(() => app.getWeek('WK 44'), /Column I.*order metadata/);
  }
  const displayed = block({ stores, rows: [['Kung Pao', 0, 0, 0, 0, 0, 143, 0, 143]] });
  displayed[1][8] = '';
  const app = load({ tabs: [{ name: 'Other orders', displayed }] });
  assert.throws(() => app.get('2026-11-04'), /Missing destination for nonzero cases in column 9/);
  assert.throws(() => app.getWeek('WK 44'), /Missing destination for nonzero cases in column 9/);
});

test('column I week markers survive total exclusion and retain week boundaries', () => {
  const stores = ['MINOOKA', 'DAYTONA', 'HARTFORD', 'CAROLINE', 'KENTUCKY', 'Nazareth', 'NEW YORK'];
  for (const split of [false, true]) {
    function weekly(week, po, rows) {
      const data = block({ stores, pos: stores.map(() => po),
        rows: rows.map(([name, cases]) => [name, 0, 0, 0, 0, 0, cases, 0, cases]) });
      data[0] = Array(9).fill('');
      data[0][8] = split ? String(week) : 'WK ' + week;
      if (split) data[0][7] = 'WK';
      data[1][8] = '';
      return data;
    }
    const displayed = [
      ...weekly(43, '00043', [['Kung Pao', 143]]),
      ...weekly(44, '0159543664', holidayRows),
      ...weekly(45, '00045', [['Kung Pao', 286]])
    ];
    const before = JSON.stringify(displayed);
    const app = load({ tabs: [{ name: 'EAST COAST', displayed }] });
    const result = app.getWeek('WK 44');
    assert.equal(result.totalPallets, 21);
    assert.equal(result.stores.length, 1);
    assert.equal(result.stores[0].destination, 'Nazareth');
    assert.deepEqual(result.stores[0].products.map(order => order.totalPallets), [4, 2, 1, 2, 1, 1, 1, 3, 3, 1, 2]);
    assert.ok(result.stores[0].products.every(order => order.poNumber === '0159543664'));
    assert.equal(app.getWeek('WK 43').totalPallets, 1);
    assert.equal(app.getWeek('WK 45').totalPallets, 2);
    assert.equal(JSON.stringify(displayed), before);
  }
});

test('invalid and conflicting column I week markers cannot be erased into a successful batch', () => {
  for (const marker of ['WK 54', 'WK440', 'WK']) {
    const displayed = block();
    displayed[0] = Array(9).fill('');
    displayed[0][8] = marker;
    const app = load({ tabs: [{ name: 'EAST COAST', displayed }] });
    assert.throws(() => app.getWeek('WK 44'), /Week marker.*WK 1 through WK 53/);
  }
  const displayed = block();
  displayed[0][8] = 'WK 45';
  const app = load({ tabs: [{ name: 'EAST COAST', displayed }] });
  assert.throws(() => app.getWeek('WK 44'), /Conflicting week markers/);
});

test('date loading still excludes column I markers and totals without requiring weeks', () => {
  const displayed = block({ rows: [['Kung Pao', 143, 0, 0, 0, 0, 0, 0, 143]] });
  displayed[0] = Array(9).fill('');
  displayed[0][8] = 'WK 44';
  const app = load({ tabs: [{ name: 'EAST COAST', displayed }] });
  assert.equal(app.get('2026-11-04').totalPallets, 1);
});

test('EAST COAST ignores unrelated J+ quantities, headers, metadata and week markers', () => {
  const stores = ['MINOOKA', 'DAYTONA', 'HARTFORD', 'CAROLINE', 'KENTUCKY', 'Nazareth', 'NEW YORK'];
  const displayed = block({ stores, rows: holidayRows.map(([name, cases]) =>
    [name, 0, 143, 0, 0, 0, cases, 143, cases + 286]) });
  displayed[0] = Array(9).fill('');
  displayed[0][8] = 'WK 44';
  displayed[1][8] = '';
  const baseline = load({ tabs: [{ name: 'EAST COAST', displayed }] });
  const date = baseline.get('2026-11-04');
  const week = baseline.getWeek('WK 44');
  const extended = displayed.map((row, index) => {
    const copy = row.slice();
    copy.length = Math.max(copy.length, 11);
    copy[10] = index === 1 ? '' : 999;
    copy[11] = index === 0 ? 'WK 54' : 'Unrelated content';
    copy[12] = index === 0 ? 'WK 45' : '#REF!';
    copy[13] = index === 1 ? 'OTHER STORE' : 'PO#';
    return copy;
  });
  const before = JSON.stringify(extended);
  const app = load({ tabs: [{ name: 'EAST COAST', displayed: extended }] });
  assert.deepEqual(app.get('2026-11-04'), date);
  assert.deepEqual(app.getWeek('WK 44'), week);
  assert.equal(week.stores.find(store => store.destination === 'Nazareth').totalPallets, 21);
  assert.equal(JSON.stringify(extended), before);
});

test('EAST COAST cannot take its week reference from columns beyond I', () => {
  const displayed = block();
  displayed[0] = Array(10).fill('');
  displayed[0][9] = 'WK 44';
  const app = load({ tabs: [{ name: 'EAST COAST', displayed }] });
  assert.throws(() => app.getWeek('WK 44'), /no recognizable worksheet week markers/);
  assert.equal(app.get('2026-11-04').totalPallets, 1);
});

test('exact Nazareth breakdown uses per-item PALLET# totals, 21 only for store summary, and leading-zero PO', () => {
  const { get } = load({ tabs: [{ name: 'Holiday', displayed: block({ rows: holidayRows }) }] });
  const result = get('2026-11-04');
  assert.equal(result.date, '2026-11-04');
  assert.equal(result.totalPallets, 21);
  assert.equal(result.stores.length, 1);
  const store = result.stores[0];
  assert.equal(store.destination, 'Nazareth');
  assert.equal(store.totalCases, 2897);
  assert.equal(store.totalPallets, 21);
  assert.deepEqual(store.products.map(entry => entry.totalPallets), [4, 2, 1, 2, 1, 1, 1, 3, 3, 1, 2]);
  assert.ok(store.products.every(entry => entry.totalPallets !== store.totalPallets));
  assert.deepEqual(store.products.map(entry => [entry.totalCases, entry.product.casesPerPallet, entry.totalPallets]),
    [[572, 143, 4], [286, 143, 2], [143, 143, 1], [286, 143, 2], [143, 143, 1], [143, 143, 1],
      [143, 143, 1], [429, 143, 3], [429, 143, 3], [143, 143, 1], [180, 90, 2]]);
  assert.ok(store.products.every(entry => entry.poNumber === '0159543664'));
  assert.equal(store.products[8].product.name, 'Holiday Vanilla Bean Whipped Sw…');
  assert.equal(store.products[8].product.size, '');
  assert.equal(store.products[4].product.size, '14 oz');
  assert.ok(result.warnings.some(warning => warning.includes('truncated')));
  assert.equal(result.warnings.filter(warning => warning.startsWith('Size unavailable')).length, 7);
  assert.ok(!JSON.stringify(result).includes('Sweet'));
});

test('authoritative catalog keeps exactly six original manual products and verified capacities', () => {
  const { context } = load();
  const catalog = JSON.parse(JSON.stringify(context.getEastCoastProducts_()));
  assert.equal(catalog.filter(product => product.manual).length, 6);
  assert.equal(catalog.find(product => product.name === 'SHELLED EDAMAME').casesPerPallet, 160);
  assert.ok(catalog.every(product => Array.isArray(product.aliases)));
  assert.ok(catalog.filter(product => !product.manual).every(product => product.size === ''));
  assert.ok(catalog.filter(product => !product.manual).every(product => product.nameClass === 'name--holiday'));
});

test('defaults-only catalog ignores malformed overrides so unrelated manual apps still render', () => {
  const { context, get } = load({ properties: { EAST_COAST_PRODUCTS: 'invalid JSON' } });
  const defaults = JSON.parse(JSON.stringify(context.getEastCoastProducts_(true)));
  assert.equal(defaults.filter(product => product.manual).length, 6);
  assert.deepEqual(defaults, JSON.parse(JSON.stringify(load().context.getEastCoastProducts_())));
  assert.throws(() => context.getEastCoastProducts_(), /EAST_COAST_PRODUCTS/);
  assert.throws(() => context.getEastCoastProducts_(false), /EAST_COAST_PRODUCTS/);
  assert.throws(() => get('2026-11-04'), /EAST_COAST_PRODUCTS/);
  context.PropertiesService.getScriptProperties = () => { throw new Error('Properties unavailable'); };
  assert.equal(context.getEastCoastProducts_(true).length, defaults.length);
});

test('all matching stores are included, wrong dates excluded and summary columns ignored', () => {
  const { get } = load({ tabs: [{ name: 'Weekly', displayed: block({
    stores: ['Nazareth', 'VIRGINIA', 'CAROLINE'],
    pos: ['0159543664', '002', '003'],
    arrivals: ['11/04/2026', '2026-11-04', '11/05/2026'],
    rows: [['Kung Pao', 143, 286, 429, 858, 6, 9], ['Garlic Asiago Dip', 0, 1, 143, 144, 2, 3]]
  }) }] });
  const result = get('2026-11-04');
  assert.deepEqual(result.stores.map(store => store.destination), ['Nazareth', 'VIRGINIA']);
  assert.equal(result.totalPallets, 4);
  assert.deepEqual(get('2026-11-06').stores, []);
});

test('repeated weeks use nearest local headers, including CAROLINE instead of top VIRGINIA', () => {
  const displayed = [
    ...block({ stores: ['VIRGINIA'], arrivals: ['10/28/2026'], rows: [['Kung Pao', 143]] }),
    ...block({ stores: ['CAROLINE'], pos: ['0008'], rows: [['Garlic Asiago Dip', 286]] }),
    ...block({ stores: ['Nazareth'], pos: ['0009'], rows: [['Kung Pao', 143]] })
  ];
  const result = load({ tabs: [{ name: 'Weeks', displayed }] }).get('2026-11-04');
  assert.deepEqual(result.stores.map(store => store.destination), ['CAROLINE', 'Nazareth']);
  assert.equal(result.totalPallets, 3);
});

test('same store, PO and product aggregate before rounding partial pallets', () => {
  const displayed = [
    ...block({ stores: ['CAROLINE'], pos: ['0008'], rows: [['Kung Pao', 50], [' kung  pao ', 50]] }),
    ...block({ stores: ['caroline'], pos: ['0008'], rows: [['KUNG PAO', 43]] }),
    ...block({ stores: ['CAROLINE'], pos: ['0009'], rows: [['Kung Pao', 1]] })
  ];
  const result = load({ tabs: [{ name: 'Weeks', displayed }] }).get('2026-11-04');
  assert.equal(result.stores.length, 1);
  assert.deepEqual(result.stores[0].products.map(entry => [entry.poNumber, entry.totalCases, entry.totalPallets]),
    [['0008', 143, 1], ['0009', 1, 1]]);
  assert.equal(result.totalPallets, 2);
});

test('JIT EDAMAME rows use 160 cases per pallet, and blank/zero product rows skip', () => {
  const displayed = [...block({ rows: [['JIT EDAMAME', ''], ['Shelled edamame', 161], ['Unknown zero', '0'], ['Unknown blank', '']] }),
    ...block({ stores: ['VIRGINIA'], pos: ['0010'], rows: [['JIT EDAMAME', 160]] })];
  const result = load({ tabs: [{ name: 'JIT', displayed }] }).get('2026-11-04');
  assert.equal(result.totalPallets, 3);
  assert.equal(result.stores[0].products[0].product.name, 'SHELLED EDAMAME');
  assert.equal(result.stores[1].products[0].totalPallets, 1);
});

test('case, nonbreaking-space, whitespace and explicit screenshot aliases resolve', () => {
  const displayed = block({ rows: [
    ['  Stuffed\u00a0 peppers   beef  ', 143],
    [' STUFFED POBLANOS CHICKEN ', 143],
    ['roasted VEGETABLES with garlic', 143],
    ['Holiday Vanilla Bean Whipped Sw...', 143]
  ] });
  const result = load({ tabs: [{ name: 'Aliases', displayed }] }).get('2026-11-04');
  assert.deepEqual(result.stores[0].products.map(entry => entry.product.name),
    ['STUFFED PEPPERS WITH BEEF', 'CHICKEN STUFFED POBLANO', 'Roasted vegetables w/ garlic', 'Holiday Vanilla Bean Whipped Sw…']);
});

test('columns dynamically follow a shifted destination header, not fixed B:H', () => {
  const displayed = block({ stores: ['Nazareth', 'CAROLINE'], pos: ['0001', '0002'], rows: [['Kung Pao', 143, 286]] })
    .map(row => ['', '', ...row]);
  const result = load({ tabs: [{ name: 'Shifted', displayed }] }).get('2026-11-04');
  assert.equal(result.totalPallets, 3);
  assert.deepEqual(result.stores.map(store => store.destination), ['Nazareth', 'CAROLINE']);
});

test('product descriptions may occupy an extra column before the destination columns', () => {
  const displayed = [
    ['', 'Product', 'Nazareth', 'TOTAL CASES'],
    ['PO#', '', '0012', ''],
    ['ARRIVE (RDD)', '', '11/04/2026', ''],
    ['12345', 'Kung Pao', 143, 143],
    ['', 'Garlic Asiago Dip', 143, 143],
    ['total pallets', '', 2, 2]
  ];
  const result = load({ tabs: [{ name: 'Extra column', displayed }] }).get('2026-11-04');
  assert.equal(result.totalPallets, 2);
});

test('typed dates use spreadsheet timezone and displayed PO text, never RPC Date objects', () => {
  const displayed = block({ pos: ['0159543664'], arrivals: ['11/03/2026'] });
  const typed = displayed.map(row => row.slice());
  typed[2][1] = 159543664;
  typed[3][1] = new Date('2026-11-02T05:00:00Z');
  typed[4][1] = new Date('2026-11-04T01:00:00Z');
  const { get, calls } = load({ tabs: [{ name: 'Dates', displayed, typed }] });
  assert.equal(get('2026-11-03').stores[0].products[0].poNumber, '0159543664');
  assert.equal(get('2026-11-04').totalPallets, 0);
  assert.ok(calls.some(call => call.timezone === 'America/New_York' && call.format === 'yyyy-MM-dd'));
});

test('valid leap days and US string dates normalize without Date.parse', () => {
  const { context } = load();
  assert.equal(context.eastCoastCellDate_('', '2/29/2028', 'UTC', 'test'), '2028-02-29');
  assert.equal(context.eastCoastCellDate_('', '2026-11-04', 'UTC', 'test'), '2026-11-04');
});

test('strict input date rejects impossible, ambiguous and nonstring values before opening sheet', () => {
  for (const date of ['2026-02-29', '2026-11-31', '0000-01-01', '2026-13-01', '2026-11-4', '11/04/2026', '2026-11-04T00:00:00Z', '', null, new Date()]) {
    const { get, calls } = load();
    assert.throws(() => get(date), /date/i);
    assert.equal(calls.length, 0);
  }
});

test('malformed arrival dates fail clearly rather than appearing to have no orders', () => {
  for (const arrival of ['11/31/2026', '2026-02-29', '11/04/26', '4 Nov 2026', '#VALUE!']) {
    const { get } = load({ tabs: [{ name: 'Bad dates', displayed: block({ arrivals: [arrival] }) }] });
    assert.throws(() => get('2026-11-04'), /ARRIVE \(RDD\).*Bad dates.*Nazareth/);
  }
});

test('invalid typed arrival date fails but irrelevant pickup metadata never rejects arrival batch', () => {
  const displayed = block();
  const typed = displayed.map(row => row.slice());
  typed[4][1] = new Date(NaN);
  assert.throws(() => load({ tabs: [{ name: 'Invalid', displayed, typed }] }).get('2026-11-04'), /invalid date/);
  const badPickup = block();
  badPickup[3][1] = '11/31/2026';
  badPickup.splice(4, 0, ['PICK UP DATE', 'TBD']);
  assert.equal(load({ tabs: [{ name: 'Invalid', displayed: badPickup }] }).get('2026-11-04').totalPallets, 1);
});

test('missing arrival row, missing arrival date and duplicate arrival rows fail', () => {
  const missingRow = block().filter(row => row[0] !== 'ARRIVE (RDD)');
  assert.throws(() => load({ tabs: [{ name: 'Missing', displayed: missingRow }] }).get('2026-11-04'), /Missing ARRIVE \(RDD\) row/);
  assert.throws(() => load({ tabs: [{ name: 'Missing', displayed: block({ arrivals: [''] }) }] }).get('2026-11-04'), /Missing ARRIVE \(RDD\) date/);
  const duplicated = block();
  duplicated.splice(5, 0, ['ARRIVE (RDD)', '11/04/2026']);
  assert.throws(() => load({ tabs: [{ name: 'Duplicate', displayed: duplicated }] }).get('2026-11-04'), /Multiple ARRIVE/);
});

test('missing or whitespace-only PO cannot silently drop positive cases', () => {
  for (const po of ['', '   ']) {
    const { get } = load({ tabs: [{ name: 'Missing', displayed: block({ pos: [po] }) }] });
    assert.throws(() => get('2026-11-04'), /PO#/);
  }
});

test('displayed alphanumeric PO strings are preserved without a digits-only restriction', () => {
  for (const po of ['PO-0159543664-A', '0012B', 'FHH / HOLIDAY 44', '#ORDER-12']) {
    const result = load({ tabs: [{ name: 'PO strings', displayed: block({ pos: [po] }) }] }).get('2026-11-04');
    assert.equal(result.stores[0].products[0].poNumber, po);
  }
});

test('spreadsheet PO error values reject positive orders without rejecting valid alphanumeric POs', () => {
  for (const po of ['#REF!', '#N/A', '#VALUE!', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!', '#ERROR!', '#SPILL!', '#CALC!']) {
    const { get } = load({ tabs: [{ name: 'PO error', displayed: block({ pos: [po] }) }] });
    assert.throws(() => get('2026-11-04'), /Spreadsheet error.*PO#.*Correct the PO cell/);
    assert.equal(get('2026-11-05').totalPallets, 0);
  }
  assert.equal(load({ tabs: [{ name: 'Unused PO error', displayed: block({ pos: ['#REF!'], rows: [['Kung Pao', 0]] }) }] })
    .get('2026-11-04').totalPallets, 0);
});

test('unknown nonzero product fails the whole date, but unknown zero and other dates skip', () => {
  const { get } = load({ tabs: [{ name: 'Unknown', displayed: block({ rows: [['Kung Pao', 143], ['Mystery dish', 1]] }) }] });
  assert.throws(() => get('2026-11-04'), /Unknown product "Mystery dish".*EAST_COAST_PRODUCTS/);
  assert.equal(get('2026-11-05').totalPallets, 0);
  assert.equal(load({ tabs: [{ name: 'Zero', displayed: block({ rows: [['Mystery dish', 0]] }) }] }).get('2026-11-04').totalPallets, 0);
});

test('blank product label with positive cases fails rather than losing an order', () => {
  assert.throws(() => load({ tabs: [{ name: 'Blank', displayed: block({ rows: [['', 143]] }) }] }).get('2026-11-04'), /blank product name/);
});

test('negative, fractional, nonfinite, formula errors and malformed text quantities fail', () => {
  for (const quantity of [-1, 1.5, Infinity, NaN, '-1', '1.5', '#N/A', '143 cases', '1,43', true]) {
    const { get } = load({ tabs: [{ name: 'Quantity', displayed: block({ rows: [['Kung Pao', quantity]] }) }] });
    assert.throws(() => get('2026-11-04'), /invalid case quantity/);
  }
});

test('formatted numeric case counts and zero are valid', () => {
  const { get } = load({ tabs: [{ name: 'Formatted', displayed: block({ rows: [['Kung Pao', '1,001'], ['Garlic Asiago Dip', '0']] }) }] });
  assert.equal(get('2026-11-04').totalPallets, 7);
});

test('blank and zero text forms skip even unknown products; integer decimal text is valid', () => {
  const zeroRows = ['', ' ', '\u00a0', '0', '000', ' 0 ', '0.0', '0.00', '-0', '+0'].map(value => ['Unknown zero product', value]);
  const displayed = block({ rows: [...zeroRows, ['Kung Pao', '143.00']] });
  const result = load({ tabs: [{ name: 'Zeros', displayed }] }).get('2026-11-04');
  assert.equal(result.stores[0].products.length, 1);
  assert.equal(result.totalPallets, 1);
});

test('case aggregation rejects unsafe overflow within a block, across blocks and across products', () => {
  const max = Number.MAX_SAFE_INTEGER;
  const variants = [
    block({ rows: [['Kung Pao', max], ['Kung Pao', 1]] }),
    [...block({ rows: [['Kung Pao', max]] }), ...block({ rows: [['Kung Pao', 1]] })],
    block({ rows: [['Kung Pao', max], ['Garlic Asiago Dip', 1]] })
  ];
  for (const displayed of variants) {
    assert.throws(() => load({ tabs: [{ name: 'Overflow', displayed }] }).get('2026-11-04'), /exceed the safe whole-number range/);
  }
});

test('pallet aggregation rejects unsafe overflow across stores', () => {
  const properties = { EAST_COAST_PRODUCTS: JSON.stringify([{ name: 'One case per pallet', casesPerPallet: 1 }]) };
  const displayed = block({ stores: ['Nazareth', 'CAROLINE'], pos: ['001', '002'],
    rows: [['One case per pallet', Number.MAX_SAFE_INTEGER, 1]] });
  assert.throws(() => load({ properties, tabs: [{ name: 'Overflow', displayed }] }).get('2026-11-04'), /pallet totals.*safe whole-number range/);
});

test('missing or invalid catalog capacities fail even in pure parser', () => {
  const { context } = load();
  for (const capacity of [undefined, 0, -1, Infinity, 1.5, '143']) {
    assert.throws(() => context.parseEastCoastSheet_(block(), block(), 'UTC', '2026-11-04',
      [{ name: 'Kung Pao', size: '', casesPerPallet: capacity, aliases: [] }], 'Capacity'), /pallet capacity/);
  }
});

test('missing per-block header never reuses top VIRGINIA header', () => {
  const second = block({ stores: ['CAROLINE'], pos: ['0010'] });
  second.splice(1, 1);
  const displayed = [...block({ stores: ['VIRGINIA'] }), ...second];
  assert.throws(() => load({ tabs: [{ name: 'Header', displayed }] }).get('2026-11-04'), /Missing destination header/);
});

test('a malformed quantity row cannot masquerade as a later block destination header', () => {
  const second = block({ stores: ['CAROLINE'], pos: ['0010'] });
  second.splice(1, 1);
  const displayed = [...block({ rows: [['Kung Pao', 'bad quantity']] }), ...second];
  assert.throws(() => load({ tabs: [{ name: 'Header', displayed }] }).get('2026-11-04'), /invalid case quantity|Missing destination header/);
});

test('unmatched order columns and fully absent headers fail', () => {
  const displayed = block({ stores: ['Nazareth', 'CAROLINE'], pos: ['001', '002'], rows: [['Kung Pao', 143, 286]] });
  displayed[1][2] = '';
  assert.throws(() => load({ tabs: [{ name: 'Header', displayed }] }).get('2026-11-04'), /Missing destination for column 3/);
  const absent = block();
  absent.splice(1, 1);
  assert.throws(() => load({ tabs: [{ name: 'Header', displayed: absent }] }).get('2026-11-04'), /Missing destination header/);
});

test('orphan nonzero cases without destination, PO or arrival are not silently lost', () => {
  const displayed = block({ stores: ['Nazareth', 'CAROLINE'], pos: ['001', ''], arrivals: ['11/04/2026', ''],
    rows: [['Kung Pao', 143, 286]] });
  displayed[1][2] = '';
  displayed[3][2] = '';
  assert.throws(() => load({ tabs: [{ name: 'Orphan', displayed }] }).get('2026-11-04'), /Missing destination for nonzero cases/);
});

test('a populated PO with a blank arrival fails even when its product rows are zero', () => {
  assert.throws(() => load({ tabs: [{ name: 'Malformed', displayed: block({ arrivals: [''], rows: [['Kung Pao', 0]] }) }] })
    .get('2026-11-04'), /Missing ARRIVE \(RDD\) date for PO#/);
});

test('explicit selected tabs read all matching stores or only the specified tab', () => {
  const tabs = [
    { name: 'One', displayed: block({ stores: ['Nazareth'] }) },
    { name: 'Two', displayed: block({ stores: ['CAROLINE'], pos: ['0002'] }) }
  ];
  assert.equal(load({ tabs, properties: { EAST_COAST_SHEET_NAMES: '["One"]' } }).get('2026-11-04').stores.length, 1);
  assert.equal(load({ tabs, properties: { EAST_COAST_SHEET_NAMES: '["One","Two"]' } }).get('2026-11-04').stores.length, 2);
});

test('identical store/PO snapshots across selected tabs deduplicate with warning', () => {
  const displayed = block({ rows: [['Kung Pao', 70], ['Kung Pao', 73]] });
  const { get } = load({
    tabs: [{ name: 'One', displayed }, { name: 'Two', displayed: block() }],
    properties: { EAST_COAST_SHEET_NAMES: '["One","Two"]' }
  });
  const result = get('2026-11-04');
  assert.equal(result.totalPallets, 1);
  assert.equal(result.stores[0].totalCases, 143);
  assert.ok(result.warnings.some(warning => warning.includes('duplicated across tabs')));
});

test('identical complete store/PO snapshots within a tab deduplicate weekly/JIT repeated views', () => {
  const displayed = [
    ...block({ rows: [['Shelled Edamame', 160]] }),
    ...block({ rows: [['JIT EDAMAME', 80], ['JIT EDAMAME', 80]] })
  ];
  const result = load({ tabs: [{ name: 'Weekly and JIT', displayed }] }).get('2026-11-04');
  assert.equal(result.stores[0].totalCases, 160);
  assert.equal(result.totalPallets, 1);
  assert.ok(result.warnings.some(warning => warning.includes('duplicated within tab Weekly and JIT')));
});

test('JIT repeated product snapshot deduplicates even when weekly view has additional products', () => {
  const displayed = [
    ...block({ rows: [['Shelled Edamame', 160], ['Kung Pao', 143]] }),
    ...block({ rows: [['JIT EDAMAME', 160]] })
  ];
  const result = load({ tabs: [{ name: 'Weekly and JIT', displayed }] }).get('2026-11-04');
  assert.equal(result.stores[0].totalCases, 303);
  assert.equal(result.totalPallets, 2);
  assert.ok(result.warnings.some(warning => warning.includes('SHELLED EDAMAME duplicated within tab')));
});

test('JIT title/store header schema uses its own PO# and ARRIVE (RDD) rows', () => {
  const displayed = [
    ['JIT EDAMAME', 'Nazareth', 'VIRGINIA', 'TOTAL CASES'],
    ['PO#', 'JIT-001', 'JIT-002', ''],
    ['PICK UP DATE', 'TBD', '#VALUE!', ''],
    ['ARRIVE (RDD)', '11/04/2026', '11/04/2026', ''],
    ['JIT EDAMAME', 160, 320, 480],
    ['total pallets', 1, 2, 3]
  ];
  const result = load({ tabs: [{ name: 'JIT schema', displayed }] }).get('2026-11-04');
  assert.equal(result.totalPallets, 3);
  assert.deepEqual(result.stores.map(store => store.products[0].poNumber), ['JIT-001', 'JIT-002']);
});

test('different quantities across selected tabs aggregate, not discarded as duplicates', () => {
  const { get } = load({
    tabs: [{ name: 'One', displayed: block({ rows: [['Kung Pao', 70]] }) },
      { name: 'Two', displayed: block({ rows: [['Kung Pao', 73]] }) }],
    properties: { EAST_COAST_SHEET_NAMES: '["One","Two"]' }
  });
  assert.equal(get('2026-11-04').stores[0].totalCases, 143);
  assert.equal(get('2026-11-04').totalPallets, 1);
});

test('helpful missing config, spreadsheet access, missing tab and read errors', () => {
  assert.throws(() => load({ properties: { EAST_COAST_SPREADSHEET_ID: '' } }).get('2026-11-04'), /Set the EAST_COAST_SPREADSHEET_ID/);
  assert.throws(() => load({ accessError: 'Permission denied' }).get('2026-11-04'), /grant.*read access.*Permission denied/);
  assert.throws(() => load({ properties: { EAST_COAST_SHEET_NAMES: '["Missing"]' } }).get('2026-11-04'), /tab not found.*Missing/);
  assert.throws(() => load({ tabs: [{ name: 'Read', error: 'Denied' }] }).get('2026-11-04'), /Cannot read.*Read.*read access/);
  assert.throws(() => load({ tabs: [] }).get('2026-11-04'), /EAST_COAST_SHEET_NAMES/);
});

test('malformed tab configuration and empty/unrecognizable data fail helpfully', () => {
  for (const value of ['not JSON', '{}', '[]', '["Weekly","Weekly"]', '[null]', '[""]']) {
    assert.throws(() => load({ properties: { EAST_COAST_SHEET_NAMES: value } }).get('2026-11-04'), /EAST_COAST_SHEET_NAMES/);
  }
  assert.throws(() => load({ tabs: [{ name: 'Empty', displayed: [['', '']] }] }).get('2026-11-04'), /no order data/);
  assert.throws(() => load({ tabs: [{ name: 'Wrong', displayed: [['some unrelated table']] }] }).get('2026-11-04'), /no recognizable PO#/);
});

test('configured verified extra product supplies capacity and size without adding manual selection', () => {
  const properties = { EAST_COAST_PRODUCTS: JSON.stringify([
    { name: 'Verified dish', size: '18 oz', casesPerPallet: 100, aliases: ['Special dish'] }
  ]) };
  const { get, context } = load({ properties, tabs: [{ name: 'Extras', displayed: block({ rows: [['Special dish', 101]] }) }] });
  assert.equal(get('2026-11-04').totalPallets, 2);
  assert.equal(get('2026-11-04').stores[0].products[0].product.size, '18 oz');
  assert.equal(context.getEastCoastProducts_().filter(product => product.manual).length, 6);
});

test('verified configuration may supply missing holiday size but not change known capacity', () => {
  const { context } = load({ properties: { EAST_COAST_PRODUCTS: JSON.stringify([
    { name: 'Holiday Creamed Corn Pudding', size: 'verified size', casesPerPallet: 143 }
  ]) } });
  assert.equal(context.getEastCoastProducts_().find(product => product.name === 'Holiday Creamed Corn Pudding').size, 'verified size');
  assert.throws(() => load({ properties: { EAST_COAST_PRODUCTS: JSON.stringify([
    { name: 'SHELLED EDAMAME', casesPerPallet: 143 }
  ]) } }).get('2026-11-04'), /cannot change.*capacity/);
});

test('invalid extra definitions and ambiguous aliases fail clearly', () => {
  for (const extra of ['{}', 'oops', '[{}]', '[{"name":"Dish","casesPerPallet":0}]',
    '[{"name":"Dish","casesPerPallet":100,"aliases":[1]}]',
    '[{"name":"Dish","casesPerPallet":100,"nameClass":"unsafe class"}]',
    '[{"name":"Dish","casesPerPallet":100,"aliases":["Kung Pao"]}]']) {
    assert.throws(() => load({ properties: { EAST_COAST_PRODUCTS: extra } }).get('2026-11-04'), /EAST_COAST_PRODUCTS|Ambiguous East Coast product alias/);
  }
});

test('both sheet APIs require explicit East Coast tab configuration before opening or reading', () => {
  for (const names of [undefined, null, '', ' ', '[]', '["Weekly","Weekly"]', '[" "]', '[1]', '{}', 'bad JSON']) {
    const { get, getWeek, calls } = load({ properties: { EAST_COAST_SHEET_NAMES: names },
      tabs: [{ name: 'TJ-OCT', error: 'Must never read first tab' }] });
    assert.throws(() => get('2026-11-04'), /EAST_COAST_SHEET_NAMES/);
    assert.throws(() => getWeek('WK 44'), /EAST_COAST_SHEET_NAMES/);
    assert.equal(calls.length, 0);
    if (!names || names === ' ') {
      assert.throws(() => getWeek('44'), /East Coast tab configuration.*first-tab fallback is not allowed/);
    }
  }
});

test('configured East Coast tab ignores the wrong first TJ-OCT tab in both modes', () => {
  const { get, getWeek } = load({
    properties: { EAST_COAST_SHEET_NAMES: '["EAST COAST"]' },
    tabs: [{ name: 'TJ-OCT', error: 'Wrong first tab read' }, { name: 'EAST COAST', displayed: block() }]
  });
  assert.equal(get('2026-11-04').totalPallets, 1);
  assert.equal(getWeek('44').totalPallets, 1);
});

test('week API normalizes canonical labels and aliases including boundary weeks before sheet access', () => {
  for (const [input, expected] of [['WK 44', 'WK 44'], ['44', 'WK 44'], [' week : 44 ', 'WK 44'],
    ['wK\u00a0 44', 'WK 44'], ['1', 'WK 1'], ['Week 53', 'WK 53'], ['WK 01', 'WK 1']]) {
    const displayed = block();
    displayed[0] = [expected];
    const result = load({ tabs: [{ name: 'Weeks', displayed }] }).getWeek(input);
    assert.equal(result.week, expected);
    assert.equal(result.totalPallets, 1);
    assert.ok(!Object.hasOwn(result, 'date'));
  }
  for (const week of ['', null, undefined, 44, '0', '54', 'WK 440', 'WK440', '44x', '2026-W44', 'WK -1', '4.4']) {
    const { getWeek, calls } = load();
    assert.throws(() => getWeek(week), /WK 1 through WK 53/);
    assert.equal(calls.length, 0);
  }
});

test('literal week 44 includes screenshot Nazareth 21 and Daytona/New York despite different arrivals', () => {
  const displayed = [
    ...block({ rows: holidayRows }),
    ...block({ stores: ['Daytona', 'New York'], pos: ['0002', '0003'],
      arrivals: ['11/05/2026', '11/06/2026'], rows: [['Kung Pao', 143, 286]] }).slice(1)
  ];
  const { getWeek, get } = load({ tabs: [{ name: 'EAST COAST', displayed }] });
  const result = getWeek('WK 44');
  assert.equal(result.week, 'WK 44');
  assert.deepEqual(result.stores.map(store => store.destination), ['Nazareth', 'Daytona', 'New York']);
  assert.equal(result.totalPallets, 24);
  const nazareth = result.stores[0];
  assert.equal(nazareth.totalCases, 2897);
  assert.equal(nazareth.totalPallets, 21);
  const palletBreakdown = nazareth.products.map(entry => entry.totalPallets);
  assert.ok(!nazareth.products.some(entry => entry.product.name === 'CHICKEN STUFFED POBLANO'));
  palletBreakdown.splice(10, 0, 0);
  assert.deepEqual(palletBreakdown, [4, 2, 1, 2, 1, 1, 1, 3, 3, 1, 0, 2]);
  assert.ok(nazareth.products.every(entry => entry.poNumber === '0159543664'));
  assert.deepEqual(result.stores.slice(1).map(store => store.products[0].arrivalDate), ['2026-11-05', '2026-11-06']);
  assert.equal(get('2026-11-04').stores.length, 1);
  assert.ok(result.warnings.some(warning => warning.includes('truncated')));
});

test('week markers are anchored, case/whitespace insensitive and may be split in any column', () => {
  for (const marker of [[' wk : 44 '], ['', '', 'WEEK', 44], ['', 'Wk:', '44'],
    ['Week 44', '', 'WK : 44']]) {
    const displayed = block().map(row => ['', '', ...row]);
    displayed[0] = marker;
    assert.equal(load({ tabs: [{ name: 'Split', displayed }] }).getWeek('44').totalPallets, 1);
  }
  const displayed = block();
  displayed[0] = ['Orders for WK 44'];
  assert.throws(() => load({ tabs: [{ name: 'Anchored', displayed }] }).getWeek('44'), /no recognizable worksheet week markers/);
});

test('valid weeks absent from configured tabs return empty results without calendar/date inference', () => {
  const displayed = block({ arrivals: ['11/04/2026'] });
  displayed[0] = ['Week 43'];
  const result = load({ tabs: [{ name: 'Only 43', displayed }] }).getWeek('44');
  assert.deepEqual(result, { week: 'WK 44', stores: [], totalPallets: 0, warnings: [] });
});

test('missing, malformed, out-of-range and conflicting sheet markers reject the full batch', () => {
  for (const marker of [['WK'], ['WK', ''], ['WK 0'], ['WK 54'], ['WK 440'], ['WK440'],
    ['Week 4.4'], ['WK', 'nope'], ['WK 44', 'Week 45']]) {
    const displayed = [...block(), ...block()];
    displayed[8] = marker;
    assert.throws(() => load({ tabs: [{ name: 'Bad marker', displayed }] }).getWeek('44'),
      /Week marker.*Bad marker|Conflicting week markers.*Bad marker/);
  }
  for (const displayed of [[['', '']], block().slice(1), [['some unrelated table']]]) {
    assert.throws(() => load({ tabs: [{ name: 'No markers', displayed }] }).getWeek('44'),
      /No markers.*no recognizable worksheet week markers.*ARRIVE dates do not determine weeks/);
  }
});

test('week boundaries isolate local headers and products even when the next week has no PO', () => {
  const first = block({ stores: ['VIRGINIA'], rows: [['Kung Pao', 143]] });
  first[0] = ['WK 43'];
  const selected = block({ stores: ['CAROLINE'], pos: ['0008'], rows: [['Garlic Asiago Dip', 286]] });
  const displayed = [...first, ...selected, ['WK', '45'], ['', 'Unknown next-week store'],
    ['Mystery product', 'bad quantity']];
  const { getWeek } = load({ tabs: [{ name: 'Regions', displayed }] });
  assert.deepEqual(getWeek('44').stores.map(store => store.destination), ['CAROLINE']);
  assert.equal(getWeek('44').totalPallets, 2);
  assert.throws(() => getWeek('45'), /invalid case quantity/);
  const missing = block({ stores: ['CAROLINE'] });
  missing.splice(1, 1);
  assert.throws(() => load({ tabs: [{ name: 'Missing local header', displayed: [...first, ...missing] }] })
    .getWeek('44'), /Missing destination header/);
});

test('week mode does not parse irrelevant weeks malformed metadata, products, quantities or headers', () => {
  const irrelevant = block({ arrivals: ['not a date'], pos: ['#REF!'], rows: [['Mystery', 'bad quantity']] });
  irrelevant[0] = ['WK 43'];
  irrelevant.splice(1, 1);
  irrelevant.push(['ARRIVE (RDD)', 'another bad date']);
  const displayed = [...irrelevant, ...block()];
  assert.equal(load({ tabs: [{ name: 'Isolation', displayed }] }).getWeek('44').totalPallets, 1);
});

test('week mode accepts missing ARRIVE rows and blank dates, including blank dates with populated PO', () => {
  for (const displayed of [block().filter(row => row[0] !== 'ARRIVE (RDD)'), block({ arrivals: [''] })]) {
    const result = load({ tabs: [{ name: 'No arrival', displayed }] }).getWeek('44');
    assert.equal(result.totalPallets, 1);
    assert.equal(result.stores[0].products[0].arrivalDate, '');
    assert.equal(result.stores[0].products[0].poNumber, '0159543664');
  }
});

test('week mode serializes typed dates in spreadsheet timezone and preserves displayed leading-zero PO', () => {
  const displayed = block();
  const typed = displayed.map(row => row.slice());
  typed[2][1] = 159543664;
  typed[4][1] = new Date('2026-11-04T01:00:00Z');
  const { getWeek, calls } = load({ tabs: [{ name: 'Dates', displayed, typed }] });
  const result = getWeek('44');
  assert.equal(result.stores[0].products[0].arrivalDate, '2026-11-03');
  assert.equal(result.stores[0].products[0].poNumber, '0159543664');
  assert.ok(calls.some(call => call.timezone === 'America/New_York'));
});

test('week mode rejects malformed or multiple supplied arrival rows in selected blocks', () => {
  for (const arrival of ['11/31/2026', '11/04/26', '#VALUE!']) {
    assert.throws(() => load({ tabs: [{ name: 'Bad arrival', displayed: block({ arrivals: [arrival] }) }] })
      .getWeek('44'), /ARRIVE \(RDD\)/);
  }
  const displayed = block();
  displayed.splice(5, 0, ['ARRIVE (RDD)', '11/04/2026']);
  assert.throws(() => load({ tabs: [{ name: 'Duplicate arrival', displayed }] }).getWeek('44'), /Multiple ARRIVE/);
  const typed = block();
  typed[4][1] = new Date(NaN);
  assert.throws(() => load({ tabs: [{ name: 'Invalid typed date', displayed: block(), typed }] }).getWeek('44'), /invalid date/);
});

test('week aggregate keeps different arrival orders separate rather than merging partial pallets', () => {
  const displayed = [
    ...block({ arrivals: ['11/04/2026'], rows: [['Kung Pao', 50], ['Kung Pao', 20]] }),
    ...block({ arrivals: ['11/05/2026'], rows: [['Kung Pao', 73]] }).slice(1),
    ...block({ arrivals: [''], rows: [['Kung Pao', 1]] }).slice(1),
    ...block({ arrivals: ['11/04/2026'], pos: ['0009'], rows: [['Kung Pao', 1]] }).slice(1)
  ];
  const result = load({ tabs: [{ name: 'Arrivals', displayed }] }).getWeek('44');
  assert.equal(result.totalPallets, 4);
  assert.deepEqual(result.stores[0].products.map(entry => [entry.poNumber, entry.arrivalDate, entry.totalCases]),
    [['0159543664', '2026-11-04', 70], ['0159543664', '2026-11-05', 73],
      ['0159543664', '', 1], ['0009', '2026-11-04', 1]]);
});

test('week mode sums same-arrival partial lines and deduplicates repeated per-product views within/across tabs', () => {
  const displayed = [
    ...block({ rows: [['Kung Pao', 50], ['Kung Pao', 20], ['Shelled Edamame', 160]] }),
    ...block({ rows: [['Kung Pao', 73]] }).slice(1),
    ...block({ rows: [['JIT EDAMAME', 160]] }).slice(1)
  ];
  const result = load({ tabs: [{ name: 'Weekly', displayed },
    { name: 'Repeated', displayed: block({ rows: [['JIT EDAMAME', 160]] }) }] }).getWeek('44');
  assert.equal(result.totalPallets, 2);
  assert.deepEqual(result.stores[0].products.map(entry => entry.totalCases), [143, 160]);
  assert.ok(result.warnings.some(warning => warning.includes('duplicated within tab')));
  assert.ok(result.warnings.some(warning => warning.includes('duplicated across tabs')));
});

test('identical cases on different arrivals are not deduplicated, while repeated blank-arrival views are', () => {
  const displayed = [
    ...block({ arrivals: ['11/04/2026'], rows: [['Kung Pao', 70]] }),
    ...block({ arrivals: ['11/05/2026'], rows: [['Kung Pao', 70]] }).slice(1),
    ...block({ arrivals: [''], rows: [['Kung Pao', 70]] }).slice(1),
    ...block({ arrivals: [''], rows: [['Kung Pao', 70]] }).slice(1)
  ];
  const result = load({ tabs: [{ name: 'Date deduplication', displayed }] }).getWeek('44');
  assert.equal(result.totalPallets, 3);
  assert.equal(result.stores[0].totalCases, 210);
  assert.equal(result.warnings.filter(warning => warning.includes('Ignored identical order')).length, 1);
});

test('each configured tab must contain valid week markers even if another has selected orders', () => {
  for (const displayed of [block().slice(1), [['WK 54']]]) {
    assert.throws(() => load({ tabs: [{ name: 'Valid', displayed: block() }, { name: 'Invalid', displayed }] })
      .getWeek('44'), /Invalid.*no recognizable worksheet week markers|Week marker.*Invalid/);
  }
});

test('a marker-only week cannot lend its local header to a later selected region', () => {
  const missing = block();
  missing.splice(1, 1);
  const displayed = [['WK 43'], ['', 'VIRGINIA'], ['Mystery', 123], ...missing];
  assert.throws(() => load({ tabs: [{ name: 'Boundary', displayed }] }).getWeek('44'), /Missing destination header/);
});

test('selected week invalid quantities, PO, unknown/blank product or missing destinations fail the whole batch', () => {
  const badBlocks = [
    block({ rows: [['Unknown product', 1]] }), block({ rows: [['', 1]] }),
    block({ pos: [''] }), block({ pos: ['#REF!'] }),
    ...[-1, 1.5, Infinity, NaN, '#N/A', '143 cases'].map(value => block({ rows: [['Kung Pao', value]] }))
  ];
  const missingHeader = block();
  missingHeader.splice(1, 1);
  badBlocks.push(missingHeader);
  const orphan = block({ stores: ['Nazareth', 'CAROLINE'], pos: ['0001', ''], arrivals: ['', ''],
    rows: [['Kung Pao', 143, 1]] });
  orphan[1][2] = '';
  orphan[3][2] = '';
  badBlocks.push(orphan);
  for (const bad of badBlocks) {
    assert.throws(() => load({ tabs: [{ name: 'Bad selected week', displayed: [...block(), ...bad] }] }).getWeek('44'),
      /Unknown product|blank product name|PO#|invalid case quantity|Missing destination/);
  }
});

test('week mode preserves zero skips, extra product aliases, capacities and JIT schema', () => {
  const displayed = [
    ['WK', 44],
    ['JIT EDAMAME', 'Nazareth', 'CAROLINE', 'TOTAL CASES'],
    ['PO#', '001', '002', ''],
    ['JIT EDAMAME', 161, 160, 321],
    ['Kung Pao', 144, 143, 287],
    ['Holiday Entree Turkey Pot Pie', 91, 90, 181],
    ['Mystery zero', 0, '', 0],
    ['total pallets', 6, 3, 9]
  ];
  const result = load({ tabs: [{ name: 'JIT', displayed }] }).getWeek('44');
  assert.equal(result.totalPallets, 9);
  assert.deepEqual(result.stores[0].products.map(entry => [entry.product.casesPerPallet, entry.totalPallets]),
    [[160, 2], [143, 2], [90, 2]]);
  const properties = { EAST_COAST_PRODUCTS: '[{"name":"Verified","casesPerPallet":100,"aliases":["Custom"]}]' };
  assert.equal(load({ properties, tabs: [{ name: 'Custom', displayed: block({ rows: [['Custom', 101]] }) }] })
    .getWeek('44').totalPallets, 2);
});

test('week mode rejects unsafe aggregate counts and invalid configured capacities', () => {
  const displayed = block({ rows: [['Kung Pao', Number.MAX_SAFE_INTEGER], ['Kung Pao', 1]] });
  assert.throws(() => load({ tabs: [{ name: 'Overflow', displayed }] }).getWeek('44'), /safe whole-number range/);
  assert.throws(() => load({ properties: {
    EAST_COAST_PRODUCTS: '[{"name":"Kung Pao","casesPerPallet":0}]'
  } }).getWeek('44'), /EAST_COAST_PRODUCTS/);
});

test('week API provides helpful spreadsheet open, missing-tab and read-access failures', () => {
  assert.throws(() => load({ properties: { EAST_COAST_SPREADSHEET_ID: '' } }).getWeek('44'), /EAST_COAST_SPREADSHEET_ID/);
  assert.throws(() => load({ accessError: 'Permission denied' }).getWeek('44'), /grant.*read access.*Permission denied/);
  assert.throws(() => load({ properties: { EAST_COAST_SHEET_NAMES: '["Missing"]' } }).getWeek('44'), /tab not found.*Missing/);
  assert.throws(() => load({ tabs: [{ name: 'Read', error: 'Denied' }] }).getWeek('44'), /Cannot read.*Read.*read access/);
});

test('week marker and local destinations can share a row, including split WK and adjacent number', () => {
  for (const [header, metadata, product] of [
    [['WK 44', 'NAZARETH', 'DAYTONA'], ['PO#', '0001', '0002'], ['Kung Pao', 143, 286]],
    [['WK', 44, 'NAZARETH', 'DAYTONA'], ['PO#', '', '0001', '0002'], ['Kung Pao', '', 143, 286]],
    [['', 'NAZARETH', 'DAYTONA', 'Week : 44'], ['PO#', '0001', '0002', ''], ['Kung Pao', 143, 286, '']]
  ]) {
    const displayed = [header, metadata, product];
    const result = load({ tabs: [{ name: 'EAST COAST', displayed }] }).getWeek('WK 44');
    assert.deepEqual(result.stores.map(store => store.destination), ['NAZARETH', 'DAYTONA']);
    assert.equal(result.totalPallets, 3);
  }
});

test('week parser reports original worksheet PO and product row numbers after slicing', () => {
  for (const [rows, expected] of [
    [[['Unknown', 1]], /Unknown product.*PO# row 11.*row 13/],
    [[['Kung Pao', 'bad quantity']], /PO# row 11.*row 13.*invalid case quantity/]
  ]) {
    const first = block();
    first[0] = ['WK 43'];
    const second = block({ rows }).filter(row => row[0] !== 'ARRIVE (RDD)');
    assert.throws(() => load({ tabs: [{ name: 'EAST COAST', displayed: [...first, ...second] }] }).getWeek('44'), expected);
  }
});

test('selected week nonzero products without PO blocks fail rather than producing incomplete empty results', () => {
  for (const label of ['Kung Pao', 'Unknown product']) {
    const displayed = [['WK', 44], ['', 'NAZARETH'], [label, 143]];
    assert.throws(() => load({ tabs: [{ name: 'EAST COAST', displayed }] }).getWeek('44'),
      /Missing PO# order block.*EAST COAST.*row 3/);
  }
  const displayed = [['WK 44'], ['Kung Pao', 143], ...block().slice(1)];
  assert.throws(() => load({ tabs: [{ name: 'EAST COAST', displayed }] }).getWeek('44'), /Missing PO# order block/);
  for (const displayed of [[['WK 44']], [['WK', 44]], [['WK 44'], ['', 'NAZARETH'], ['Kung Pao', 0]]]) {
    assert.equal(load({ tabs: [{ name: 'EAST COAST', displayed }] }).getWeek('44').totalPallets, 0);
  }
});

test('descriptive week-start metadata and summaries are not numeric worksheet markers or product rows', () => {
  const displayed = [
    ['Week of 10/26/2026'],
    ['Week starts', '10/26/2026'],
    ...block({ rows: [
      ['Kung Pao', 143],
      ['Week starts 10/26/2026', 999],
      ['Week of 10/26/2026', 999],
      ['Week total', 999],
      ['WK planning notes', 999],
      ['total cases', 999]
    ] })
  ];
  const result = load({ tabs: [{ name: 'EAST COAST', displayed }] }).getWeek('WK 44');
  assert.equal(result.totalPallets, 1);
  assert.equal(result.stores[0].products.length, 1);
  assert.equal(result.stores[0].totalCases, 143);
  assert.throws(() => load({ tabs: [{ name: 'Metadata only', displayed: displayed.slice(0, 2) }] }).getWeek('44'),
    /no recognizable worksheet week markers/);
});

test('unlabeled nonzero cells without a PO or before the first PO cannot silently disappear', () => {
  for (const displayed of [
    [['WK44'], ['', 'NAZARETH'], ['', 143]],
    [['WK44'], ['', 143], ...block().slice(1)]
  ]) {
    assert.throws(() => load({ tabs: [{ name: 'EAST COAST', displayed }] }).getWeek('WK 44'),
      /Missing PO# order block for nonzero cases.*EAST COAST.*row [23]/);
  }
  const displayed = [['WK', 44], ['', 'NAZARETH'], ['', 0], ['Week starts', '10/26/2026'], ['total cases', 143]];
  assert.equal(load({ tabs: [{ name: 'EAST COAST', displayed }] }).getWeek('WK 44').totalPallets, 0);
});

test('invalid nonblank quantities in local store columns fail without a PO and before the first PO', () => {
  for (const label of ['Kung Pao', 'Unknown product', 'JIT EDAMAME', '']) {
    for (const appendPo of [false, true]) {
      const displayed = [['WK44'], ['', 'NAZARETH'], [label, 'bad']];
      if (appendPo) displayed.push(...block().slice(1));
      assert.throws(() => load({ tabs: [{ name: 'EAST COAST', displayed }] }).getWeek('WK 44'),
        /EAST COAST.*row 3.*invalid case quantity.*bad/);
    }
  }
  for (const displayed of [
    [['WK44'], ['', 'NAZARETH']],
    [['WK', 44, 'NAZARETH', 'DAYTONA']],
    [['WK44'], ['JIT EDAMAME', 'NAZARETH']],
    [['WK44'], ['', 'NAZARETH'], ['', 'NAZARETH'], ['Unknown product', 0]]
  ]) {
    assert.equal(load({ tabs: [{ name: 'EAST COAST', displayed }] }).getWeek('WK 44').totalPallets, 0);
  }
});
