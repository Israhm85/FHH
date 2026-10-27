const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// Minimal DOM for exercising the real client event handlers without dependencies.
class Element {
  constructor() {
    this.children = [];
    this.listeners = {};
    this.attributes = {};
    this.classList = {
      add() {},
      remove() {}
    };
    this.value = '';
    this.disabled = false;
    this.hidden = false;
    this._text = '';
    this._html = '';
  }
  set textContent(value) { this._text = value; this.children = []; this._html = ''; }
  get textContent() { return this._text; }
  set innerHTML(value) { this._html = value; this.children = []; }
  get innerHTML() { return this._html; }
  appendChild(child) {
    if (child.fragment) { this.children.push(...child.children); }
    else { this.children.push(child); }
    return child;
  }
  setAttribute(name, value) { this.attributes[name] = value; }
  getAttribute(name) { return this.attributes[name]; }
  addEventListener(name, callback) {
    (this.listeners[name] ||= []).push(callback);
  }
  fire(name) {
    (this.listeners[name] || []).forEach(callback => callback({ preventDefault() {} }));
  }
  checkValidity() { return true; }
  scrollIntoView() {}
}

function setup() {
  const elements = {};
  const el = id => elements[id] ||= new Element();
  const product = { name: 'SHELLED EDAMAME', size: '9 oz', casesPerPallet: 160, manual: true };
  el('eastCoastCatalog').textContent = JSON.stringify([product]);
  el('btnPrint').disabled = true;
  const body = new Element();
  body.setAttribute('data-client', 'east-coast');
  const pending = [];
  let prints = 0;
  const context = vm.createContext({
    document: {
      getElementById: el,
      createElement: () => new Element(),
      createDocumentFragment: () => Object.assign(new Element(), { fragment: true }),
      body
    },
    window: { print() { prints++; } },
    google: { script: { run: {
      withSuccessHandler(success) {
        return {
          withFailureHandler(failure) {
            return {
              getEastCoastOrders(date) { pending.push({ success, failure, date }); },
              getEastCoastOrdersByWeek(week) { pending.push({ success, failure, week }); }
            };
          }
        };
      }
    } } }
  });
  ['PlacardApp.html', 'SheetPlacardApp.html'].forEach(file => {
    const source = fs.readFileSync(path.join(__dirname, '..', file), 'utf8');
    vm.runInContext(source.replace(/<\/?script>/gi, ''), context);
  });
  function load(date = '2026-11-04') {
    el('arrivalDate').value = date;
    el('arrivalDate').fire('input');
    el('sheetForm').fire('submit');
    return pending.at(-1);
  }
  function result(date = '2026-11-04') {
    return {
      date, totalPallets: 3, warnings: [],
      stores: ['NAZARETH', 'DAYTONA'].map((destination, i) => ({
        destination, totalPallets: i + 1,
        products: [{ product, poNumber: '0159543664', totalCases: 160 * (i + 1), totalPallets: i + 1 }]
      }))
    };
  }
  function loadWeek(week = '44') {
    el('orderWeek').value = week;
    el('orderWeek').fire('input');
    el('sheetForm').fire('submit');
    return pending.at(-1);
  }
  function weekResult(week = 'WK 44') {
    const value = result();
    delete value.date;
    value.week = week;
    value.stores.forEach((store, i) => {
      store.products[0].arrivalDate = i === 0 ? '2026-11-04' : '2026-11-06';
    });
    return value;
  }
  return { el, body, pending, context, load, result, loadWeek, weekResult, prints: () => prints };
}

test('catalog injection cannot close its script tag and manual uses defaults only', () => {
  let defaultsOnly = false;
  const context = vm.createContext({
    getEastCoastProducts_(defaults) {
      defaultsOnly = defaults;
      return [{ name: '</script><script>alert(1)</script>', manual: true }];
    }
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), context);
  const json = context.getEastCoastCatalogJson_();
  assert.equal(defaultsOnly, true);
  assert.doesNotMatch(json, /</);
  assert.equal(JSON.parse(json)[0].name, '</script><script>alert(1)</script>');
});

test('manual generation and per-product numbering still work', () => {
  const app = setup();
  app.el('poNumber').value = '0007';
  app.el('destination').value = 'NAZARETH';
  app.el('totalCases').value = '161';
  app.el('productSelect').value = '0';
  app.el('placardForm').fire('submit');
  const cards = app.el('printArea').children;
  assert.equal(cards.length, 2);
  assert.match(cards[1].innerHTML, /\(2\) OF \(2\)/);
  assert.match(cards[0].innerHTML, /PO#0007/);
  app.el('btnPrint').fire('click');
  assert.equal(app.prints(), 1);
});

test('placard builder prints short labels, escapes them and falls back to the original name', () => {
  const app = setup();
  const labels = ['corn pudding', 'roasted vegetables', 'cranberry stuffing',
    'scalloped potatoes', 'vainilla sweet potatoes', 'pot pie', '<img onerror=x>'];
  labels.forEach(label => {
    const card = app.context.buildPlacard({
      palletNumber: 1, totalPallets: 3, poNumber: '0159543664', destination: 'NAZARETH',
      product: { name: 'Original worksheet name', placardName: label, size: '' }
    });
    assert.ok(card.innerHTML.includes(app.context.escapeHtml(label)));
    assert.doesNotMatch(card.innerHTML, /Original worksheet name|<img/i);
    assert.match(card.innerHTML, /\(1\) OF \(3\)/);
    assert.match(card.innerHTML, /0159543664/);
    assert.match(card.innerHTML, /name--short-label/);
  });
  const fallback = app.context.buildPlacard({
    palletNumber: 1, totalPallets: 1, poNumber: '0001', destination: 'DAYTONA',
    product: { name: 'SHELLED EDAMAME', size: '9 oz' }
  });
  assert.match(fallback.innerHTML, /SHELLED EDAMAME/);
});

test('sheet mode groups all stores, selects isolated batch, preserves PO and escapes strings', () => {
  const app = setup();
  app.el('ecSheetMode').fire('click');
  const request = app.load();
  assert.equal(app.el('btnLoadOrders').disabled, true);
  assert.equal(app.el('btnPrintSheet').disabled, true);
  assert.equal(app.prints(), 0);
  request.success(app.result());
  assert.equal(app.el('printArea').children.length, 3);
  assert.equal(app.el('sheetSummary').children.length, 2);
  assert.equal(app.el('btnLoadOrders').disabled, false);
  assert.match(app.el('arrivalDisplay').textContent, /11\/04\/2026/);
  app.el('sheetStore').value = '0';
  app.el('sheetStore').fire('change');
  assert.equal(app.el('printArea').children.length, 1);
  assert.match(app.el('printArea').children[0].innerHTML, /PO#0159543664/);
  assert.doesNotMatch(app.el('printArea').children[0].innerHTML, /DAYTONA/);
  app.el('btnPrintSheet').fire('click');
  assert.equal(app.prints(), 1);
  app.body.setAttribute('data-client', 'ids');
  app.el('btnPrintSheet').fire('click');
  assert.equal(app.prints(), 1);
  const card = app.context.buildPlacard({
    palletNumber: 1, totalPallets: 1, destination: '<script>', poNumber: '"&',
    product: { name: '<img onerror=x>', nameClass: '" onclick="x' }
  });
  assert.doesNotMatch(card.innerHTML, /<script>|<img|class="[^"]*" onclick=/i);
  assert.match(card.innerHTML, /&lt;img/);
  assert.doesNotMatch(card.innerHTML, /undefined/);
});

test('date changes, failures, mode changes and out-of-order responses cannot print stale cards', () => {
  const app = setup();
  app.el('ecSheetMode').fire('click');
  const first = app.load();
  first.success(app.result());
  const second = app.load('2026-11-05');
  assert.equal(app.el('printArea').children.length, 0);
  app.el('btnPrintSheet').fire('click');
  assert.equal(app.prints(), 0);
  const third = app.load('2026-11-06');
  second.success(app.result('2026-11-05'));
  second.failure({ message: 'late failure' });
  assert.equal(app.el('btnLoadOrders').disabled, true);
  third.failure({ message: '<bad sheet>' });
  assert.equal(app.el('sheetStatus').textContent, '<bad sheet>');
  assert.equal(app.el('btnPrintSheet').disabled, true);
  assert.equal(app.el('printArea').children.length, 0);
  const fourth = app.load();
  app.el('arrivalDate').fire('input');
  app.el('ecManualMode').fire('click');
  fourth.success(app.result());
  assert.equal(app.el('printArea').children.length, 0);
  assert.equal(app.el('placardForm').hidden, false);
});

test('pallet numbering restarts for each item, never uses the store/order total', () => {
  const app = setup();
  app.el('ecSheetMode').fire('click');
  const result = app.result();
  result.stores = [{
    destination: 'NAZARETH', totalPallets: 6,
    products: [
      { product: { name: 'Holiday Creamed Corn Pudding', casesPerPallet: 143 },
        poNumber: '0159543664', totalCases: 572, totalPallets: 4 },
      { product: { name: 'Roasted vegetables w/ garlic', casesPerPallet: 143 },
        poNumber: '0159543664', totalCases: 286, totalPallets: 2 }
    ]
  }];
  result.totalPallets = 6;
  app.load().success(result);
  const cards = app.el('printArea').children;
  assert.equal(cards.length, 6);
  assert.match(cards[0].innerHTML, /\(1\) OF \(4\)/);
  assert.match(cards[3].innerHTML, /\(4\) OF \(4\)/);
  assert.match(cards[4].innerHTML, /\(1\) OF \(2\)/);
  assert.match(cards[5].innerHTML, /\(2\) OF \(2\)/);
  cards.forEach(card => assert.doesNotMatch(card.innerHTML, /OF \(6\)/));
});

test('no orders and invalid/mismatched responses leave print disabled', () => {
  const app = setup();
  app.el('ecSheetMode').fire('click');
  app.load().success({ date: '2026-11-04', stores: [], totalPallets: 0 });
  assert.match(app.el('sheetStatus').textContent, /No hay órdenes/);
  assert.equal(app.el('btnPrintSheet').disabled, true);
  app.load().success(app.result('2026-04-11'));
  assert.match(app.el('sheetStatus').textContent, /inválida/);
  assert.equal(app.el('printArea').children.length, 0);
});

test('week mode requests worksheet WK and prints all stores regardless of arrival dates', () => {
  const app = setup();
  app.el('ecWeekMode').fire('click');
  assert.equal(app.el('arrivalField').hidden, true);
  assert.equal(app.el('arrivalDate').disabled, true);
  assert.equal(app.el('orderWeek').required, true);
  assert.equal(app.el('weekField').hidden, false);
  assert.equal(app.el('placardForm').hidden, true);
  const request = app.loadWeek('044');
  assert.equal(request.week, 'WK 44');
  assert.equal(request.date, undefined);
  assert.equal(app.el('ecWeekMode').disabled, true);
  assert.equal(app.el('btnPrintSheet').disabled, true);
  request.success(app.weekResult());
  assert.equal(app.el('weekDisplay').textContent, 'WK 44');
  assert.equal(app.el('printArea').children.length, 3);
  assert.match(app.el('sheetStatus').textContent, /WK 44 — 3 placards en 2 tiendas/);
  assert.match(app.el('sheetSummary').children[1].children[1].children[0].textContent, /2026-11-06/);
  app.el('sheetStore').value = '1';
  app.el('sheetStore').fire('change');
  const cards = app.el('printArea').children;
  assert.equal(cards.length, 2);
  assert.match(cards[0].innerHTML, /DAYTONA/);
  assert.match(cards[0].innerHTML, /PO#0159543664/);
  assert.match(cards[0].innerHTML, /\(1\) OF \(2\)/);
  assert.match(cards[1].innerHTML, /\(2\) OF \(2\)/);
  assert.doesNotMatch(cards[0].innerHTML, /NAZARETH/);
  assert.equal(app.prints(), 0);
  app.el('btnPrintSheet').fire('click');
  assert.equal(app.prints(), 1);
  app.body.setAttribute('data-client', 'reporteras');
  app.el('btnPrintSheet').fire('click');
  assert.equal(app.prints(), 1);
});

test('week changes and cross-mode out-of-order responses never restore stale batches', () => {
  const app = setup();
  app.el('ecWeekMode').fire('click');
  app.loadWeek().success(app.weekResult());
  const oldWeek = app.loadWeek('45');
  assert.equal(app.el('printArea').children.length, 0);
  app.el('btnPrintSheet').fire('click');
  assert.equal(app.prints(), 0);
  const newWeek = app.loadWeek('46');
  oldWeek.success(app.weekResult('WK 45'));
  oldWeek.failure({ message: 'late' });
  assert.equal(app.el('btnLoadOrders').disabled, true);
  app.el('orderWeek').fire('input');
  app.el('ecSheetMode').fire('click');
  assert.equal(app.el('arrivalDate').required, true);
  assert.equal(app.el('orderWeek').disabled, true);
  const date = app.load();
  newWeek.success(app.weekResult('WK 46'));
  assert.equal(app.el('printArea').children.length, 0);
  date.success(app.result());
  assert.equal(app.el('printArea').children.length, 3);
  app.el('ecWeekMode').fire('click');
  assert.equal(app.el('printArea').children.length, 0);
  date.failure({ message: 'late date failure' });
  const week = app.loadWeek();
  week.failure({ message: 'Bad EAST COAST tab' });
  assert.equal(app.el('sheetStatus').textContent, 'Bad EAST COAST tab');
  assert.equal(app.el('btnPrintSheet').disabled, true);
  app.el('ecManualMode').fire('click');
  assert.equal(app.el('placardForm').hidden, false);
});

test('WK 44 Nazareth renders exactly 21 cards with item-specific totals', () => {
  const app = setup();
  app.el('ecWeekMode').fire('click');
  const result = app.weekResult();
  const quantities = [572, 286, 143, 286, 143, 143, 143, 429, 429, 143, 180];
  const counts = [4, 2, 1, 2, 1, 1, 1, 3, 3, 1, 2];
  result.stores[0].products = quantities.map((totalCases, i) => ({
    product: { name: `Fixture item ${i}`, size: '', casesPerPallet: i === 10 ? 90 : 143 },
    poNumber: '0159543664', arrivalDate: '2026-11-04',
    totalCases, totalPallets: counts[i]
  }));
  result.stores[0].totalPallets = 21;
  result.totalPallets = 23;
  app.loadWeek().success(result);
  assert.equal(app.el('printArea').children.length, 23);
  app.el('sheetStore').value = '0';
  app.el('sheetStore').fire('change');
  const cards = app.el('printArea').children;
  assert.equal(cards.length, 21);
  let offset = 0;
  counts.forEach((count, item) => {
    for (let pallet = 1; pallet <= count; pallet++) {
      const html = cards[offset++].innerHTML;
      assert.ok(html.includes(`(${pallet}) OF (${count})`));
      assert.ok(html.includes(`Fixture item ${item}`));
      assert.match(html, /PO#0159543664/);
      assert.doesNotMatch(html, /OF \(21\)/);
    }
  });
});

test('invalid, missing, mismatched and empty weeks leave printing disabled', () => {
  const app = setup();
  app.el('ecWeekMode').fire('click');
  ['', '0', '54', '-1', '44.5', 'WK 44', 'NaN'].forEach(week => {
    app.loadWeek(week);
    assert.equal(app.pending.length, 0);
    assert.match(app.el('sheetStatus').textContent, /semana válida/);
    assert.equal(app.el('btnPrintSheet').disabled, true);
  });
  app.loadWeek('53').success({ week: 'WK 53', stores: [], totalPallets: 0 });
  assert.match(app.el('sheetStatus').textContent, /No hay órdenes.*WK 53/);
  app.loadWeek().success(app.weekResult('WK 45'));
  assert.match(app.el('sheetStatus').textContent, /inválida/);
  assert.equal(app.el('printArea').children.length, 0);
  app.loadWeek().success(app.result());
  assert.match(app.el('sheetStatus').textContent, /inválida/);
  assert.equal(app.el('btnPrintSheet').disabled, true);
});
