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
            return { getEastCoastOrders(date) { pending.push({ success, failure, date }); } };
          }
        };
      }
    } } }
  });
  ['PlacardApp.html', 'SheetPlacardApp.html'].forEach(file => {
    const source = fs.readFileSync(path.join(__dirname, '..', file), 'utf8');
    vm.runInContext(source.replace(/<\/?script>/g, ''), context);
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
  return { el, body, pending, context, load, result, prints: () => prints };
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
  assert.doesNotMatch(card.innerHTML, /<script>|<img|class="[^"]*" onclick=/);
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
