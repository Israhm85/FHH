var TJ_HEADERS = ['PROD_DATE', 'ARRIVE_DATE', 'LOCATION', 'PO_NUMBER', 'ORDER_TYPE', 'PRODUCT', 'QTY', 'UOM', 'NOTES'];
var TJ_LABEL = '[TRJO]';
var TJ_PRODUCTS = {
  'HOL SIDE CORN PUDDING': 'Holiday Side Corn Pudding',
  'HOL SIDE VANILLA BEAN WHIPPED SWEET POTATOES': 'Holiday Vanilla Bean Whipped Sweet Potatoes',
  'HOL SIDE VAINILLA BEAN WHIPPED SWEET POTATOES': 'Holiday Vanilla Bean Whipped Sweet Potatoes'
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Trader Joe’s')
    .addItem('Importar PO de Gmail', 'importarPedidosTraderJoes')
    .addItem('Dar formato a filas existentes', 'formatearPedidosTraderJoes')
    .addToUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    ss.toast('No se aplicó el formato porque otra operación está en curso.', 'Trader Joe’s', 5);
    return;
  }
  try {
    if (tjFormatExistingRows_(ss) === false) {
      ss.toast('No se aplicó el formato: las filas 2 y 3 deben ser pedidos Regular con el formato deseado.',
        'Trader Joe’s', 8);
    }
  } finally {
    lock.releaseLock();
  }
}

function formatearPedidosTraderJoes() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('Otra operación está en curso. Inténtalo nuevamente.');
  try {
    var count = tjFormatExistingRows_(SpreadsheetApp.getActiveSpreadsheet());
    if (count === false) {
      throw new Error('Las filas 2 y 3 deben ser pedidos Regular con el formato que quieres copiar.');
    }
    ui.alert('Formato aplicado', count
      ? 'Se formatearon ' + count + ' filas desde la fila 4.'
      : 'No hay filas de pedidos para formatear.', ui.ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}

function tjFormatExistingRows_(ss) {
  var sheet = tjDatos_(ss);
  var count = sheet.getLastRow() - 3;
  if (count < 1) return 0;
  var types = sheet.getRange(4, 5, count, 1).getDisplayValues().map(function(row) { return row[0]; });
  return tjFormatRows_(sheet, 4, types, true) ? count : false;
}

function tjDatos_(ss) {
  var sheet = ss.getSheetByName('Datos');
  if (!sheet) throw new Error('No existe la pestaña Datos.');
  var headers = sheet.getRange(1, 1, 1, TJ_HEADERS.length).getDisplayValues()[0];
  if (headers.some(function(value, i) { return value.trim() !== TJ_HEADERS[i]; })) {
    throw new Error('Las primeras nueve columnas de Datos no coinciden con los encabezados esperados.');
  }
  return sheet;
}

function tjFormatRows_(sheet, first, types, preserveNumberFormats) {
  var count = types.length;
  if (!count) return true;
  if (first < 4 || sheet.getRange(2, 5, 2, 1).getDisplayValues().some(function(row) {
    return String(row[0]).trim().toLowerCase() !== 'regular';
  })) return false;
  var target = sheet.getRange(first, 1, count, TJ_HEADERS.length);
  var numberFormats = preserveNumberFormats ? target.getNumberFormats() : null;
  var heights = [sheet.getRowHeight(2), sheet.getRowHeight(3)];
  types.forEach(function(_, i) {
    var row = first + i;
    var template = row % 2 === 0 ? 2 : 3;
    if (row !== template) {
      sheet.getRange(template, 1, 1, TJ_HEADERS.length).copyTo(
        sheet.getRange(row, 1, 1, TJ_HEADERS.length), { formatOnly: true });
    }
    sheet.setRowHeight(row, heights[template - 2]);
  });
  var alignments = ['center', 'center', 'center', 'center', 'center', 'left', 'right', 'left', 'left'];
  target.setBackgrounds(types.map(function(type, i) {
    var demo = String(type).trim().toLowerCase() === 'demo';
    var color = demo ? '#fef3c9' : ((first + i) % 2 === 0 ? '#f8fafc' : '#ffffff');
    return Array(TJ_HEADERS.length).fill(color);
  }));
  target.setHorizontalAlignments(types.map(function() { return alignments.slice(); }));
  target.setVerticalAlignment('middle');
  target.setBorder(true, true, true, true, true, true, '#cbd5e0', SpreadsheetApp.BorderStyle.SOLID);
  types.forEach(function(type, i) {
    if (String(type).trim().toLowerCase() === 'demo') {
      sheet.getRange(first + i, 1, 1, TJ_HEADERS.length).setFontWeight('bold');
    }
  });
  if (numberFormats) target.setNumberFormats(numberFormats);
  return true;
}

function importarPedidosTraderJoes() {
  var ui = SpreadsheetApp.getUi();
  var start = ui.prompt('Importar PO', 'Fecha inicial ARRIVE (AAAA-MM-DD). Se revisan correos de los últimos 10 días:', ui.ButtonSet.OK_CANCEL);
  if (start.getSelectedButton() !== ui.Button.OK) return;
  var end = ui.prompt('Importar PO', 'Fecha final ARRIVE (AAAA-MM-DD), inclusive:', ui.ButtonSet.OK_CANCEL);
  if (end.getSelectedButton() !== ui.Button.OK) return;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('Otra importación está en curso. Inténtalo nuevamente.');
  try {
    var report = tjImport_(SpreadsheetApp.getActiveSpreadsheet(), start.getResponseText().trim(), end.getResponseText().trim());
    ui.alert('Importación terminada', report, ui.ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}

function tjImport_(ss, start, end, now) {
  tjParseIso_(start);
  tjParseIso_(end);
  if (start > end) throw new Error('La fecha inicial debe ser anterior o igual a la final.');
  var sheet = tjDatos_(ss);
  var timezone = ss.getSpreadsheetTimeZone();
  if (timezone !== 'America/Los_Angeles') {
    throw new Error('Configura la zona horaria de la hoja como America/Los_Angeles antes de importar.');
  }
  var today = Utilities.formatDate(now || new Date(), timezone, 'yyyy-MM-dd');
  var receivedFrom = tjShiftDate_(today, -9);
  var query = 'label:' + TJ_LABEL + ' after:' + tjShiftDate_(receivedFrom, -1).replace(/-/g, '/') +
    ' before:' + tjShiftDate_(today, 2).replace(/-/g, '/');
  var orders = {};
  var offset = 0;
  var threads;
  do {
    threads = GmailApp.search(query, offset, 100);
    threads.forEach(function(thread) {
      thread.getMessages().forEach(function(message) {
        var date = Utilities.formatDate(message.getDate(), timezone, 'yyyy-MM-dd');
        if (date < receivedFrom || date > today) return;
        var body = message.getPlainBody();
        if (!/PURCHASE AGREEMENT/i.test(body)) return;
        var order;
        try {
          var arrive = tjArrivalFromBody_(body);
          if (arrive < start || arrive > end) return;
          order = tjParseOrder_(body);
        } catch (e) {
          throw new Error('No se pudo leer el correo ' + message.getId() + ': ' + e.message);
        }
        var key = tjPoKey_(order.po);
        if (orders[key] && tjOrderSignature_(orders[key]) !== tjOrderSignature_(order)) {
          throw new Error('El PO ' + order.po + ' aparece con datos distintos en los correos. Revisa antes de importar.');
        }
        if (!orders[key]) orders[key] = order;
      });
    });
    offset += threads.length;
  } while (threads.length === 100);

  var lastRow = sheet.getLastRow();
  var existing = {};
  if (lastRow > 1) {
    var range = sheet.getRange(2, 1, lastRow - 1, TJ_HEADERS.length);
    var values = range.getValues();
    var displayed = range.getDisplayValues();
    values.forEach(function(row, i) {
      var po = String(displayed[i][3]).trim();
      if (!po) return;
      var key = tjPoKey_(po);
      if (!existing[key]) existing[key] = [];
      existing[key].push(tjExistingSignature_(row, displayed[i], timezone));
    });
  }
  var rows = [];
  var skipped = [];
  var changed = [];
  var formatWarning = '';
  Object.keys(orders).sort().forEach(function(key) {
    var order = orders[key];
    if (existing[key]) {
      var oldSignature = existing[key].slice().sort().join('\n');
      var newSignature = order.items.map(function(item) { return tjItemSignature_(order, item); }).sort().join('\n');
      if (oldSignature !== newSignature) changed.push(order.po);
      else skipped.push(order.po);
      return;
    }
    order.items.forEach(function(item) {
      rows.push([tjSheetDate_(tjShiftDate_(order.arrive, -1), timezone), tjSheetDate_(order.arrive, timezone),
        order.location, order.po, order.type, item.product, item.qty, item.uom, '']);
    });
  });
  if (rows.length) {
    var first = lastRow + 1;
    var requiredRows = first + rows.length - 1 - sheet.getMaxRows();
    if (requiredRows > 0) sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows);
    if (!tjFormatRows_(sheet, first, rows.map(function(row) { return row[4]; }), false)) {
      formatWarning = ' No se aplicó el formato: las filas 2 y 3 deben ser pedidos Regular con el formato deseado.';
    }
    sheet.getRange(first, 4, rows.length, 1).setNumberFormat('@');
    sheet.getRange(first, 1, rows.length, 2).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(first, 1, rows.length, TJ_HEADERS.length).setValues(rows);
  }
  return 'Correos revisados del ' + receivedFrom + ' al ' + today + '. ARRIVE del ' + start + ' al ' + end +
    '. Filas nuevas: ' + rows.length + '. PO existentes sin cambios: ' + skipped.length +
    '. PO existentes con diferencias (no modificados): ' + (changed.join(', ') || 'ninguno') + '.' + formatWarning;
}

function tjPoKey_(po) {
  var value = String(po).trim();
  return /^\d+$/.test(value) ? value.replace(/^0+(?=\d)/, '') : value;
}

function tjParseOrder_(body) {
  var poMatches = body.match(/^[ \t]*P\.O\.[ \t]*:/gim) || [];
  if (poMatches.length !== 1) throw new Error('Se esperaba exactamente un P.O. por correo.');
  function field(name) {
    var match = body.match(new RegExp('^[ \\t]*' + name + '[ \\t]*:[ \\t]*([^\\r\\n]+)', 'im'));
    if (!match || !match[1].trim()) throw new Error('Falta ' + name + '.');
    return match[1].trim();
  }
  var po = field('P\\.O\\.');
  if (!/^\d+$/.test(po)) throw new Error('P.O. no numérico: ' + po);
  var locationText = field('DELIVER TO').toUpperCase();
  var matches = locationText.match(/\b(?:PALMDALE|FONTANA|STOCKTON|SACRAMENTO|LACEY|ARIZONA|TEXAS|COLORADO)\b/g) || [];
  matches = matches.filter(function(value, index) { return matches.indexOf(value) === index; });
  if (matches.length !== 1) throw new Error('Ubicación desconocida o ambigua: ' + locationText);
  var arrive = tjArrivalFromBody_(body);

  var sections = body.split(/^[ \t]*PRODUCT NAME[ \t]*:[ \t]*/im);
  if (sections.length < 2) throw new Error('No se encontraron productos.');
  var productLines = body.match(/^[ \t]*PRODUCT NAME\b/gim) || [];
  if (productLines.length !== sections.length - 1) throw new Error('Hay encabezados de producto inválidos.');
  var seen = {};
  var items = sections.slice(1).map(function(section) {
    var name = section.split(/\r?\n/)[0].trim();
    var sku = section.match(/^[ \t]*SKU[ \t]*:[ \t]*([^\s\r\n]+)/im);
    var qtyLines = section.match(/^[ \t]*(?:Case|Each)[ \t]+QTY[ \t]*:/gim) || [];
    var qty = section.match(/^[ \t]*(Case|Each)[ \t]+QTY[ \t]*:[ \t]*([\d,]+)[ \t]*$/im);
    if (!name || name.charAt(0) === '=' || !sku || qtyLines.length !== 1 || !qty || !/^\d{1,3}(?:,\d{3})*$|^\d+$/.test(qty[2])) {
      throw new Error('Producto, SKU o cantidad inválidos: ' + name);
    }
    var quantity = Number(qty[2].replace(/,/g, ''));
    if (!Number.isSafeInteger(quantity) || quantity <= 0) throw new Error('Cantidad inválida: ' + name);
    var key = sku[1].toUpperCase();
    if (seen[key]) throw new Error('SKU repetido en el PO: ' + key);
    seen[key] = true;
    var normalizedName = name.replace(/\s+/g, ' ').toUpperCase();
    return { sku: key, product: TJ_PRODUCTS[normalizedName] || name.replace(/\s+/g, ' '),
      qty: quantity, uom: qty[1].toUpperCase() === 'CASE' ? 'CASES' : 'UNITS' };
  });
  return { po: po, location: matches[0], arrive: arrive, type: items.length < 8 ? 'Demo' : 'Regular', items: items };
}

function tjArrivalFromBody_(body) {
  var poMatches = body.match(/^[ \t]*P\.O\.[ \t]*:/gim) || [];
  if (poMatches.length !== 1) throw new Error('Se esperaba exactamente un P.O. por correo.');
  var matches = body.match(/^[ \t]*ARRIVE[ \t]*:[ \t]*([^\r\n]*)/gim) || [];
  if (matches.length !== 1) throw new Error('Se esperaba exactamente una fecha ARRIVE por correo.');
  var rawDate = matches[0].replace(/^[ \t]*ARRIVE[ \t]*:[ \t]*/i, '').trim();
  var dateMatch = rawDate.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (!dateMatch) throw new Error('Fecha ARRIVE inválida: ' + rawDate);
  var year = Number(dateMatch[3]);
  if (dateMatch[3].length === 2) year += 2000;
  var arrive = year + '-' + ('0' + dateMatch[1]).slice(-2) + '-' + ('0' + dateMatch[2]).slice(-2);
  tjParseIso_(arrive);
  return arrive;
}

function tjParseIso_(value) {
  var match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) throw new Error('Fecha inválida: ' + value + '. Usa AAAA-MM-DD.');
  var date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12));
  if (date.getUTCFullYear() !== Number(match[1]) || date.getUTCMonth() !== Number(match[2]) - 1 ||
      date.getUTCDate() !== Number(match[3])) throw new Error('Fecha inválida: ' + value);
  return date;
}

function tjShiftDate_(value, days) {
  var date = tjParseIso_(value);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function tjSheetDate_(value, timezone) {
  return Utilities.parseDate(value + ' 12:00', timezone, 'yyyy-MM-dd HH:mm');
}

function tjItemSignature_(order, item) {
  return JSON.stringify([order.arrive, order.location, order.type, item.product, item.qty, item.uom]);
}

function tjOrderSignature_(order) {
  return order.items.map(function(item) { return item.sku + '|' + tjItemSignature_(order, item); }).sort().join('\n');
}

function tjExistingSignature_(row, displayed, timezone) {
  var date = row[1] instanceof Date ? Utilities.formatDate(row[1], timezone, 'yyyy-MM-dd') : String(displayed[1]).trim();
  return JSON.stringify([date, String(displayed[2]).trim(), String(displayed[4]).trim(),
    String(displayed[5]).trim(), Number(row[6]), String(displayed[7]).trim()]);
}
