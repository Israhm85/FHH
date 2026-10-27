function getEastCoastProducts_(defaultsOnly) {
  var products = [
    { name: 'SHELLED EDAMAME', size: '9 oz', casesPerPallet: 160, nameClass: 'name--edamame', aliases: ['Edamame', 'JIT EDAMAME'], manual: true },
    { name: 'STUFFED PEPPERS WITH BEEF', size: '15 oz', casesPerPallet: 143, nameClass: 'name--stuffed-peppers', aliases: ['Stuffed peppers beef', 'Stuffed peppers w/ beef'], manual: true },
    { name: 'KUNG PAO BRUSSEL SPROUTS', size: '14 oz', casesPerPallet: 143, nameClass: 'name--kung-pao', aliases: ['Kung Pao', 'Kung Pao Brussels Sprouts'], manual: true },
    { name: 'GARLIC ASIAGO DIP', size: '10 oz', casesPerPallet: 143, nameClass: 'name--garlic', aliases: [], manual: true },
    { name: 'CHICKEN STUFFED POBLANO', size: '15 oz', casesPerPallet: 143, nameClass: 'name--chicken', aliases: ['Stuffed poblanos chicken', 'Chicken stuffed poblanos', 'Stuffed poblano chicken'], manual: true },
    { name: 'ROASTED GREEN VEGETABLES', size: '12 oz', casesPerPallet: 143, nameClass: 'name--roasted', aliases: [], manual: true },
    { name: 'Holiday Creamed Corn Pudding', placardName: 'corn pudding', size: '', casesPerPallet: 143, nameClass: 'name--holiday', aliases: [], manual: false },
    { name: 'Roasted vegetables w/ garlic', placardName: 'roasted vegetables', size: '', casesPerPallet: 143, nameClass: 'name--holiday', aliases: ['Roasted vegetables with garlic'], manual: false },
    { name: 'Holiday Apple Cranberry Stuffing', placardName: 'cranberry stuffing', size: '', casesPerPallet: 143, nameClass: 'name--holiday', aliases: [], manual: false },
    { name: 'Four cheese scalloped potatoes', placardName: 'scalloped potatoes', size: '', casesPerPallet: 143, nameClass: 'name--holiday', aliases: [], manual: false },
    { name: 'Loaded mashed potatoes', size: '', casesPerPallet: 143, nameClass: 'name--holiday', aliases: [], manual: false },
    { name: 'Holiday Vanilla Bean Whipped Sw…', placardName: 'vainilla sweet potatoes', size: '', casesPerPallet: 143, nameClass: 'name--holiday', aliases: ['Holiday Vanilla Bean Whipped Sw...', 'Holiday Vanilla Bean Whipped Sw', 'Holiday Vanilla Bean Whipped Sweet Potatoes'], manual: false },
    { name: 'Holiday Entree Turkey Pot Pie', placardName: 'pot pie', size: '', casesPerPallet: 90, nameClass: 'name--holiday', aliases: [], manual: false }
  ];
  var raw = defaultsOnly === true ? null : PropertiesService.getScriptProperties().getProperty('EAST_COAST_PRODUCTS');
  if (raw) {
    var extras = eastCoastJsonArray_(raw, 'EAST_COAST_PRODUCTS');
    extras.forEach(function (extra) {
      if (!extra || typeof extra.name !== 'string' || !extra.name.trim() ||
          !eastCoastCapacity_(extra.casesPerPallet) ||
          (extra.size !== undefined && typeof extra.size !== 'string') ||
          (extra.nameClass !== undefined && !/^name--[a-z0-9-]+$/.test(extra.nameClass)) ||
          (extra.aliases !== undefined && (!Array.isArray(extra.aliases) || extra.aliases.some(function (alias) { return typeof alias !== 'string' || !alias.trim(); })))) {
        throw new Error('EAST_COAST_PRODUCTS must contain definitions with a name, positive integer casesPerPallet, and optional size, nameClass and string aliases.');
      }
      var existing = products.filter(function (product) { return eastCoastText_(product.name) === eastCoastText_(extra.name); })[0];
      if (existing) {
        if (existing.casesPerPallet !== extra.casesPerPallet) {
          throw new Error('EAST_COAST_PRODUCTS cannot change the known pallet capacity for ' + existing.name + '.');
        }
        if (extra.size !== undefined) existing.size = extra.size;
        existing.aliases = existing.aliases.concat(extra.aliases || []);
      } else {
        products.push({ name: extra.name.trim(), size: extra.size || '', casesPerPallet: extra.casesPerPallet,
          nameClass: extra.nameClass || 'name--holiday', aliases: extra.aliases || [], manual: false });
      }
    });
  }
  eastCoastProductIndex_(products);
  return products;
}

function getEastCoastOrders(date) {
  eastCoastIsoDate_(date, 'Arrival date');
  return eastCoastFetchOrders_({ date: date });
}

function getEastCoastOrdersByWeek(week) {
  return eastCoastFetchOrders_({ week: eastCoastWeek_(week, 'Order week') });
}

function eastCoastFetchOrders_(selection) {
  var properties = PropertiesService.getScriptProperties();
  var id = properties.getProperty('EAST_COAST_SPREADSHEET_ID');
  if (!id || !id.trim()) {
    throw new Error('Set the EAST_COAST_SPREADSHEET_ID Script Property to the East Coast orders spreadsheet ID.');
  }
  var namesRaw = properties.getProperty('EAST_COAST_SHEET_NAMES');
  if (!namesRaw || !namesRaw.trim()) {
    throw new Error('Set EAST_COAST_SHEET_NAMES to a nonempty JSON array of explicit East Coast orders tab names, for example ["EAST COAST"]. East Coast tab configuration is required; first-tab fallback is not allowed.');
  }
  var names = eastCoastJsonArray_(namesRaw, 'EAST_COAST_SHEET_NAMES');
  if (!names.length || names.some(function (name) { return typeof name !== 'string' || !name.trim(); }) ||
      names.some(function (name, index) { return names.indexOf(name) !== index; })) {
    throw new Error('EAST_COAST_SHEET_NAMES must be a nonempty JSON array of unique sheet names.');
  }
  var spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(id.trim());
  } catch (error) {
    throw new Error('Cannot open the East Coast orders spreadsheet. Check EAST_COAST_SPREADSHEET_ID and grant the web app executing account read access. ' + error.message);
  }
  var catalog = getEastCoastProducts_();
  var sheets = names.map(function (name) {
    var sheet = spreadsheet.getSheetByName(name);
    if (!sheet) throw new Error('East Coast orders tab not found: "' + name + '". Check EAST_COAST_SHEET_NAMES.');
    return sheet;
  });
  if (!sheets.length) throw new Error('The East Coast orders spreadsheet has no sheets.');
  var timezone = spreadsheet.getSpreadsheetTimeZone();
  var shipments = [];
  sheets.forEach(function (sheet) {
    var typed, displayed;
    try {
      var range = sheet.getDataRange();
      typed = range.getValues();
      displayed = range.getDisplayValues();
    } catch (error) {
      throw new Error('Cannot read East Coast orders tab "' + sheet.getName() + '". Check spreadsheet read access. ' + error.message);
    }
    if (sheet.getName() === 'EAST COAST') {
      // The confirmed orders layout ends at I; unrelated J+ cells are not orders.
      typed = typed.map(function (row) { return row.slice(0, 9); });
      displayed = displayed.map(function (row) { return row.slice(0, 9); });
    }
    // Capture week boundaries before removing column I item totals and labels.
    var weekMarkers = selection.week ? eastCoastWeekMarkers_(displayed, sheet.getName()) : null;
    if (sheet.getName() === 'EAST COAST') {
      var orders = eastCoastWithoutItemTotals_(typed, displayed, sheet.getName());
      typed = orders.typed;
      displayed = orders.displayed;
    }
    shipments = shipments.concat(selection.week ?
      parseEastCoastWeekSheet_(typed, displayed, timezone, selection.week, catalog, sheet.getName(), weekMarkers) :
      parseEastCoastSheet_(typed, displayed, timezone, selection.date, catalog, sheet.getName()));
  });
  return eastCoastAggregate_(selection.date, shipments, selection.week);
}

function eastCoastWithoutItemTotals_(typed, displayed, sheetName) {
  // In the confirmed EAST COAST layout, I contains item totals, not store orders.
  var totalColumn = 8;
  displayed.forEach(function (row, index) {
    var marker = eastCoastRowMarker_(row);
    if (!marker) return;
    var raw = (typed[index] || [])[totalColumn];
    var text = String(row[totalColumn] === undefined ? '' : row[totalColumn]).trim();
    if (text || raw !== undefined && raw !== null && raw !== '') {
      throw new Error('Column I is configured as item totals in "' + sheetName +
        '", but contains order metadata at row ' + (index + 1) +
        '. Check the sheet layout before generating placards.');
    }
  });
  function withoutTotals(rows) {
    return rows.map(function (row) {
      var copy = row.slice();
      if (copy.length > totalColumn) copy[totalColumn] = '';
      return copy;
    });
  }
  return { typed: withoutTotals(typed), displayed: withoutTotals(displayed) };
}

function eastCoastWeek_(value, context) {
  var match = typeof value === 'string' && /^(?:(?:wk|week)\s*:?\s*)?(\d{1,2})$/i.exec(eastCoastText_(value));
  if (!match || Number(match[1]) < 1 || Number(match[1]) > 53) {
    throw new Error(context + ' must be a worksheet week from WK 1 through WK 53 (for example "WK 44"). No calendar week is inferred.');
  }
  return 'WK ' + Number(match[1]);
}

function eastCoastWeekMarkers_(displayed, sheetName) {
  var markers = [];
  displayed.forEach(function (row, rowIndex) {
    var week = '';
    var columns = [];
    row.forEach(function (cell, column) {
      var text = eastCoastText_(cell);
      if (!/^(?:wk|week)(?:\b|\d)/.test(text)) return;
      if (!/^(?:wk|week)\s*:?\s*$/.test(text) && !/^(?:wk|week)\s*:?\s*[+-]?\d/.test(text)) return;
      var context = 'Week marker in "' + sheetName + '", row ' + (rowIndex + 1);
      if (/^(?:wk|week)\s*:?\s*$/.test(text)) {
        text += ' ' + eastCoastText_(row[column + 1]);
        columns.push(column + 1);
      }
      var found = eastCoastWeek_(text, context);
      if (week && week !== found) throw new Error('Conflicting week markers in "' + sheetName + '", row ' + (rowIndex + 1) + ': ' + week + ' and ' + found + '.');
      week = found;
      columns.push(column);
    });
    if (week) markers.push({ row: rowIndex, week: week, columns: columns });
  });
  if (!markers.length) throw new Error('East Coast orders tab "' + sheetName + '" has no recognizable worksheet week markers. Add an explicit WK 44 or Week 44 marker; ARRIVE dates do not determine weeks.');
  return markers;
}

function parseEastCoastWeekSheet_(typed, displayed, timezone, week, catalog, sheetName, weekMarkers) {
  week = eastCoastWeek_(week, 'Order week');
  sheetName = sheetName || 'Orders';
  var markers = weekMarkers || eastCoastWeekMarkers_(displayed, sheetName);
  var shipments = [];
  markers.forEach(function (marker, index) {
    if (marker.week !== week) return;
    var start = marker.row;
    var end = index + 1 < markers.length ? markers[index + 1].row : displayed.length;
    var rows = displayed.slice(start, end);
    var values = typed.slice(start, end);
    rows[0] = rows[0].slice();
    values[0] = (values[0] || []).slice();
    marker.columns.forEach(function (column) { rows[0][column] = ''; values[0][column] = ''; });
    if (!rows.some(function (row) { var value = eastCoastRowMarker_(row); return value && value.kind === 'po'; })) {
      eastCoastCheckMissingWeekPo_(values, rows, catalog, sheetName, start);
      return;
    }
    // Slice at literal markers so neither headers nor products can cross week boundaries.
    shipments = shipments.concat(parseEastCoastSheet_(values, rows, timezone, '', catalog, sheetName, true, start));
  });
  return shipments;
}

function eastCoastCheckMissingWeekPo_(typed, displayed, catalog, sheetName, rowOffset) {
  var products = eastCoastProductIndex_(catalog);
  var header = [];
  displayed.forEach(function (row, rowIndex) {
    if (eastCoastRowMarker_(row)) return;
    var label = eastCoastProductLabel_(row, header.length ? header[0].column : row.length, products);
    if (label.summary) return;
    var candidate = eastCoastHeader_(row, header.length ? header[0].column - 1 : 0, products);
    if (candidate.length && (!header.length || candidate.every(function (store) {
      return header.some(function (existing) { return existing.column === store.column && eastCoastText_(existing.destination) === eastCoastText_(store.destination); });
    }))) {
      header = candidate;
      return;
    }
    var productColumn = -1;
    row.forEach(function (cell, column) { if (products[eastCoastText_(cell)]) productColumn = column; });
    for (var column = 0; column < row.length; column++) {
      var raw = (typed[rowIndex] || [])[column];
      var text = String(row[column] === undefined ? '' : row[column]).trim();
      var orderColumn = header.some(function (store) { return store.column === column; }) ||
        productColumn >= 0 && column > productColumn;
      if (!orderColumn && typeof raw !== 'number' && !/^[+-]?(?:\d|#)/.test(text)) continue;
      var context = '"' + sheetName + '", row ' + (rowOffset + rowIndex + 1) + ', column ' + (column + 1);
      if (eastCoastQuantity_(raw, text, context)) {
        throw new Error('Missing PO# order block for nonzero cases in selected week in ' + context + '.');
      }
    }
  });
}

function eastCoastJsonArray_(raw, property) {
  var value;
  try { value = JSON.parse(raw); } catch (error) { throw new Error(property + ' must be a valid JSON array.'); }
  if (!Array.isArray(value)) throw new Error(property + ' must be a JSON array.');
  return value;
}

function eastCoastText_(value) {
  return String(value === undefined || value === null ? '' : value).replace(/\u00a0/g, ' ').trim().replace(/\s+/g, ' ').toLowerCase();
}

function eastCoastCapacity_(value) {
  return typeof value === 'number' && isFinite(value) && value > 0 && Math.floor(value) === value && value <= Number.MAX_SAFE_INTEGER;
}

function eastCoastProductIndex_(catalog) {
  var index = Object.create(null);
  catalog.forEach(function (product) {
    [product.name].concat(product.aliases || []).forEach(function (alias) {
      var key = eastCoastText_(alias);
      if (index[key] && index[key] !== product) throw new Error('Ambiguous East Coast product alias: ' + alias + '.');
      index[key] = product;
    });
  });
  return index;
}

function eastCoastIsoDate_(value, context) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    throw new Error(context + ' must be a strict ISO date (YYYY-MM-DD).');
  }
  var year = Number(value.slice(0, 4)), month = Number(value.slice(5, 7)), day = Number(value.slice(8, 10));
  var leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  var days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (year < 1 || month < 1 || month > 12 || day < 1 || day > days[month - 1]) {
    throw new Error(context + ' contains an invalid calendar date: ' + value + '.');
  }
  return value;
}

function eastCoastCellDate_(typed, displayed, timezone, context) {
  if (Object.prototype.toString.call(typed) === '[object Date]') {
    if (!isFinite(typed.getTime())) throw new Error(context + ' contains an invalid date.');
    return eastCoastIsoDate_(Utilities.formatDate(typed, timezone, 'yyyy-MM-dd'), context);
  }
  var text = String(displayed === undefined || displayed === null || displayed === '' ? typed || '' : displayed).trim();
  if (!text) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return eastCoastIsoDate_(text, context);
  var match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (match) return eastCoastIsoDate_(match[3] + '-' + ('0' + match[1]).slice(-2) + '-' + ('0' + match[2]).slice(-2), context);
  throw new Error(context + ' must be MM/DD/YYYY, YYYY-MM-DD, or a spreadsheet date; found "' + text + '".');
}

function eastCoastMarker_(value) {
  var text = eastCoastText_(value);
  if (/^po\s*(?:#|number|no\.?)\s*:?\s*$/.test(text)) return 'po';
  if (/^(?:pick\s*up|pickup)\s*date\s*:?\s*$/.test(text)) return 'pickup';
  if (/^(?:arrive(?:\s*\(rdd\))?|arrival(?:\s*date)?|rdd)\s*:?\s*$/.test(text)) return 'arrival';
  if (/^(?:total(?:\s+(?:cases|pallets))?|pallets|cases|v[\s-]*boards?|week(?:\s.*)?|wk(?:\s.*)?)\s*:?\s*$/.test(text)) return 'summary';
  return '';
}

function eastCoastRowMarker_(row) {
  for (var column = 0; column < row.length; column++) {
    var kind = eastCoastMarker_(row[column]);
    if (kind === 'po' || kind === 'arrival' || kind === 'pickup') return { kind: kind, column: column };
  }
  return null;
}

function eastCoastHeader_(row, labelColumn, products) {
  var header = [];
  for (var label = 0; label <= labelColumn; label++) {
    var heading = eastCoastText_(row[label]);
    if (heading && !/^(?:products?|items?|description|sku|destination|stores?|dc|jit edamame|east coast(?: orders)?|week(?:\s.*)?|wk(?:\s.*)?)$/.test(heading)) return header;
  }
  for (var column = labelColumn + 1; column < row.length; column++) {
    var text = String(row[column] || '').trim();
    var normalized = eastCoastText_(text);
    if (!text || eastCoastMarker_(text) || products[normalized] ||
        /^(?:product|item|description|size|pack|capacity|cases per pallet|total.*|pallet.*|v[\s-]*board.*|week.*|wk\s*\d.*)$/i.test(text)) continue;
    if (/[a-z]/i.test(text) && !/^\d/.test(text) && text.charAt(0) !== '=') {
      header.push({ column: column, destination: text });
    }
  }
  return header;
}

function eastCoastProductLabel_(row, firstStoreColumn, products) {
  var label = '';
  for (var column = 0; column < firstStoreColumn; column++) {
    var text = String(row[column] || '').trim();
    if (!text) continue;
    if (eastCoastMarker_(text) === 'summary' || /^(?:total\b|v[\s-]*boards?\b|week\b|wk\b)/i.test(text)) return { summary: true, label: text };
    if (products[eastCoastText_(text)]) return { summary: false, label: text };
    if (!/^[\d,.]+$/.test(text)) label = text;
  }
  return { summary: false, label: label };
}

function eastCoastQuantity_(typed, displayed, context) {
  var value = typed === undefined || typed === null || typed === '' ? displayed : typed;
  if (value === undefined || value === null || String(value).trim() === '') return 0;
  if (typeof value !== 'number') {
    var text = String(value).trim();
    if (!/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(text)) throw new Error(context + ' has an invalid case quantity: "' + text + '". Use a nonnegative whole number.');
    value = Number(text.replace(/,/g, ''));
  }
  if (!isFinite(value) || value < 0 || Math.floor(value) !== value || value > Number.MAX_SAFE_INTEGER) {
    throw new Error(context + ' has an invalid case quantity. Use a finite nonnegative whole number.');
  }
  return value;
}

function parseEastCoastSheet_(typed, displayed, timezone, date, catalog, sheetName, weekMode, rowOffset) {
  if (!weekMode) eastCoastIsoDate_(date, 'Arrival date');
  rowOffset = rowOffset || 0;
  sheetName = sheetName || 'Orders';
  var products = eastCoastProductIndex_(catalog);
  if (!displayed.length || !displayed.some(function (row) { return row.some(function (cell) { return String(cell).trim() !== ''; }); })) {
    throw new Error('East Coast orders tab "' + sheetName + '" has no order data.');
  }
  var blocks = [];
  displayed.forEach(function (row, rowIndex) {
    var marker = eastCoastRowMarker_(row);
    if (marker && marker.kind === 'po') blocks.push({ row: rowIndex, labelColumn: marker.column });
  });
  if (!blocks.length) throw new Error('East Coast orders tab "' + sheetName + '" has no recognizable PO# order blocks.');
  if (weekMode) {
    eastCoastCheckMissingWeekPo_(typed.slice(0, blocks[0].row), displayed.slice(0, blocks[0].row), catalog, sheetName, rowOffset);
  }
  var shipments = [];
  blocks.forEach(function (block, blockIndex) {
    var context = '"' + sheetName + '", PO# row ' + (rowOffset + block.row + 1);
    var previous = blockIndex ? blocks[blockIndex - 1].row : -1;
    var end = blockIndex + 1 < blocks.length ? blocks[blockIndex + 1].row : displayed.length;
    var header = [], headerRow = -1;
    for (var row = block.row - 1; row > previous; row--) {
      if (eastCoastRowMarker_(displayed[row])) continue;
      header = eastCoastHeader_(displayed[row], block.labelColumn, products);
      if (header.length) { headerRow = row; break; }
    }
    if (!header.length) throw new Error('Missing destination header for East Coast orders ' + context + '. Each PO# block needs its own store header.');
    // The next block's local destination header ends this block, not its PO# row.
    if (blockIndex + 1 < blocks.length) {
      for (var next = end - 1; next > block.row; next--) {
        if (!eastCoastRowMarker_(displayed[next]) && eastCoastHeader_(displayed[next], blocks[blockIndex + 1].labelColumn, products).length) {
          end = next;
          break;
        }
      }
    }
    var arrivalRow = -1;
    for (var metadataRow = block.row + 1; metadataRow < end; metadataRow++) {
      var marker = eastCoastRowMarker_(displayed[metadataRow]);
      if (marker && marker.kind === 'arrival') {
        if (arrivalRow !== -1) throw new Error('Multiple ARRIVE (RDD) rows in ' + context + '.');
        arrivalRow = metadataRow;
      }
    }
    if (arrivalRow < 0 && !weekMode) throw new Error('Missing ARRIVE (RDD) row in East Coast orders ' + context + '.');
    var arrivalCells = displayed[arrivalRow] || [];
    var columns = Object.create(null);
    header.forEach(function (store) { columns[store.column] = true; });
    var firstStoreColumn = header[0].column;
    // An order column without a local destination must never be silently dropped.
    var width = Math.max(displayed[block.row].length, arrivalCells.length);
    for (var c = block.labelColumn + 1; c < width; c++) {
      if (columns[c] || eastCoastMarker_(displayed[headerRow][c]) ||
          /^(?:total.*|pallet.*|v[\s-]*board.*)$/i.test(String(displayed[headerRow][c] || '').trim())) continue;
      if (String(displayed[block.row][c] || '').trim() || String(arrivalCells[c] || '').trim()) {
        throw new Error('Missing destination for column ' + (c + 1) + ' in ' + context + '.');
      }
      if (c < firstStoreColumn) continue;
      for (var orphanRow = block.row + 1; orphanRow < end; orphanRow++) {
        if (eastCoastRowMarker_(displayed[orphanRow]) || eastCoastProductLabel_(displayed[orphanRow], firstStoreColumn, products).summary) continue;
        var orphanValue = (typed[orphanRow] || [])[c];
        var orphanDisplay = displayed[orphanRow][c];
        if (eastCoastQuantity_(orphanValue, orphanDisplay, context + ', column ' + (c + 1) + ', row ' + (rowOffset + orphanRow + 1))) {
          throw new Error('Missing destination for nonzero cases in column ' + (c + 1) + ' in ' + context + '.');
        }
      }
    }
    header.forEach(function (store) {
      var column = store.column;
      var storeContext = context + ', ' + store.destination;
      var arrival = eastCoastCellDate_((typed[arrivalRow] || [])[column], arrivalCells[column], timezone, 'ARRIVE (RDD) in ' + storeContext);
      var po = String(displayed[block.row][column] === undefined ? '' : displayed[block.row][column]).trim();
      if (!weekMode && !arrival && po) throw new Error('Missing ARRIVE (RDD) date for PO# ' + po + ' in ' + storeContext + '.');
      var entries = [];
      for (var productRow = block.row + 1; productRow < end; productRow++) {
        var productCells = displayed[productRow];
        if (eastCoastRowMarker_(productCells)) continue;
        var productLabel = eastCoastProductLabel_(productCells, firstStoreColumn, products);
        var label = productLabel.label;
        if (productLabel.summary) continue;
        var raw = (typed[productRow] || [])[column];
        var display = productCells[column];
        var nonblank = raw !== undefined && raw !== null && raw !== '' || display !== undefined && String(display).trim() !== '';
        if (!nonblank) continue;
        if (!weekMode && !arrival) {
          var missingDateQuantity = eastCoastQuantity_(raw, display, storeContext + ', row ' + (rowOffset + productRow + 1));
          if (missingDateQuantity) throw new Error('Missing ARRIVE (RDD) date for nonzero order in ' + storeContext + '.');
          continue;
        }
        if (!weekMode && arrival !== date) continue;
        var quantity = eastCoastQuantity_(raw, display, storeContext + ', row ' + (rowOffset + productRow + 1));
        if (!quantity) continue;
        if (!po) throw new Error('Missing PO# for nonzero order in ' + storeContext + '.');
        if (/^#(?:REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|ERROR!|SPILL!|CALC!|FIELD!|DATA!|GETTING_DATA|BLOCKED!|CONNECT!|UNKNOWN!|BUSY!|LOADING!)$/i.test(po)) {
          throw new Error('Spreadsheet error "' + po + '" in PO# for nonzero order in ' + storeContext + '. Correct the PO cell before generating placards.');
        }
        var product = products[eastCoastText_(label)];
        if (!product) throw new Error('Unknown product "' + (label || '(blank product name)') + '" with nonzero cases in ' + storeContext + ', row ' + (rowOffset + productRow + 1) + '. Add a verified definition to EAST_COAST_PRODUCTS.');
        if (!eastCoastCapacity_(product.casesPerPallet)) throw new Error('Missing or invalid pallet capacity for ' + product.name + ' in ' + storeContext + '.');
        entries.push({ product: { name: product.name, placardName: product.placardName || product.name, size: product.size || '', casesPerPallet: product.casesPerPallet, nameClass: product.nameClass },
          poNumber: po, totalCases: quantity, totalPallets: 0 });
      }
      if (entries.length) shipments.push({ destination: store.destination, poNumber: po, date: arrival, products: entries, sheet: sheetName });
    });
  });
  return shipments;
}

function eastCoastAdd_(left, right, context) {
  var sum = left + right;
  if (!Number.isSafeInteger(left) || !Number.isSafeInteger(right) || !Number.isSafeInteger(sum) || left < 0 || right < 0) {
    throw new Error('East Coast orders ' + context + ' exceed the safe whole-number range. Correct the quantities before generating placards.');
  }
  return sum;
}

function eastCoastAggregate_(date, shipments, week) {
  var result = week ? { week: week, stores: [], totalPallets: 0, warnings: [] } :
    { date: date, stores: [], totalPallets: 0, warnings: [] };
  var stores = Object.create(null), seen = Object.create(null), warningSet = Object.create(null);
  function warn(message) {
    if (!warningSet[message]) { warningSet[message] = true; result.warnings.push(message); }
  }
  shipments.forEach(function (shipment) {
    var fingerprintProducts = Object.create(null);
    var snapshotProducts = Object.create(null);
    shipment.products.forEach(function (entry) {
      var key = entry.product.name;
      fingerprintProducts[key] = eastCoastAdd_(fingerprintProducts[key] || 0, entry.totalCases, 'case totals for ' + shipment.destination + ', PO ' + shipment.poNumber + ', ' + key);
      snapshotProducts[key] = entry.product;
    });
    Object.keys(fingerprintProducts).forEach(function (name) {
      var entry = { product: snapshotProducts[name], poNumber: shipment.poNumber, totalCases: fingerprintProducts[name] };
      var fingerprint = JSON.stringify([eastCoastText_(shipment.destination), shipment.poNumber, shipment.date, name, entry.totalCases]);
      // Sum lines within a block; identical product snapshots in other blocks are repeated views.
      if (seen[fingerprint]) {
        var location = seen[fingerprint] === shipment.sheet ? 'within tab ' + shipment.sheet : 'across tabs ' + seen[fingerprint] + ' and ' + shipment.sheet;
        warn('Ignored identical order for ' + shipment.destination + ', PO ' + shipment.poNumber + ', ' + name + ' duplicated ' + location + '.');
        return;
      }
      seen[fingerprint] = shipment.sheet;
      var storeKey = eastCoastText_(shipment.destination);
      var store = stores[storeKey];
      if (!store) {
        store = { destination: shipment.destination, totalCases: 0, totalPallets: 0, products: [] };
        stores[storeKey] = store;
        result.stores.push(store);
      }
      var existing = store.products.filter(function (value) {
        return value.poNumber === entry.poNumber && value.product.name === entry.product.name && (!week || value.arrivalDate === shipment.date);
      })[0];
      if (existing) existing.totalCases = eastCoastAdd_(existing.totalCases, entry.totalCases, 'case totals for ' + shipment.destination + ', PO ' + entry.poNumber + ', ' + entry.product.name);
      else {
        var order = { product: entry.product, poNumber: entry.poNumber, totalCases: entry.totalCases, totalPallets: 0 };
        if (week) order.arrivalDate = shipment.date;
        store.products.push(order);
      }
      if (!entry.product.size) warn('Size unavailable for ' + entry.product.name + '; verify packaging before printing.');
      if (entry.product.name === 'Holiday Vanilla Bean Whipped Sw…') warn('Product name is truncated in the source: Holiday Vanilla Bean Whipped Sw…; no longer name has been inferred.');
    });
  });
  result.stores.forEach(function (store) {
    store.products.forEach(function (entry) {
      entry.totalPallets = Math.ceil(entry.totalCases / entry.product.casesPerPallet);
      store.totalCases = eastCoastAdd_(store.totalCases, entry.totalCases, 'case totals for ' + store.destination);
      store.totalPallets = eastCoastAdd_(store.totalPallets, entry.totalPallets, 'pallet totals for ' + store.destination);
    });
    result.totalPallets = eastCoastAdd_(result.totalPallets, store.totalPallets, 'pallet totals for ' + (week || date));
  });
  return result;
}
