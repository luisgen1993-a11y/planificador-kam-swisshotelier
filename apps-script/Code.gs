/**
 * Swiss Hotelier – Planificador KAM (backend en Google Apps Script)
 *
 * Hojas del Spreadsheet:
 *   - Primera hoja (la existente): un registro por día con las tareas completadas.
 *   - "Ventas": un registro por venta (se crea sola y se llena con las ventas
 *     que había guardadas dentro de los días la primera vez que se usa).
 *
 * Seguridad: todas las peticiones deben traer el parámetro `token`, que se compara
 * con la propiedad de script TOKEN (Configuración del proyecto → Propiedades de secuencia de comandos).
 */

var SALES_SHEET_NAME = 'Ventas';
var SALES_HEADERS = ['ID', 'Fecha', 'Cliente', 'Tipo', 'Monto', 'Diferido', 'Fecha Setup', 'Actualizado'];
var PAID_TYPES = ['activation', 'upsell'];

var DEFAULT_SETTINGS = {
  baseSalary: 1016,
  internet: 50,
  recurrentHotels: 50,
  recurrentRate: 12,
  bonusGoal: 4,
  bonusAmount: 200,
  monthlyTarget: 4583.33,
  hoursPerMonth: 160
};

// ---------------------------------------------------------------- entrada

function doGet(e) {
  return handle_(e);
}

function doPost(e) {
  return handle_(e);
}

function handle_(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var expected = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!expected) {
      return json_({ status: 'error', code: 'no_token_configured',
        message: 'Falta la propiedad TOKEN en el Apps Script.' });
    }
    if (params.token !== expected) {
      return json_({ status: 'error', code: 'unauthorized', message: 'Clave incorrecta.' });
    }

    var action = params.action;
    ensureSalesSheet_();

    switch (action) {
      case 'getDay':       return json_({ status: 'success', completedTasks: getDay_(params.date) });
      case 'saveDay':      return withLock_(function () { saveDay_(params.date, parse_(params.completedTasks, {})); });
      case 'getSales':     return json_({ status: 'success', sales: readSales_() });
      case 'upsertSale':   return withLock_(function () { upsertSale_(parse_(params.sale, null)); });
      case 'deleteSale':   return withLock_(function () { deleteSale_(params.id); });
      case 'getSettings':  return json_({ status: 'success', settings: getSettings_() });
      case 'saveSettings': return withLock_(function () { saveSettings_(parse_(params.settings, {})); });
      default:             return json_({ status: 'active' });
    }
  } catch (error) {
    return json_({ status: 'error', message: String(error) });
  }
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    fn();
    return json_({ status: 'success' });
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- días

function getDaysSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== SALES_SHEET_NAME) return sheets[i];
  }
  return ss.insertSheet('Días');
}

function findDayRow_(sheet, date) {
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] && formatCellDate_(rows[i][0]) === date) return { index: i + 1, values: rows[i] };
  }
  return null;
}

function getDay_(date) {
  if (!isDate_(date)) throw new Error('Fecha inválida');
  var row = findDayRow_(getDaysSheet_(), date);
  if (!row || !row.values[4]) return {};
  var payload = parse_(row.values[4], {});
  return payload.completedTasks || {};
}

function saveDay_(date, completedTasks) {
  if (!isDate_(date)) throw new Error('Fecha inválida');
  var sheet = getDaysSheet_();
  var count = Object.keys(completedTasks).filter(function (k) { return completedTasks[k] === true; }).length;
  var summary = daySalesSummary_(date);
  var values = ["'" + date, 'Tareas: ' + count, summary.total, summary.detail,
    JSON.stringify({ date: date, completedTasks: completedTasks })];

  var row = findDayRow_(sheet, date);
  if (row) {
    sheet.getRange(row.index, 1, 1, 5).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
}

/** Actualiza las columnas resumen (C y D) de un día, si ese día ya existe en la hoja. */
function refreshDaySummary_(date) {
  if (!isDate_(date)) return;
  var sheet = getDaysSheet_();
  var row = findDayRow_(sheet, date);
  if (!row) return;
  var summary = daySalesSummary_(date);
  sheet.getRange(row.index, 3, 1, 2).setValues([[summary.total, summary.detail]]);
}

function daySalesSummary_(date) {
  var sales = readSales_().filter(function (s) { return s.date === date; });
  var total = sales.reduce(function (sum, s) {
    return PAID_TYPES.indexOf(s.type) >= 0 ? sum + Number(s.amount || 0) : sum;
  }, 0);
  var detail = sales.map(function (s) {
    return (s.client || 'Cliente') + ' ($' + (s.amount || 0) + ' - ' + (s.type || 'setup') + ')';
  }).join(', ') || 'Sin ventas';
  return { total: total, detail: detail };
}

// ---------------------------------------------------------------- ventas

function getSalesSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SALES_SHEET_NAME);
}

/**
 * Crea la hoja "Ventas" si no existe y la llena una sola vez con las ventas que estaban
 * guardadas dentro de cada día. Si la misma venta aparece en varios días (por ejemplo,
 * registrada en Setup y luego activada otro día), gana la versión del día más reciente.
 */
function ensureSalesSheet_() {
  if (getSalesSheet_()) return;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (getSalesSheet_()) return;

    var daysSheet = getDaysSheet_();
    var rows = daysSheet.getDataRange().getValues().slice(1)
      .filter(function (r) { return r[0] && r[4]; })
      .map(function (r) { return { date: formatCellDate_(r[0]), json: r[4] }; })
      .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });

    var byId = {};
    var order = [];
    rows.forEach(function (row) {
      var payload = parse_(row.json, {});
      (payload.salesList || []).forEach(function (sale) {
        var id = sale.id || ('sale_' + String(sale.client || '').replace(/\s+/g, '_') + '_' + row.date);
        if (!byId[id]) order.push(id);
        byId[id] = normalizeSale_(sale, id, row.date);
      });
    });

    var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SALES_SHEET_NAME);
    sheet.appendRow(SALES_HEADERS);
    sheet.setFrozenRows(1);
    if (order.length) {
      var data = order.map(function (id) { return saleToRow_(byId[id]); });
      sheet.getRange(2, 1, data.length, SALES_HEADERS.length).setValues(data);
    }
  } finally {
    lock.releaseLock();
  }
}

function normalizeSale_(sale, id, fallbackDate) {
  return {
    id: String(id || sale.id),
    date: String(sale.date || fallbackDate || ''),
    client: String(sale.client || ''),
    type: String(sale.type || 'setup'),
    amount: Number(sale.amount || 0),
    deferredToNextMonth: sale.deferredToNextMonth === true || sale.deferredToNextMonth === 'true',
    setupDate: String(sale.setupDate || '')
  };
}

function saleToRow_(s) {
  return [s.id, "'" + s.date, s.client, s.type, s.amount, s.deferredToNextMonth,
    s.setupDate ? "'" + s.setupDate : '', new Date()];
}

function readSales_() {
  var sheet = getSalesSheet_();
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        id: String(r[0]),
        date: formatCellDate_(r[1]),
        client: String(r[2]),
        type: String(r[3]),
        amount: Number(r[4] || 0),
        deferredToNextMonth: r[5] === true || r[5] === 'TRUE' || r[5] === 'true',
        setupDate: r[6] ? formatCellDate_(r[6]) : ''
      };
    });
}

function findSaleRow_(sheet, id) {
  var ids = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (var i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function upsertSale_(sale) {
  if (!sale || !sale.id) throw new Error('Venta sin ID');
  var clean = normalizeSale_(sale, sale.id, '');
  if (!isDate_(clean.date)) throw new Error('Fecha de venta inválida');

  var sheet = getSalesSheet_();
  var rowIndex = findSaleRow_(sheet, clean.id);
  var previousDate = null;

  if (rowIndex > 0) {
    previousDate = formatCellDate_(sheet.getRange(rowIndex, 2).getValue());
    sheet.getRange(rowIndex, 1, 1, SALES_HEADERS.length).setValues([saleToRow_(clean)]);
  } else {
    sheet.appendRow(saleToRow_(clean));
  }

  refreshDaySummary_(clean.date);
  if (previousDate && previousDate !== clean.date) refreshDaySummary_(previousDate);
}

function deleteSale_(id) {
  var sheet = getSalesSheet_();
  var rowIndex = findSaleRow_(sheet, id);
  if (rowIndex < 0) return;
  var date = formatCellDate_(sheet.getRange(rowIndex, 2).getValue());
  sheet.deleteRow(rowIndex);
  refreshDaySummary_(date);
}

// ---------------------------------------------------------------- ajustes

function getSettings_() {
  var saved = parse_(PropertiesService.getScriptProperties().getProperty('SETTINGS'), {});
  var result = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
    result[k] = (saved[k] !== undefined && saved[k] !== null) ? Number(saved[k]) : DEFAULT_SETTINGS[k];
  });
  return result;
}

function saveSettings_(incoming) {
  var current = getSettings_();
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
    var v = Number(incoming[k]);
    if (incoming[k] !== undefined && !isNaN(v) && v >= 0) current[k] = v;
  });
  PropertiesService.getScriptProperties().setProperty('SETTINGS', JSON.stringify(current));
}

// ---------------------------------------------------------------- utilidades

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function parse_(raw, fallback) {
  if (raw === undefined || raw === null || raw === '') return fallback;
  try { return JSON.parse(raw); } catch (err) { return fallback; }
}

function isDate_(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
}

function formatCellDate_(cellValue) {
  if (!cellValue) return '';
  if (cellValue instanceof Date) {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(cellValue, tz, 'yyyy-MM-dd');
  }
  return String(cellValue).replace(/^'/, '').trim();
}
