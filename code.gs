// ╔══════════════════════════════════════════════════════════════╗
// ║  Сімейний бюджет — Google Apps Script v3.0 (Final)         ║
// ║  Файл: «Наш семейный бюджет – 2026»                        ║
// ╚══════════════════════════════════════════════════════════════╝

// ── Замініть на реальні email ──
var ALLOWED_USERS = [
  'xxxx.911@gmail.com',
  'yyyy@gmail.com'
];
var USER_PREFIX = {
  'xxxx@gmail.com': 'xx',
  'yyyy@gmail.com': 'yy'
};

// ── Листи ──
var SH = {
  expenses: 'Витрати',
  income: 'Доходи',
  business: 'Бізнес',
  settings: 'Settings',
  monthly: 'Monthly Overview',
  annual: 'Annual Overview',
  analyz: 'Аналіз',
  dMonthInc: 'Diagram Monthly Income',
  dMonthExp: 'Diagram Monthly Expenses',
  dMonthSub: 'Diagram Monthly Expenses Subcategory',
  dAnnBal: 'Diagram Annual Balance',
  dAnnInc: 'Diagram Annual Income',
  dAnnExp: 'Diagram Annual Expenses',
  dAnlInc: 'Diagram Analys Income',
  dAnlExp: 'Diagram Analys Expenses',
  dAnlSub: 'Diagram Analys Expenses Subcategory'
};

// ── Структура колонок ──
// Витрати/Доходи: A=UUID B=Created C=Updated D=Status E=Дата F=Категорія G=Сума H=Опис I=Валюта J=Курс K=USD(formula)
// Бізнес:         A=UUID B=Created C=Updated D=Status E=Дата F=Категорія G=Сума H=Опис I=Валюта J=Курс→USD K=USD(formula) L=Курс2→UAH M=UAH(formula)

// ── Клітинки фільтрів аналітики ──
var FC = {
  monthly: { month: 'K4', year: 'K6', category: 'K8' },
  annual: { month: 'B16', year: 'B20' },
  analyz: { month: 'C5', year: 'F5', category: 'B47' }
};

// ══════════════════════════════════════════
// ROUTING
// ══════════════════════════════════════════
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var email = verifyToken(data._token || '');
    if (!email) return respond({ status: 'error', message: 'Невалідний токен' });
    if (!isAllowed(email)) return respond({ status: 'error', message: 'Доступ заборонено: ' + email });

    var res;
    switch (data.action) {
      case 'addRecord': res = addRecord(data, email); break;
      case 'resolveRate': res = resolveRate(data); break;
      case 'resolveOfflineRate': res = resolveOfflineRate(data); break;
      case 'getRecords': res = getRecords(data); break;
      case 'updateRecord': res = updateRecord(data); break;
      case 'deleteRecord': res = deleteRecord(data); break;
      case 'getSettings': res = getSettings(); break;
      case 'whoAmI': res = { status: 'ok', prefix: getPrefix(email) }; break;
      case 'getFilters': res = getFilters(); break;
      case 'applyFilter': res = applyFilter(data); break;
      case 'getRateSync': res = getRateSync(data); break;
      default: res = { status: 'error', message: 'Unknown action: ' + data.action };
    }
    return respond(res);
  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  try {
    if (e.parameter.action === 'getSettings') return respond(getSettings());
    return respond({ status: 'error', message: 'Unknown GET action' });
  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

function isAllowed(email) {
  var e = (email || '').toLowerCase();
  return ALLOWED_USERS.some(function (u) { return u.toLowerCase() === e; });
}

function getPrefix(email) {
  var e = (email || '').toLowerCase();
  for (var k in USER_PREFIX) {
    if (k.toLowerCase() === e) return USER_PREFIX[k];
  }
  return 'XX';
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════
// 1. ДОДАТИ ЗАПИС
// ══════════════════════════════════════════
function addRecord(data, email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetKey = data.sheet;
  var sheet = ss.getSheetByName(SH[sheetKey]);
  if (!sheet) return { status: 'error', message: 'Лист не знайдено: ' + SH[sheetKey] };

  var now = new Date();
  var uuid = data.uuid || (getPrefix(email) + '_' + Utilities.getUuid());
  var currency = (data.currency || 'USD').toUpperCase();
  var amount = parseFloat(data.amount) || 0;
  var dateISO = (data.date || '').substring(0, 10);
  var row = getNextRow(sheet);

  sheet.getRange(row, 1).setValue(uuid);            // A: UUID
  sheet.getRange(row, 2).setValue(now);             // B: Created
  sheet.getRange(row, 3).setValue(now);             // C: Updated
  sheet.getRange(row, 4).setValue('ACTIVE');        // D: Status
  sheet.getRange(row, 5).setValue(formatDate(dateISO)); // E: Дата
  sheet.getRange(row, 6).setValue(data.category || ''); // F: Категорія
  sheet.getRange(row, 7).setValue(amount);          // G: Сума
  sheet.getRange(row, 8).setValue(data.desc || ''); // H: Опис
  sheet.getRange(row, 9).setValue(currency);        // I: Валюта

  // J: Курс→USD
  if (currency === 'USD') {
    sheet.getRange(row, 10).setValue(1);
  } else if (data.offline) {
    sheet.getRange(row, 10).setValue(''); // заповниться при синхронізації
  } else {
    sheet.getRange(row, 10).setFormula(gfFormula(currency, 'USD', dateISO));
  }
  // K: USD = formula (не чіпаємо)

  if (sheetKey === 'business') {
    // L: Курс2→UAH
    if (currency === 'UAH') {
      sheet.getRange(row, 12).setValue(1);
    } else if (data.offline) {
      sheet.getRange(row, 12).setValue('');
    } else {
      sheet.getRange(row, 12).setFormula(gfFormula(currency, 'UAH', dateISO));
    }
    // M: UAH = formula (не чіпаємо)
  }

  return { status: 'ok', row: row, uuid: uuid };
}

// ══════════════════════════════════════════
// 2. ЗАФІКСУВАТИ КУРС (фоновий виклик через 4 сек)
// ══════════════════════════════════════════
function resolveRate(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH[data.sheet]);
  if (!sheet) return { status: 'error', message: 'Лист не знайдено' };

  var row = parseInt(data.row);
  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  var cellJ = sheet.getRange(row, 10);
  var valJ = cellJ.getValue();
  if (typeof valJ === 'number' && valJ > 0) cellJ.setValue(valJ);

  if (data.sheet === 'business') {
    var cellL = sheet.getRange(row, 12);
    var valL = cellL.getValue();
    if (typeof valL === 'number' && valL > 0) cellL.setValue(valL);
  }

  return { status: 'ok' };
}

// ══════════════════════════════════════════
// 2b. ЗАПОВНИТИ КУРС ДЛЯ ОФЛАЙН-ЗАПИСУ
// ══════════════════════════════════════════
function resolveOfflineRate(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH[data.sheet]);
  if (!sheet) return { status: 'error', message: 'Лист не знайдено' };

  var row = parseInt(data.row);
  var currency = (data.currency || 'USD').toUpperCase();
  var dateISO = (data.date || '').substring(0, 10);

  if (currency === 'USD') {
    sheet.getRange(row, 10).setValue(1);
  } else {
    sheet.getRange(row, 10).setFormula(gfFormula(currency, 'USD', dateISO));
  }

  if (data.sheet === 'business') {
    if (currency === 'UAH') {
      sheet.getRange(row, 12).setValue(1);
    } else {
      sheet.getRange(row, 12).setFormula(gfFormula(currency, 'UAH', dateISO));
    }
  }

  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  var cellJ = sheet.getRange(row, 10);
  var valJ = cellJ.getValue();
  if (typeof valJ === 'number' && valJ > 0) cellJ.setValue(valJ);

  if (data.sheet === 'business') {
    var cellL = sheet.getRange(row, 12);
    var valL = cellL.getValue();
    if (typeof valL === 'number' && valL > 0) cellL.setValue(valL);
  }

  return { status: 'ok' };
}

// ══════════════════════════════════════════
// 3. ОТРИМАТИ ЗАПИСИ
// ══════════════════════════════════════════
function getRecords(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH[data.sheet]);
  if (!sheet) return { status: 'error', message: 'Лист не знайдено' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'ok', records: [] };

  var numCols = data.sheet === 'business' ? 13 : 11;
  var values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var records = [];

  values.forEach(function (row, i) {
    if (row[3] === 'DELETED') return;  // D
    if (!row[4] && !row[6]) return;    // E, G

    var d = row[4];
    var dateStr = d instanceof Date
      ? Utilities.formatDate(d, 'Europe/Kiev', 'dd.MM.yyyy')
      : String(d || '');

    var created = row[1];
    var createdStr = created instanceof Date
      ? Utilities.formatDate(created, 'Europe/Kiev', 'dd.MM.yyyy HH:mm')
      : '';

    records.push({
      row: i + 2,
      uuid: row[0] || '',
      created: createdStr,
      date: dateStr,
      category: row[5] || '',
      amount: row[6] || 0,
      desc: row[7] || '',
      currency: row[8] || '',
      rateUSD: row[9] || 0,
      usd: row[10] || 0,   // K: базова валюта (USD) — для коректних підсумків
      rateUAH: data.sheet === 'business' ? (row[11] || 0) : null,
      uah: data.sheet === 'business' ? (row[12] || 0) : null  // M
    });
  });

  records.reverse();
  return { status: 'ok', records: records.slice(0, parseInt(data.limit) || 100) };
}

// ══════════════════════════════════════════
// 4. РЕДАГУВАТИ ЗАПИС
// ══════════════════════════════════════════
function updateRecord(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH[data.sheet]);
  if (!sheet) return { status: 'error', message: 'Лист не знайдено' };

  var row = parseInt(data.row);
  var currency = (data.currency || 'USD').toUpperCase();
  var dateISO = (data.date || '').substring(0, 10);

  sheet.getRange(row, 3).setValue(new Date());             // C: Updated
  sheet.getRange(row, 5).setValue(formatDate(dateISO));    // E: Дата
  sheet.getRange(row, 6).setValue(data.category || '');    // F
  sheet.getRange(row, 7).setValue(parseFloat(data.amount) || 0); // G
  sheet.getRange(row, 8).setValue(data.desc || '');        // H
  sheet.getRange(row, 9).setValue(currency);               // I

  if (currency === 'USD') {
    sheet.getRange(row, 10).setValue(1);
  } else {
    sheet.getRange(row, 10).setFormula(gfFormula(currency, 'USD', dateISO));
  }

  if (data.sheet === 'business') {
    if (currency === 'UAH') {
      sheet.getRange(row, 12).setValue(1);
    } else {
      sheet.getRange(row, 12).setFormula(gfFormula(currency, 'UAH', dateISO));
    }
  }

  // Для редагування фіксуємо одразу (не критично затримати відповідь)
  SpreadsheetApp.flush();
  Utilities.sleep(800);

  var vJ = sheet.getRange(row, 10).getValue();
  if (typeof vJ === 'number' && vJ > 0) sheet.getRange(row, 10).setValue(vJ);

  if (data.sheet === 'business') {
    var vL = sheet.getRange(row, 12).getValue();
    if (typeof vL === 'number' && vL > 0) sheet.getRange(row, 12).setValue(vL);
  }

  return { status: 'ok' };
}

// ══════════════════════════════════════════
// 5. М'ЯКЕ ВИДАЛЕННЯ
// ══════════════════════════════════════════
function deleteRecord(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH[data.sheet]);
  if (!sheet) return { status: 'error', message: 'Лист не знайдено' };
  var row = parseInt(data.row);
  sheet.getRange(row, 3).setValue(new Date()); // C: Updated
  sheet.getRange(row, 4).setValue('DELETED');   // D: Status
  return { status: 'ok' };
}

// ══════════════════════════════════════════
// 6. НАЛАШТУВАННЯ
// Settings: A=ВитрCats B=ДохCats C=БізCats D=Currencies E=MainCurrency
//           F=ВитрCatsFreq G=ВитрDescs H=ДохCatsFreq I=ДохDescs
//           J=БізCatsFreq K=БізDescs M=Months N=SchemaVersion
// ══════════════════════════════════════════
function getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH.settings);
  if (!sheet) return { status: 'error', message: 'Settings не знайдено' };

  var n = Math.max(sheet.getLastRow(), 13);

  function col(c) {
    return sheet.getRange(2, c, n - 1, 1).getValues()
      .map(function (r) { return r[0]; }).filter(Boolean);
  }

  return {
    status: 'ok',
    expenseCategories: col(6),   // F
    expenseDescs: col(7),   // G
    incomeCategories: col(8),   // H
    incomeDescs: col(9),   // I
    businessCategories: col(10),  // J
    businessDescs: col(11),  // K
    currencies: col(4),   // D
    mainCurrency: sheet.getRange(2, 5).getValue() || 'USD',
    months: col(13),  // M
    expenseSubcategories: col(15), // O
    expenseCategoriesOrder: col(1), // A — фіксований порядок (для діаграм)
    incomeCategoriesOrder: col(2), // B — фіксований порядок (для діаграм)
  };
}

// ══════════════════════════════════════════
// 7. ОТРИМАТИ ПОТОЧНІ ФІЛЬТРИ
// ══════════════════════════════════════════
function getFilters() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function v(shName, cell) {
    var s = ss.getSheetByName(shName);
    return s ? String(s.getRange(cell).getValue() || '') : '';
  }

  return {
    status: 'ok',
    monthly: {
      month: v(SH.monthly, FC.monthly.month),
      year: v(SH.monthly, FC.monthly.year),
      category: v(SH.monthly, FC.monthly.category)
    },
    annual: {
      month: v(SH.annual, FC.annual.month),
      year: v(SH.annual, FC.annual.year)
    },
    analyz: {
      month: v(SH.analyz, FC.analyz.month),
      year: v(SH.analyz, FC.analyz.year),
      category: v(SH.analyz, FC.analyz.category)
    }
  };
}

// ══════════════════════════════════════════
// 8. ЗАСТОСУВАТИ ФІЛЬТР + ПОВЕРНУТИ ДІАГРАМИ
// (один виклик = запис фільтру + flush + читання даних)
// ══════════════════════════════════════════
function applyFilter(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(data.sheetName);
  if (!sheet) return { status: 'error', message: 'Лист не знайдено: ' + data.sheetName };

  (data.cells || []).forEach(function (item) {
    sheet.getRange(item.cell).setValue(item.value);
  });

  SpreadsheetApp.flush();
  Utilities.sleep(800);

  return getDiagrams({ tab: data.tab });
}

// ══════════════════════════════════════════
// 9. ДАНІ ДІАГРАМ
// ══════════════════════════════════════════
function getDiagrams(data) {
  var tab = data.tab;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function read(name) {
    var s = ss.getSheetByName(name);
    if (!s) return { headers: [], rows: [] };
    var lr = s.getLastRow();
    var lc = s.getLastColumn();
    if (lr < 1 || lc < 1) return { headers: [], rows: [] };
    var vals = s.getRange(1, 1, lr, lc).getValues();
    return {
      headers: vals[0].map(function (h) { return h !== '' ? String(h) : ''; }),
      rows: vals.slice(1).filter(function (r) { return r[0] !== ''; })
    };
  }

  if (tab === 'monthly') return {
    status: 'ok',
    income: read(SH.dMonthInc),
    expenses: read(SH.dMonthExp),
    subcat: read(SH.dMonthSub)
  };

  if (tab === 'annual') return {
    status: 'ok',
    balance: read(SH.dAnnBal),
    income: read(SH.dAnnInc),
    expenses: read(SH.dAnnExp)
  };

  if (tab === 'analyz') return {
    status: 'ok',
    income: read(SH.dAnlInc),
    expenses: read(SH.dAnlExp),
    subcat: read(SH.dAnlSub)
  };

  return { status: 'error', message: 'Unknown tab: ' + tab };
}

// ══════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════
function getNextRow(sheet) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    return Math.max(sheet.getLastRow() + 1, 2);
  } finally {
    lock.releaseLock();
  }
}

function gfFormula(from, to, dateISO) {
  // Отримуємо сьогоднішню дату у форматі YYYY-MM-DD (за місцевим часом)
  var today = new Date();
  var todayISO = today.getFullYear() + '-' +
    String(today.getMonth() + 1).padStart(2, '0') + '-' +
    String(today.getDate()).padStart(2, '0');

  // 1. Якщо запитувана дата — це СЬОГОДНІ:
  if (dateISO === todayISO) {
    // Намагаємося взяти "живий" курс без дат.
    // Якщо ринок закритий або помилка — беремо останню історію за 9 днів.
    return '=IFERROR(GOOGLEFINANCE("CURRENCY:' + from + to + '"); ' +
      'INDEX(SORT(GOOGLEFINANCE("CURRENCY:' + from + to + '";"close";DATEVALUE("' + dateISO + '")-9;"' + dateISO + '");1;FALSE);2;2))';
  }

  // 2. Якщо це будь-яка МИНУЛА дата:
  // Шукаємо точний курс або останній доступний за 9 днів до цієї дати.
  else {
    return '=IFERROR(INDEX(SORT(GOOGLEFINANCE("CURRENCY:' + from + to + '";"close";DATEVALUE("' + dateISO + '")-9;"' + dateISO + '");1;FALSE);2;2);1)';
  }
}

// ══════════════════════════════════════════
// СИНХРОННИЙ КУРС — для Firebase (не пише в реальні листи)
// ══════════════════════════════════════════
function getRateSync(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RateScratch');
  if (!sheet) return { status: 'error', message: 'Лист RateScratch не знайдено' };

  var currency = (data.currency || 'USD').toUpperCase();
  var toCur = data.to || 'USD';
  var dateISO = (data.date || '').substring(0, 10);
  var cell = sheet.getRange('A1');

  if (currency === toCur) return { status: 'ok', rate: 1 };

  cell.setFormula(gfFormula(currency, toCur, dateISO));
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  var val = cell.getValue();
  cell.clearContent();

  return { status: 'ok', rate: (typeof val === 'number' && val > 0) ? val : 1 };
}

function formatDate(iso) {
  if (!iso) return '';
  var p = iso.split('-');
  return p.length === 3 ? p[2] + '.' + p[1] + '.' + p[0] : iso;
}

function verifyToken(token) {
  try {
    if (!token) return null;
    var parts = token.split('.');
    if (parts.length < 2) return null;
    var padded = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    while (padded.length % 4) padded += '=';
    var payload = JSON.parse(
      Utilities.newBlob(Utilities.base64Decode(padded)).getDataAsString()
    );
    var now = Math.floor(Date.now() / 1000);
    if (payload.exp && payload.exp < now) return null;
    return payload.email || null;
  } catch (e) {
    Logger.log('Token error: ' + e);
    return null;
  }
}

// ══════════════════════════════════════════
// ОДНОРАЗОВИЙ ІМПОРТ ДОВІДНИКІВ У FIRESTORE
// Лист "export", колонки A:E (ID, Тип, Назва, No, Місяць)
// Запустіть вручну ОДИН РАЗ
// ══════════════════════════════════════════
function importDictionariesToFirestore() {
  var PROJECT_ID = 'expensesa';
  var token = ScriptApp.getOAuthToken();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('export');
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 5).getValues(); // A:E

  var typeMap = {
    'Категорія витрат': { doc: 'expenseCategories', prefix: 'ExCat' },
    'Категорія доходів': { doc: 'incomeCategories', prefix: 'InCat' },
    'Категорія бізнесу': { doc: 'businessCategories', prefix: 'BzCat' },
    'Опис витрат': { doc: 'expenseDescriptions', prefix: 'ExDsc' },
    'Опис доходів': { doc: 'incomeDescriptions', prefix: 'InDsc' },
    'Опис бізнесу': { doc: 'BusinessDescriptions', prefix: 'BzDsc' },
    'Підкатегорія витрат': { doc: 'expenseSubcategories', prefix: 'ExSub' }
  };

  var collected = {};
  var counters = {};
  var months = [];

  // ── Валюти обробляємо окремо: дедуплікація за кодом (USD/UAH/...) ──
  var currencyCodeToId = {}; // "USD" -> "Cur01"
  var currencyCounter = 0;
  var baseCode = 'USD';
  var localCode = 'UAH';

  function getOrCreateCurrencyId(code) {
    if (currencyCodeToId[code]) return currencyCodeToId[code];
    currencyCounter++;
    var id = 'Cur' + String(currencyCounter).padStart(2, '0');
    currencyCodeToId[code] = id;
    return id;
  }

  values.forEach(function (row) {
    var id = row[0], type = row[1], name = row[2];
    var monthNo = row[3], monthName = row[4];

    if (monthNo && monthName) {
      months[parseInt(monthNo, 10) - 1] = monthName;
    }

    if (!type || !name) return;

    // ── Валюти: дедуплікуємо за кодом, base/local запам'ятовуємо окремо ──
    if (type === 'Валюта') {
      getOrCreateCurrencyId(name);
      return;
    }
    if (type === 'Базова валюта') { baseCode = name; getOrCreateCurrencyId(name); return; }
    if (type === 'Локальна валюта') { localCode = name; getOrCreateCurrencyId(name); return; }

    if (!id) return;

    var cfg = typeMap[type];
    if (!cfg) return;
    if (!collected[cfg.doc]) collected[cfg.doc] = {};
    collected[cfg.doc][id] = name;

    var num = parseInt(String(id).replace(/[^0-9]/g, ''), 10);
    if (!counters[cfg.prefix] || num > counters[cfg.prefix]) counters[cfg.prefix] = num;
  });

  // ── Формуємо документ currencies: Cur01:"USD", ..., base:"USD", local:"UAH" ──
  var currenciesDoc = {};
  Object.keys(currencyCodeToId).forEach(function (code) {
    currenciesDoc[currencyCodeToId[code]] = code;
  });
  currenciesDoc.baseCurrency = baseCode;
  currenciesDoc.localCurrency = localCode;
  collected['currencies'] = currenciesDoc;
  counters['Cur'] = currencyCounter;

  // ── Записуємо всі довідники ──
  Object.keys(collected).forEach(function (docName) {
    firestorePatchDoc_('dictionaries/' + docName, collected[docName], PROJECT_ID, token);
  });

  firestorePatchDoc_('dictionaries/counters', counters, PROJECT_ID, token);
  firestorePatchDoc_('dictionaries/months', { list: months }, PROJECT_ID, token);

  Logger.log('Довідники: ' + JSON.stringify(Object.keys(collected)));
  Logger.log('Лічильники: ' + JSON.stringify(counters));
  Logger.log('Валюти: ' + JSON.stringify(currenciesDoc));
  Logger.log('Місяці: ' + JSON.stringify(months));
}

function firestorePatchDoc_(path, obj, projectId, token) {
  var fields = {};
  Object.keys(obj).forEach(function (key) {
    var val = obj[key];
    if (Array.isArray(val)) {
      fields[key] = { arrayValue: { values: val.map(function (v) { return { stringValue: String(v) }; }) } };
    } else if (typeof val === 'number') {
      fields[key] = { integerValue: val };
    } else {
      fields[key] = { stringValue: String(val) };
    }
  });

  var url = 'https://firestore.googleapis.com/v1/projects/' + projectId
    + '/databases/(default)/documents/' + path;

  var resp = UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ fields: fields }),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    Logger.log('Помилка запису ' + path + ': ' + resp.getContentText());
  }
}

// ══════════════════════════════════════════
// ОДНОРАЗОВИЙ ЕКСПОРТ ІСТОРИЧНИХ ЗАПИСІВ У FIRESTORE
// (з конвертацією назв категорій/підкатегорій/описів у ID)
// ══════════════════════════════════════════
function exportRecordsToFirestore() {
  var PROJECT_ID = 'expensesa';
  var token = ScriptApp.getOAuthToken();

  // 1. Спочатку очищаємо три колекції у Firestore
  var collectionsToClear = ['expenses', 'income', 'business'];
  Logger.log('--- Початок очищення колекцій ---');
  collectionsToClear.forEach(function (collectionName) {
    clearCollection_(collectionName, PROJECT_ID, token);
  });
  Logger.log('--- Очищення завершено. Початок експорту ---');

  // 2. Завантажуємо довідники з Firestore, будуємо зворотній пошук
  var expCatMap = buildReverseMap_(getFirestoreDoc_('dictionaries/expenseCategories', PROJECT_ID, token));
  var incCatMap = buildReverseMap_(getFirestoreDoc_('dictionaries/incomeCategories', PROJECT_ID, token));
  var bizCatMap = buildReverseMap_(getFirestoreDoc_('dictionaries/businessCategories', PROJECT_ID, token));
  var expDscMap = buildReverseMap_(getFirestoreDoc_('dictionaries/expenseDescriptions', PROJECT_ID, token));
  var incDscMap = buildReverseMap_(getFirestoreDoc_('dictionaries/incomeDescriptions', PROJECT_ID, token));
  var bizDscMap = buildReverseMap_(getFirestoreDoc_('dictionaries/businessDescriptions', PROJECT_ID, token));
  var expSubMap = buildReverseMap_(getFirestoreDoc_('dictionaries/expenseSubcategories', PROJECT_ID, token));
  var curMap = buildCurrencyReverseMap_(getFirestoreDoc_('dictionaries/currencies', PROJECT_ID, token));

  var unmatched = [];

  // 3. Записуємо нові дані
  exportSheet_({
    sheetName: 'Витрати', collection: 'expenses', numCols: 12, // A:L, підкатегорія в L
    catMap: expCatMap, dscMap: expDscMap, subMap: expSubMap, curMap: curMap,
    isBusiness: false, hasSubcat: true
  }, PROJECT_ID, token, unmatched);

  exportSheet_({
    sheetName: 'Доходи', collection: 'income', numCols: 11, // A:K
    catMap: incCatMap, dscMap: incDscMap, subMap: null, curMap: curMap,
    isBusiness: false, hasSubcat: false
  }, PROJECT_ID, token, unmatched);

  exportSheet_({
    sheetName: 'Бізнес', collection: 'business', numCols: 13, // A:M
    catMap: bizCatMap, dscMap: bizDscMap, subMap: null, curMap: curMap,
    isBusiness: true, hasSubcat: false
  }, PROJECT_ID, token, unmatched);

  Logger.log('Експорт завершено.');
  if (unmatched.length) {
    Logger.log('⚠ Незіставлені значення (' + unmatched.length + '):');
    Logger.log(JSON.stringify(unmatched, null, 2));
  } else {
    Logger.log('✓ Усі категорії/підкатегорії/описи/валюти успішно зіставлені з ID.');
  }
}

/**
 * Оптимізоване очищення колекції Firestore через :commit (batch write)
 */
function clearCollection_(collection, projectId, token) {
  var deletedCount = 0;
  var getUrl = 'https://firestore.googleapis.com/v1/projects/' + projectId
    + '/databases/(default)/documents/' + collection + '?pageSize=300';
  var commitUrl = 'https://firestore.googleapis.com/v1/projects/' + projectId
    + '/databases/(default)/documents:commit';

  while (true) {
    var resp = UrlFetchApp.fetch(getUrl, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('Помилка при отриманні списку документів з "' + collection + '": ' + resp.getContentText());
      break;
    }

    var data = JSON.parse(resp.getContentText());
    var docs = data.documents || [];

    if (docs.length === 0) break;

    // Формуємо масив видалень для одного пакетного POST-запиту (:commit)
    var writes = docs.map(function (doc) {
      return { delete: doc.name };
    });

    // Видаляємо до 300 документів за 1 HTTP-запит!
    var commitResp = UrlFetchApp.fetch(commitUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ writes: writes }),
      muteHttpExceptions: true
    });

    if (commitResp.getResponseCode() === 200) {
      deletedCount += docs.length;
    } else {
      Logger.log('Помилка пакетного видалення в "' + collection + '": ' + commitResp.getContentText());
      break;
    }

    // Пауза 500 мс для запобігання лімітам частоти Google
    Utilities.sleep(500);
  }

  Logger.log('Колекцію "' + collection + '" очищено. Видалено документів: ' + deletedCount);
}

/**
 * Оптимізований експорт аркуша пакетами по 30 записів
 */
function exportSheet_(cfg, projectId, token, unmatched) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var values = sheet.getRange(2, 1, lastRow - 1, cfg.numCols).getValues();
  var requests = [];

  values.forEach(function (row) {
    if (row[3] === 'DELETED') return;
    if (!row[4] && !row[6]) return;

    var dateVal = row[4];
    var dateStr = dateVal instanceof Date
      ? Utilities.formatDate(dateVal, 'Europe/Kiev', 'yyyy-MM-dd')
      : '';

    var catText = String(row[5] || '').trim();
    var catId = catMapLookup_(cfg.catMap, catText);
    if (catText && !catId) unmatched.push({ sheet: cfg.sheetName, type: 'category', text: catText });

    var dscText = String(row[7] || '').trim();
    var dscId = catMapLookup_(cfg.dscMap, dscText);
    if (dscText && !dscId) unmatched.push({ sheet: cfg.sheetName, type: 'desc', text: dscText });

    var curText = String(row[8] || '').trim();
    var curId = catMapLookup_(cfg.curMap, curText);
    if (curText && !curId) unmatched.push({ sheet: cfg.sheetName, type: 'currency', text: curText });

    var subId = '';
    if (cfg.hasSubcat) {
      var subText = String(row[11] || '').trim(); // колонка L
      subId = catMapLookup_(cfg.subMap, subText);
      if (subText && !subId) unmatched.push({ sheet: cfg.sheetName, type: 'subcategory', text: subText });
    }

    var doc = {
      fields: {
        categoryId: { stringValue: catId },
        subcategoryId: { stringValue: subId },
        amount: { doubleValue: row[6] || 0 },
        descId: { stringValue: dscId },
        currencyId: { stringValue: curId },
        currency: { stringValue: curText },
        rateUSD: { doubleValue: row[9] || 1 },
        amountUSD: { doubleValue: row[10] || 0 },
        date: { stringValue: dateStr },
        author: { stringValue: 'xxxx@gmail.com' },
        status: { stringValue: 'ACTIVE' },
        createdAt: { timestampValue: toFirestoreTimestamp_(row[1], dateStr) },
        updatedAt: { timestampValue: toFirestoreTimestamp_(row[2], dateStr) },
        imported: { booleanValue: true }
      }
    };

    if (cfg.isBusiness) {
      doc.fields.rateUAH = { doubleValue: row[11] || 1 };
      doc.fields.amountUAH = { doubleValue: row[12] || 0 };
    }

    var url = 'https://firestore.googleapis.com/v1/projects/' + projectId
      + '/databases/(default)/documents/' + cfg.collection;

    requests.push({
      url: url,
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(doc),
      muteHttpExceptions: true
    });
  });

  var count = 0;
  var BATCH_SIZE = 30; // Безпечна кількість запитів за один раз

  for (var i = 0; i < requests.length; i += BATCH_SIZE) {
    var chunk = requests.slice(i, i + BATCH_SIZE);
    var responses = UrlFetchApp.fetchAll(chunk);

    responses.forEach(function (resp) {
      if (resp.getResponseCode() === 200) {
        count++;
      } else {
        Logger.log('Помилка рядка (' + cfg.sheetName + '): ' + resp.getContentText());
      }
    });

    // Пауза 500 мс між пакетами, щоб Google не блокував по квоті
    if (i + BATCH_SIZE < requests.length) {
      Utilities.sleep(500);
    }
  }

  Logger.log(cfg.sheetName + ': експортовано ' + count + ' записів');
}

function catMapLookup_(map, text) {
  if (!map || !text) return '';
  return map[text.toLowerCase()] || '';
}

function toFirestoreTimestamp_(dateVal, fallbackDateStr) {
  if (dateVal instanceof Date) return dateVal.toISOString();
  if (fallbackDateStr) return fallbackDateStr + 'T00:00:00.000Z';
  return new Date().toISOString();
}

function buildCurrencyReverseMap_(dict) {
  var map = {};
  Object.keys(dict).forEach(function (key) {
    if (key.indexOf('Cur') !== 0) return;
    var code = dict[key];
    if (typeof code === 'string') map[code.toLowerCase()] = key;
  });
  return map;
}

// ── Читає документ Firestore, повертає JS-об'єкт { "ExCat00001": "Житло", ... } ──
function getFirestoreDoc_(path, projectId, token) {
  var url = 'https://firestore.googleapis.com/v1/projects/' + projectId
    + '/databases/(default)/documents/' + path;
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    Logger.log('Не вдалось прочитати ' + path + ': ' + resp.getContentText());
    return {};
  }
  var json = JSON.parse(resp.getContentText());
  var result = {};
  var fields = json.fields || {};
  Object.keys(fields).forEach(function (key) {
    var f = fields[key];
    if (f.stringValue !== undefined) result[key] = f.stringValue;
  });
  return result;
}

// ── { "ExCat00001": "Житло" } → { "житло": "ExCat00001" } (для пошуку за назвою) ──
function buildReverseMap_(dict) {
  var map = {};
  Object.keys(dict).forEach(function (id) {
    var name = dict[id];
    if (typeof name === 'string') map[name.toLowerCase()] = id;
  });
  return map;
}

function fixCurrencyCounter() {
  var PROJECT_ID = 'expensesa';
  var token = ScriptApp.getOAuthToken();
  firestorePatchDoc_('dictionaries/counters', { BzCat: 3, BzDsc: 2, Cur: 5, ExCat: 11, ExSub: 18, ExDsc: 44, InCat: 7, InDsc: 2 }, PROJECT_ID, token);
  Logger.log('Готово');
}

function formatFirestoreTimestamp(val, fallbackDate) {
  // 1. Якщо значення вже є і це валідна дата/Timestamp
  if (val) {
    if (val instanceof Date && !isNaN(val.getTime())) {
      return val.toISOString();
    }
    var parsed = new Date(val);
    if (!isNaN(parsed.getTime())) {
      return parsed.toISOString();
    }
  }

  // 2. Якщо значення відсутнє — беремо fallbackDate (dateStr) і встановлюємо 12:00:00
  if (fallbackDate) {
    var d;
    if (fallbackDate instanceof Date && !isNaN(fallbackDate.getTime())) {
      d = new Date(fallbackDate);
    } else if (typeof fallbackDate === 'string' && fallbackDate.trim() !== '') {
      var cleanStr = fallbackDate.trim().substring(0, 10);

      // Обробка формату DD.MM.YYYY
      if (cleanStr.includes('.')) {
        var parts = cleanStr.split('.'); // ['DD', 'MM', 'YYYY']
        cleanStr = parts[2] + '-' + parts[1] + '-' + parts[0]; // 'YYYY-MM-DD'
      }

      d = new Date(cleanStr + 'T12:00:00Z');
    }

    if (d && !isNaN(d.getTime())) {
      d.setUTCHours(12, 0, 0, 0); // встановлюємо 12:00:00 UTC
      return d.toISOString();
    }
  }

  // 3. Резервний варіант, якщо і dateStr порожній — поточний час
  return new Date().toISOString();
}

// ── Тести ──
function testAdd() {
  Logger.log(JSON.stringify(addRecord({
    sheet: 'expenses', uuid: 'SL_test-001',
    date: '2026-07-12', category: 'Продукти та хозтовари',
    amount: 99, desc: 'тест', currency: 'UAH', offline: false
  }, 'test@gmail.com')));
}
function testSettings() { Logger.log(JSON.stringify(getSettings())); }
function testFilters() { Logger.log(JSON.stringify(getFilters())); }
function testDiagrams() { Logger.log(JSON.stringify(getDiagrams({ tab: 'monthly' }))); }
