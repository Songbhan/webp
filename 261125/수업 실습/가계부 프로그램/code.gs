const SPREADSHEET_ID = 'example';
const DAILY_LIMIT = 40;
const MAX_AMOUNT = 999999999999;
const TYPES = ['현금 소비', '수입', 'ATM 인출'];
const CURRENCIES = ['JPY', 'KRW'];
const PAYMENT_METHODS = ['현금', '카드'];
const DEFAULT_CATEGORIES = ['식비', '교통', '쇼핑', '장보기', '공과금'];

const RAW_HEADERS = [
  '기록 ID', '수정한 기록 ID', '날짜', '구분', '통화',
  '카테고리', '금액', '메모', '상태', '결제수단'
];

const TAB_HEADERS = {
  '거래 원본': RAW_HEADERS,
  '카테고리': ['카테고리'],
  '가계부': ['전체 합계'],
  '월별 요약': ['월별 합계']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('여행 가계부')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 편집기에서 실행하면 시트를 준비하고 집계를 갱신합니다.
function setupLedger() {
  return refreshLedgerReports();
}

function refreshLedgerReports() {
  return refreshLedgerReports_(openLedger_());
}

// 독립형 Apps Script 프로젝트에서 자동 갱신을 설치할 때 실행합니다.
function installLedgerEditTrigger() {
  const handler = 'handleLedgerSheetEdit';
  const installed = ScriptApp.getProjectTriggers().some((trigger) =>
    [handler, 'refreshLedgerOnEdit_'].includes(trigger.getHandlerFunction()) &&
    trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
    trigger.getTriggerSourceId() === SPREADSHEET_ID
  );

  if (!installed) {
    ScriptApp.newTrigger(handler)
      .forSpreadsheet(SPREADSHEET_ID)
      .onEdit()
      .create();
  }

  console.log(installed ? '시트 편집 트리거 확인 완료' : '시트 편집 트리거 설치 완료');
}

// 이전에 설치한 트리거도 동작하도록 유지합니다.
function refreshLedgerOnEdit_(event) {
  handleLedgerSheetEdit(event);
}

// 같은 스프레드시트에 연결된 프로젝트에서 직접 편집할 때 실행됩니다.
function onEdit(event) {
  handleLedgerSheetEdit_(event, 5000);
}

function handleLedgerSheetEdit(event) {
  handleLedgerSheetEdit_(event, 30000);
}

function handleLedgerSheetEdit_(event, timeoutMs) {
  const range = event?.range;
  const book = event?.source;
  const statusColumn = RAW_HEADERS.indexOf('상태') + 1;

  if (!range || !book || book.getId() !== SPREADSHEET_ID ||
      range.getSheet().getName() !== '거래 원본' ||
      range.getLastRow() < 2 ||
      range.getColumn() > statusColumn ||
      range.getLastColumn() < statusColumn) {
    return;
  }

  refreshLedgerReports_(book, timeoutMs);
}

function refreshLedgerReports_(book, timeoutMs = 30000) {
  const result = withLedgerLock_(() => {
    const tabs = ensureTabs_(book);
    const ledger = readLedger_(tabs['거래 원본']);
    refreshLedgerViews_(tabs, ledger);

    return {
      url: book.getUrl(),
      deleted: ledger.all.filter((item) => item.deleted).length,
      active: ledger.active.length
    };
  }, timeoutMs);

  if (result.ok === false) throw new Error(result.message);

  console.log('집계 갱신 완료 | 삭제 ' + result.deleted + '건 | 유효 ' + result.active + '건');
  return result;
}

function getLedgerData(includeRecent) {
  const book = openLedger_();
  const categories = readCategories_(book.getSheetByName('카테고리'));
  const raw = book.getSheetByName('거래 원본');
  const recent = includeRecent && raw
    ? readLedger_(raw).active.slice(-40).reverse()
    : [];

  return {
    categories,
    recent: recent.map(({ id, date, type, currency, category, amount, memo, payment }) =>
      ({ id, date, type, currency, category, amount, memo, payment }))
  };
}

function saveLedgerEntry(input) {
  const entry = validateEntry_(input);
  if (!entry) {
    return { ok: false, message: '날짜, 통화, 금액과 입력 항목을 확인해 주세요.' };
  }

  return withLedgerLock_(() => {
    const book = openLedger_();
    const tabs = ensureTabs_(book);
    const raw = tabs['거래 원본'];
    const ledger = readLedger_(raw);
    const saved = ledger.all.find((item) => item.id === entry.id);

    if (saved) {
      if (saved.deleted) {
        return { ok: false, message: '이미 삭제된 기록이에요.' };
      }

      const keys = ['source', 'date', 'type', 'currency', 'category', 'amount', 'memo', 'payment'];
      if (keys.some((key) => saved[key] !== entry[key])) {
        return { ok: false, message: '이미 처리된 요청이에요. 화면을 새로고침해 주세요.' };
      }

      return finishSave_(book, tabs, saved, ledger);
    }

    if (entry.source && !ledger.active.some((item) => item.id === entry.source)) {
      return {
        ok: false,
        message: '이미 수정되거나 삭제된 기록이에요. 재입력 목록에서 다시 선택해 주세요.'
      };
    }

    const remaining = ledger.active.filter((item) => item.id !== entry.source);
    const dayCount = remaining.filter((item) => item.date === entry.date).length;

    if (dayCount >= DAILY_LIMIT) {
      return { ok: false, message: entry.date + '에는 이미 40건이 있어요.' };
    }

    const row = raw.getLastRow() + 1;
    ensureRows_(raw, row);
    raw.getRange(row, 1, 1, RAW_HEADERS.length)
      .setNumberFormats([['@', '@', '@', '@', '@', '@', '#,##0', '@', '0', '@']])
      .setValues([[
        entry.id,
        entry.source,
        sheetText_(entry.date),
        displayType_(entry.type),
        entry.currency,
        sheetText_(entry.category),
        entry.amount,
        sheetText_(entry.memo),
        '=IF(COUNTIF($B$2:$B,A' + row + ')>0,"수정됨","유효")',
        entry.payment
      ]]);

    SpreadsheetApp.flush();

    entry.row = row;
    return finishSave_(book, tabs, entry, readLedger_(raw));
  });
}

function finishSave_(book, tabs, entry, ledger) {
  const raw = tabs['거래 원본'];
  const rowUrl = book.getUrl().split('#')[0] +
    '#gid=' + raw.getSheetId() + '&range=C' + entry.row + ':J' + entry.row;
  let reportReady = true;

  try {
    refreshLedgerViews_(tabs, ledger);
    SpreadsheetApp.flush();
  } catch (error) {
    console.error(error);
    reportReady = false;
  }

  return {
    ok: true,
    rowUrl,
    reportReady,
    date: entry.date,
    dayCount: ledger.active.filter((item) => item.date === entry.date).length,
    categories: readCategories_(tabs['카테고리'])
  };
}

function openLedger_() {
  if (!/^[A-Za-z0-9_-]+$/.test(SPREADSHEET_ID)) {
    throw new Error('Code.gs에 구글시트 ID를 먼저 입력해 주세요.');
  }

  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function withLedgerLock_(work, timeoutMs = 10000) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(timeoutMs)) {
    return { ok: false, message: '다른 저장이 진행 중이에요. 잠시 후 다시 시도해 주세요.' };
  }

  try {
    return work();
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

function ensureTabs_(book) {
  const tabs = {};

  for (const [name, headers] of Object.entries(TAB_HEADERS)) {
    const sheet = book.getSheetByName(name);

    if (sheet && sheet.getLastRow()) {
      const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      const matches = headers.every((header, index) =>
        header === firstRow[index] ||
        (name === '거래 원본' && index === headers.length - 1 && firstRow[index] === '')
      );

      if (!matches) {
        throw new Error('기존 "' + name + '" 탭 이름을 바꾸거나 다른 구글시트를 연결해 주세요.');
      }
    }

    tabs[name] = sheet;
  }

  for (const [name, headers] of Object.entries(TAB_HEADERS)) {
    const sheet = tabs[name] || book.insertSheet(name);
    tabs[name] = sheet;

    if (sheet.getLastRow()) {
      if (name === '거래 원본') {
        const paymentHeader = sheet.getRange(1, RAW_HEADERS.length);

        if (paymentHeader.isBlank()) {
          paymentHeader.setValue('결제수단').setFontWeight('bold');
        }
      }
      continue;
    }

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold');

    if (name === '거래 원본') {
      sheet.setFrozenRows(1);
      sheet.hideColumns(1, 2);
      sheet.setColumnWidth(8, 280);
    }

    if (name === '카테고리') {
      sheet.getRange(2, 1, DEFAULT_CATEGORIES.length, 1)
        .setValues(DEFAULT_CATEGORIES.map((category) => [category]));
    }

    if (name === '가계부' || name === '월별 요약') {
      sheet.setFrozenRows(name === '가계부' ? 4 : 2);
      sheet.setColumnWidths(1, 5, 132);
      sheet.setColumnWidth(5, 260);
    }
  }

  return tabs;
}

function readCategories_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return DEFAULT_CATEGORIES.slice();

  const custom = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .map(([name]) => name.trim())
    .filter(Boolean);

  return [...new Set([...DEFAULT_CATEGORIES, ...custom])];
}

function readLedger_(sheet) {
  if (sheet.getLastRow() < 2) return { all: [], active: [] };

  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, RAW_HEADERS.length)
    .getValues()
    .map((values, index) => {
      const [id, source, date, storedType, currency, category, amount, memo, status, storedPayment] = values;
      const type = storedType === '소비' ? TYPES[0] : storedType;
      const payment = type === TYPES[0] ? (storedPayment || '현금') : '';
      const deleted = String(status).trim() === '삭제';

      return {
        id, source, date, type, currency, category, amount, memo,
        payment, deleted, row: index + 2
      };
    })
    .filter((item) => item.id);

  // 삭제된 수정본도 원본을 대체한 이력을 유지합니다.
  const replaced = new Set(all.map((item) => item.source).filter(Boolean));
  const active = all.filter((item) => !item.deleted && !replaced.has(item.id));

  return { all, active };
}

function validateEntry_(input) {
  const { id, source, date, type, currency, amount } = input || {};
  const uuid = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/;

  if (typeof id !== 'string' || !uuid.test(id) || typeof source !== 'string' ||
      (source && (!uuid.test(source) || source === id))) return null;

  if (typeof date !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(date) ||
      date < '1900-01-01' || date > '9999-12-31') return null;

  const parsed = new Date(date + 'T00:00:00Z');
  if (!Number.isFinite(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== date ||
      !TYPES.includes(type) || !CURRENCIES.includes(currency) ||
      !Number.isSafeInteger(amount) || amount < 1 || amount > MAX_AMOUNT) return null;

  const category = type === TYPES[0] && typeof input.category === 'string'
    ? input.category.trim()
    : '';
  const memo = typeof input.memo === 'string' ? input.memo.trim() : '';
  const payment = type === TYPES[0] ? (input.payment ?? '현금') : '';

  if ((type === TYPES[0] && !category) || category.length > 30 || memo.length > 300 ||
      /[\r\n]/.test(category)) return null;

  if (type === TYPES[0] && !PAYMENT_METHODS.includes(payment)) return null;

  return { id, source, date, type, currency, category, amount, memo, payment };
}

function displayType_(type) {
  return type === TYPES[0] ? '소비' : type;
}

function sheetText_(value) {
  return value ? "'" + value : '';
}

function ensureRows_(sheet, required) {
  const current = sheet.getMaxRows();
  if (required > current) sheet.insertRowsAfter(current, required - current);
}

function refreshLedgerViews_(tabs, ledger) {
  renderReports_(tabs, ledger.active);
  syncLedgerCategories_(tabs['카테고리'], ledger);
}

function syncLedgerCategories_(sheet, ledger) {
  const initialRows = sheet.getLastRow();
  const names = initialRows > 1
    ? sheet.getRange(2, 1, initialRows - 1, 1)
        .getDisplayValues()
        .map(([name]) => name.trim())
    : [];

  const known = new Set(names);
  const history = ledger.all
    .filter((entry) => entry.type === TYPES[0])
    .map((entry) => entry.category);
  const missing = [...new Set([...DEFAULT_CATEGORIES, ...history])]
    .filter((name) => name && !known.has(name));

  if (missing.length) {
    ensureRows_(sheet, initialRows + missing.length);
    sheet.getRange(initialRows + 1, 1, missing.length, 1)
      .setNumberFormat('@')
      .setValues(missing.map((name) => [sheetText_(name)]));
    names.push(...missing);
  }

  const counts = new Map();

  for (const entry of ledger.active) {
    if (entry.type === TYPES[0] && entry.category) {
      counts.set(entry.category, (counts.get(entry.category) || 0) + 1);
    }
  }

  sheet.getRange(1, 2).setValue('유효 지출 건수').setFontWeight('bold');

  if (names.length) {
    sheet.getRange(2, 2, names.length, 1)
      .setValues(names.map((name) => [name ? (counts.get(name) || 0) : '']));
  }
}

function renderReports_(tabs, entries) {
  const totals = new Map();
  const days = new Map();
  const daily = new Map();
  const monthly = new Map();
  const categories = new Map();

  const add = (groups, period, item) => {
    const key = JSON.stringify([period, item.currency]);
    if (!groups.has(key)) groups.set(key, [0, 0, 0]);
    groups.get(key)[TYPES.indexOf(item.type)] += item.amount;
  };

  for (const entry of entries) {
    if (!days.has(entry.date)) days.set(entry.date, []);
    days.get(entry.date).push(entry);

    add(totals, '', entry);
    add(daily, entry.date, entry);
    add(monthly, entry.date.slice(0, 7), entry);

    if (entry.type === TYPES[0]) {
      add(categories, entry.date.slice(0, 7) + ' · ' + entry.category, entry);
    }
  }

  const sorted = (groups) => [...groups.entries()].sort(([a], [b]) => b.localeCompare(a));
  const currencyName = (currency) => currency === 'JPY' ? '엔화' : '원화';
  const headings = ['통화', '지출', '수입', 'ATM 인출'];
  const main = [];
  const summary = [];

  const push = (rows, values, style, amounts = []) => {
    rows.push({ values, style, amounts });
  };

  const pushSum = (rows, period, currency, sums) => {
    push(rows, [period, currencyName(currency), ...sums], '',
      [[2, currency], [3, currency], [4, currency]]);
  };

  push(main, ['전체 합계'], 'title');
  push(main, ['구분', ...headings], 'header');

  for (const currency of CURRENCIES) {
    const sums = totals.get(JSON.stringify(['', currency])) || [0, 0, 0];
    pushSum(main, '전체', currency, sums);
  }

  for (const date of [...days.keys()].sort().reverse()) {
    const items = days.get(date);
    push(main, []);
    push(main, [date + ' · ' + items.length + '/' + DAILY_LIMIT + '건'], 'section');
    push(main, ['구분', '카테고리', '통화', '금액', '메모', '결제수단'], 'header');

    for (const item of items.slice().reverse()) {
      push(main, [
        displayType_(item.type),
        sheetText_(item.category),
        currencyName(item.currency),
        item.amount,
        sheetText_(item.memo),
        item.payment
      ], '', [[3, item.currency]]);
    }
  }

  push(main, []);
  push(main, ['일별 합계'], 'title');
  push(main, ['날짜', ...headings], 'header');

  for (const [key, sums] of sorted(daily)) {
    pushSum(main, ...JSON.parse(key), sums);
  }

  push(summary, ['월별 합계'], 'title');
  push(summary, ['월', ...headings], 'header');

  for (const [key, sums] of sorted(monthly)) {
    pushSum(summary, ...JSON.parse(key), sums);
  }

  push(summary, []);
  push(summary, ['월별 카테고리 지출'], 'title');
  push(summary, ['월 · 카테고리', '통화', '지출'], 'header');

  for (const [key, sums] of sorted(categories)) {
    const [label, currency] = JSON.parse(key);
    push(summary, [sheetText_(label), currencyName(currency), sums[0]], '', [[2, currency]]);
  }

  writeReport_(tabs['가계부'], main);
  writeReport_(tabs['월별 요약'], summary);
}

function writeReport_(sheet, rows) {
  const previousLength = sheet.getLastRow();
  ensureRows_(sheet, rows.length);

  const width = 6;
  const range = sheet.getRange(1, 1, rows.length, width);
  const matrix = (makeCell) => rows.map((row, index) =>
    Array.from({ length: width }, (_, col) => makeCell(row, col, index)));

  range.setNumberFormats(matrix((row, col) => {
    const amount = row.amounts.find(([index]) => index === col);
    if (!amount) return '@';
    return amount[1] === 'JPY' ? '"¥"#,##0' : '"₩"#,##0';
  }));

  range.setValues(matrix((row, col) => row.values[col] ?? ''));

  range.setBackgrounds(matrix((row, col, index) => {
    if (row.style === 'title') return '#174b42';
    if (row.style === 'section') return '#c7e9d9';
    if (row.style === 'header') return '#e7f2ec';
    return row.values.length && index % 2 ? '#f6faf8' : '#ffffff';
  }));

  range.setFontColors(matrix((row) => row.style === 'title' ? '#ffffff' : '#17342b'));
  range.setFontWeights(matrix((row) => row.style ? 'bold' : 'normal'));
  range.setWrap(true);

  if (previousLength > rows.length) {
    sheet.getRange(rows.length + 1, 1, previousLength - rows.length, width)
      .clearContent()
      .clearFormat();
  }
}
