const SPREADSHEET = '1JFlgph2B8ZiSbGUh4BwUwSrp7tIbCzuAu9W5wVc9sKs';
const SHEET_NAME = '신청';
const HEADERS = ['이름', '음식', '하고싶은말'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('11/25 저녁 신청받기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function testConnection() {
  const sheet = getResponseSheet_();
  checkHeaders_(sheet);
  console.log('시트 연결 및 헤더 확인 완료');
}

function saveApplication(form) {
  const name = typeof form?.name === 'string' ? form.name.trim() : '';
  const food = typeof form?.food === 'string' ? form.food.trim() : '';
  const comment = typeof form?.comment === 'string' ? form.comment.trim() : '';

  if (!name || !food) {
    throw new Error('이름과 음식을 입력해 주세요.');
  }

  if (name.length > 100 || food.length > 100 || comment.length > 2000) {
    throw new Error('이름과 음식은 100자, 하고싶은말은 2000자 이내로 입력해 주세요.');
  }

  const sheet = getResponseSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    checkHeaders_(sheet);

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }

    // 입력값을 텍스트로 저장해 수식 실행을 방지합니다.
    const row = [name, food, comment].map(value => value ? "'" + value : '');
    sheet.appendRow(row);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function getResponseSheet_() {
  const input = SPREADSHEET.trim();
  const urlMatch = input.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  const id = urlMatch ? urlMatch[1] : input;

  if (!/^[A-Za-z0-9_-]+$/.test(id)) {
    throw new Error('code.gs의 SPREADSHEET에 구글 스프레드시트 주소를 넣어 주세요.');
  }

  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(id);
  } catch (error) {
    throw new Error('스프레드시트를 열 수 없습니다. 주소와 실행 계정의 시트 접근 권한을 확인해 주세요.');
  }

  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('응답 탭을 찾을 수 없습니다. SHEET_NAME을 시트 하단의 탭 이름과 맞춰 주세요.');
  }

  return sheet;
}

function checkHeaders_(sheet) {
  if (sheet.getLastRow() === 0) return;

  const [actual] = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues();
  if (!HEADERS.every((header, index) => header === actual[index].trim())) {
    throw new Error('시트의 A1, B1, C1을 각각 이름, 음식, 하고싶은말로 설정해 주세요.');
  }
}
