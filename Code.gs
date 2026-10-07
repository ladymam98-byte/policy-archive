/**
 * 도시·주거공간 정책 연구 아카이브 - Google Apps Script 백엔드
 * 스프레드시트 [확장 프로그램] > [Apps Script]에 붙여넣으세요.
 */

const SHEET_NAME = '기고';
const HEADERS = ['id', 'date', 'category', 'title', 'author', 'email', 'content', 'submittedAt'];
const CATEGORIES = ['도시재생', '주거복지', '공간·GIS'];

/** 시트가 없으면 만들고 헤더를 작성합니다. */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#0f3e36')
      .setFontColor('#ffffff');
  }
  return sheet;
}

/** 수식 주입(=, +, -, @로 시작) 방지 */
function clean_(value, maxLen) {
  let s = String(value == null ? '' : value).trim().slice(0, maxLen);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

/** 기고 폼 POST 수신 → 시트에 한 행 추가 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = JSON.parse(e.postData.contents);

    if (!data.title || !data.author || !data.content || CATEGORIES.indexOf(data.category) === -1) {
      return json_({ ok: false, error: 'invalid payload' });
    }

    const now = new Date();
    const row = [
      clean_(data.id || 'u' + now.getTime(), 40),
      Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd'),
      data.category,
      clean_(data.title, 120),
      clean_(data.author, 40),
      clean_(data.email, 120),
      clean_(data.content, 2000),
      now.toISOString()
    ];
    getSheet_().appendRow(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 저장된 기고 목록을 JSON으로 반환 (이메일은 공개하지 않음) */
function doGet() {
  const values = getSheet_().getDataRange().getValues();
  const header = values.shift();
  const rows = values.map(function (r) {
    const o = {};
    header.forEach(function (h, i) { o[h] = r[i]; });
    if (o.date instanceof Date) o.date = Utilities.formatDate(o.date, 'Asia/Seoul', 'yyyy-MM-dd');
    delete o.email;
    return o;
  }).reverse();
  return json_(rows);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
