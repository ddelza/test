const SPREADSHEET_ID = '1Ovg82HKLbmVEJdvC9u7yO9zMCjgFjlv2wv4VJ6_Iv4Y';
const SHEET_NAME = '시트1';
const ADMIN_PASSWORD = '1234';

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';

  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('학생 건의함 - 관리자')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('학생 건의함')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 학생 건의 제출
function submitSuggestion(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const rawName = (data && data.nickname) ? String(data.nickname).trim() : '';
    const rawContent = (data && data.content) ? String(data.content).trim() : '';

    if (!rawContent) {
      throw new Error('건의 내용을 입력해주세요.');
    }

    const nickname = sanitizeForSheet(rawName || '익명');
    const content = sanitizeForSheet(rawContent);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['제출 일시', '이름(별명)', '건의 내용']);
      const headerRange = sheet.getRange(1, 1, 1, 3);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f1f5f9');
      sheet.setFrozenRows(1);
    }

    const timestamp = Utilities.formatDate(
      new Date(),
      'Asia/Seoul',
      'yyyy-MM-dd HH:mm:ss'
    );

    sheet.appendRow([timestamp, nickname, content]);

    return { success: true };
  } catch (err) {
    throw new Error(err.message || '저장 중 오류가 발생했습니다.');
  } finally {
    lock.releaseLock();
  }
}

// 관리자용 건의 목록 조회 (비밀번호 검증)
function getSuggestions(password) {
  if (String(password) !== ADMIN_PASSWORD) {
    throw new Error('비밀번호가 일치하지 않습니다.');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();

  return values
    .map((row, idx) => ({
      number: idx + 1,
      timestamp: row[0],
      nickname: row[1],
      content: row[2]
    }))
    .filter(item => item.timestamp || item.nickname || item.content);
}

function sanitizeForSheet(value) {
  if (/^[=+\-@]/.test(value)) {
    return "'" + value;
  }
  return value;
}
