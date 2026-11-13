const SPREADSHEET_ID = '1Ovg82HKLbmVEJdvC9u7yO9zMCjgFjlv2wv4VJ6_Iv4Y';
const SHEET_NAME = '시트1';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('학생 건의함')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function submitSuggestion(data) {
  const lock = LockService.getScriptLock();
  try {
    // 여러 학생이 동시에 제출할 때 데이터가 꼬이지 않도록 10초간 대기
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

    // 시트가 비어 있으면 헤더 행 자동 생성
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

// 시트 수식 오작동(=, +, -, @ 로 시작하는 입력) 방지
function sanitizeForSheet(value) {
  if (/^[=+\-@]/.test(value)) {
    return "'" + value;
  }
  return value;
}
