const SS_PREP_ID = '17HV5qLlOhlxsll8ZZdAUxe3pqAmwA7n1ysmg4gCukU4';

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;

  if (!action) {
    if (e.parameter.page === 'youtube') {
      return HtmlService.createHtmlOutputFromFile('youtube')
        .setTitle('발표 유튜브 링크 모아보기')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('발표회 청중 보고서 조회')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  let result;
  if (action === 'getStudents') {
    result = getStudents();
  } else if (action === 'getSubmissions') {
    result = getSubmissions(e.parameter.반, e.parameter.번호);
  } else if (action === 'getPrepStatus') {
    result = getPrepStatus(e.parameter.반, e.parameter.성명);
  } else if (action === 'getYoutubeByClass') {
    result = getYoutubeByClass(e.parameter.반);
  } else {
    result = { error: 'unknown action' };
  }

  const json = JSON.stringify(result);

  // JSONP: callback 파라미터가 있으면 함수 호출 형태로 반환
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const infoSheet = ss.getSheetByName('학생 정보');
  const respSheet = ss.getSheetByName('설문지 응답 시트1');

  // 제출 횟수 집계: key = "반_번호"
  const countMap = {};
  const respData = respSheet.getDataRange().getValues();
  for (let i = 1; i < respData.length; i++) {
    const row = respData[i];
    if (!row[0]) continue;
    const key = String(row[1]) + '_' + String(row[2]);
    countMap[key] = (countMap[key] || 0) + 1;
  }

  const data = infoSheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const key = String(row[1]) + '_' + String(row[2]);
    students.push({
      학번: String(row[0]),
      반: String(row[1]),
      번호: String(row[2]),
      성명: String(row[3]),
      모둠: String(row[4]),
      제출수: countMap[key] || 0
    });
  }
  return students;
}

function norm(s) {
  return String(s || '').replace(/\s+/g, '').trim();
}

function getPrepStatus(반, 성명) {
  const sheet = SpreadsheetApp.openById(SS_PREP_ID).getSheetByName('시트1');
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const members = [row[1], row[2], row[3], row[4]].map(String).filter(Boolean);
    const isMatch = String(row[0]) === String(반) &&
      members.some(m => norm(m) === norm(성명));
    if (isMatch) {
      results.push({
        반: String(row[0]),
        모둠원: members,
        주제: String(row[5] || ''),
        대본: String(row[6] || ''),
        캔바링크: String(row[7] || ''),
        예상질문: [row[8], row[9], row[10], row[11]].map(String).filter(Boolean)
      });
    }
  }
  return results;
}

function getYoutubeByClass(반) {
  const sheet = SpreadsheetApp.openById(SS_PREP_ID).getSheetByName('시트1');
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[0]) !== String(반)) continue;
    const members = [row[1], row[2], row[3], row[4]].map(String).filter(Boolean);
    results.push({
      반: String(row[0]),
      모둠원: members,
      주제: String(row[5] || ''),
      유튜브링크: String(row[12] || '')
    });
  }
  return results;
}

function getSubmissions(반, 번호) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('설문지 응답 시트1');
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[1]) === String(반) && String(row[2]) === String(번호)) {
      results.push({
        타임스탬프: row[0] instanceof Date
          ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
          : String(row[0]),
        주제: String(row[3] || ''),
        메모: String(row[4] || ''),
        질문: String(row[5] || '')
      });
    }
  }
  return results;
}
