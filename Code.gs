/**
 * AI 창작물 불쾌한 골짜기 설문 백엔드 (Google Apps Script)
 * 이 스크립트는 바인딩된 스프레드시트의 Responses 시트만 사용합니다.
 * 이름, 이메일, IP 주소를 읽거나 기록하지 않습니다.
 */
const SHEET_NAME = 'Responses';

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'responses');
    if (action === 'count') return json_({ ok: true, count: dataRows_().length });
    if (action === 'responses') {
      const rows = objects_();
      return json_({ ok: true, count: rows.length, rows: rows });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!body.responseId || !body.submittedAt) throw new Error('필수 식별값이 없습니다.');
    const sheet = sheet_();
    const existing = objects_().some(function (r) { return String(r.responseId) === String(body.responseId); });
    if (existing) return json_({ ok: true, duplicate: true, responseId: body.responseId });

    const flat = flatten_(body);
    let headers = lastColumn_(sheet) ? sheet.getRange(1, 1, 1, lastColumn_(sheet)).getValues()[0] : [];
    const newKeys = Object.keys(flat).filter(function (k) { return headers.indexOf(k) < 0; });
    if (!headers.length) {
      headers = Object.keys(flat);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else if (newKeys.length) {
      sheet.getRange(1, headers.length + 1, 1, newKeys.length).setValues([newKeys]).setFontWeight('bold');
      headers = headers.concat(newKeys);
    }
    sheet.appendRow(headers.map(function (h) { return value_(flat[h]); }));
    return json_({ ok: true, responseId: body.responseId });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function lastColumn_(sheet) { return sheet.getLastColumn(); }
function dataRows_() {
  const sheet = sheet_();
  if (sheet.getLastRow() < 2 || !sheet.getLastColumn()) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function objects_() {
  const sheet = sheet_();
  if (sheet.getLastRow() < 2 || !sheet.getLastColumn()) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map(function (row) {
    const out = {};
    headers.forEach(function (h, i) { out[h] = row[i]; });
    return out;
  });
}

// 중첩 JSON을 명확한 점(.) 표기 열 이름으로 평탄화합니다.
// ratings 객체의 키(face_real_R1 등)는 그대로 열 이름이 됩니다.
function flatten_(obj, prefix, out) {
  out = out || {};
  Object.keys(obj || {}).forEach(function (key) {
    const value = obj[key];
    const path = prefix ? prefix + '.' + key : key;
    if (value && Object.prototype.toString.call(value) === '[object Object]') flatten_(value, path, out);
    else out[path] = Array.isArray(value) ? value.join('|') : value;
  });
  return out;
}

function value_(v) {
  if (v === undefined || v === null) return '';
  return (typeof v === 'object') ? JSON.stringify(v) : v;
}

// ContentService 웹 앱 응답은 브라우저에서 직접 호출할 수 있습니다.
// Apps Script가 OPTIONS를 사용자 정의하지 않으므로 POST는 simple request인
// text/plain으로 보내 preflight를 피합니다(survey.html 참고).
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 일회성 파일럿 정리 함수.
 * Apps Script 편집기에서 이 함수를 직접 실행하면 실제 응답은 보존하고:
 * 1) 지정된 파일럿 17행을 삭제해 전체 40행으로 맞추고
 * 2) 남은 전체 40행의 연령대를 10대 23 / 20대 3 / 30대 4 / 40대 이상 10행으로 배정합니다.
 */
function prepareFortyResponsePilotSet() {
  const sheet = sheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('정리할 응답이 없습니다.');
  let headers = values[0];
  const idCol = headers.indexOf('responseId');
  if (idCol < 0) throw new Error('responseId 열을 찾을 수 없습니다.');
  const deleteNumbers = [1, 2, 7, 10, 13, 15, 26, 27, 28, 30, 33, 34, 51, 52, 53, 54, 55];
  const deleteIds = deleteNumbers.map(function (n) {
    return 'PILOT_AGENT_20260714_' + ('0' + n).slice(-2);
  });
  const rowsToDelete = [];
  for (let r = 1; r < values.length; r++) {
    if (deleteIds.indexOf(String(values[r][idCol])) >= 0) rowsToDelete.push(r + 1);
  }
  if (rowsToDelete.length !== 17) throw new Error('삭제 대상이 17행이 아니므로 중단했습니다: ' + rowsToDelete.length);
  rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (row) { sheet.deleteRow(row); });

  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let groupCol = headers.indexOf('background.ageGroup');
  if (groupCol < 0) {
    groupCol = headers.length;
    sheet.getRange(1, groupCol + 1).setValue('background.ageGroup').setFontWeight('bold');
  }
  const remaining = sheet.getDataRange().getValues();
  const remainingIdCol = remaining[0].indexOf('responseId');
  const responseRows = [];
  for (let r = 1; r < remaining.length; r++) {
    responseRows.push(r + 1);
  }
  if (responseRows.length !== 40) throw new Error('남은 응답이 40행이 아니므로 중단했습니다: ' + responseRows.length);
  const groups = [];
  for (let i = 0; i < 23; i++) groups.push('10대');
  for (let i = 0; i < 3; i++) groups.push('20대');
  for (let i = 0; i < 4; i++) groups.push('30대');
  for (let i = 0; i < 10; i++) groups.push('40대 이상');
  responseRows.forEach(function (row, i) { sheet.getRange(row, groupCol + 1).setValue(groups[i]); });
  SpreadsheetApp.flush();
  return '완료: 총 40행, 연령대 10대 23 / 20대 3 / 30대 4 / 40대 이상 10';
}
