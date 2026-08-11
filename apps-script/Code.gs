/**
 * 슬기로운 이스린생활 — 백엔드
 * 이 파일을 구글시트의 확장 프로그램 > Apps Script 에 그대로 붙여넣고 웹 앱으로 배포하세요.
 * 자세한 순서는 SETUP.md 참고.
 */

var SHEET_NAME = 'Data';

// 페이지 잠금 암호와 동일하게 맞춰두세요. 나중에 암호를 바꾸면 여기도 같이 바꿔야 해요.
var SHARED_KEY = '1234';

function doGet(e) {
  if (e.parameter.key !== SHARED_KEY) {
    return jsonOutput({ error: 'unauthorized' });
  }
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var header = rows.shift();
  var items = rows
    .filter(function (row) { return row[0]; })
    .map(function (row) {
      var obj = {};
      header.forEach(function (key, i) { obj[key] = cellToValue(row[i]); });
      return obj;
    });
  return jsonOutput({ items: items });
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput({ error: 'bad_request' });
  }
  if (data.key !== SHARED_KEY) {
    return jsonOutput({ error: 'unauthorized' });
  }
  var sheet = getSheet();
  // "2026-08-15" 같은 날짜 형식 문자열을 시트가 자동으로 날짜 셀로 바꿔버리는 걸 막기 위해
  // 값을 쓰기 전에 셀 서식을 텍스트로 고정해둡니다.
  var range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, 6);
  range.setNumberFormat('@');
  range.setValues([[
    new Date().toISOString(),
    data.category || '',
    data.author || '',
    data.title || '',
    data.content || '',
    data.link || ''
  ]]);
  return jsonOutput({ ok: true });
}

// 시트에 이미 날짜 셀로 저장된 값이 있어도 안전하게 문자열로 돌려줍니다.
function cellToValue(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['timestamp', 'category', 'author', 'title', 'content', 'link']);
  }
  return sheet;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
