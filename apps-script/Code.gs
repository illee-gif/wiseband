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

  if (data.action === 'delete') {
    return handleDelete(data);
  }

  if (data.action === 'update') {
    return handleUpdate(data);
  }

  var link = data.link || '';
  if (data.fileData) {
    try {
      var blob = Utilities.newBlob(
        Utilities.base64Decode(data.fileData),
        data.mimeType || 'application/octet-stream',
        data.filename || 'file'
      );
      var file = getUploadFolder(data.category).createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      link = file.getUrl();
    } catch (err) {
      return jsonOutput({ error: 'upload_failed', message: String(err) });
    }
  }

  var sheet = getSheet();
  // "2026-08-15" 같은 날짜 형식 문자열을 시트가 자동으로 날짜 셀로 바꿔버리는 걸 막기 위해
  // 값을 쓰기 전에 셀 서식을 텍스트로 고정해둡니다.
  var range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, 7);
  range.setNumberFormat('@');
  range.setValues([[
    new Date().toISOString(),
    data.category || '',
    data.author || '',
    data.title || '',
    data.content || '',
    link,
    data.part || ''
  ]]);
  return jsonOutput({ ok: true, link: link });
}

function handleDelete(data) {
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.timestamp)) {
      trashDriveFileIfAny(values[i][5]);
      sheet.deleteRow(i + 1);
      return jsonOutput({ ok: true });
    }
  }
  return jsonOutput({ error: 'not_found' });
}

function handleUpdate(data) {
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var editableFields = ['author', 'title', 'content', 'link', 'part'];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.timestamp)) {
      var rowIndex = i + 1;
      editableFields.forEach(function (field) {
        if (data[field] === undefined) return;
        var col = header.indexOf(field) + 1;
        if (col > 0) {
          sheet.getRange(rowIndex, col).setNumberFormat('@').setValue(data[field]);
        }
      });
      return jsonOutput({ ok: true });
    }
  }
  return jsonOutput({ error: 'not_found' });
}

function trashDriveFileIfAny(link) {
  if (!link) return;
  var m = String(link).match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (!m) return;
  try {
    DriveApp.getFileById(m[1]).setTrashed(true);
  } catch (err) {
    // 이미 지워졌거나 접근 권한이 없어도 시트 행 삭제는 계속 진행합니다.
  }
}

var UPLOAD_FOLDERS = {
  score: '이스린생활_악보',
  recording: '이스린생활_연습기록'
};

function getUploadFolder(category) {
  var name = UPLOAD_FOLDERS[category] || '이스린생활_기타자료';
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
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
    sheet.appendRow(['timestamp', 'category', 'author', 'title', 'content', 'link', 'part']);
    return sheet;
  }
  // 예전에 만들어진 시트에는 'part' 칸이 없을 수 있어서, 없으면 자동으로 추가해줍니다.
  var lastCol = sheet.getLastColumn();
  var header = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (header.indexOf('part') === -1) {
    sheet.getRange(1, header.length + 1).setValue('part');
  }
  return sheet;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 드라이브 쓰기 권한을 승인받기 위한 용도입니다. 에디터에서 이 함수를 선택해 한 번 실행해주세요.
function grantDrivePermission() {
  var folder = DriveApp.createFolder('__permission_test__');
  folder.setTrashed(true);
}
