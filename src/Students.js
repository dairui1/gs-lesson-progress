/**
 * Students.js
 * 「學生資料」表的查找。以「學生姓名」為事實主鍵。
 */

function Students_all_() {
  var ss = SpreadsheetApp.openById(getMasterId_());
  var sh = ss.getSheetByName(SHEET_STUDENTS);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var head = readHeaders_(sh);
  if (head.lastCol < 1) return [];
  var values = sh.getRange(2, 1, last - 1, head.lastCol).getValues();
  return values.map(function (row, i) {
    var o = { _rowIndex: i + 2 };
    head.headers.forEach(function (h, idx) {
      if (h) o[h] = row[idx];
    });
    return o;
  });
}

function Students_find(key, value) {
  if (value === null || value === undefined || value === '') return null;
  var target = String(value).trim();
  var list = Students_all_();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i][key] || '').trim() === target) return list[i];
  }
  return null;
}

/**
 * 依學生姓名查家長 Email；找不到或空字串回傳 ''。
 */
function lookupParentEmailByName_(name) {
  var s = Students_find('學生姓名', name);
  if (!s) return '';
  return String(s['家長Email'] || '').trim();
}

/**
 * 依 Zoom meeting_id 找學生；若學生資料沒有 zoom_meeting_id 欄位，自然回傳 null。
 */
function findByZoomMeetingId(meetingId) {
  if (!meetingId) return null;
  var normalized = String(meetingId).replace(/\s+/g, '');
  var list = Students_all_();
  for (var i = 0; i < list.length; i++) {
    var stored = String(list[i].zoom_meeting_id || '').replace(/\s+/g, '');
    if (stored && stored === normalized) return list[i];
  }
  return null;
}

/**
 * 以會議主題反推學生：嘗試以「學生姓名」子字串匹配。
 */
function findByMeetingTopic_(topic) {
  if (!topic) return null;
  var t = String(topic);
  var list = Students_all_();
  for (var i = 0; i < list.length; i++) {
    var name = String(list[i]['學生姓名'] || '').trim();
    if (name && t.indexOf(name) >= 0) return list[i];
  }
  return null;
}
