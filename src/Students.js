/**
 * Students.js
 * 「學生資料」表的查找與更新。
 */

function Students_all_() {
  var ss = SpreadsheetApp.openById(getMasterId_());
  var sh = ss.getSheetByName(SHEET_STUDENTS);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, STUDENTS_HEADERS.length).getValues();
  return values.map(function (row, i) {
    var o = { _rowIndex: i + 2 };
    STUDENTS_HEADERS.forEach(function (h, idx) {
      o[h] = row[idx];
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
 * 依 Zoom meeting_id 找到學生；兩邊都正規化為無空白字串。
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
 * 以會議名稱 fallback 反推學生：找「S\d+」或「學生姓名」子字串。
 */
function findByMeetingTopic_(topic) {
  if (!topic) return null;
  var t = String(topic);
  var m = /\b(S\d{3,})\b/.exec(t);
  if (m) {
    var byId = Students_find('student_id', m[1]);
    if (byId) return byId;
  }
  var list = Students_all_();
  for (var i = 0; i < list.length; i++) {
    var name = String(list[i]['學生姓名'] || '').trim();
    if (name && t.indexOf(name) >= 0) return list[i];
  }
  return null;
}

function Students_update(studentId, patch) {
  var ss = SpreadsheetApp.openById(getMasterId_());
  var sh = ss.getSheetByName(SHEET_STUDENTS);
  var s = Students_find('student_id', studentId);
  if (!s) throw new Error('student not found: ' + studentId);
  STUDENTS_HEADERS.forEach(function (h, idx) {
    if (Object.prototype.hasOwnProperty.call(patch, h)) {
      sh.getRange(s._rowIndex, idx + 1).setValue(patch[h]);
    }
  });
}

function Students_nextId_() {
  var list = Students_all_();
  var max = list.reduce(function (acc, s) {
    var m = /^S(\d+)$/.exec(String(s.student_id || ''));
    return m ? Math.max(acc, parseInt(m[1], 10)) : acc;
  }, 0);
  return 'S' + String(max + 1).padStart(3, '0');
}
