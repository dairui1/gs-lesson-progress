/**
 * Lessons.js
 * 「上課紀錄」表的 upsert 與欄位工具。
 *
 * 注意：客戶當前 schema 不含 Meeting UUID / Zoom主題，Webhook upsert 將無法去重定位，
 *       實際會每次 insert 新列。待未來客戶加上這兩欄後才具備真正的去重能力。
 */

function _getLessonsSheet_() {
  var ss = SpreadsheetApp.openById(getMasterId_());
  var sh = ss.getSheetByName(SHEET_LESSONS);
  if (!sh) throw new Error('找不到「' + SHEET_LESSONS + '」分頁');
  return sh;
}

function findLessonRowByUuid_(uuid) {
  if (!uuid) return null;
  var sh = _getLessonsSheet_();
  var head = readHeaders_(sh);
  if (head.idx['Meeting UUID'] === undefined) return null;
  var last = sh.getLastRow();
  if (last < 2) return null;
  var uuidCol = head.idx['Meeting UUID'] + 1;
  var values = sh.getRange(2, uuidCol, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === String(uuid).trim()) {
      return { rowIndex: i + 2 };
    }
  }
  return null;
}

function upsertLesson(payload) {
  var sh = _getLessonsSheet_();
  var head = readHeaders_(sh);
  var existing = payload.uuid ? findLessonRowByUuid_(payload.uuid) : null;

  if (existing) {
    _updateLessonRow_(sh, head, existing.rowIndex, payload);
    return { action: 'update', rowIndex: existing.rowIndex };
  }

  var rowIndex = _insertLessonRow_(sh, head, payload);
  var counter = _bumpStudentLessonCount_(payload.student && payload.student['學生姓名']);
  return { action: 'insert', rowIndex: rowIndex, counter: counter };
}

/**
 * 新增一堂課時，到「學生資料」對應列：已上課堂數 +1、剩餘堂數 -1。
 * 允許剩餘堂數變成負數，但會 logEvent_ 提醒。
 */
function _bumpStudentLessonCount_(studentName) {
  if (!studentName) return null;
  var ss = SpreadsheetApp.openById(getMasterId_());
  var sh = ss.getSheetByName(SHEET_STUDENTS);
  if (!sh) return null;
  var head = readHeaders_(sh);
  var nameCol = head.idx['學生姓名'];
  var doneCol = head.idx['已上課堂數'];
  var leftCol = head.idx['剩餘堂數'];
  if (nameCol === undefined || (doneCol === undefined && leftCol === undefined)) return null;

  var last = sh.getLastRow();
  if (last < 2) return null;
  var values = sh.getRange(2, 1, last - 1, head.lastCol).getValues();
  var target = String(studentName).trim();
  var rowIndex = -1;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][nameCol] || '').trim() === target) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex < 0) {
    logEvent_('lesson_count', '找不到學生，跳過堂數遞增', { name: studentName });
    return null;
  }

  var done = null;
  var left = null;
  if (doneCol !== undefined) {
    var dRaw = values[rowIndex - 2][doneCol];
    var dNum = Number(dRaw);
    done = (isNaN(dNum) ? 0 : dNum) + 1;
    sh.getRange(rowIndex, doneCol + 1).setValue(done);
  }
  if (leftCol !== undefined) {
    var lRaw = values[rowIndex - 2][leftCol];
    var lNum = Number(lRaw);
    left = (isNaN(lNum) ? 0 : lNum) - 1;
    sh.getRange(rowIndex, leftCol + 1).setValue(left);
    if (left < 0) {
      logEvent_('lesson_count', '剩餘堂數已為負數，請補堂', {
        name: studentName, 已上課堂數: done, 剩餘堂數: left
      });
    }
  }
  return { rowIndex: rowIndex, 已上課堂數: done, 剩餘堂數: left };
}

function _insertLessonRow_(sh, head, payload) {
  var row = new Array(head.lastCol).fill('');
  _setIfHas_(row, head, 'ID', Utilities.getUuid());
  _setIfHas_(row, head, '日期', _formatLessonDate_(payload.startTime));
  _setIfHas_(row, head, '學生姓名', payload.student && payload.student['學生姓名']);
  _setIfHas_(row, head, '家長Email', payload.student && payload.student['家長Email']);
  _setIfHas_(row, head, '上課內容', payload.lessonContent || '');
  _setIfHas_(row, head, '作業', payload.homeworkFromAI || '');
  _setIfHas_(row, head, '上課時數(hr)', _normalizeHr_(payload.durationHr));
  _setIfHas_(row, head, '學習狀況/建議(會議摘要)', payload.summaryFullText || '');
  _setIfHas_(row, head, '會議摘要連結', payload.summaryDocUrl || '');
  _setIfHas_(row, head, 'Zoom主題', payload.topic || '');
  _setIfHas_(row, head, 'Meeting UUID', payload.uuid || '');
  _setIfHas_(row, head, '發信', false);
  _setIfHas_(row, head, 'EmailSent', false);
  sh.appendRow(row);
  return sh.getLastRow();
}

function _updateLessonRow_(sh, head, rowIndex, payload) {
  var row = sh.getRange(rowIndex, 1, 1, head.lastCol).getValues()[0];
  var writes = [];

  function setIfEmpty(header, value) {
    if (value === undefined || value === null || value === '') return;
    if (head.idx[header] === undefined) return;
    var i = head.idx[header];
    if (row[i] === '' || row[i] === null || row[i] === undefined) {
      writes.push({ col: i + 1, value: value });
    }
  }

  function setAlways(header, value) {
    if (value === undefined || value === null || value === '') return;
    if (head.idx[header] === undefined) return;
    writes.push({ col: head.idx[header] + 1, value: value });
  }

  setIfEmpty('日期', _formatLessonDate_(payload.startTime));
  setIfEmpty('上課時數(hr)', _normalizeHr_(payload.durationHr));
  setIfEmpty('Zoom主題', payload.topic);
  setIfEmpty('學生姓名', payload.student && payload.student['學生姓名']);
  setIfEmpty('家長Email', payload.student && payload.student['家長Email']);

  if (payload.source === 'summary') {
    setAlways('上課內容', payload.lessonContent);
    setAlways('學習狀況/建議(會議摘要)', payload.summaryFullText);
    setAlways('會議摘要連結', payload.summaryDocUrl);
    setIfEmpty('作業', payload.homeworkFromAI);
  }

  writes.forEach(function (w) {
    sh.getRange(rowIndex, w.col).setValue(w.value);
  });
}

function _setIfHas_(row, head, header, value) {
  if (value === undefined || value === null) return;
  if (head.idx[header] === undefined) return;
  row[head.idx[header]] = value;
}

function _formatLessonDate_(startTime) {
  if (!startTime) return '';
  var d = new Date(startTime);
  if (isNaN(d.getTime())) return '';
  return formatDateTz_(d, 'yyyy-MM-dd');
}

function _normalizeHr_(hr) {
  if (hr === null || hr === undefined || hr === '') return '';
  var n = Number(hr);
  if (isNaN(n)) return '';
  return Math.round(n * 10) / 10;
}

function computeDurationHr_(startTime, endTime, fallbackMinutes) {
  if (startTime && endTime) {
    var ms = new Date(endTime).getTime() - new Date(startTime).getTime();
    if (!isNaN(ms) && ms > 0) return ms / 3600000;
  }
  if (fallbackMinutes != null && fallbackMinutes !== '') {
    var m = Number(fallbackMinutes);
    if (!isNaN(m)) return m / 60;
  }
  return '';
}
