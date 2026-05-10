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
  return { action: 'insert', rowIndex: rowIndex };
}

/**
 * 一次性：把學生資料的「已上課堂數 / 剩餘堂數」改成公式驅動。
 *   已上課堂數 = COUNTIF(上課紀錄!學生姓名列, 本行學生姓名)
 *   剩餘堂數   = 總堂數 - 已上課堂數
 * 之後新增 / 刪除上課紀錄列時會自動同步，不需手動維護。
 * 在 Apps Script Editor 選擇此函式按 Run 即可。
 */
function installLessonCountFormulas() {
  var ss = SpreadsheetApp.openById(getMasterId_());
  var stuSh = ss.getSheetByName(SHEET_STUDENTS);
  var lesSh = ss.getSheetByName(SHEET_LESSONS);
  if (!stuSh || !lesSh) throw new Error('找不到必要分頁');

  var stuHead = readHeaders_(stuSh);
  var lesHead = readHeaders_(lesSh);
  var nameStuIdx = stuHead.idx['學生姓名'];
  var totalIdx = stuHead.idx['總堂數'];
  var doneIdx = stuHead.idx['已上課堂數'];
  var leftIdx = stuHead.idx['剩餘堂數'];
  var nameLesIdx = lesHead.idx['學生姓名'];

  if (nameStuIdx === undefined || doneIdx === undefined || nameLesIdx === undefined) {
    throw new Error('表頭缺欄位：學生姓名 / 已上課堂數 / 上課紀錄學生姓名');
  }

  var lastRow = stuSh.getLastRow();
  if (lastRow < 2) return 'no student rows';

  var nameCol = _colLetter_(nameStuIdx + 1);
  var lesNameCol = _colLetter_(nameLesIdx + 1);
  var doneCol = _colLetter_(doneIdx + 1);
  var totalCol = totalIdx !== undefined ? _colLetter_(totalIdx + 1) : null;
  var n = lastRow - 1;

  var doneFormulas = [];
  var leftFormulas = [];
  for (var r = 2; r <= lastRow; r++) {
    doneFormulas.push(["=COUNTIF('" + SHEET_LESSONS + "'!" + lesNameCol + ':' + lesNameCol + ', ' + nameCol + r + ')']);
    if (leftIdx !== undefined && totalCol) {
      leftFormulas.push(['=' + totalCol + r + '-' + doneCol + r]);
    }
  }
  stuSh.getRange(2, doneIdx + 1, n, 1).setFormulas(doneFormulas);
  if (leftFormulas.length) {
    stuSh.getRange(2, leftIdx + 1, n, 1).setFormulas(leftFormulas);
  }
  return 'installed for ' + n + ' student rows';
}

function _colLetter_(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
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
