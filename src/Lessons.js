/**
 * Lessons.js
 * 「上課紀錄」表的 upsert 與欄位工具。
 *
 * upsertLesson(payload)：
 *   - 若已有相同 Meeting UUID 的列：就地更新「未填」欄位；對摘要相關欄位以 summary 事件覆寫為主。
 *   - 否則：新增一列，並指派新的 lesson_id。
 *
 * payload 結構：
 *   {
 *     source: 'summary' | 'ended',
 *     student: { student_id, '學生姓名', '家長email' },
 *     uuid, topic, startTime, durationHr,
 *     summaryDocUrl, homeworkFromAI
 *   }
 */

function ensureLessonsSheet_() {
  var ss = SpreadsheetApp.openById(getMasterId_());
  var sh = ss.getSheetByName(SHEET_LESSONS);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_LESSONS);
  sh.appendRow(LESSONS_HEADERS);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, LESSONS_HEADERS.length).setFontWeight('bold');
  // 欄寬微調
  var widths = [0, 180, 80, 140, 120, 180, 180, 280, 80, 260, 180, 220, 60, 70, 150, 220];
  for (var c = 1; c < widths.length; c++) {
    if (widths[c]) sh.setColumnWidth(c, widths[c]);
  }
  return sh;
}

function _lessonsIndex_() {
  var idx = {};
  LESSONS_HEADERS.forEach(function (h, i) { idx[h] = i; });
  return idx;
}

function findLessonRowByUuid_(uuid) {
  if (!uuid) return null;
  var sh = ensureLessonsSheet_();
  var last = sh.getLastRow();
  if (last < 2) return null;
  var idx = _lessonsIndex_();
  var uuidCol = idx['Meeting UUID'] + 1;
  var values = sh.getRange(2, uuidCol, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === String(uuid).trim()) {
      return { rowIndex: i + 2 };
    }
  }
  return null;
}

function upsertLesson(payload) {
  var sh = ensureLessonsSheet_();
  var idx = _lessonsIndex_();
  var existing = payload.uuid ? findLessonRowByUuid_(payload.uuid) : null;

  if (existing) {
    _updateLessonRow_(sh, idx, existing.rowIndex, payload);
    return { action: 'update', rowIndex: existing.rowIndex };
  }

  var rowIndex = _insertLessonRow_(sh, idx, payload);
  return { action: 'insert', rowIndex: rowIndex };
}

function _insertLessonRow_(sh, idx, payload) {
  var row = new Array(LESSONS_HEADERS.length).fill('');
  row[idx['lesson_id']] = Utilities.getUuid();
  row[idx['student_id']] = (payload.student && payload.student.student_id) || '';
  row[idx['日期']] = _formatLessonDate_(payload.startTime);
  row[idx['學生姓名']] = (payload.student && payload.student['學生姓名']) || '';
  row[idx['家長email']] = (payload.student && payload.student['家長email']) || '';
  row[idx['上課內容']] = ''; // 老師在 AppSheet 補
  row[idx['作業']] = payload.homeworkFromAI || '';
  row[idx['上課時數(hr)']] = _normalizeHr_(payload.durationHr);
  row[idx['學習狀況/建議(會議摘要)']] = payload.summaryDocUrl || '';
  row[idx['Zoom主題']] = payload.topic || '';
  row[idx['Meeting UUID']] = payload.uuid || '';
  row[idx['發信']] = false;
  row[idx['EmailSent']] = false;
  row[idx['SentAt']] = '';
  row[idx['ErrorMessage']] = '';
  sh.appendRow(row);
  return sh.getLastRow();
}

/**
 * 更新策略：
 *   - 日期、時數、Zoom主題、學生姓名/Email：僅在目前為空時填入（避免覆寫老師已修正的值）
 *   - 摘要相關（學習狀況/建議(會議摘要)、作業）：只要 source=summary 就覆寫；source=ended 不動
 */
function _updateLessonRow_(sh, idx, rowIndex, payload) {
  var row = sh.getRange(rowIndex, 1, 1, LESSONS_HEADERS.length).getValues()[0];
  var writes = [];

  function setIfEmpty(header, value) {
    if (value === undefined || value === null || value === '') return;
    var i = idx[header];
    if (row[i] === '' || row[i] === null || row[i] === undefined) {
      writes.push({ col: i + 1, value: value });
    }
  }

  function setAlways(header, value) {
    if (value === undefined || value === null || value === '') return;
    writes.push({ col: idx[header] + 1, value: value });
  }

  setIfEmpty('日期', _formatLessonDate_(payload.startTime));
  setIfEmpty('上課時數(hr)', _normalizeHr_(payload.durationHr));
  setIfEmpty('Zoom主題', payload.topic);
  setIfEmpty('學生姓名', payload.student && payload.student['學生姓名']);
  setIfEmpty('家長email', payload.student && payload.student['家長email']);
  setIfEmpty('student_id', payload.student && payload.student.student_id);

  if (payload.source === 'summary') {
    setAlways('學習狀況/建議(會議摘要)', payload.summaryDocUrl);
    // 作業只在目前為空才填 AI 版本，避免覆蓋老師已手動改寫的內容
    setIfEmpty('作業', payload.homeworkFromAI);
  }

  writes.forEach(function (w) {
    sh.getRange(rowIndex, w.col).setValue(w.value);
  });
}

function _formatLessonDate_(startTime) {
  if (!startTime) return '';
  var d = new Date(startTime);
  if (isNaN(d.getTime())) return '';
  return formatDateTz_(d, 'yyyy-MM-dd HH:mm');
}

function _normalizeHr_(hr) {
  if (hr === null || hr === undefined || hr === '') return '';
  var n = Number(hr);
  if (isNaN(n)) return '';
  return Math.round(n * 10) / 10; // 1 位小數
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
