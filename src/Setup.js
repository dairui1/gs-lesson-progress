/**
 * Setup.js
 * 一次性初始化、老師端選單、新增學生等互動入口。
 */

function initMasterSpreadsheet() {
  var masterId = null;
  try { masterId = getMasterId_(); } catch (e) { /* not set yet */ }

  var ss;
  if (masterId) {
    ss = SpreadsheetApp.openById(masterId);
  } else {
    ss = SpreadsheetApp.create('上課進度 Master（老師）');
    masterId = ss.getId();
    setProp_(PROP_MASTER_ID, masterId);
  }

  // 學生資料
  if (!ss.getSheetByName(SHEET_STUDENTS)) {
    var stu = ss.insertSheet(SHEET_STUDENTS);
    stu.appendRow(STUDENTS_HEADERS);
    stu.setFrozenRows(1);
    stu.getRange(1, 1, 1, STUDENTS_HEADERS.length).setFontWeight('bold');
  }

  // 上課紀錄
  ensureLessonsSheet_();

  // Log
  if (!ss.getSheetByName(SHEET_LOG)) {
    var log = ss.insertSheet(SHEET_LOG);
    log.appendRow(LOG_HEADERS);
    log.setFrozenRows(1);
    log.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
  }

  // 清掉預設空白分頁
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('工作表1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(def);
  }

  // 安裝 triggers：onOpen（選單） + sendPendingEmails（每 5 分鐘）
  installAllTriggers_(masterId);

  // 自動產生 URL token（若尚未設定）
  if (!getUrlToken_()) {
    setProp_(PROP_URL_TOKEN, Utilities.getUuid().replace(/-/g, ''));
  }

  console.log('Master ready:', ss.getUrl());
  console.log('URL token:', getUrlToken_());
  return { id: masterId, url: ss.getUrl(), urlToken: getUrlToken_() };
}

function installAllTriggers_(ssId) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onOpen' || fn === 'sendPendingEmails') {
      try { ScriptApp.deleteTrigger(t); } catch (e) { /* noop */ }
    }
  });
  ScriptApp.newTrigger('onOpen').forSpreadsheet(ssId).onOpen().create();
  ScriptApp.newTrigger('sendPendingEmails').timeBased().everyMinutes(5).create();
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('老師工具')
      .addItem('新增學生', 'uiAddStudent')
      .addItem('立即掃描並寄送待發 Email', 'uiSendNow')
      .addSeparator()
      .addItem('顯示 Webhook 設定資訊', 'uiShowWebhookInfo')
      .addItem('初始化／檢查設定', 'initMasterSpreadsheet')
      .addToUi();
  } catch (e) {
    console.warn('onOpen menu failed:', e && e.message);
  }
}

function uiAddStudent() {
  var ui = SpreadsheetApp.getUi();
  var nameRes = ui.prompt('新增學生（1/3）', '學生姓名：', ui.ButtonSet.OK_CANCEL);
  if (nameRes.getSelectedButton() !== ui.Button.OK) return;
  var emailRes = ui.prompt('新增學生（2/3）', '家長 Email：', ui.ButtonSet.OK_CANCEL);
  if (emailRes.getSelectedButton() !== ui.Button.OK) return;
  var zoomRes = ui.prompt('新增學生（3/3）', 'Zoom Meeting ID（純數字）：', ui.ButtonSet.OK_CANCEL);
  if (zoomRes.getSelectedButton() !== ui.Button.OK) return;

  try {
    var res = addStudent({
      name: nameRes.getResponseText().trim(),
      parentEmail: emailRes.getResponseText().trim(),
      zoomMeetingId: zoomRes.getResponseText().trim()
    });
    ui.alert('已新增：' + res['學生姓名'] + '（' + res.student_id + '）');
  } catch (e) {
    ui.alert('失敗：' + (e && e.message ? e.message : e));
  }
}

function uiSendNow() {
  var ui = SpreadsheetApp.getUi();
  try {
    var res = sendPendingEmails();
    ui.alert('掃描完成：寄出 ' + res.sent + '、失敗 ' + res.failed + '、跳過 ' + res.skipped);
  } catch (e) {
    ui.alert('失敗：' + (e && e.message ? e.message : e));
  }
}

function uiShowWebhookInfo() {
  var ui = SpreadsheetApp.getUi();
  var token = getUrlToken_() || '(尚未產生，請先執行 initMasterSpreadsheet)';
  var secretSet = !!getZoomSecret_();
  ui.alert([
    '— Webhook 設定 —',
    'URL token（附加到 /exec?token=）：',
    token,
    '',
    'ZOOM_WEBHOOK_SECRET 是否已設定：' + (secretSet ? '是' : '否（請到 Apps Script 專案設定 → Script Properties 加入）'),
    '',
    '部署：clasp deploy → 取得 /exec URL',
    'Zoom Endpoint URL 填：<YOUR_EXEC_URL>?token=' + token,
    '訂閱事件：Meeting → End Meeting、Meeting → Summary completed'
  ].join('\n'));
}

/**
 * 新增學生：只寫入「學生資料」表，不再建立 mirror。
 */
function addStudent(opts) {
  var name = String((opts && opts.name) || '').trim();
  var parentEmail = String((opts && opts.parentEmail) || '').trim();
  var zoomMeetingId = String((opts && opts.zoomMeetingId) || '').replace(/\s+/g, '');

  if (!name) throw new Error('學生姓名必填');
  if (parentEmail.indexOf('@') < 0) throw new Error('家長 Email 格式錯誤');

  if (zoomMeetingId) {
    var exist = findByZoomMeetingId(zoomMeetingId);
    if (exist) {
      throw new Error('此 Zoom Meeting ID 已綁定到 ' + exist.student_id + ' / ' + exist['學生姓名']);
    }
  }

  var ss = SpreadsheetApp.openById(getMasterId_());
  var registry = ss.getSheetByName(SHEET_STUDENTS);
  var studentId = Students_nextId_();

  registry.appendRow([
    studentId,
    name,
    parentEmail,
    zoomMeetingId,
    ''
  ]);

  return {
    student_id: studentId,
    '學生姓名': name,
    '家長email': parentEmail,
    zoom_meeting_id: zoomMeetingId
  };
}
