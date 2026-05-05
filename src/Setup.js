/**
 * Setup.js
 * Container-bound 版：此腳本綁在客戶既有 Sheet 上。
 * 首次部署時由使用者手動執行 attachToExistingSpreadsheet()
 * → 寫入 Script Property、校驗必需表頭、安裝排程觸發器。
 */

function attachToExistingSpreadsheet() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('請在綁定的 Google Sheet 內執行此函式');

  var ssId = ss.getId();
  setProp_(PROP_MASTER_ID, ssId);

  var missingSheets = [];
  [SHEET_STUDENTS, SHEET_LESSONS].forEach(function (name) {
    if (!ss.getSheetByName(name)) missingSheets.push(name);
  });
  if (missingSheets.length) {
    throw new Error('找不到必要分頁：' + missingSheets.join(', '));
  }

  var warnings = [];
  var stuHead = readHeaders_(ss.getSheetByName(SHEET_STUDENTS));
  var lesHead = readHeaders_(ss.getSheetByName(SHEET_LESSONS));

  STUDENTS_HEADERS.forEach(function (h) {
    if (stuHead.idx[h] === undefined) warnings.push('「' + SHEET_STUDENTS + '」缺欄位：' + h);
  });
  LESSONS_HEADERS.forEach(function (h) {
    if (lesHead.idx[h] === undefined) warnings.push('「' + SHEET_LESSONS + '」缺欄位：' + h);
  });

  installAllTriggers_(ssId);

  if (!getUrlToken_()) {
    setProp_(PROP_URL_TOKEN, Utilities.getUuid().replace(/-/g, ''));
  }

  var info = {
    spreadsheetId: ssId,
    url: ss.getUrl(),
    urlToken: getUrlToken_(),
    warnings: warnings
  };
  console.log('attachToExistingSpreadsheet OK:', JSON.stringify(info));
  if (warnings.length) {
    console.warn('表頭缺失（寄信流程所需欄位）：\n' + warnings.join('\n'));
  }
  return info;
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
      .addItem('使用說明', 'uiShowHelp')
      .addSeparator()
      .addItem('立即掃描並寄送待發 Email', 'uiSendNow')
      .addSeparator()
      .addItem('顯示 Webhook 設定資訊', 'uiShowWebhookInfo')
      .addItem('連結此 Sheet 並啟用寄信', 'attachToExistingSpreadsheet')
      .addItem('鎖定所有歷史列（EmailSent=TRUE）', 'markAllHistoricalAsSent')
      .addToUi();
  } catch (e) {
    console.warn('onOpen menu failed:', e && e.message);
  }
}

function uiShowHelp() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('家教上課紀錄系統 — 使用說明', [
    '【主要流程】',
    '1. Zoom 上完課，AI Companion 摘要產生後（約 5–10 分鐘），',
    '   「上課紀錄」分頁會自動新增一列，並自動填入：',
    '     ・日期 / 上課時數',
    '     ・上課內容（Zoom 摘要的 Overview）',
    '     ・作業(Zoom 摘要的 Next Steps）',
    '     ・學習狀況/建議（會議摘要全文）',
    '     ・會議摘要連結',
    '',
    '2. 檢查 AI 帶入的內容，可直接在表格修改。',
    '',
    '3. 確認無誤後，把該列的「發信」欄位勾起來。',
    '',
    '4. 系統每 5 分鐘自動掃描，把勾選的列寄信給家長。',
    '   寄出後：EmailSent 自動勾起 / SentAt 填入時間。',
    '',
    '【常用操作】',
    '・想立刻寄、不等 5 分鐘：選單「立即掃描並寄送待發 Email」',
    '・想重寄某列：把該列 EmailSent 取消勾選，再勾「發信」',
    '・「家長Email」空白時，系統會用「學生姓名」到「學生資料」分頁反查',
    '',
    '【寄件人】',
    '寄信使用本 Google 帳號（家長按回覆會回到老師信箱）。',
    '想改顯示名稱：Apps Script → 專案設定 → 指令碼屬性 →',
    '新增 EMAIL_SENDER_NAME = 你想顯示的名字',
    '',
    '【寄信額度】',
    '免費 Gmail：100 封/天　|　Google Workspace：1500 封/天'
  ].join('\n'), ui.ButtonSet.OK);
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

/**
 * 一次性工具：把「上課紀錄」現有所有列標為 EmailSent=TRUE，
 * 讓歷史資料（含老師已手動寄過、或覺得不用寄的）全部被鎖住，
 * 即便之後有人誤勾「發信」也不會觸發寄信。
 *
 * 需要重發某一列？將該列 EmailSent 清為 FALSE / 空，再勾發信即可。
 */
function markAllHistoricalAsSent() {
  var ss = SpreadsheetApp.openById(getMasterId_());
  var sh = ss.getSheetByName(SHEET_LESSONS);
  if (!sh) throw new Error('找不到「' + SHEET_LESSONS + '」分頁');

  var head = readHeaders_(sh);
  if (head.idx['EmailSent'] === undefined) {
    throw new Error('「' + SHEET_LESSONS + '」缺少 EmailSent 欄位');
  }

  var last = sh.getLastRow();
  if (last < 2) {
    console.log('markAllHistoricalAsSent: 沒有現有資料可鎖');
    return { locked: 0 };
  }

  var nRows = last - 1;
  var emailSentCol = head.idx['EmailSent'] + 1;
  var values = [];
  for (var i = 0; i < nRows; i++) values.push([true]);
  sh.getRange(2, emailSentCol, nRows, 1).setValues(values);

  var info = { locked: nRows };
  console.log('markAllHistoricalAsSent: 已鎖 ' + nRows + ' 列');
  try {
    SpreadsheetApp.getUi().alert('已將「' + SHEET_LESSONS + '」' + nRows + ' 列設為 EmailSent=TRUE。\n未來新增的列不受影響。');
  } catch (e) { /* 從編輯器直跑不會有 UI */ }
  return info;
}

/**
 * Debug：模擬 Zoom meeting.ended 事件，看會發生什麼。
 * 在 Apps Script 編輯器選此函式 → 執行 → 看執行紀錄輸出。
 * 之後再去 Sheet 開 _log_ 分頁看完整事件。
 */
function debugSimulateMeetingEnded() {
  var students = Students_all_();
  var firstName = students.length ? students[0]['學生姓名'] : '';
  var fakeBody = {
    event: 'meeting.ended',
    payload: {
      object: {
        uuid: 'debug-uuid-' + new Date().getTime(),
        id: '999999999',
        topic: firstName ? ('DEBUG ' + firstName + ' 測試') : 'DEBUG 沒有任何學生',
        start_time: new Date(Date.now() - 3600000).toISOString(),
        end_time: new Date().toISOString(),
        duration: 60
      }
    }
  };
  console.log('students count:', students.length);
  console.log('first student:', JSON.stringify(students[0] || {}));
  console.log('fake topic:', fakeBody.payload.object.topic);
  handleMeetingEnded_(fakeBody);
  console.log('done — 去 Sheet 看 _log_ 與 上課紀錄 分頁');
  return { studentsCount: students.length, fakeTopic: fakeBody.payload.object.topic };
}

function uiShowWebhookInfo() {
  var ui = SpreadsheetApp.getUi();
  var token = getUrlToken_() || '(尚未產生，請先執行 attachToExistingSpreadsheet)';
  var secretSet = !!getZoomSecret_();
  ui.alert([
    '— Webhook 設定（未來接 Zoom 才需要）—',
    'URL token（附加到 /exec?token=）：',
    token,
    '',
    'ZOOM_WEBHOOK_SECRET 是否已設定：' + (secretSet ? '是' : '否'),
    '',
    '部署：clasp deploy → 取得 /exec URL',
    'Zoom Endpoint URL 填：<YOUR_EXEC_URL>?token=' + token,
    '訂閱事件：Meeting → End Meeting、Meeting → Summary completed',
    '',
    '另需在客戶表補上以下欄位才能去重 upsert：',
    '  學生資料：zoom_meeting_id',
    '  上課紀錄：Meeting UUID、Zoom主題'
  ].join('\n'));
}
