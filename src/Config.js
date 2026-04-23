/**
 * Config.js
 * 常量、Script Properties 存取集中於此。
 * 資料架構：客戶既有 Sheet，以「學生姓名」為事實主鍵，「上課紀錄」單表為 AppSheet 資料源。
 */

// Script Properties keys
var PROP_MASTER_ID = 'MASTER_SPREADSHEET_ID';
var PROP_ZOOM_SECRET = 'ZOOM_WEBHOOK_SECRET';           // Zoom Secret Token，用於 CRC（未來啟用 Webhook 才需要）
var PROP_URL_TOKEN = 'ZOOM_URL_TOKEN';                  // 自訂共用字串，放在 Web App URL ?token=（同上）
var PROP_EMAIL_SENDER_NAME = 'EMAIL_SENDER_NAME';       // 寄件人顯示名稱，預設「老師」
var PROP_EMAIL_CC = 'EMAIL_CC';                         // 選填：自動 CC 的 email

var TZ = 'Asia/Taipei';

// 分頁名稱
var SHEET_STUDENTS = '學生資料';
var SHEET_LESSONS = '上課紀錄';

// 「學生資料」必需表頭（客戶既有 schema）
var STUDENTS_HEADERS = [
  '學生姓名',
  '家長Email',
  '總堂數',
  '已上課堂數',
  '剩餘堂數'
];

// 「上課紀錄」必需表頭（客戶既有 7 列 + 我們新增的 4 控制列）
var LESSONS_HEADERS = [
  '日期',
  '學生姓名',
  '家長Email',
  '上課內容',
  '作業',
  '上課時數(hr)',
  '學習狀況/建議(會議摘要)',
  '發信',
  'EmailSent',
  'SentAt',
  'ErrorMessage'
];

// 未來若要啟用 Zoom Webhook，需在客戶表補上這些欄位：
//   學生資料：zoom_meeting_id
//   上課紀錄：Meeting UUID、Zoom主題
// 程式會動態讀表頭；欄位存在即自動啟用，不存在即靜默略過。

function getMasterId_() {
  var id = PropertiesService.getScriptProperties().getProperty(PROP_MASTER_ID);
  if (id) return id;
  try {
    var active = SpreadsheetApp.getActive();
    if (active) return active.getId();
  } catch (e) { /* noop */ }
  throw new Error('尚未初始化：請先執行 attachToExistingSpreadsheet()');
}

function getZoomSecret_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_ZOOM_SECRET) || '';
}

function getUrlToken_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_URL_TOKEN) || '';
}

function getEmailSenderName_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_EMAIL_SENDER_NAME) || '老師';
}

function getEmailCc_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_EMAIL_CC) || '';
}

function setProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/**
 * 讀取指定分頁第 1 行表頭；回傳 { headers, idx, lastCol }。
 * 所有讀寫 Sheet 的邏輯都動態從此取得欄位索引，避免寫死欄位順序。
 */
function readHeaders_(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return { headers: [], idx: {}, lastCol: 0 };
  var row = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var headers = row.map(function (v) { return String(v || '').trim(); });
  var idx = {};
  headers.forEach(function (h, i) { if (h) idx[h] = i; });
  return { headers: headers, idx: idx, lastCol: lastCol };
}

/**
 * 將各種可能代表「TRUE」的值判定為 true：
 * 布林 true、數字 1、AppSheet Y/N 字串、中文「是 / 勾選 / ✓」等。
 */
function isTruthyFlag_(v) {
  if (v === true) return true;
  if (typeof v === 'number') return v === 1;
  if (typeof v === 'string') {
    return /^(true|yes|y|1|✓|✔|是|勾選)$/i.test(v.trim());
  }
  return false;
}
