/**
 * Config.js
 * 常量、Script Properties 存取集中於此。
 * 資料架構為單一「上課紀錄」表（AppSheet 資料源）+ 單一「學生資料」表。
 */

// Script Properties keys
var PROP_MASTER_ID = 'MASTER_SPREADSHEET_ID';
var PROP_ZOOM_SECRET = 'ZOOM_WEBHOOK_SECRET';           // Zoom Secret Token，用於 CRC
var PROP_URL_TOKEN = 'ZOOM_URL_TOKEN';                  // 自訂共用字串，放在 Web App URL ?token=
var PROP_EMAIL_SENDER_NAME = 'EMAIL_SENDER_NAME';       // 寄件人顯示名稱，預設「老師」
var PROP_EMAIL_CC = 'EMAIL_CC';                         // 選填：自動 CC 的 email

var TZ = 'Asia/Taipei';

// 分頁名稱
var SHEET_STUDENTS = '學生資料';
var SHEET_LESSONS = '上課紀錄';
var SHEET_LOG = 'Log';

// 「學生資料」表頭
var STUDENTS_HEADERS = [
  'student_id',
  '學生姓名',
  '家長email',
  'zoom_meeting_id',
  'notes'
];

// 「上課紀錄」表頭（AppSheet 顯示欄位 + 內部追蹤欄位）
var LESSONS_HEADERS = [
  'lesson_id',
  'student_id',
  '日期',
  '學生姓名',
  '家長email',
  '上課內容',
  '作業',
  '上課時數(hr)',
  '學習狀況/建議(會議摘要)',
  'Zoom主題',
  'Meeting UUID',
  '發信',          // 老師在 AppSheet 勾選（TRUE）觸發寄信
  'EmailSent',    // 系統寫回 TRUE 表示已寄出
  'SentAt',       // 寄出時間
  'ErrorMessage'  // 寄信失敗訊息
];

var LOG_HEADERS = ['時間', '事件', '摘要', '原始 payload'];

function getMasterId_() {
  var id = PropertiesService.getScriptProperties().getProperty(PROP_MASTER_ID);
  if (!id) {
    throw new Error('尚未初始化：請先於 Apps Script 編輯器執行 initMasterSpreadsheet()');
  }
  return id;
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
