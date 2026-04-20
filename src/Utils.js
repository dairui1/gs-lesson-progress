/**
 * Utils.js
 * HMAC / 日期 / JSON / 去重等輔助。
 */

function hmacHex_(message, secret) {
  var raw = Utilities.computeHmacSha256Signature(message, secret);
  return raw.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function formatDateTz_(date, pattern) {
  return Utilities.formatDate(date, TZ, pattern || 'yyyy-MM-dd HH:mm');
}

function safeJsonParse_(s) {
  try {
    return JSON.parse(s);
  } catch (e) {
    return null;
  }
}

/**
 * 寫入 Log 分頁；避免本身拋錯中斷主流程。
 */
function logEvent_(eventName, summary, payload) {
  try {
    var ss = SpreadsheetApp.openById(getMasterId_());
    var log = ss.getSheetByName(SHEET_LOG);
    if (!log) {
      log = ss.insertSheet(SHEET_LOG);
      log.appendRow(LOG_HEADERS);
      log.setFrozenRows(1);
    }
    log.appendRow([
      new Date(),
      eventName || '',
      summary || '',
      typeof payload === 'string' ? payload : JSON.stringify(payload || {})
    ]);
  } catch (e) {
    console.error('logEvent_ failed:', e && e.message);
  }
}

/**
 * Zoom meeting uuid 可能含 "=" 與 "/"，轉為 cache-safe key。
 */
function uuidLockKey_(uuid) {
  return 'zoom_uuid_' + Utilities.base64EncodeWebSafe(String(uuid || ''));
}

/**
 * 將任意值轉為 ContentService JSON 回應。
 * 備註：Apps Script Web App 無法自訂 HTTP status code，
 * 非預期狀況一律回 200 + { ok:false, reason } 讓 Zoom 不重送。
 */
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
