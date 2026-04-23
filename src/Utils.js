/**
 * Utils.js
 * HMAC / 日期 / JSON / 事件日誌等輔助。
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
 * 事件日誌：寫到 Apps Script Stackdriver（clasp logs 可看），不落 Sheet。
 */
function logEvent_(eventName, summary, payload) {
  try {
    var body = (eventName || '') + ' | ' + (summary || '');
    var p = typeof payload === 'string' ? payload : JSON.stringify(payload || {});
    console.log(body + ' | ' + p);
  } catch (e) {
    console.error('logEvent_ failed:', e && e.message);
  }
}

function uuidLockKey_(uuid) {
  return 'zoom_uuid_' + Utilities.base64EncodeWebSafe(String(uuid || ''));
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
