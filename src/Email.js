/**
 * Email.js
 * 掃描「上課紀錄」表，針對 發信=TRUE 且 EmailSent!=TRUE 的列寄送家長信。
 * 由 5 分鐘時間觸發器週期性執行（在 initMasterSpreadsheet 時安裝）。
 *
 * 發件：GmailApp.sendEmail，免費 Gmail 每日額度 100 封、Workspace 1500 封。
 */

function sendPendingEmails() {
  var ss = SpreadsheetApp.openById(getMasterId_());
  var sh = ss.getSheetByName(SHEET_LESSONS);
  if (!sh) return { sent: 0, failed: 0, skipped: 0 };

  var last = sh.getLastRow();
  if (last < 2) return { sent: 0, failed: 0, skipped: 0 };

  var idx = {};
  LESSONS_HEADERS.forEach(function (h, i) { idx[h] = i; });

  var values = sh.getRange(2, 1, last - 1, LESSONS_HEADERS.length).getValues();
  var senderName = getEmailSenderName_();
  var cc = getEmailCc_();

  var sent = 0, failed = 0, skipped = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowIndex = i + 2;

    if (!isTruthyFlag_(row[idx['發信']])) { skipped++; continue; }
    if (isTruthyFlag_(row[idx['EmailSent']])) { skipped++; continue; }

    var parentEmail = String(row[idx['家長email']] || '').trim();
    if (!parentEmail || parentEmail.indexOf('@') < 0) {
      _writeError_(sh, rowIndex, idx, '家長email 欄位為空或格式錯誤');
      failed++;
      continue;
    }

    var payload = {
      studentName: row[idx['學生姓名']] || '',
      date: row[idx['日期']] || '',
      content: row[idx['上課內容']] || '',
      homework: row[idx['作業']] || '',
      durationHr: row[idx['上課時數(hr)']],
      summaryUrl: row[idx['學習狀況/建議(會議摘要)']] || '',
      senderName: senderName
    };
    var subject = (payload.studentName || '學生') + ' 的上課紀錄 - ' + _subjectDate_(payload.date);
    var html = _buildEmailHtml_(payload);
    var plain = _htmlToPlain_(html);
    var options = { htmlBody: html, name: senderName };
    if (cc) options.cc = cc;

    try {
      GmailApp.sendEmail(parentEmail, subject, plain, options);
      sh.getRange(rowIndex, idx['EmailSent'] + 1).setValue(true);
      sh.getRange(rowIndex, idx['SentAt'] + 1).setValue(new Date());
      sh.getRange(rowIndex, idx['ErrorMessage'] + 1).setValue('');
      sent++;
      logEvent_('email_sent', parentEmail + ' / ' + payload.studentName, { date: payload.date });
    } catch (err) {
      _writeError_(sh, rowIndex, idx, String(err && err.message || err));
      failed++;
      logEvent_('email_failed', parentEmail, { err: String(err && err.message || err) });
    }
  }

  return { sent: sent, failed: failed, skipped: skipped };
}

function _writeError_(sh, rowIndex, idx, msg) {
  try {
    sh.getRange(rowIndex, idx['ErrorMessage'] + 1).setValue(msg);
  } catch (e) { /* noop */ }
}

function _subjectDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return formatDateTz_(v, 'yyyy-MM-dd');
  }
  var s = String(v);
  // 盡量抓前 10 或 16 個字元作為日期主體
  return s.length > 16 ? s.substring(0, 16) : s;
}

function _buildEmailHtml_(p) {
  var hrText = (p.durationHr !== '' && p.durationHr !== null && p.durationHr !== undefined)
    ? (p.durationHr + ' 小時') : '';
  var lines = [
    '<p>您好，</p>',
    '<p>以下是 <strong>' + _esc_(p.studentName) + '</strong> 本次的上課紀錄：</p>',
    '<table style="border-collapse:collapse;font-size:14px;line-height:1.6">',
      _kv_('日期', _esc_(p.date)),
      hrText ? _kv_('上課時數', hrText) : '',
      _kv_('上課內容', _esc_(p.content).replace(/\n/g, '<br>')),
      _kv_('作業', _esc_(p.homework).replace(/\n/g, '<br>')),
      p.summaryUrl ? _kv_('會議摘要',
        '<a href="' + _esc_(p.summaryUrl) + '" target="_blank">' + _esc_(p.summaryUrl) + '</a>'
      ) : '',
    '</table>',
    '<p style="margin-top:16px">如有任何問題，歡迎直接回覆此信。</p>',
    '<p>— ' + _esc_(p.senderName) + '</p>'
  ];
  return lines.filter(Boolean).join('\n');
}

function _kv_(k, v) {
  if (v === '' || v === null || v === undefined) return '';
  return '<tr>'
    + '<td style="padding:4px 12px 4px 0;color:#666;vertical-align:top;white-space:nowrap">' + k + '</td>'
    + '<td style="padding:4px 0">' + v + '</td>'
    + '</tr>';
}

function _esc_(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function _htmlToPlain_(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .trim();
}
