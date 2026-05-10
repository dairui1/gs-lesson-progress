/**
 * Email.js
 * 掃描「上課紀錄」表，針對 發信=TRUE 且 EmailSent!=TRUE 的列寄送家長信。
 * 由 5 分鐘時間觸發器週期性執行（attachToExistingSpreadsheet 會安裝）。
 *
 * 家長 Email 取值順序：該行「家長Email」→ 若空則以「學生姓名」去「學生資料」反查。
 * 發件：GmailApp.sendEmail，免費 Gmail 每日額度 100 封、Workspace 1500 封。
 */

function sendPendingEmails() {
  var ss = SpreadsheetApp.openById(getMasterId_());
  var sh = ss.getSheetByName(SHEET_LESSONS);
  if (!sh) {
    console.warn('sendPendingEmails: 找不到分頁「' + SHEET_LESSONS + '」');
    return { sent: 0, failed: 0, skipped: 0 };
  }

  var head = readHeaders_(sh);
  var required = ['發信', 'EmailSent', '學生姓名', '日期'];
  var missing = required.filter(function (h) { return head.idx[h] === undefined; });
  if (missing.length) {
    console.warn('sendPendingEmails: 表頭缺少欄位 ' + missing.join(', '));
    return { sent: 0, failed: 0, skipped: 0 };
  }

  var last = sh.getLastRow();
  if (last < 2) return { sent: 0, failed: 0, skipped: 0 };

  var values = sh.getRange(2, 1, last - 1, head.lastCol).getValues();
  var senderName = getEmailSenderName_();
  var cc = getEmailCc_();
  var sent = 0, failed = 0, skipped = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowIndex = i + 2;

    if (!isTruthyFlag_(row[head.idx['發信']])) { skipped++; continue; }
    if (isTruthyFlag_(row[head.idx['EmailSent']])) { skipped++; continue; }

    var studentName = String(row[head.idx['學生姓名']] || '').trim();
    var parentEmail = head.idx['家長Email'] !== undefined
      ? String(row[head.idx['家長Email']] || '').trim() : '';

    if (!parentEmail) {
      parentEmail = lookupParentEmailByName_(studentName);
    }

    if (!parentEmail || parentEmail.indexOf('@') < 0) {
      _writeError_(sh, head, rowIndex, '家長Email 欄位為空或格式錯誤（也找不到學生「' + studentName + '」的聯絡信箱）');
      failed++;
      continue;
    }

    var payload = {
      studentName: studentName,
      date: row[head.idx['日期']] || '',
      content: head.idx['上課內容'] !== undefined ? (row[head.idx['上課內容']] || '') : '',
      homework: head.idx['作業'] !== undefined ? (row[head.idx['作業']] || '') : '',
      durationHr: head.idx['上課時數(hr)'] !== undefined ? row[head.idx['上課時數(hr)']] : '',
      summaryText: head.idx['學習狀況/建議(會議摘要)'] !== undefined
        ? (row[head.idx['學習狀況/建議(會議摘要)']] || '') : '',
      summaryDocUrl: head.idx['會議摘要連結'] !== undefined
        ? (row[head.idx['會議摘要連結']] || '') : '',
      senderName: senderName
    };
    var subject = (payload.studentName || '學生') + ' 的上課紀錄 - ' + _subjectDate_(payload.date);
    var html = _buildEmailHtml_(payload);
    var plain = _htmlToPlain_(html);
    var options = { htmlBody: html, name: senderName };
    if (cc) options.cc = cc;

    try {
      GmailApp.sendEmail(parentEmail, subject, plain, options);
      if (head.idx['EmailSent'] !== undefined) sh.getRange(rowIndex, head.idx['EmailSent'] + 1).setValue(true);
      if (head.idx['SentAt'] !== undefined) sh.getRange(rowIndex, head.idx['SentAt'] + 1).setValue(new Date());
      if (head.idx['ErrorMessage'] !== undefined) sh.getRange(rowIndex, head.idx['ErrorMessage'] + 1).setValue('');
      sent++;
      logEvent_('email_sent', parentEmail + ' / ' + payload.studentName, { date: payload.date });
    } catch (err) {
      _writeError_(sh, head, rowIndex, String(err && err.message || err));
      failed++;
      logEvent_('email_failed', parentEmail, { err: String(err && err.message || err) });
    }
  }

  return { sent: sent, failed: failed, skipped: skipped };
}

function _writeError_(sh, head, rowIndex, msg) {
  try {
    if (head.idx['ErrorMessage'] !== undefined) {
      sh.getRange(rowIndex, head.idx['ErrorMessage'] + 1).setValue(msg);
    }
  } catch (e) { /* noop */ }
}

function _subjectDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return formatDateTz_(v, 'yyyy-MM-dd');
  }
  var s = String(v);
  return s.length > 16 ? s.substring(0, 16) : s;
}

function _buildEmailHtml_(p) {
  var hrText = (p.durationHr !== '' && p.durationHr !== null && p.durationHr !== undefined)
    ? (p.durationHr + ' 小時') : '';
  var summaryHtml = _mdToHtml_(p.summaryText);
  var docLink = p.summaryDocUrl
    ? '<p style="margin:18px 0"><a href="' + _esc_(p.summaryDocUrl) + '" target="_blank" style="color:#1a73e8;text-decoration:none">查看完整會議摘要文件 →</a></p>'
    : '';
  var lines = [
    '<div style="font-family:-apple-system,BlinkMacSystemFont,Helvetica,Arial,sans-serif;font-size:14px;color:#222;line-height:1.6;max-width:680px">',
    '<p>您好，</p>',
    '<p>以下是 <strong>' + _esc_(p.studentName) + '</strong> 本次的上課紀錄：</p>',
    '<table style="border-collapse:collapse;font-size:14px;line-height:1.6">',
      _kv_('日期', _esc_(p.date)),
      hrText ? _kv_('上課時數', hrText) : '',
      _kv_('上課內容', _esc_(p.content).replace(/\n/g, '<br>')),
      _kv_('作業', _esc_(p.homework).replace(/\n/g, '<br>')),
    '</table>',
    summaryHtml ? '<h2 style="margin:28px 0 12px;font-size:20px;font-weight:bold;border-top:1px solid #e8e8e8;padding-top:20px;color:#111">Meeting summary</h2>' : '',
    summaryHtml,
    docLink,
    '<p style="margin-top:20px">如有任何問題，歡迎直接回覆此信。</p>',
    '<p>— ' + _esc_(p.senderName) + '</p>',
    '</div>'
  ];
  return lines.filter(Boolean).join('\n');
}

/**
 * 把 Zoom summary_content 那種 markdown 轉成郵件用的 inline-styled HTML。
 * 只處理 ## / ### 標題、- 列表、空行分段；夠覆蓋 Zoom AI summary 的格式。
 */
function _mdToHtml_(md) {
  if (!md) return '';
  var lines = String(md).split(/\r?\n/);
  var out = [];
  var inList = false;
  function closeList() { if (inList) { out.push('</ul>'); inList = false; } }

  for (var i = 0; i < lines.length; i++) {
    var trimmed = lines[i].replace(/^\s+|\s+$/g, '');
    if (!trimmed) { closeList(); continue; }

    var m;
    if ((m = trimmed.match(/^###\s+(.*)$/))) {
      closeList();
      out.push('<h4 style="margin:16px 0 6px;font-size:15px;font-weight:bold;color:#222">' + _esc_(m[1]) + '</h4>');
    } else if ((m = trimmed.match(/^##\s+(.*)$/))) {
      closeList();
      out.push('<h3 style="margin:24px 0 10px;font-size:17px;font-weight:bold;color:#111">' + _esc_(m[1]) + '</h3>');
    } else if ((m = trimmed.match(/^[-*]\s+(.*)$/))) {
      if (!inList) { out.push('<ul style="margin:6px 0;padding-left:24px">'); inList = true; }
      out.push('<li style="margin:4px 0">' + _esc_(m[1]) + '</li>');
    } else {
      closeList();
      out.push('<p style="margin:8px 0">' + _esc_(trimmed) + '</p>');
    }
  }
  closeList();
  return out.join('\n');
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
