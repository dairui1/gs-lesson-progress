/**
 * Webhook.js
 * Zoom webhook 入口。
 *
 * 主要訂閱事件：
 *   - meeting.summary_completed（AI Companion 摘要）：upsert 並覆寫摘要相關欄位
 *   - meeting.ended：fallback；若無摘要時先把日期 / 時長 / Zoom 主題寫入
 *
 * 驗證：
 *   - endpoint.url_validation：以 ZOOM_WEBHOOK_SECRET 回應 encryptedToken
 *   - 後續事件：因 Apps Script doPost 無法讀 HTTP header，改用 URL query token
 *     （Zoom Endpoint URL 需為 <EXEC_URL>?token=<ZOOM_URL_TOKEN>）
 */

function doGet() {
  return jsonOut_({ ok: true, name: 'gs-lesson-progress webhook' });
}

function doPost(e) {
  var raw = (e && e.postData && e.postData.contents) || '';
  var body = safeJsonParse_(raw) || {};

  // 排查期：先把原始 body 整段記下來，不管已知/未知事件
  logEvent_('raw_body', body.event || '(no event)', body);

  try {
    // 1. URL 驗證（CRC）
    if (body.event === 'endpoint.url_validation') {
      var plainToken = (body.payload && body.payload.plainToken) || '';
      var secret = getZoomSecret_();
      if (!secret) {
        logEvent_('url_validation_failed', 'ZOOM_WEBHOOK_SECRET 尚未設定', { plainToken: plainToken });
        return jsonOut_({ ok: false, reason: 'secret not set' });
      }
      var encryptedToken = hmacHex_(plainToken, secret);
      logEvent_('url_validation', 'CRC ok', { plainToken: plainToken });
      return jsonOut_({ plainToken: plainToken, encryptedToken: encryptedToken });
    }

    // 2. URL token 驗證
    var expected = getUrlToken_();
    var got = (e && e.parameter && e.parameter.token) || '';
    if (!expected || got !== expected) {
      logEvent_('unauthorized', 'bad url token', { event: body.event });
      return jsonOut_({ ok: false, reason: 'unauthorized' });
    }

    // 3. 事件分派
    if (body.event === 'meeting.summary_completed') {
      handleSummaryCompleted_(body);
      return jsonOut_({ ok: true });
    }
    if (body.event === 'meeting.ended') {
      handleMeetingEnded_(body);
      return jsonOut_({ ok: true });
    }

    logEvent_('unhandled_event', body.event || '(none)', body);
    return jsonOut_({ ok: true, ignored: true });
  } catch (err) {
    logEvent_('webhook_error', err && err.message, { stack: err && err.stack, raw: raw });
    return jsonOut_({ ok: false, error: String(err && err.message) });
  }
}

function handleMeetingEnded_(body) {
  var obj = (body.payload && body.payload.object) || {};
  var uuid = obj.uuid || '';
  var meetingId = obj.id || '';
  var topic = obj.topic || '';
  var startTime = obj.start_time || '';
  var endTime = obj.end_time || '';
  var durationMinutes = obj.duration;

  // 去重
  if (uuid && _seen_(uuid + ':ended')) {
    logEvent_('duplicate', 'meeting.ended duplicate', { uuid: uuid });
    return;
  }

  var student = _resolveStudent_(meetingId, topic);
  if (!student) {
    logEvent_('unknown_meeting', 'ended: meeting_id=' + meetingId + ' topic=' + topic, body);
    return;
  }

  var durationHr = computeDurationHr_(startTime, endTime, durationMinutes);

  var res = upsertLesson({
    source: 'ended',
    student: student,
    uuid: uuid,
    topic: topic,
    startTime: startTime || new Date().toISOString(),
    durationHr: durationHr
  });

  logEvent_('lesson_ended',
    student['學生姓名'] + ' @ ' + topic,
    { meeting_id: meetingId, uuid: uuid, action: res.action }
  );
}

function handleSummaryCompleted_(body) {
  var obj = (body.payload && body.payload.object) || {};
  // Zoom 文件欄位（以 meeting_ 前綴居多）；部分環境使用舊命名
  var uuid = obj.meeting_uuid || obj.uuid || '';
  var meetingId = obj.meeting_id || obj.id || '';
  var topic = obj.meeting_topic || obj.topic || '';
  var startTime = obj.meeting_start_time || obj.start_time || '';
  var endTime = obj.meeting_end_time || obj.end_time || '';
  var summaryDocUrl = obj.summary_doc_url || obj.summary_url || '';
  var summaryOverview = obj.summary_overview || '';
  var summaryDetails = obj.summary_details || [];
  var nextSteps = obj.next_steps || [];

  if (uuid && _seen_(uuid + ':summary')) {
    logEvent_('duplicate', 'meeting.summary_completed duplicate', { uuid: uuid });
    return;
  }

  var student = _resolveStudent_(meetingId, topic);
  if (!student) {
    logEvent_('unknown_meeting', 'summary: meeting_id=' + meetingId + ' topic=' + topic, body);
    return;
  }

  var durationHr = computeDurationHr_(startTime, endTime);
  var homeworkFromAI = _formatNextSteps_(nextSteps);
  var summaryFullText = _formatSummaryDetails_(summaryDetails, summaryOverview);

  var res = upsertLesson({
    source: 'summary',
    student: student,
    uuid: uuid,
    topic: topic,
    startTime: startTime || new Date().toISOString(),
    durationHr: durationHr,
    lessonContent: summaryOverview,
    summaryFullText: summaryFullText,
    summaryDocUrl: summaryDocUrl,
    homeworkFromAI: homeworkFromAI
  });

  logEvent_('lesson_summary',
    student['學生姓名'] + ' @ ' + topic,
    { meeting_id: meetingId, uuid: uuid, action: res.action, hasNextSteps: nextSteps.length }
  );
}

function _resolveStudent_(meetingId, topic) {
  return findByZoomMeetingId(meetingId) || findByMeetingTopic_(topic);
}

function _seen_(key) {
  var cache = CacheService.getScriptCache();
  var k = 'zoom_' + Utilities.base64EncodeWebSafe(key);
  if (cache.get(k)) return true;
  cache.put(k, '1', 600);
  return false;
}

function _formatSummaryDetails_(details, overview) {
  var parts = [];
  if (details && details.length) {
    details.forEach(function (d) {
      if (!d) return;
      var label = d.label || d.section || d.title || '';
      var text = d.summary || d.content || d.text || '';
      if (typeof d === 'string') text = d;
      if (!text) return;
      parts.push((label ? '【' + label + '】\n' : '') + text);
    });
  }
  if (!parts.length && overview) return overview;
  return parts.join('\n\n');
}

function _formatNextSteps_(nextSteps) {
  if (!nextSteps || !nextSteps.length) return '';
  return nextSteps.map(function (s) {
    if (!s) return '';
    if (typeof s === 'string') return '- ' + s;
    return '- ' + (s.content || s.text || JSON.stringify(s));
  }).filter(Boolean).join('\n');
}
