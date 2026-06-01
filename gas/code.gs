// ============================================================
// karasuletters — Google Apps Script backend
// ============================================================
// デプロイ設定:
//   実行ユーザー: 自分
//   アクセス: 全員
//
// 必要なトリガー設定 (トリガーページから手動で追加):
//   1. checkAndSendScheduledPosts — 時間ベース、毎分
//   2. onCalendarEventUpdated     — カレンダーから、カレンダーの更新時
//      ※ カレンダー削除で予約が自動取り消しされる
//
// スクリプトプロパティに必要な値:
//   SUBSCRIBERS_SHEET_ID — 読者リストのスプレッドシートID
// ============================================================

const SCHEDULE_KEY        = 'schedules_v2';
const SUBSCRIBERS_SHEET   = 'subscribers';
const DRAFT_FOLDER_NAME   = 'karasuletters_drafts';

// ===== エントリーポイント =====
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (!action) return sendMailToAll(body);

    switch (action) {
      case 'schedule':          return handleSchedule(body);
      case 'getScheduleStatus': return handleGetScheduleStatus();
      case 'cancelSchedule':    return handleCancelSchedule(body.id);
      case 'updateSchedule':    return handleUpdateSchedule(body.id, body.scheduledAt);
      case 'saveDraft':         return handleSaveDraft(body.data);
      case 'listDrafts':        return handleListDrafts();
      case 'loadDraft':         return handleLoadDraft(body.issue);
      case 'deleteDraft':       return handleDeleteDraft(body.issue);
      default:
        return jsonRes({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return jsonRes({ ok: false, error: err.message });
  }
}

function jsonRes(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 即時送信 =====
function sendMailToAll(body) {
  const rows  = getSubscriberRows();
  let sent    = 0;
  for (let i = 1; i < rows.length; i++) {
    const email = rows[i][0];
    if (email && String(email).includes('@')) {
      GmailApp.sendEmail(email, body.subject, '', { htmlBody: body.html });
      sent++;
    }
  }
  return jsonRes({ ok: true, sent });
}

// ===== 予約作成 =====
function handleSchedule(body) {
  const props     = PropertiesService.getScriptProperties();
  const schedules = loadSchedules(props);
  const id        = Utilities.getUuid();

  let calEventId = null;
  try {
    const start = new Date(body.scheduledAt);
    const end   = new Date(start.getTime() + 30 * 60 * 1000);
    const ev    = CalendarApp.getDefaultCalendar().createEvent(
      '[karasuletters] ' + body.subject, start, end,
      { description: 'karasuletters 予約投稿\nID: ' + id }
    );
    calEventId = ev.getId();
  } catch (calErr) {
    Logger.log('Calendar create error: ' + calErr.message);
  }

  schedules[id] = {
    id,
    subject:     body.subject,
    html:        body.html,
    scheduledAt: body.scheduledAt,
    calEventId:  calEventId
  };
  saveSchedules(props, schedules);

  return jsonRes({ ok: true, id, hasCalendar: !!calEventId });
}

// ===== 予約一覧取得 =====
function handleGetScheduleStatus() {
  const schedules = loadSchedules(PropertiesService.getScriptProperties());
  const list = Object.values(schedules)
    .map(s => ({ id: s.id, subject: s.subject, scheduledAt: s.scheduledAt, html: s.html }))
    .sort((a, b) => new Date(a.scheduledAt) - new Date(b.scheduledAt));
  return jsonRes({ ok: true, schedules: list });
}

// ===== 予約取り消し（カレンダーイベントも削除） =====
function handleCancelSchedule(id) {
  const props     = PropertiesService.getScriptProperties();
  const schedules = loadSchedules(props);
  const schedule  = schedules[id];

  if (!schedule) {
    return jsonRes({ ok: false, error: 'Schedule not found: ' + id });
  }

  // Google カレンダーのイベントを削除
  if (schedule.calEventId) {
    deleteCalEvent(schedule.calEventId);
  }

  delete schedules[id];
  saveSchedules(props, schedules);
  return jsonRes({ ok: true });
}

// ===== 予約日時変更（カレンダーイベントも更新） =====
function handleUpdateSchedule(id, scheduledAt) {
  const props     = PropertiesService.getScriptProperties();
  const schedules = loadSchedules(props);
  const schedule  = schedules[id];

  if (!schedule) {
    return jsonRes({ ok: false, error: 'Schedule not found: ' + id });
  }

  if (schedule.calEventId) {
    try {
      const ev = CalendarApp.getEventById(schedule.calEventId);
      if (ev) {
        const start = new Date(scheduledAt);
        const end   = new Date(start.getTime() + 30 * 60 * 1000);
        ev.setTime(start, end);
      }
    } catch (calErr) {
      Logger.log('Calendar update error: ' + calErr.message);
    }
  }

  schedules[id].scheduledAt = scheduledAt;
  saveSchedules(props, schedules);
  return jsonRes({ ok: true });
}

// ===== カレンダートリガー: カレンダー側の削除を予約に反映 =====
// トリガー設定: イベントソース「カレンダーから」→「カレンダーの更新時」
function onCalendarEventUpdated() {
  const props     = PropertiesService.getScriptProperties();
  const schedules = loadSchedules(props);
  let changed     = false;

  for (const id in schedules) {
    const calEventId = schedules[id].calEventId;
    if (!calEventId) continue;

    let exists = false;
    try {
      const ev = CalendarApp.getEventById(calEventId);
      exists = !!ev;
    } catch (err) {
      exists = false;
    }

    if (!exists) {
      Logger.log('Calendar event deleted → cancel schedule: ' + id);
      delete schedules[id];
      changed = true;
    }
  }

  if (changed) saveSchedules(props, schedules);
}

// ===== 時間トリガー: 予約時刻になったらメール送信 =====
// トリガー設定: 時間ベース → 毎分
function checkAndSendScheduledPosts() {
  const props     = PropertiesService.getScriptProperties();
  const schedules = loadSchedules(props);
  const now       = new Date();

  for (const id in schedules) {
    const s = schedules[id];
    if (new Date(s.scheduledAt) > now) continue;

    try {
      const rows = getSubscriberRows();
      for (let i = 1; i < rows.length; i++) {
        const email = rows[i][0];
        if (email && String(email).includes('@')) {
          GmailApp.sendEmail(email, s.subject, '', { htmlBody: s.html });
        }
      }
      if (s.calEventId) deleteCalEvent(s.calEventId);
      delete schedules[id];
      saveSchedules(props, schedules);
      Logger.log('Sent scheduled post: ' + s.subject);
    } catch (err) {
      Logger.log('Failed to send ' + id + ': ' + err.message);
    }
  }
}

// ===== 下書き =====
function handleSaveDraft(data) {
  const folder   = getDraftFolder();
  const filename = 'draft_' + (data.issue || 'unknown') + '.json';
  const content  = JSON.stringify({ ...data, updated: new Date().toISOString() });
  const iter     = folder.getFilesByName(filename);

  if (iter.hasNext()) {
    iter.next().setContent(content);
  } else {
    folder.createFile(filename, content, MimeType.PLAIN_TEXT);
  }
  return jsonRes({ ok: true });
}

function handleListDrafts() {
  const iter   = getDraftFolder().getFiles();
  const drafts = [];
  while (iter.hasNext()) {
    const f = iter.next();
    if (!f.getName().endsWith('.json')) continue;
    try {
      const d = JSON.parse(f.getBlob().getDataAsString());
      drafts.push({ issue: d.issue, updated: d.updated });
    } catch (e) {}
  }
  drafts.sort((a, b) => new Date(b.updated || 0) - new Date(a.updated || 0));
  return jsonRes({ ok: true, drafts });
}

function handleLoadDraft(issue) {
  const iter = getDraftFolder().getFilesByName('draft_' + issue + '.json');
  if (!iter.hasNext()) return jsonRes({ ok: false, error: 'Draft not found' });
  const data = JSON.parse(iter.next().getBlob().getDataAsString());
  return jsonRes({ ok: true, data });
}

function handleDeleteDraft(issue) {
  const iter = getDraftFolder().getFilesByName('draft_' + issue + '.json');
  if (!iter.hasNext()) return jsonRes({ ok: false, error: 'Draft not found' });
  iter.next().setTrashed(true);
  return jsonRes({ ok: true });
}

// ===== ユーティリティ =====
function loadSchedules(props) {
  const raw = props.getProperty(SCHEDULE_KEY);
  return raw ? JSON.parse(raw) : {};
}

function saveSchedules(props, schedules) {
  props.setProperty(SCHEDULE_KEY, JSON.stringify(schedules));
}

function deleteCalEvent(calEventId) {
  try {
    const ev = CalendarApp.getEventById(calEventId);
    if (ev) {
      ev.deleteEvent();
      Logger.log('Deleted calendar event: ' + calEventId);
    }
  } catch (err) {
    Logger.log('deleteCalEvent error (' + calEventId + '): ' + err.message);
  }
}

function getDraftFolder() {
  const iter = DriveApp.getFoldersByName(DRAFT_FOLDER_NAME);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(DRAFT_FOLDER_NAME);
}

function getSubscriberRows() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SUBSCRIBERS_SHEET_ID');
  return SpreadsheetApp.openById(sheetId)
    .getSheetByName(SUBSCRIBERS_SHEET)
    .getDataRange()
    .getValues();
}
