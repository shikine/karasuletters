// karasuletters — Google Apps Script (完全版)
// デプロイ設定: ウェブアプリ → 全員がアクセス可能
//
// ★ 必要なトリガー設定 (GASエディタ > トリガーページから手動で追加):
//   1. checkAndSendScheduled    — 時間主導型、毎10分
//   2. onCalendarEventUpdated   — カレンダーから、カレンダーの更新時
//      ↑ これを設定するとカレンダー側の削除が予約取り消しに反映される

var NOTIFY_EMAIL  = 'watanabeshikine@gmail.com';
var UNSUB_EMAIL   = 'watanabeshikine@gmail.com';
var SHEET_NAME    = 'フォームの回答 1';
var UNSUB_LABEL   = 'karasuletters-unsubscribed';
var DRAFT_FOLDER  = 'karasuletters_drafts';
var CALENDAR_ID   = 'urj4s4v32702jrsemq4aope5d0@group.calendar.google.com';

// ===== ウェブアプリエントリーポイント =====
function doPost(e) {
  var data   = JSON.parse(e.postData.contents);
  var action = data.action;

  // actionなし + subject/html あり → 即時送信
  if (!action && data.subject && data.html) return handleSend(data.subject, data.html);

  if (action === 'register')          return handleRegister(data.name, data.email);
  if (action === 'checkEmail')        return respond({ duplicate: isDuplicate(data.email) });
  if (action === 'schedule')          return handleSchedule(data.subject, data.html, data.scheduledAt);
  if (action === 'getScheduleStatus') return handleGetSchedules();
  if (action === 'cancelSchedule')    return handleCancelSchedule(data.id);
  if (action === 'updateSchedule')    return handleUpdateSchedule(data.id, data.subject, data.scheduledAt);
  if (action === 'saveDraft')         return handleSaveDraft(data.data);
  if (action === 'listDrafts')        return handleListDrafts();
  if (action === 'loadDraft')         return handleLoadDraft(data.issue);
  if (action === 'deleteDraft')       return handleDeleteDraft(data.issue);

  return respond({ error: 'unknown action' });
}

function doGet(e) {
  return respond({ status: 'ok' });
}

// ===== 即時送信 =====
function handleSend(subject, html) {
  var subscribers = getSubscribers();
  subscribers.forEach(function(row) {
    var name  = row[1] || '';
    var email = row[2];
    if (!email) return;
    GmailApp.sendEmail(email, subject, '', {
      htmlBody: html,
      name: 'karasuletters'
    });
  });
  return respond({ ok: true, sent: subscribers.length });
}

// ===== スケジュール管理 =====
function handleSchedule(subject, html, scheduledAt) {
  var props = PropertiesService.getScriptProperties();
  var id    = 'sched_' + new Date().getTime();
  props.setProperty(id, JSON.stringify({ subject: subject, html: html, scheduledAt: scheduledAt }));
  ensureScheduleTrigger();

  // Google カレンダーにイベントを登録
  try {
    var sendDate = new Date(scheduledAt);
    var endDate  = new Date(sendDate.getTime() + 30 * 60 * 1000);
    var cal = CalendarApp.getCalendarById(CALENDAR_ID);
    if (cal) {
      var event = cal.createEvent('[karasuletters]予約配信: ' + subject, sendDate, endDate,
        { description: '件名: ' + subject + '\n配信ID: ' + id });
      if (event) {
        var stored = JSON.parse(props.getProperty(id));
        stored.calEventId = event.getId();
        props.setProperty(id, JSON.stringify(stored));
      }
    }
  } catch(e) { Logger.log('handleSchedule calendar error: ' + e.message); }

  return respond({ ok: true, id: id });
}

// カレンダーイベントの存在もチェックし、削除済みなら予約も自動取り消し
function handleGetSchedules() {
  var props = PropertiesService.getScriptProperties();
  var all   = props.getProperties();
  var cal   = CalendarApp.getCalendarById(CALENDAR_ID);
  var list  = [];

  Object.keys(all).forEach(function(key) {
    if (key.indexOf('sched_') !== 0) return;
    try {
      var v = JSON.parse(all[key]);
      v.id  = key;

      // カレンダーイベントが削除されていたら予約も自動取り消し
      if (v.calEventId && cal) {
        var ev = null;
        try { ev = cal.getEventById(v.calEventId); } catch(e) {}
        if (!ev) {
          props.deleteProperty(key);
          return; // 一覧に含めない
        }
      }

      list.push(v);
    } catch(e) {}
  });

  return respond({ schedules: list });
}

// ===== 予約取り消し — カレンダーイベントも削除 =====
// 修正: CalendarApp.getEventById() → cal.getEventById() で特定カレンダーを指定
function handleCancelSchedule(id) {
  var props = PropertiesService.getScriptProperties();
  try {
    var stored = JSON.parse(props.getProperty(id) || '{}');
    if (stored.calEventId) {
      deleteCalendarEvent(stored.calEventId);
    }
  } catch(e) { Logger.log('handleCancelSchedule error: ' + e.message); }
  props.deleteProperty(id);
  return respond({ ok: true });
}

// ===== 日時変更 — カレンダーイベントも更新 =====
// 修正: CalendarApp.getEventById() → cal.getEventById() で特定カレンダーを指定
function handleUpdateSchedule(id, subject, scheduledAt) {
  var props  = PropertiesService.getScriptProperties();
  var stored = JSON.parse(props.getProperty(id) || '{}');
  if (subject)     stored.subject     = subject;
  if (scheduledAt) stored.scheduledAt = scheduledAt;
  props.setProperty(id, JSON.stringify(stored));
  try {
    if (stored.calEventId) {
      var cal = CalendarApp.getCalendarById(CALENDAR_ID);
      if (cal) {
        var ev = cal.getEventById(stored.calEventId);
        if (ev) {
          var sendDate = new Date(stored.scheduledAt);
          var endDate  = new Date(sendDate.getTime() + 30 * 60 * 1000);
          ev.setTime(sendDate, endDate);
          ev.setTitle('[karasuletters]予約配信: ' + stored.subject);
        }
      }
    }
  } catch(e) { Logger.log('handleUpdateSchedule calendar error: ' + e.message); }
  return respond({ ok: true });
}

// ===== 時間トリガー: 予約時刻にメール送信 + カレンダー削除 =====
function checkAndSendScheduled() {
  var now   = new Date().getTime();
  var props = PropertiesService.getScriptProperties();
  var all   = props.getProperties();
  Object.keys(all).forEach(function(key) {
    if (key.indexOf('sched_') !== 0) return;
    try {
      var v = JSON.parse(all[key]);
      if (new Date(v.scheduledAt).getTime() <= now) {
        handleSend(v.subject, v.html);
        // 送信後にカレンダーイベントも削除
        if (v.calEventId) deleteCalendarEvent(v.calEventId);
        props.deleteProperty(key);
      }
    } catch(e) { Logger.log('checkAndSendScheduled error: ' + e.message); }
  });
}

// ===== カレンダートリガー: カレンダー側の削除を予約に反映 =====
// このトリガーを有効にするには:
// GASエディタ > トリガー > トリガーを追加
//   関数: onCalendarEventUpdated
//   イベントのソース: カレンダーから
//   イベントの種類: カレンダーの更新時
function onCalendarEventUpdated() {
  var props   = PropertiesService.getScriptProperties();
  var all     = props.getProperties();
  var changed = false;

  Object.keys(all).forEach(function(key) {
    if (key.indexOf('sched_') !== 0) return;
    try {
      var v = JSON.parse(all[key]);
      if (!v.calEventId) return;

      // カレンダーイベントの存在確認
      var cal     = CalendarApp.getCalendarById(CALENDAR_ID);
      var exists  = false;
      if (cal) {
        try {
          var ev = cal.getEventById(v.calEventId);
          exists = !!ev;
        } catch(e) { exists = false; }
      }

      if (!exists) {
        Logger.log('Calendar event deleted → cancel schedule: ' + key);
        props.deleteProperty(key);
        changed = true;
      }
    } catch(e) { Logger.log('onCalendarEventUpdated error for ' + key + ': ' + e.message); }
  });

  if (changed) Logger.log('onCalendarEventUpdated: cancelled orphaned schedules');
}

// ===== カレンダーイベント削除ヘルパー =====
function deleteCalendarEvent(calEventId) {
  try {
    var cal = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!cal) return;
    var ev = cal.getEventById(calEventId);
    if (ev) {
      ev.deleteEvent();
      Logger.log('Deleted calendar event: ' + calEventId);
    }
  } catch(e) { Logger.log('deleteCalendarEvent error (' + calEventId + '): ' + e.message); }
}

// ===== トリガー自動生成 =====
function ensureScheduleTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAndSendScheduled') return;
  }
  ScriptApp.newTrigger('checkAndSendScheduled').timeBased().everyMinutes(10).create();
}

// ===== トリガー初期設定（手動実行） =====
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'processUnsubscribeEmails' || fn === 'checkAndSendScheduled') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processUnsubscribeEmails').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('checkAndSendScheduled').timeBased().everyMinutes(10).create();
}

// ===== 下書き CRUD (Google Drive) =====
function getDraftFolder() {
  var folders = DriveApp.getFoldersByName(DRAFT_FOLDER);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRAFT_FOLDER);
}

function handleSaveDraft(data) {
  var folder  = getDraftFolder();
  var issue   = data.issue;
  data.updated = new Date().toISOString();
  var files   = folder.getFilesByName(issue + '.json');
  var content = JSON.stringify(data);
  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    folder.createFile(issue + '.json', content, MimeType.PLAIN_TEXT);
  }
  return respond({ ok: true });
}

function handleListDrafts() {
  var folder = getDraftFolder();
  var files  = folder.getFiles();
  var list   = [];
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().match(/\.json$/)) {
      try { list.push(JSON.parse(f.getBlob().getDataAsString())); } catch(e) {}
    }
  }
  return respond({ drafts: list });
}

function handleLoadDraft(issue) {
  var folder = getDraftFolder();
  var files  = folder.getFilesByName(issue + '.json');
  if (!files.hasNext()) return respond({ error: 'not found' });
  try {
    return respond({ data: JSON.parse(files.next().getBlob().getDataAsString()) });
  } catch(e) {
    return respond({ error: 'parse error' });
  }
}

function handleDeleteDraft(issue) {
  var folder = getDraftFolder();
  var files  = folder.getFilesByName(issue + '.json');
  if (files.hasNext()) files.next().setTrashed(true);
  return respond({ ok: true });
}

// ===== 購読者管理 =====
function handleRegister(name, email) {
  if (isDuplicate(email)) {
    return respond({ duplicate: true });
  }
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  sheet.appendRow([new Date(), name, email]);

  // 運営者への通知メール
  MailApp.sendEmail({
    to:      NOTIFY_EMAIL,
    subject: '【karasuletters】新しい読者登録がありました',
    body:    '新しい読者登録がありました。\n\nお名前: ' + name + '\nメール: ' + email + '\n\n登録日時: ' + new Date().toLocaleString('ja-JP')
  });

  // 登録者へのサンクスメール
  MailApp.sendEmail({
    to:      email,
    subject: '【karasuletters】登録ありがとうございます',
    body:    name + ' さま\n\nkarasuletters へようこそ。\n\n読者登録ありがとうございます。\n毎回一通、手紙のように丁寧に、そんなことをお届けしたいと思っています。\n\nどうかお楽しみに。\n\nwith love,\nShikine Watanabe\n—\n配信停止をご希望の方は ' + UNSUB_EMAIL + ' に空メールをお送りください。'
  });

  return respond({ ok: true });
}

// ===== 配信停止メール処理（受信トレイで自動処理） =====
function processUnsubscribeEmails() {
  var label = GmailApp.getUserLabelByName(UNSUB_LABEL);
  if (!label) label = GmailApp.createLabel(UNSUB_LABEL);

  var threads = GmailApp.search('is:unread in:inbox');

  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    var lastMsg  = messages[messages.length - 1];
    var body     = lastMsg.getPlainBody().trim();
    var from     = lastMsg.getFrom();

    var isUnsubRequest = body === '' ||
      /(配信停止|unsubscribe|配信解除|登録解除)$/i.test(body);

    if (!isUnsubRequest) return;

    var emailMatch  = from.match(/<(.+?)>/) || [null, from];
    var senderEmail = emailMatch[1].trim().toLowerCase();

    var removed = removeSubscriber(senderEmail);

    if (removed) {
      MailApp.sendEmail({
        to:      senderEmail,
        subject: '【karasuletters】配信停止が完了しました',
        body:    'karasuletters の配信停止が完了しました。\n\n配信リストのメール: ' + senderEmail + '\n\nまたいつかお会いしましょう。 https://shikine.github.io/ から再購読いただけます。\n\nShikine Watanabe'
      });

      MailApp.sendEmail({
        to:      NOTIFY_EMAIL,
        subject: '【karasuletters】配信停止がありました',
        body:    '配信停止のリクエストを処理しました。\n\nメール: ' + senderEmail + '\n\n実施日時: ' + new Date().toLocaleString('ja-JP')
      });
    }

    thread.addLabel(label);
    thread.markRead();
    thread.moveToArchive();
  });
}

// ===== ユーティリティ =====
function getSubscribers() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  return data.slice(1); // ヘッダー行を除く
}

function isDuplicate(email) {
  var lower = email.toLowerCase();
  var rows  = getSubscribers();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][2]).toLowerCase() === lower) return true;
  }
  return false;
}

function removeSubscriber(email) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  var lower = email.toLowerCase();

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][2]).toLowerCase() === lower) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
