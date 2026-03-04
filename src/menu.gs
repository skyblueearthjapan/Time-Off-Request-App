// ====== エントリポイント ======

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'top';
  var settings = getSettings_();
  var appUrl = normalize_(settings['APP_URL'] || '');
  var portalUrl = normalize_(settings['PORTAL_URL'] || '');

  switch (page) {
    case 'admin':
      if (!canAccessAdmin_()) {
        var t = HtmlService.createTemplateFromFile('no_auth');
        t.message = 'このページは管理者・総務のみアクセスできます。';
        t.APP_URL = appUrl;
        return t.evaluate()
          .setTitle('アクセス権限エラー')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
      var adminTpl = HtmlService.createTemplateFromFile('admin');
      adminTpl.APP_URL = appUrl;
      adminTpl.PORTAL_URL = portalUrl;
      return adminTpl.evaluate()
        .setTitle('休暇届 管理')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'preview':
      var prevTpl = HtmlService.createTemplateFromFile('preview');
      prevTpl.REQ_ID = (e && e.parameter && e.parameter.reqId) ? e.parameter.reqId : '';
      prevTpl.APP_URL = appUrl;
      prevTpl.PORTAL_URL = portalUrl;
      return prevTpl.evaluate()
        .setTitle('休暇届 プレビュー')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    case 'sign':
      var signTpl = HtmlService.createTemplateFromFile('sign');
      signTpl.REQ_ID = (e && e.parameter && e.parameter.reqId) ? e.parameter.reqId : '';
      signTpl.APP_URL = appUrl;
      return signTpl.evaluate()
        .setTitle('休暇届 サイン')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    default: // 'top'
      var topTpl = HtmlService.createTemplateFromFile('top');
      topTpl.APP_URL = appUrl;
      topTpl.PORTAL_URL = portalUrl;
      return topTpl.evaluate()
        .setTitle('休暇届 申請TOP')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====== スプレッドシートメニュー ======

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('休暇届アプリ')
    .addItem('マスタ同期（手動実行）', 'syncAllMasters')
    .addItem('サマリー再構築', 'rebuildLeaveSummary_')
    .addSeparator()
    .addItem('有給警告メール送信（手動）', 'checkAndSendWarnMails_')
    .addSeparator()
    .addItem('全トリガー初期セットアップ', 'setupAllTriggers_')
    .addToUi();
}

// ====== トリガー設定 ======

function setupAllTriggers_() {
  setupSyncTrigger_();
  setupWarnMailTrigger_();
  Logger.log('全トリガーをセットアップしました。');
}

function setupSyncTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.some(function(t) { return t.getHandlerFunction() === 'syncAllMasters'; })) return;
  ScriptApp.newTrigger('syncAllMasters')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(0)
    .create();
}

function setupWarnMailTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.some(function(t) { return t.getHandlerFunction() === 'checkAndSendWarnMails_'; })) return;
  ScriptApp.newTrigger('checkAndSendWarnMails_')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(0)
    .create();
}
