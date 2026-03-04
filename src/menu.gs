// ====== エントリポイント（休暇届申請システム） ======

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
    .addItem('初期セットアップ（シート自動生成）', 'setupAllSheets_')
    .addSeparator()
    .addItem('マスタ同期（手動実行）', 'syncAllMasters')
    .addItem('サマリー再構築', 'rebuildLeaveSummary_')
    .addSeparator()
    .addItem('有給警告メール送信（手動）', 'checkAndSendWarnMails_')
    .addSeparator()
    .addItem('全トリガー初期セットアップ', 'setupAllTriggers_')
    .addToUi();
}

// ====== 初期セットアップ（全シート自動生成） ======

function setupAllSheets_() {
  var ss = SpreadsheetApp.getActive();
  var created = [];
  var skipped = [];

  // ---------- 1. M_SYSTEM_SETTING ----------
  var settingSh = ensureSheet_(ss, 'M_SYSTEM_SETTING');
  if (settingSh.getLastRow() < 2) {
    var settingRows = [
      ['KEY', 'VALUE', 'MEMO'],
      ['APP_URL', '', 'WebアプリURL（デプロイ後に設定）'],
      ['PORTAL_URL', '', '社内ポータルURL（任意）'],
      ['ADMIN_EMAILS', '', '管理者メール（カンマ区切り）'],
      ['SOMU_EMAILS', '', '総務メール（カンマ区切り）'],
      ['ALLOW_ADMIN_PAGE_EMAILS', '', '管理ページ認可リスト（管理者＋総務のメール、カンマ区切り）'],
      ['MAIL_TO', '', '通知先メール（管理者＋総務、カンマ区切り）'],
      ['PDF_FOLDER_ID', '', 'PDF保存先DriveフォルダID'],
      ['PDF_TEMPLATE_SSID', '', 'PDFテンプレートスプレッドシートID'],
      ['MASTER_SOURCE_SSID', '', '同期元マスタスプレッドシートID'],
      ['MASTER_SOURCE_DEPT_SHEET', '', '同期元の部署シート名'],
      ['MASTER_SOURCE_WORKER_SHEET', '', '同期元の作業員シート名'],
    ];
    settingSh.getRange(1, 1, settingRows.length, 3).setValues(settingRows);
    formatHeaderRow_(settingSh);
    settingSh.setColumnWidth(1, 250);
    settingSh.setColumnWidth(2, 350);
    settingSh.setColumnWidth(3, 300);
    created.push('M_SYSTEM_SETTING');
  } else {
    skipped.push('M_SYSTEM_SETTING（既存）');
  }

  // ---------- 2. M_DEPT ----------
  var deptSh = ensureSheet_(ss, 'M_DEPT');
  if (deptSh.getLastRow() < 1) {
    var deptHeader = [['部署ID', '部署名', 'IS_ACTIVE', 'SORT']];
    deptSh.getRange(1, 1, 1, 4).setValues(deptHeader);
    // サンプルデータ
    var deptSample = [
      ['DEPT01', 'サンプル部署A', 1, 1],
      ['DEPT02', 'サンプル部署B', 1, 2],
    ];
    deptSh.getRange(2, 1, deptSample.length, 4).setValues(deptSample);
    formatHeaderRow_(deptSh);
    created.push('M_DEPT');
  } else {
    skipped.push('M_DEPT（既存）');
  }

  // ---------- 3. M_WORKER ----------
  var workerSh = ensureSheet_(ss, 'M_WORKER');
  if (workerSh.getLastRow() < 1) {
    var workerHeader = [['作業員ID', '氏名', '部署ID', 'Googleアカウント', 'IS_ACTIVE', 'SORT']];
    workerSh.getRange(1, 1, 1, 6).setValues(workerHeader);
    // サンプル（現在のユーザーで1件）
    var email = Session.getActiveUser().getEmail() || '';
    var workerSample = [
      ['W001', 'サンプル太郎', 'DEPT01', email, 1, 1],
    ];
    workerSh.getRange(2, 1, workerSample.length, 6).setValues(workerSample);
    formatHeaderRow_(workerSh);
    workerSh.setColumnWidth(4, 250);
    created.push('M_WORKER');
  } else {
    skipped.push('M_WORKER（既存）');
  }

  // ---------- 4. M_LOOKUP ----------
  var lookupSh = ensureSheet_(ss, 'M_LOOKUP');
  if (lookupSh.getLastRow() < 1) {
    var lookupRows = [
      ['カテゴリ', '値', 'SORT'],
      ['特別休暇理由', '結婚休暇', 1],
      ['特別休暇理由', '忌引休暇', 2],
      ['特別休暇理由', '出産休暇', 3],
      ['特別休暇理由', '生理休暇', 4],
      ['特別休暇理由', '裁判員休暇', 5],
      ['特別休暇理由', 'その他', 99],
    ];
    lookupSh.getRange(1, 1, lookupRows.length, 3).setValues(lookupRows);
    formatHeaderRow_(lookupSh);
    lookupSh.setColumnWidth(1, 180);
    lookupSh.setColumnWidth(2, 200);
    created.push('M_LOOKUP');
  } else {
    skipped.push('M_LOOKUP（既存）');
  }

  // ---------- 5. T_LEAVE_REQUEST ----------
  var reqSh = ensureSheet_(ss, 'T_LEAVE_REQUEST');
  if (reqSh.getLastRow() < 1) {
    var reqHeader = [[
      'REQ_ID', '年度', '部署ID', '部署名', '作業員ID', '作業員名',
      '休暇区分', '半日区分', '休暇日', '休暇種類',
      '振替元出勤日', '特別理由', '有給詳細', '追加詳細',
      '申請日時', '承認日時', '承認状態', 'サイン画像URL', 'PDF_URL', 'PDF_FILE_ID'
    ]];
    reqSh.getRange(1, 1, 1, reqHeader[0].length).setValues(reqHeader);
    formatHeaderRow_(reqSh);
    reqSh.setFrozenRows(1);
    created.push('T_LEAVE_REQUEST');
  } else {
    skipped.push('T_LEAVE_REQUEST（既存）');
  }

  // ---------- 6. V_LEAVE_SUMMARY ----------
  var sumSh = ensureSheet_(ss, 'V_LEAVE_SUMMARY');
  if (sumSh.getLastRow() < 1) {
    var sumHeader = [[
      '年度', '部署ID', '部署名', '作業員ID', '作業員名',
      '有給回数', '有給半日回数', '振替回数', '特別回数', '合計回数',
      '最終有給日', '警告レベル'
    ]];
    sumSh.getRange(1, 1, 1, sumHeader[0].length).setValues(sumHeader);
    formatHeaderRow_(sumSh);
    sumSh.setFrozenRows(1);
    created.push('V_LEAVE_SUMMARY');
  } else {
    skipped.push('V_LEAVE_SUMMARY（既存）');
  }

  // ---------- 7. BatchLogs ----------
  var logSh = ensureSheet_(ss, 'BatchLogs');
  if (logSh.getLastRow() < 1) {
    var logHeader = [['実行日時', 'バッチ名', '対象日', '成功', 'スキップ', '失敗', 'エラー詳細']];
    logSh.getRange(1, 1, 1, logHeader[0].length).setValues(logHeader);
    formatHeaderRow_(logSh);
    logSh.setFrozenRows(1);
    created.push('BatchLogs');
  } else {
    skipped.push('BatchLogs（既存）');
  }

  SpreadsheetApp.flush();

  // 結果表示
  var msg = '=== シート自動生成 完了 ===\n\n';
  if (created.length) msg += '作成: ' + created.join(', ') + '\n';
  if (skipped.length) msg += 'スキップ: ' + skipped.join(', ') + '\n';
  msg += '\n次のステップ:\n';
  msg += '1. M_SYSTEM_SETTING の VALUE列を設定してください\n';
  msg += '2. M_DEPT / M_WORKER のサンプルデータを実データに置き換えてください\n';
  msg += '3. Webアプリとしてデプロイしてください';

  SpreadsheetApp.getUi().alert(msg);
  Logger.log(msg);
}

/** シートが無ければ作成して返す */
function ensureSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  return sh;
}

/** ヘッダ行を太字＋背景色で書式設定 */
function formatHeaderRow_(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  var range = sh.getRange(1, 1, 1, lastCol);
  range.setFontWeight('bold');
  range.setBackground('#4a86c8');
  range.setFontColor('#ffffff');
  sh.setFrozenRows(1);
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
