// Cutover.gs - VPS 版への切り替え（移転モード）
//
// スクリプトのプロパティ MIGRATED_TO_URL に新しいアプリのアドレスが入っている間だけ「移転モード」。
// プロパティが無ければ、今までと同じ動き（この判定でプロパティを 1 回読む以外は何も変えない）。
//
// 移転モードのとき:
// - doGet: getSettings_() より前で、新しいアドレスへの案内ページを返す（page・reqId を引き継ぐ）。
// - getDb_()（config.gs）の先頭で止める。画面の関数も保守の関数もここを通る（罠 47）。
// - getDb_() を通らない入口（setupAllSheets・トリガーを作る関数など）と、google.script.run で呼べる
//   公開関数（api_* ほか）の先頭でも止める（api_getApproverList のように例外を握る関数があるため）。
// - 時刻のトリガーの関数（MIGRATION_TRIGGER_FUNCTIONS）は何もせずに戻る。
// - シートのメニューは「移転済み（新しいアドレス）」の 1 項目だけ。
// - 内部用の入口（Internal.gs の doPost）は止めない。token を確かめた実行の中だけ
//   __KYUKA_INTERNAL_OK__ を true にし、getDb_() を通す。
//
// 下の運用関数は、スクリプトの持ち主が GAS エディタから手で実行する（罠 48）。
// 公開の名前は「持ち主の確かめ → 末尾 _ の実体」を呼ぶだけ（google.script.run から実体を直接呼べない）。

var MIGRATED_TO_URL_KEY = 'MIGRATED_TO_URL';                       // 移転モードの印＝新しいアプリのアドレス
var MIGRATION_REMOVED_TRIGGERS_KEY = 'MIGRATION_REMOVED_TRIGGERS'; // 消したトリガーの控え（戻すときに使う）
var INTERNAL_ONLY_DEPLOYMENT_ID_KEY = 'INTERNAL_ONLY_DEPLOYMENT_ID'; // 第 2 案（匿名の内部用の版）のときだけ入れる

// 移転先として受けるアドレス（休暇届の VPS 版だけ）。既定値としては使わない（罠 24d）。
var KYUKA_MIGRATION_URL_PATTERN_ = /^https:\/\/leave\.lineworks-local\.net\/?$/;
var KYUKA_MIGRATION_URL_ = 'https://leave.lineworks-local.net';

// 移転ページが引き継ぐ page（7 ページ）と reqId の形
var MIGRATION_PAGES_ = ['top', 'admin', 'preview', 'sign', 'deptAdmin', 'somuAdmin', 'leaveStatus'];
var MIGRATION_REQ_ID_PATTERN_ = /^LV-\d{8}-\d{3,}$/;

// このプロジェクトの時刻のトリガーの関数（コードが作るのは syncAllMasters だけ。ほかは昔のトリガーが
// 残っているおそれがある＝inv-3 T3）
var MIGRATION_TRIGGER_FUNCTIONS = [
  'syncAllMasters',
  'checkAndSendWarnMails_',
  'notifyTodayLeaves_',
  'notifyPendingApproval1_',
  'notifyPendingApproval2_',
];

// 戻すときに作り直すトリガーの時刻（menu.gs の setup*Trigger_ と同じ）
var MIGRATION_TRIGGER_SCHEDULES_ = {
  syncAllMasters: { hour: 6, minute: 0 },
  checkAndSendWarnMails_: { hour: 8, minute: 0 },
  notifyTodayLeaves_: { hour: 7, minute: 0 },
  notifyPendingApproval1_: { hour: 17, minute: 15 },
  notifyPendingApproval2_: { hour: 18, minute: 0 },
};

// 実行の中だけの印。Internal.gs の入口が token を確かめた後でだけ true にする（finally で戻す）。
var __KYUKA_INTERNAL_OK__ = false;

// 実行ごとに 1 回だけプロパティを読む（getDb_ は 1 回の実行で何度も呼ばれるため）。
var _migratedToUrlCache_ = null;

function getMigratedToUrl_() {
  if (_migratedToUrlCache_ !== null) return _migratedToUrlCache_;
  try {
    var v = String(PropertiesService.getScriptProperties().getProperty(MIGRATED_TO_URL_KEY) || '').trim();
    _migratedToUrlCache_ = v;
    return v;
  } catch (e) {
    // プロパティが読めないときは今までと同じ動きにする（キャッシュしない）。
    Logger.log('getMigratedToUrl_: プロパティを読めませんでした: ' + e.message);
    return '';
  }
}

function invalidateMigrationCache_() {
  _migratedToUrlCache_ = null;
}

function isMigrationMode_() {
  return !!getMigratedToUrl_();
}

function isValidMigrationUrl_(url) {
  return KYUKA_MIGRATION_URL_PATTERN_.test(String(url || '').trim());
}

// 'https://leave.lineworks-local.net/' の形（末尾 / 1 つ）
function migrationBaseUrl_(url) {
  return String(url || '').trim().replace(/\/+$/, '') + '/';
}

function migrationMessage_() {
  var url = getMigratedToUrl_();
  return '休暇届は新しいアドレスへ移りました。' +
    (isValidMigrationUrl_(url) ? migrationBaseUrl_(url) + ' を開いてください。' : '新しいアドレスは管理者にお問い合わせください。') +
    '（この画面からは操作できません）';
}

// 画面の関数・保守の関数の先頭で呼ぶ。移転モードなら例外にする。
function assertNotMigrated_() {
  if (isMigrationMode_()) {
    throw new Error(migrationMessage_());
  }
}

// getDb_() から呼ぶ。内部用の入口（token を確かめた実行）だけは通す。
function assertDbOpenAllowed_() {
  if (__KYUKA_INTERNAL_OK__ === true) return;
  assertNotMigrated_();
}

// トリガーの関数の先頭で呼ぶ。移転モードなら true（呼び出し側は何もせずに戻る）。
function skipBecauseMigrated_(fnName) {
  if (!isMigrationMode_()) return false;
  Logger.log(fnName + ': 移転モード（' + getMigratedToUrl_() + '）のため何もしません');
  return true;
}

// ===== doGet の移転ページ =====

function escapeHtmlCutover_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// page は 7 ページの名前だけ、reqId は LV-yyyyMMdd-NNN の形だけを引き継ぐ（ほかは捨てる）。
function migratedTargetUrl_(url, page, reqId) {
  var parts = [];
  var p = String(page == null ? '' : page);
  var r = String(reqId == null ? '' : reqId);
  if (MIGRATION_PAGES_.indexOf(p) >= 0) parts.push('page=' + encodeURIComponent(p));
  if (MIGRATION_REQ_ID_PATTERN_.test(r)) parts.push('reqId=' + encodeURIComponent(r));
  var base = migrationBaseUrl_(url);
  return parts.length ? base + '?' + parts.join('&') : base;
}

// GAS の iframe の外へ出るため <base target="_top"> とボタンの target="_top" を使う。
// 自動の移動（window.top.location）は「動けばよい」扱い（ブラウザが止めることがある。罠 29）。
function renderMigratedPage_(url, page, reqId) {
  var title = '休暇届（移転しました）';
  if (!isValidMigrationUrl_(url)) {
    Logger.log('renderMigratedPage_: MIGRATED_TO_URL の形が正しくないため、固定の文言を出します');
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="utf-8"></head>' +
      '<body style="font-family:sans-serif;margin:40px 16px">' +
      '<p>休暇届は新しいアドレスへ移りました。新しいアドレスは管理者にお問い合わせください。</p>' +
      '</body></html>')
      .setTitle(title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var target = migratedTargetUrl_(url, page, reqId);
  var html = [
    '<!DOCTYPE html><html><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<base target="_top">',
    '<style>body{font-family:sans-serif;margin:40px 16px;line-height:1.7;color:#222}',
    '.go{display:inline-block;padding:18px 36px;font-size:22px;font-weight:bold;color:#fff;background:#1a365d;',
    'border-radius:8px;text-decoration:none}.url{font-size:14px;color:#555;word-break:break-all}</style></head><body>',
    '<h1 style="font-size:20px">休暇届は新しいアドレスへ移りました</h1>',
    '<p>下のボタンを押して、新しいアドレスで開いてください。</p>',
    '<p><a class="go" href="' + escapeHtmlCutover_(target) + '" target="_top">新しいアドレスで開く</a></p>',
    '<p class="url">新しいアドレス: ' + escapeHtmlCutover_(migrationBaseUrl_(url)) +
      '<br>ブックマークをお使いの方は、差し替えをお願いします。</p>',
    '<script>setTimeout(function(){try{window.top.location.href=' +
      JSON.stringify(target).replace(/</g, '\\u003c') + ';}catch(err){}},1000);</script>',
    '</body></html>',
  ].join('');
  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// doGet の先頭で呼ぶ。返り値が null なら今までと同じ動き。
// ① 移転モードなら移転ページ。
// ② 第 2 案（匿名の内部用の版）を使うときだけ: プロパティ INTERNAL_ONLY_DEPLOYMENT_ID があり、
//    この版の URL がその ID を含み、ログインした人が分からない（匿名の GET）なら、画面を出さない。
//    プロパティが無い間（第 1 案・今）は何もしない。
function cutoverDoGetGuard_(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  if (isMigrationMode_()) {
    return renderMigratedPage_(getMigratedToUrl_(), params.page, params.reqId);
  }
  var internalOnlyId = '';
  try {
    internalOnlyId = String(PropertiesService.getScriptProperties().getProperty(INTERNAL_ONLY_DEPLOYMENT_ID_KEY) || '').trim();
  } catch (err) {
    internalOnlyId = '';
  }
  if (!internalOnlyId) return null;
  var serviceUrl = '';
  var email = '';
  try { serviceUrl = String(ScriptApp.getService().getUrl() || ''); } catch (err2) { serviceUrl = ''; }
  try { email = String(Session.getActiveUser().getEmail() || ''); } catch (err3) { email = ''; }
  if (serviceUrl.indexOf(internalOnlyId) >= 0 && !email) {
    return HtmlService.createHtmlOutput('このアドレスでは画面を開けません。').setTitle('休暇届');
  }
  return null;
}

// onOpen の先頭で呼ぶ。移転モードならメニューを 1 項目だけ作って true を返す。
function addMigratedMenuIfMigrated_() {
  if (!isMigrationMode_()) return false;
  SpreadsheetApp.getUi()
    .createMenu('休暇届アプリ')
    .addItem('移転済み（新しいアドレス）', 'showMigratedAddress')
    .addToUi();
  return true;
}

// 移転モードのメニューの 1 項目から呼ぶ。新しいアドレスを出すだけ（何も書かない）。
function showMigratedAddress() {
  var msg = migrationMessage_();
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return msg;
}

// ===== 運用関数を動かしてよい人の確認（罠 48） =====
// Web アプリは持ち主として実行されるが、画面（google.script.run）から呼ぶと「操作している人」は
// 利用者のアカウントになるので、持ち主以外は通らない。エディタから実行すれば成り立つ。
// さらに、持ち主のメールそのものと照合する（スクリプトの写しを別の人が持っても動かないように）。
var CUTOVER_OWNER_EMAIL_ = 'imaizumi@lineworks-local.info';

function assertCutoverOperator_() {
  var active = '';
  var effective = '';
  try { active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e1) { active = ''; }
  try { effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e2) { effective = ''; }
  if (!active || !effective || active !== effective || active !== CUTOVER_OWNER_EMAIL_) {
    throw new Error('この操作は、スクリプトの持ち主が GAS エディタから実行してください。');
  }
  return active;
}

function triggerHandlers_() {
  return ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });
}

// ===== GAS エディタから実行する運用関数 =====

// 例: enterMigrationMode('https://leave.lineworks-local.net')
// アドレスは必ず引数で渡す（既定値を持たない。罠 24d）。休暇届の VPS 版のアドレス以外は受けない。
function enterMigrationMode(url) {
  assertCutoverOperator_();
  return enterMigrationMode_(url);
}

// エディタの実行ボタンは引数を渡せないので、当日はこれを実行する（確定のアドレスで包むだけ）。
function enterMigrationModeForKyuka() {
  assertCutoverOperator_();
  return enterMigrationMode_(KYUKA_MIGRATION_URL_);
}

function enterMigrationMode_(url) {
  var value = String(url == null ? '' : url).trim();
  if (!isValidMigrationUrl_(value)) {
    throw new Error('移転先は https://leave.lineworks-local.net の形で、引数に渡してください: ' + value);
  }
  value = value.replace(/\/+$/, '');
  PropertiesService.getScriptProperties().setProperty(MIGRATED_TO_URL_KEY, value);
  invalidateMigrationCache_();
  Logger.log('移転モードにしました: ' + value);
  return value;
}

// 戻すとき。移転モードを外し、今までと同じ動きに戻す。トリガーは restore* で戻す。
function exitMigrationMode() {
  assertCutoverOperator_();
  return exitMigrationMode_();
}

function exitMigrationMode_() {
  PropertiesService.getScriptProperties().deleteProperty(MIGRATED_TO_URL_KEY);
  invalidateMigrationCache_();
  Logger.log('移転モードを外しました。');
  return true;
}

// VPS_INTERNAL_TOKEN を保存する。引数なしなら 64 字の乱数を作る（VPS の .env の GAS_INTERNAL_TOKEN へは、
// スクリプト プロパティの画面で値を見て手で写す）。引数で渡すときは 32 字以上だけ受ける。
// ログと返り値は長さだけ（値は出さない）。
function setInternalToken(token) {
  assertCutoverOperator_();
  return setInternalToken_(token);
}

function setInternalToken_(token) {
  var value = token == null || token === '' ? '' : String(token).trim();
  if (!value) {
    value = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); // 64 字（16 進）
  }
  if (value.length < 32) throw new Error('VPS_INTERNAL_TOKEN は 32 字以上にしてください（値は表示しません）。');
  PropertiesService.getScriptProperties().setProperty('VPS_INTERNAL_TOKEN', value);
  Logger.log('VPS_INTERNAL_TOKEN を保存しました（' + value.length + ' 字）。値はスクリプト プロパティの画面で確かめてください。');
  return { saved: true, length: value.length };
}

// このスクリプトのトリガー（実行している人の分）を全部控えてから、全部消す。
// 控えはプロパティ MIGRATION_REMOVED_TRIGGERS に足していく（2 回実行しても前の控えを消さない）。
// 見えるのは実行している人のトリガーだけ。ほかの人の分は GAS の「トリガー」画面で確かめる。
function removeAllTriggersForMigration() {
  assertCutoverOperator_();
  return removeAllTriggersForMigration_();
}

// 同じ処理の別名（依頼書の名前）
function removeAllTriggersForCutover() {
  assertCutoverOperator_();
  return removeAllTriggersForMigration_();
}

function removeAllTriggersForMigration_() {
  var props = PropertiesService.getScriptProperties();
  var record = [];
  try { record = JSON.parse(props.getProperty(MIGRATION_REMOVED_TRIGGERS_KEY) || '[]'); } catch (e0) { record = []; }
  var triggers = ScriptApp.getProjectTriggers();
  var listed = triggers.map(function(t) {
    var type = '';
    var id = '';
    try { type = String(t.getEventType()); } catch (e1) { type = ''; }
    try { id = String(t.getUniqueId()); } catch (e2) { id = ''; }
    return { fn: t.getHandlerFunction(), type: type, id: id, removedAt: new Date().toISOString() };
  });
  // 先に控える（消している途中で止まっても、控えは残る）
  props.setProperty(MIGRATION_REMOVED_TRIGGERS_KEY, JSON.stringify(record.concat(listed)));
  var removed = [];
  triggers.forEach(function(t, i) {
    ScriptApp.deleteTrigger(t);
    removed.push(listed[i].fn + (listed[i].id ? ' (' + listed[i].id + ')' : ''));
  });
  var left = triggerHandlers_();
  Logger.log('削除したトリガー: ' + (removed.length ? removed.join(', ') : 'なし') +
    '\n残っているトリガー: ' + (left.length ? left.join(', ') : 'なし'));
  return { removed: removed, left: left };
}

function createTriggerForRollback_(fn) {
  var s = MIGRATION_TRIGGER_SCHEDULES_[fn];
  if (!s) return false;
  if (triggerHandlers_().indexOf(fn) >= 0) return true; // 同じ関数のトリガーがあれば作らない
  ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(s.hour).nearMinute(s.minute).create();
  return true;
}

function assertNotMigratedForRollback_() {
  invalidateMigrationCache_();
  if (isMigrationMode_()) throw new Error('移転モードのままです。先に exitMigrationMode を実行してください。');
}

// 戻すとき（exitMigrationMode の後）。今のコード（setupAllTriggers）が作る syncAllMasters の 1 本だけ作り直す。
function restoreTriggersAfterRollback() {
  assertCutoverOperator_();
  return restoreTriggersAfterRollback_();
}

function restoreTriggersAfterRollback_() {
  assertNotMigratedForRollback_();
  var before = triggerHandlers_();
  createTriggerForRollback_('syncAllMasters');
  var after = triggerHandlers_();
  Logger.log('トリガー（前）: ' + (before.join(', ') || 'なし') + '\nトリガー（後）: ' + (after.join(', ') || 'なし'));
  return after;
}

// 戻すとき（exitMigrationMode の後）。removeAllTriggersForMigration で控えたトリガーを、関数ごとに 1 本ずつ
// 今のコードと同じ時刻で作り直す（同じ関数のトリガーがあれば作らない）。時刻を知らない関数は作らずに
// 一覧に出す（GAS の「トリガー」画面で手で作る）。作り直した後、控えを消す。
function restoreAllTriggersAfterRollback() {
  assertCutoverOperator_();
  return restoreAllTriggersAfterRollback_();
}

function restoreAllTriggersAfterRollback_() {
  assertNotMigratedForRollback_();
  var props = PropertiesService.getScriptProperties();
  var record = [];
  try { record = JSON.parse(props.getProperty(MIGRATION_REMOVED_TRIGGERS_KEY) || '[]'); } catch (e0) { record = []; }
  var before = triggerHandlers_();
  var seen = {};
  var unknown = [];
  record.forEach(function(r) {
    var fn = String(r && r.fn || '');
    if (!fn || seen[fn]) return;
    seen[fn] = true;
    if (!createTriggerForRollback_(fn)) unknown.push(fn + (r.type ? ' (' + r.type + ')' : ''));
  });
  if (!unknown.length) props.deleteProperty(MIGRATION_REMOVED_TRIGGERS_KEY);
  var after = triggerHandlers_();
  Logger.log('トリガー（前）: ' + (before.join(', ') || 'なし') + '\nトリガー（後）: ' + (after.join(', ') || 'なし') +
    '\n作り直せなかった（手で作る）: ' + (unknown.length ? unknown.join(', ') : 'なし'));
  return { triggers: after, unknown: unknown };
}

// 今の状態をログに出す。何も変えない。getDb_() の止めを通らない（DB を開かない）。
function showMigrationStatus() {
  assertCutoverOperator_();
  return showMigrationStatus_();
}

// 同じ処理の別名（依頼書の名前）
function cutoverStatus() {
  assertCutoverOperator_();
  return showMigrationStatus_();
}

function showMigrationStatus_() {
  var props = PropertiesService.getScriptProperties();
  invalidateMigrationCache_();
  var removed = [];
  try { removed = JSON.parse(props.getProperty(MIGRATION_REMOVED_TRIGGERS_KEY) || '[]'); } catch (e) { removed = []; }
  var status = {
    apiVersion: INTERNAL_API_VERSION,
    migratedTo: getMigratedToUrl_(),
    internalTokenSet: !!props.getProperty('VPS_INTERNAL_TOKEN'),
    internalOnlyDeploymentId: String(props.getProperty(INTERNAL_ONLY_DEPLOYMENT_ID_KEY) || ''),
    triggers: triggerHandlers_(),
    removedTriggers: removed.map(function(r) { return r.fn; }),
  };
  Logger.log(
    '移転モード: ' + (status.migratedTo || '（なし。今までと同じ動き）') +
    '\n内部用の入口の版: ' + status.apiVersion +
    '\n内部用の token: ' + (status.internalTokenSet ? 'あり' : 'なし') +
    '\n匿名の内部用の版: ' + (status.internalOnlyDeploymentId || '（なし）') +
    '\nトリガー: ' + (status.triggers.length ? status.triggers.join(', ') : 'なし') +
    '\n控えてあるトリガー: ' + (status.removedTriggers.length ? status.removedTriggers.join(', ') : 'なし'));
  return status;
}
