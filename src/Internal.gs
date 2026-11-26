// Internal.gs - VPS の worker から呼ぶ内部用の入口（実装仕様 7 章・決定 B・K-5）
//
// 本文: {"internal":true,"token":"…","action":"internal.<名前>","jobId":"<uuid>","reqId":"LV-…","expect":{…},"args":{…}}
// 返事: いつも HTTP 200 の JSON。成功 {ok:true, action, jobId, reqId, sign?, mail?, pdf?, trash?, ledger?, durationMs, replayed?}
//       失敗 {ok:false, error, message, retryable, 済んだ段の結果}
//
// 認証: 本文の token をスクリプトのプロパティ VPS_INTERNAL_TOKEN（32 字以上）と長さ一定の比較で比べる。
//       未設定・短い・違うときは何もせず unauthorized。
// ロック: スクリプトのロックを tryLock(20000) で取り、最後まで持つ（今の generateLeavePdf_・updateLedger_ は
//       ロックを取らず、呼び出し側が持つ前提）。取れなければ busy。
// 冪等: jobId ごとに、済んだ段（sign・mail・pdf・trash・ledger）の結果をスクリプトのプロパティ
//       ij:<jobId>:<段> に段を終えた直後に書く。同じ jobId の再試行は記録のある段を飛ばす（メール 2 通・
//       サイン画像 2 つ・PDF 2 つを防ぐ）。全段が済んだ返事は CacheService に 6 時間置き、同じ jobId には
//       処理せず replayed:true で返す。失敗した段は記録しない（再試行でやり直す）。
// 移転モード: token を確かめた実行の中だけ __KYUKA_INTERNAL_OK__（Cutover.gs）を true にし、getDb_() を通す。
// 中身: 今の関数（saveSignImage_・sendApprovalNotification_・generateLeavePdf_・updateLedger_・
//       removeLedgerRow_・getUnderCountWorkers_・buildWarnMailBody_）を、今と同じ順序・条件で呼ぶ。
//       中で api_* を呼ばない（api_getLeaveRequestById は今の関数が中で使うだけ）。rebuildLeaveSummary_ は呼ばない。
// 台帳: 書く・消す前に LEDGER_SSID を開けるか・年度タブの見出しを確かめ、後で行を読み直して確かめる。
//       入口から呼ぶ間だけ getLedgerSs_ を「開くだけ」に差し替え、新しい台帳を作らない（画面からの今の動きは変えない）。

var INTERNAL_API_VERSION = 1;
var VPS_INTERNAL_TOKEN_KEY_ = 'VPS_INTERNAL_TOKEN';
var VPS_LOCK_WAIT_MS_ = 20000;
var VPS_STAGE_PREFIX_ = 'ij:';
var VPS_STAGE_TTL_MS_ = 7 * 24 * 60 * 60 * 1000;
var VPS_REPLAY_CACHE_SECONDS_ = 6 * 60 * 60;
var VPS_TRASH_PARENT_DEPTH_ = 3; // PDF は PDF_FOLDER_ID → 日付フォルダ → PDF の 2 段
var VPS_JOB_ID_PATTERN_ = /^[A-Za-z0-9._:-]{1,100}$/;
var VPS_REQ_ID_PATTERN_ = /^[A-Za-z0-9_-]{1,64}$/;

// ledger.gs の getLedgerFySheet_ が書く 20 列（1 字ずつ同じであること）
var VPS_LEDGER_HEADERS_ = [
  '管理番号', '作業員ID', '部署', '氏名',
  '休暇日', '曜日', '休暇区分', '半日区分', '勤務時間帯',
  '休暇種類', '理由・詳細', '振替元出勤日',
  '申請日', '1次承認日', '1次承認者', '2次承認日', '2次承認者',
  '承認状態', '有給累計取得回数', 'PDF'
];

// ===== 入口 =====

function doPost(e) {
  var body = null;
  try {
    body = JSON.parse(String((e && e.postData && e.postData.contents) || ''));
  } catch (err) {
    body = null;
  }
  var out;
  try {
    out = vpsDispatch_(body);
  } catch (err2) {
    out = vpsFail_('internal', vpsErrText_(err2), true);
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

var VPS_INTERNAL_ROUTES_ = {
  'internal.ping': { fn: function(job) { return vpsPing_(); }, readOnly: true },
  'internal.afterApprove1': { fn: function(job) { return vpsAfterApprove1_(job); }, needsReqId: true },
  'internal.afterApprove2': { fn: function(job) { return vpsAfterApprove2_(job); }, needsReqId: true },
  'internal.afterUnapprove': { fn: function(job) { return vpsAfterUnapprove_(job); }, needsReqId: true },
  'internal.sendWarnMail': { fn: function(job) { return vpsSendWarnMailAction_(job); } },
};

function vpsDispatch_(body) {
  var started = Date.now();
  if (!body || typeof body !== 'object' || body.internal !== true) {
    return vpsFail_('bad_request', 'not an internal request', false);
  }

  // 認証（何かを読む・書く前に）
  var expected = '';
  try {
    expected = String(PropertiesService.getScriptProperties().getProperty(VPS_INTERNAL_TOKEN_KEY_) || '');
  } catch (e0) {
    expected = '';
  }
  if (expected.length < 32 || !vpsTimingSafeEqual_(body.token, expected)) {
    return vpsFail_('unauthorized', 'unauthorized', false);
  }

  var action = String(body.action || '');
  var route = VPS_INTERNAL_ROUTES_[action];
  if (!route) return vpsFail_('bad_request', 'unknown action: ' + action, false, { action: action });

  __KYUKA_INTERNAL_OK__ = true;
  var lock = null;
  var locked = false;
  try {
    if (route.readOnly) {
      var r0 = route.fn(null);
      r0.action = action;
      r0.durationMs = Date.now() - started;
      return r0;
    }

    var jobId = body.jobId == null ? '' : String(body.jobId);
    if (!VPS_JOB_ID_PATTERN_.test(jobId)) {
      return vpsFail_('bad_request', 'jobId is required (1-100 chars of [A-Za-z0-9._:-])', false, { action: action });
    }
    var reqId = body.reqId == null ? '' : normalize_(body.reqId);
    if (route.needsReqId && !VPS_REQ_ID_PATTERN_.test(reqId)) {
      return vpsFail_('bad_request', 'reqId is required', false, { action: action, jobId: jobId });
    }
    var args = (body.args && typeof body.args === 'object') ? body.args : {};
    var expect = (body.expect && typeof body.expect === 'object') ? body.expect : {};

    var replay = vpsReplay_(jobId, action);
    if (replay) return replay;

    lock = LockService.getScriptLock();
    if (!lock.tryLock(VPS_LOCK_WAIT_MS_)) {
      return vpsFail_('busy', 'script lock not available', true, { action: action, jobId: jobId, reqId: reqId });
    }
    locked = true;
    replay = vpsReplay_(jobId, action); // 待っている間に同じ jobId が済んだとき
    if (replay) return replay;

    var job = { jobId: jobId, action: action, reqId: reqId, args: args, expect: expect };
    var result;
    try {
      result = route.fn(job);
    } catch (err) {
      result = err && err.vpsCode
        ? vpsFail_(err.vpsCode, vpsErrText_(err), false)
        : vpsFail_('internal', vpsErrText_(err), true);
    }
    result.action = action;
    result.jobId = jobId;
    if (route.needsReqId) result.reqId = reqId;
    result.durationMs = Date.now() - started;
    if (result.ok === true) vpsCachePut_(jobId, action, result);
    return result;
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (e1) { /* ignore */ }
    }
    __KYUKA_INTERNAL_OK__ = false;
  }
}

// ===== 小さな道具 =====

// 長さ一定の比較（長さの違いも同じ手間で比べる）
function vpsTimingSafeEqual_(a, b) {
  a = String(a == null ? '' : a);
  b = String(b == null ? '' : b);
  var len = Math.max(a.length, b.length);
  var diff = a.length ^ b.length;
  for (var i = 0; i < len; i++) {
    var ca = i < a.length ? a.charCodeAt(i) : 0;
    var cb = i < b.length ? b.charCodeAt(i) : 0;
    diff |= (ca ^ cb);
  }
  return diff === 0;
}

function vpsFail_(code, message, retryable, extra) {
  var out = { ok: false, error: code, message: String(message || code), retryable: !!retryable };
  if (extra) for (var k in extra) if (extra[k] !== undefined) out[k] = extra[k];
  return out;
}

// 例外にして入口で返す（retryable:false）
function vpsError_(code, message) {
  var e = new Error(message);
  e.vpsCode = code;
  return e;
}

function vpsErrText_(e) {
  return String((e && e.message) || e);
}

// 'yyyy-MM-dd' → その日の JST 0 時の Date。形が違えば null
function vpsParseYmd_(s) {
  var t = String(s == null ? '' : s).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) return null;
  var d = new Date(t + 'T00:00:00+09:00');
  return isNaN(d.getTime()) ? null : d;
}

// シートのセル → api_getLeaveRequestById と同じ「休暇日の文字」（Date なら yyyy-MM-dd）
function vpsLeaveDateText_(v) {
  return v instanceof Date ? fmtDate_(v) : String(v || '');
}

// Drive の URL からファイル ID を取り出す（/d/<id>/ か ?id=<id>）。無ければ ''
function vpsDriveIdFromUrl_(url) {
  var m = /\/d\/([A-Za-z0-9_-]+)/.exec(String(url || '')) || /[?&]id=([A-Za-z0-9_-]+)/.exec(String(url || ''));
  return m ? m[1] : '';
}

// ===== 段の記録（冪等） =====

function vpsStageKey_(jobId, stage) {
  return VPS_STAGE_PREFIX_ + jobId + ':' + stage;
}

function vpsStageRead_(job, stage) {
  var raw = null;
  try {
    raw = PropertiesService.getScriptProperties().getProperty(vpsStageKey_(job.jobId, stage));
  } catch (e) {
    Logger.log('vpsStageRead_: ' + e.message);
    return null;
  }
  if (!raw) return null;
  var rec = null;
  try { rec = JSON.parse(raw); } catch (e2) { return null; }
  if (!rec || typeof rec !== 'object') return null;
  if (rec.a && rec.a !== job.action) {
    throw vpsError_('bad_request', 'jobId was already used for ' + rec.a);
  }
  return rec.r || null;
}

function vpsStageWrite_(job, stage, result) {
  try {
    PropertiesService.getScriptProperties().setProperty(vpsStageKey_(job.jobId, stage),
      JSON.stringify({ t: Date.now(), a: job.action, r: result }));
  } catch (e) {
    // 記録できなくても続ける（罠 36）。再試行でこの段がもう一度動くことがある
    Logger.log('vpsStageWrite_: ' + e.message);
  }
}

// 記録があればそれを返し、無ければ fn() を動かして、失敗でなければ記録する。
// fn() の返り値に noRecord:true があれば記録しない（後の再試行で結果が変わりうるとき）。
function vpsStage_(job, stage, fn) {
  var prev = vpsStageRead_(job, stage);
  if (prev) return prev;
  var r = fn();
  var noRecord = r.noRecord === true;
  delete r.noRecord;
  if (r.status !== 'error' && !noRecord) vpsStageWrite_(job, stage, r);
  return r;
}

// 7 日より古い段の記録を消す（ping から呼ぶ）
function vpsPruneStageRecords_() {
  var n = 0;
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    var now = Date.now();
    Object.keys(all).forEach(function(k) {
      if (k.indexOf(VPS_STAGE_PREFIX_) !== 0) return;
      var t = 0;
      try { t = Number(JSON.parse(all[k]).t) || 0; } catch (e) { t = 0; }
      if (now - t > VPS_STAGE_TTL_MS_) {
        props.deleteProperty(k);
        n++;
      }
    });
  } catch (e2) {
    Logger.log('vpsPruneStageRecords_: ' + e2.message);
  }
  return n;
}

// 全段が済んだ返事（CacheService。例外は握る＝罠 36）
function vpsReplay_(jobId, action) {
  var raw = null;
  try {
    raw = CacheService.getScriptCache().get(VPS_STAGE_PREFIX_ + jobId);
  } catch (e) {
    Logger.log('vpsReplay_: ' + e.message);
    return null;
  }
  if (!raw) return null;
  var rec = null;
  try { rec = JSON.parse(raw); } catch (e2) { return null; }
  if (!rec || !rec.r) return null;
  if (rec.a !== action) {
    return vpsFail_('bad_request', 'jobId was already used for ' + rec.a, false, { action: action, jobId: jobId });
  }
  rec.r.replayed = true;
  return rec.r;
}

function vpsCachePut_(jobId, action, result) {
  try {
    CacheService.getScriptCache().put(VPS_STAGE_PREFIX_ + jobId, JSON.stringify({ a: action, r: result }), VPS_REPLAY_CACHE_SECONDS_);
  } catch (e) {
    Logger.log('vpsCachePut_: ' + e.message);
  }
}

// ===== シートの行（T_LEAVE_REQUEST。見出しの文字で列を引く。6 秒待ちをしない） =====

function vpsFindRow_(reqId) {
  var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var sh = info.sh;
  var idx = info.idx;
  var col = idx['REQ_ID'];
  if (col === undefined) throw vpsError_('bad_request', 'T_LEAVE_REQUEST に REQ_ID の列がありません');
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var readCols = Math.max(info.header.length, sh.getLastColumn());
  var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();
  for (var i = 0; i < values.length; i++) {
    if (normalize_(values[i][col]) === reqId) {
      return { sh: sh, idx: idx, rowNo: i + 2, row: values[i] };
    }
  }
  return null;
}

// 列が無ければ書かずに false（本番に PDF_FILE_ID 列は無い）
function vpsSetCell_(rc, name, value) {
  var c = rc.idx[name];
  if (c === undefined) return false;
  rc.sh.getRange(rc.rowNo, c + 1).setValue(value);
  return true;
}

function vpsGetCell_(rc, name) {
  var c = rc.idx[name];
  if (c === undefined) return '';
  return normalize_(rc.sh.getRange(rc.rowNo, c + 1).getValue());
}

var VPS_EXPECT_COLUMNS_ = {
  status: { col: '承認状態', kind: 'text' },
  approvedAt: { col: '承認日時', kind: 'datetime' },
  approvedBy1Email: { col: 'APPROVED_BY1_EMAIL', kind: 'text' },
  approvedAt2: { col: 'APPROVED_AT2', kind: 'datetime' },
  approvedBy2Email: { col: 'APPROVED_BY2_EMAIL', kind: 'text' },
  leaveDate: { col: '休暇日', kind: 'date' },
};

function vpsCellText_(v, kind) {
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return fmtDate_(v, kind === 'date' ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm:ss');
  }
  return normalize_(v);
}

// 合わない項目の名前の一覧（空なら合っている）。知らない項目は合わない扱い
function vpsExpectMismatch_(rc, expect) {
  var bad = [];
  Object.keys(expect || {}).forEach(function(k) {
    if (expect[k] === undefined) return;
    var def = VPS_EXPECT_COLUMNS_[k];
    if (!def || rc.idx[def.col] === undefined) { bad.push(k); return; }
    var want = expect[k] == null ? '' : normalize_(expect[k]);
    var got = vpsCellText_(rc.row[rc.idx[def.col]], def.kind);
    if (got !== want) bad.push(k);
  });
  return bad;
}

// 行を探して expect と照合する。合わなければ失敗の返事、合えば行
function vpsLoadRow_(job) {
  var rc = vpsFindRow_(job.reqId);
  if (!rc) return { fail: vpsFail_('row_not_found', 'request not found in sheet: ' + job.reqId, true) };
  var bad = vpsExpectMismatch_(rc, job.expect);
  if (bad.length) {
    return { fail: vpsFail_('row_not_synced', 'sheet row does not match expect: ' + bad.join(','), true, { mismatch: bad }) };
  }
  return { rc: rc };
}

// ===== 台帳の守り =====

// 入口から呼ぶ間だけ getLedgerSs_ を「開くだけ（作らない）」に差し替える
function vpsOpenLedgerNoCreate_() {
  var id = normalize_(getSettings_()['LEDGER_SSID']);
  if (!id) return null;
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    console.warn('LEDGER_SSID のSSが開けません（内部用の入口は作り直しません）: ' + e.message);
    return null;
  }
}

function vpsWithLedgerNoCreate_(fn) {
  var original = getLedgerSs_;
  getLedgerSs_ = vpsOpenLedgerNoCreate_;
  try {
    return fn();
  } finally {
    getLedgerSs_ = original;
  }
}

function vpsLedgerError_(code, message, retryable, sheetName) {
  var r = { status: 'error', error: code, message: message, retryable: retryable };
  if (sheetName) r.sheetName = sheetName;
  return r;
}

// 台帳を開けるか、年度タブ（あれば）の 1 行目が 20 列と同じかを確かめる。良ければ null
function vpsLedgerPrecheck_(fy) {
  var id = normalize_(getSettings_()['LEDGER_SSID']);
  if (!id) return vpsLedgerError_('ledger_unavailable', 'LEDGER_SSID is not set', true);
  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    return vpsLedgerError_('ledger_unavailable', vpsErrText_(e), true);
  }
  var name = fy + '年度';
  var sh = ss.getSheetByName(name);
  if (!sh) return null; // タブが無ければ今の関数が作る
  var head = sh.getRange(1, 1, 1, VPS_LEDGER_HEADERS_.length).getValues()[0];
  for (var c = 0; c < VPS_LEDGER_HEADERS_.length; c++) {
    if (String(head[c] == null ? '' : head[c]) !== VPS_LEDGER_HEADERS_[c]) {
      return vpsLedgerError_('ledger_header_mismatch', name + ' の ' + (c + 1) + ' 列目の見出しが違います', false, name);
    }
  }
  return null;
}

// 年度タブの A 列に reqId があるか。タブが無ければ false、台帳が開けなければ null
function vpsLedgerHasRow_(sheetName, reqId) {
  try {
    var ss = SpreadsheetApp.openById(normalize_(getSettings_()['LEDGER_SSID']));
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return false;
    var last = sh.getLastRow();
    if (last < 2) return false;
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (normalize_(ids[i][0]) === reqId) return true;
    return false;
  } catch (e) {
    Logger.log('vpsLedgerHasRow_: ' + e.message);
    return null;
  }
}

// updateLedger_ と同じ年度（休暇日の文字 → new Date → computeFiscalYear_）
function vpsLedgerFy_(leaveDateText) {
  return computeFiscalYear_(new Date(leaveDateText));
}

function vpsLedgerUpsert_(reqId, leaveDateText) {
  var fy = vpsLedgerFy_(leaveDateText);
  var sheetName = fy + '年度';
  var pre = vpsLedgerPrecheck_(fy);
  if (pre) return pre;
  vpsWithLedgerNoCreate_(function() { updateLedger_(reqId); });
  if (vpsLedgerHasRow_(sheetName, reqId) !== true) {
    return vpsLedgerError_('ledger_unconfirmed', 'ledger row not found after upsert: ' + reqId, true, sheetName);
  }
  return { status: 'upserted', sheetName: sheetName };
}

function vpsLedgerRemove_(reqId, leaveDateText) {
  var fy = vpsLedgerFy_(leaveDateText);
  var sheetName = fy + '年度';
  var pre = vpsLedgerPrecheck_(fy);
  if (pre) return pre;
  vpsWithLedgerNoCreate_(function() { removeLedgerRow_(reqId, leaveDateText); });
  if (vpsLedgerHasRow_(sheetName, reqId) !== false) {
    return vpsLedgerError_('ledger_unconfirmed', 'ledger row still present after remove: ' + reqId, true, sheetName);
  }
  return { status: 'removed', sheetName: sheetName };
}

// 台帳の段が失敗なら、その段のコードで失敗の返事にする（済んだ段の結果を付ける）
function vpsLedgerFail_(out) {
  var extra = {};
  ['sign', 'mail', 'pdf', 'trash', 'ledger'].forEach(function(k) { if (out[k]) extra[k] = out[k]; });
  return vpsFail_(out.ledger.error, out.ledger.message, out.ledger.retryable, extra);
}

// ===== 前の PDF をゴミ箱へ（cleanupMode='trash' のときだけ） =====

// PDF_FOLDER_ID の下（親を VPS_TRASH_PARENT_DEPTH_ 段までたどる）にある PDF だけを捨てる
// （ステージングが本番の PDF を捨てないため）。
function vpsTrashFile_(fileId) {
  var rootId = normalize_(getSettings_()['PDF_FOLDER_ID']);
  if (!rootId) return { status: 'skipped', reason: 'pdf_folder_not_set' };
  try {
    var f = DriveApp.getFileById(fileId);
    if (String(f.getMimeType()) !== 'application/pdf') return { status: 'skipped', reason: 'not_pdf' };
    if (!vpsIsUnderFolder_(f, rootId, VPS_TRASH_PARENT_DEPTH_)) return { status: 'skipped', reason: 'outside_pdf_folder' };
    f.setTrashed(true);
    return { status: 'trashed', fileId: fileId };
  } catch (e) {
    console.warn('旧PDF削除スキップ: ' + e.message);
    return { status: 'error', error: vpsErrText_(e) };
  }
}

function vpsIsUnderFolder_(item, rootId, depth) {
  if (depth <= 0) return false;
  var it = item.getParents();
  while (it.hasNext()) {
    var parent = it.next();
    if (String(parent.getId()) === rootId) return true;
    if (vpsIsUnderFolder_(parent, rootId, depth - 1)) return true;
  }
  return false;
}

function vpsCleanupMode_(args) {
  var m = args.cleanupMode == null || args.cleanupMode === '' ? 'asis' : String(args.cleanupMode);
  if (m !== 'asis' && m !== 'trash') throw vpsError_('bad_request', "cleanupMode must be 'asis' or 'trash'");
  return m;
}

// ===== internal.ping =====
// 何も書かない（古い段の記録の掃除だけ）。MAIL_TO の値は返さない。
// auth はマニフェストを読めないので、ログインした人が分かるか（bearer）／分からないか（anonymous）で推し量る。
function vpsPing_() {
  invalidateMigrationCache_();
  var props = PropertiesService.getScriptProperties();
  var migratedTo = '';
  var readError = false;
  try {
    migratedTo = String(props.getProperty(MIGRATED_TO_URL_KEY) || '').trim();
  } catch (e0) {
    readError = true;
  }
  var activeEmail = '';
  try { activeEmail = String(Session.getActiveUser().getEmail() || ''); } catch (e1) { activeEmail = ''; }
  var out = {
    ok: true,
    apiVersion: INTERNAL_API_VERSION,
    scriptId: '',
    spreadsheetId: '',
    spreadsheetTimeZone: '',
    scriptTimeZone: '',
    migrated: !!migratedTo,
    migratedTo: migratedTo,
    auth: activeEmail ? 'bearer' : 'anonymous',
    settings: {},
    ledgerKeyRows: 0,
    ledgerOpenable: false,
    pdfFolderOpenable: false,
  };
  if (readError) out.migratedReadError = true;
  try { out.scriptId = ScriptApp.getScriptId(); } catch (e2) { out.scriptIdError = vpsErrText_(e2); }
  try { out.scriptTimeZone = Session.getScriptTimeZone(); } catch (e3) { /* ignore */ }
  try {
    var ss = getDb_();
    out.spreadsheetId = ss.getId();
    try { out.spreadsheetTimeZone = String(ss.getSpreadsheetTimeZone() || ''); } catch (e4) { out.spreadsheetTimeZoneError = vpsErrText_(e4); }
    var s = getSettings_();
    out.settings = {
      PDF_FOLDER_ID: normalize_(s['PDF_FOLDER_ID']),
      LEDGER_SSID: normalize_(s['LEDGER_SSID']),
      MASTER_SOURCE_SSID: normalize_(s['MASTER_SOURCE_SSID']),
      CALENDAR_SOURCE_SSID: normalize_(s['CALENDAR_SOURCE_SSID']),
      MAIL_TO_SET: !!normalize_(s['MAIL_TO']),
    };
    // Settings に LEDGER_SSID の行がいくつあるか（2 行以上なら下の行が使われる）
    var vals = ss.getSheetByName(SHEET.SETTINGS).getDataRange().getValues();
    var n = 0;
    for (var r = 1; r < vals.length; r++) if (String(vals[r][0] || '').trim() === 'LEDGER_SSID') n++;
    out.ledgerKeyRows = n;
    if (out.settings.LEDGER_SSID) {
      try { SpreadsheetApp.openById(out.settings.LEDGER_SSID); out.ledgerOpenable = true; } catch (e5) { out.ledgerError = vpsErrText_(e5); }
    }
    if (out.settings.PDF_FOLDER_ID) {
      try { DriveApp.getFolderById(out.settings.PDF_FOLDER_ID).getId(); out.pdfFolderOpenable = true; } catch (e6) { out.pdfFolderError = vpsErrText_(e6); }
    }
  } catch (e7) {
    out.dbError = vpsErrText_(e7);
  }
  out.prunedStageRecords = vpsPruneStageRecords_();
  return out;
}

// ===== internal.afterApprove1 =====
// api_approveLeaveRequest（leaveRequest.gs:505-551）の裏の処理: サイン → メール → 台帳。集計は呼ばない。
function vpsAfterApprove1_(job) {
  var a = job.args;
  var signBase64 = a.signBase64 == null ? '' : String(a.signBase64);
  var loaded = vpsLoadRow_(job);
  if (loaded.fail) return loaded.fail;
  var rc = loaded.rc;
  var out = { ok: true };

  // ① サイン: 手書きなら保存し、URL を列に書く。電子印なら空で上書き（今のまま）
  out.sign = vpsStage_(job, 'sign', function() {
    var url = '';
    var status = 'cleared';
    var error = '';
    if (signBase64) {
      var leaveDate = vpsParseYmd_(a.leaveDate) || rc.row[rc.idx['休暇日']];
      try {
        url = saveSignImage_(job.reqId, signBase64, leaveDate) || '';
        status = url ? 'saved' : 'error';
        if (!url) error = 'sign image was not saved (PDF_FOLDER_ID is not set)';
      } catch (signErr) {
        console.error('サイン画像保存エラー（続行）: ' + signErr.message);
        url = '';
        status = 'error';
        error = vpsErrText_(signErr);
      }
    }
    if (!vpsSetCell_(rc, 'サイン画像URL', url)) {
      throw vpsError_('bad_request', 'T_LEAVE_REQUEST に サイン画像URL の列がありません');
    }
    SpreadsheetApp.flush();
    var r = { status: status, signImageUrl: url };
    if (error) r.error = error;
    return r;
  });

  // ② 承認完了メール（MAIL_TO が空なら送らない＝今の関数と同じ）
  out.mail = vpsStage_(job, 'mail', function() {
    if (!normalize_(getSettings_()['MAIL_TO'])) return { status: 'skipped' };
    try {
      sendApprovalNotification_(job.reqId);
      return { status: 'sent' };
    } catch (mailErr) {
      console.error('メール通知エラー: ' + mailErr.message);
      return { status: 'error', error: vpsErrText_(mailErr) };
    }
  });
  SpreadsheetApp.flush();

  // ③ 台帳
  out.ledger = vpsStage_(job, 'ledger', function() {
    return vpsLedgerUpsert_(job.reqId, vpsLeaveDateText_(rc.row[rc.idx['休暇日']]));
  });
  if (out.ledger.status === 'error') return vpsLedgerFail_(out);
  return out;
}

// ===== internal.afterApprove2 =====
// api_approveLeaveRequest2（leaveRequest.gs:678-720）の裏の処理: PDF →（trash）→ 台帳（PDF の成否に関係なく）。
function vpsAfterApprove2_(job) {
  var a = job.args;
  var mode = vpsCleanupMode_(a);
  var loaded = vpsLoadRow_(job);
  if (loaded.fail) return loaded.fail;
  var rc = loaded.rc;
  var out = { ok: true };

  // ① PDF。シートの PDF_URL が空でなく oldPdfUrl と違えば、もう作ってある（作らない）
  out.pdf = vpsStage_(job, 'pdf', function() {
    var cur = vpsGetCell_(rc, 'PDF_URL');
    if (cur && cur !== normalize_(a.oldPdfUrl)) {
      return { status: 'already', pdfFileId: vpsGetCell_(rc, 'PDF_FILE_ID') || vpsDriveIdFromUrl_(cur), pdfUrl: cur, pdfName: '', folderId: '' };
    }
    var approver1Email = a.approver1Email != null ? String(a.approver1Email) : vpsGetCell_(rc, 'APPROVED_BY1_EMAIL');
    var approver2Email = a.approver2Email != null ? String(a.approver2Email) : '';
    var r;
    try {
      r = generateLeavePdf_(job.reqId, approver1Email, approver2Email);
    } catch (pdfErr) {
      console.error('2次承認PDF再生成エラー（続行）: ' + pdfErr.message);
      return { status: 'error', error: vpsErrText_(pdfErr) };
    }
    if (!r || !r.ok) return { status: 'error', error: 'generateLeavePdf_ returned no result' };
    vpsSetCell_(rc, 'PDF_URL', r.pdfUrl || '');
    vpsSetCell_(rc, 'PDF_FILE_ID', r.pdfFileId || ''); // 列があるときだけ
    SpreadsheetApp.flush();
    return { status: 'created', pdfFileId: r.pdfFileId || '', pdfUrl: r.pdfUrl || '', pdfName: r.pdfName || '', folderId: r.folderId || '' };
  });

  // ② 前の PDF（'trash' で、新しい PDF ができたときだけ。'asis' は何もしない）
  out.trash = vpsStage_(job, 'trash', function() {
    var oldId = normalize_(a.oldPdfFileId);
    if (mode !== 'trash' || !oldId) return { status: 'skipped' };
    if (out.pdf.status !== 'created') return { status: 'skipped', reason: 'pdf_not_created', noRecord: out.pdf.status === 'error' };
    return vpsTrashFile_(oldId);
  });

  // ③ 台帳（PDF の成否に関係なく）
  out.ledger = vpsStage_(job, 'ledger', function() {
    return vpsLedgerUpsert_(job.reqId, vpsLeaveDateText_(rc.row[rc.idx['休暇日']]));
  });
  if (out.ledger.status === 'error') return vpsLedgerFail_(out);
  if (out.pdf.status === 'error') {
    return vpsFail_('internal', 'pdf: ' + out.pdf.error, true, { failedStage: 'pdf', pdf: out.pdf, trash: out.trash, ledger: out.ledger });
  }
  return out;
}

// ===== internal.afterUnapprove =====
// api_editLeaveRequest（:937-954）・api_deleteLeaveRequest（:986-997）の裏の処理。
// シートから休暇日・PDF を読まず、引数を使う（VPS の行の反映が先に新しい値を書くため）。
function vpsAfterUnapprove_(job) {
  var a = job.args;
  var reason = String(a.reason || '');
  if (reason !== 'edit' && reason !== 'cancel') return vpsFail_('bad_request', "reason must be 'edit' or 'cancel'", false);
  var oldLeaveDate = String(a.oldLeaveDate == null ? '' : a.oldLeaveDate).trim();
  if (!vpsParseYmd_(oldLeaveDate)) return vpsFail_('bad_request', 'oldLeaveDate must be yyyy-MM-dd', false);
  var mode = vpsCleanupMode_(a);
  var loaded = vpsLoadRow_(job);
  if (loaded.fail) return loaded.fail;
  var rc = loaded.rc;
  var out = { ok: true };

  // ① サイン: 編集で承認が外れたときは列を空に（ファイルは消さない＝今のまま）
  out.sign = vpsStage_(job, 'sign', function() {
    if (a.clearSign !== true) return { status: 'skipped' };
    if (!vpsSetCell_(rc, 'サイン画像URL', '')) {
      throw vpsError_('bad_request', 'T_LEAVE_REQUEST に サイン画像URL の列がありません');
    }
    SpreadsheetApp.flush();
    return { status: 'cleared', signImageUrl: '' };
  });

  // ② 前の PDF（'trash' のときだけ。シートの PDF_URL が oldPdfUrl と同じときだけ捨てて 2 列を空に）
  out.trash = vpsStage_(job, 'trash', function() {
    var oldId = normalize_(a.oldPdfFileId);
    if (mode !== 'trash' || !oldId) return { status: 'skipped' };
    var cur = vpsGetCell_(rc, 'PDF_URL');
    if (cur !== normalize_(a.oldPdfUrl)) return { status: 'skipped', reason: 'sheet_has_other_pdf', pdfUrl: cur };
    var t = vpsTrashFile_(oldId);
    if (t.status === 'trashed') {
      vpsSetCell_(rc, 'PDF_URL', '');
      vpsSetCell_(rc, 'PDF_FILE_ID', ''); // 列があるときだけ
      SpreadsheetApp.flush();
      t.pdfUrl = '';
      t.pdfFileId = '';
    }
    return t;
  });

  // ③ 台帳の行を消す（編集前の休暇日の年度タブから。申請中の取消でも呼ぶ＝今のまま）
  out.ledger = vpsStage_(job, 'ledger', function() {
    return vpsLedgerRemove_(job.reqId, oldLeaveDate);
  });
  if (out.ledger.status === 'error') return vpsLedgerFail_(out);
  out.reason = reason;
  return out;
}

// ===== internal.sendWarnMail =====

function vpsSendWarnMailAction_(job) {
  var out = { ok: true };
  out.mail = vpsStage_(job, 'mail', function() {
    try {
      return vpsSendWarnMail_();
    } catch (e) {
      return { status: 'error', error: vpsErrText_(e) };
    }
  });
  if (out.mail.status === 'error') return vpsFail_('internal', 'mail: ' + out.mail.error, true, { mail: out.mail });
  return out;
}

// checkAndSendWarnMails_（warnMail.gs:11-57）から rebuildLeaveSummary_() の行を除いた写し。
// 集計タブ（V_LEAVE_SUMMARY）は VPS が直前に書いたものを読む。日付の確かめ・抽出・宛先・件名・本文は同じ。
function vpsSendWarnMail_() {
  var now = new Date();
  var month = now.getMonth() + 1; // 1-12
  var day = now.getDate();

  // 当日が警告送信日か判定
  var isWarnDay = false;
  for (var i = 0; i < WARN_MAIL_DATES.length; i++) {
    if (WARN_MAIL_DATES[i].month === month && WARN_MAIL_DATES[i].day === day) {
      isWarnDay = true;
      break;
    }
  }

  if (!isWarnDay) {
    Logger.log('本日は有給警告送信日ではありません: ' + fmtDate_(now, 'yyyy/MM/dd'));
    return { status: 'not_warn_day', count: 0 };
  }

  Logger.log('有給警告メール送信開始: ' + fmtDate_(now, 'yyyy/MM/dd'));

  // 未達者抽出
  var fy = computeFiscalYear_(now);
  var underWorkers = getUnderCountWorkers_(fy);

  if (underWorkers.length === 0) {
    Logger.log('有給未達者なし');
    return { status: 'no_targets', count: 0, fiscalYear: fy };
  }

  // メール送信
  var settings = getSettings_();
  var mailTo = normalize_(settings['MAIL_TO']);
  if (!mailTo) {
    Logger.log('MAIL_TO未設定のため警告メール送信スキップ');
    return { status: 'mail_to_empty', count: underWorkers.length, fiscalYear: fy };
  }

  var subject = '【有給取得警告】有給取得義務未達者一覧 ' + fmtDate_(now, 'yyyy/MM/dd');
  var body = buildWarnMailBody_(underWorkers, fy, now);

  GmailApp.sendEmail(mailTo, subject, body);
  Logger.log('有給警告メール送信完了: ' + underWorkers.length + '名');
  return { status: 'sent', count: underWorkers.length, fiscalYear: fy };
}
