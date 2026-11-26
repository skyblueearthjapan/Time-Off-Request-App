// 移転モード（src/Cutover.gs と、今のファイルに差し込んだ行）の試験。node --test gas-tests/
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { makeWorld, publicFunctions, TOKEN } = require('./harness');

const NEW = 'https://leave.lineworks-local.net';
const MOVED = /休暇届は新しいアドレスへ移りました/;
const TRIGGER_FNS = ['syncAllMasters', 'checkAndSendWarnMails_', 'notifyTodayLeaves_', 'notifyPendingApproval1_', 'notifyPendingApproval2_'];
const OPERATOR_FNS = ['enterMigrationMode', 'enterMigrationModeForKyuka', 'exitMigrationMode', 'setInternalToken',
  'removeAllTriggersForMigration', 'removeAllTriggersForCutover', 'restoreTriggersAfterRollback',
  'restoreAllTriggersAfterRollback', 'showMigrationStatus', 'cutoverStatus'];
// 移転モードでも動いてよい公開関数（理由）
const ALLOWED_IN_MIGRATION = {
  doGet: '移転ページを返す', doPost: '内部用の入口', onOpen: 'メニュー 1 項目', showMigratedAddress: '読むだけ',
};

// ---------------------------------------------------------------- 設定が無いとき（今と同じ）
test('MIGRATED_TO_URL が無ければ doGet は今と同じ（getSettings_ を読み、各ページの題名）', () => {
  const w = makeWorld();
  const top = w.ctx.doGet({ parameter: {} });
  assert.equal(top.html, 'TEMPLATE:top');
  assert.equal(top.title, '休暇届 申請TOP');
  assert.equal(top.vars.APP_URL, 'https://script.google.com/macros/s/APPDEPLOY/exec');
  const prev = w.ctx.doGet({ parameter: { page: 'preview', reqId: 'LV-20261120-001' } });
  assert.equal(prev.title, '休暇届 プレビュー');
  assert.equal(prev.vars.REQ_ID, 'LV-20261120-001');
  const ls = w.ctx.doGet({ parameter: { page: 'leaveStatus' } });
  assert.equal(ls.title, '休暇届 取得状況');
  assert.ok(w.count('getActive') > 0);
});

test('MIGRATED_TO_URL が無ければ onOpen は今の 9 項目のメニュー', () => {
  const w = makeWorld();
  w.ctx.onOpen();
  assert.equal(w.menus.length, 1);
  assert.deepEqual(w.menus[0].items, ['setupAllSheets', 'syncAllMasters', 'syncCalendarMaster', 'syncStampMaster',
    'syncDeptApprovers', 'rebuildLeaveSummary', 'migrateLeaveRequestHeaders_', 'checkAndSendWarnMails', 'setupAllTriggers']);
});

test('MIGRATED_TO_URL が無ければ画面の関数・トリガーの関数・保守の関数は今のまま動く', () => {
  const w = makeWorld();
  assert.equal(w.ctx.getDb_(), w.db);
  assert.equal(w.ctx.api_getLeaveRequestById('LV-20261126-001').workerName, '鈴木');
  const r = w.ctx.api_deleteLeaveRequest('LV-20261126-001');
  assert.equal(r.ok, true);
  assert.equal(w.cell('LV-20261126-001', 'STATUS'), '取消');
  w.triggers = [];
  w.ctx.setupAllTriggers();
  assert.deepEqual(w.triggers.map(t => t.fn), ['syncAllMasters']);
  w.ctx.checkAndSendWarnMails_(); // 11/26 は送信日ではない（今の関数がそのまま動く）
  assert.ok(w.logs.some(l => /本日は有給警告送信日ではありません/.test(l)));
  assert.equal(w.ctx.api_getApproverList().length, 1);
});

test('INTERNAL_ONLY_DEPLOYMENT_ID が無ければ、ログインした人が分からなくても今と同じ画面', () => {
  const w = makeWorld();
  w.users.active = '';
  assert.equal(w.ctx.doGet({ parameter: {} }).html, 'TEMPLATE:top');
});

test('第 2 案: 匿名の内部用の版（ID が合う・メールが空）では画面を出さない。DOMAIN の版は今のまま', () => {
  const w = makeWorld({ props: { INTERNAL_ONLY_DEPLOYMENT_ID: 'ANONDEPLOY' } });
  w.serviceDeploymentId = 'ANONDEPLOY';
  w.users.active = '';
  const out = w.ctx.doGet({ parameter: { page: 'admin' } });
  assert.equal(out.html, 'このアドレスでは画面を開けません。');
  w.users.active = 'someone@lineworks-local.info';
  assert.notEqual(w.ctx.doGet({ parameter: {} }).html, 'このアドレスでは画面を開けません。');
  w.serviceDeploymentId = 'APPDEPLOY';
  w.users.active = '';
  assert.equal(w.ctx.doGet({ parameter: {} }).html, 'TEMPLATE:top');
});

// ---------------------------------------------------------------- 移転ページ
test('移転モード: doGet は getSettings_ より前に移転ページを返し、page・reqId を引き継ぐ', () => {
  const w = makeWorld();
  w.migrate(NEW);
  w.calls.length = 0;
  w.ctx.getSettings_ = () => { throw new Error('getSettings_ must not be called'); };
  const out = w.ctx.doGet({ parameter: { page: 'preview', reqId: 'LV-20261120-001' } });
  const target = 'https://leave.lineworks-local.net/?page=preview&amp;reqId=LV-20261120-001';
  assert.ok(out.html.includes('href="' + target + '"'), out.html);
  assert.ok(out.html.includes('<base target="_top">'));
  assert.ok(out.html.includes('target="_top"'));
  assert.ok(out.html.includes('window.top.location.href="https://leave.lineworks-local.net/?page=preview&reqId=LV-20261120-001"'));
  assert.match(out.html, MOVED);
  assert.equal(out.title, '休暇届（移転しました）');
  assert.equal(w.count('getActive'), 0, 'DB を開かない');
  const admin = w.ctx.doGet({ parameter: { page: 'admin' } });
  assert.ok(admin.html.includes('href="https://leave.lineworks-local.net/?page=admin"'));
  const none = w.ctx.doGet({ parameter: {} });
  assert.ok(none.html.includes('href="https://leave.lineworks-local.net/"'));
});

test('移転ページ: 7 ページ以外の page と、形の違う reqId は捨てる', () => {
  const w = makeWorld();
  w.migrate(NEW + '/');
  const out = w.ctx.doGet({ parameter: { page: '"><script>alert(1)</script>', reqId: 'LV-1"onmouseover=x' } });
  assert.ok(out.html.includes('href="https://leave.lineworks-local.net/"'));
  assert.ok(!out.html.includes('alert(1)'));
  assert.ok(!out.html.includes('onmouseover'));
  for (const p of ['top', 'admin', 'preview', 'sign', 'deptAdmin', 'somuAdmin', 'leaveStatus']) {
    assert.ok(w.ctx.doGet({ parameter: { page: p } }).html.includes('?page=' + p + '"'), p);
  }
  assert.ok(!w.ctx.doGet({ parameter: { page: 'unknown' } }).html.includes('?page='));
});

test('移転ページ: MIGRATED_TO_URL が正しい形でなければ、リンクを出さず固定の文言', () => {
  const w = makeWorld();
  w.migrate('https://evil.example.com');
  const out = w.ctx.doGet({ parameter: { page: 'top' } });
  assert.ok(!out.html.includes('evil.example.com'));
  assert.match(out.html, /管理者にお問い合わせください/);
});

// ---------------------------------------------------------------- 移転モードの止め
test('移転モード: getDb_() は止まり、内部用の印があるときだけ通る', () => {
  const w = makeWorld();
  w.migrate(NEW);
  assert.throws(() => w.ctx.getDb_(), MOVED);
  w.ctx.__KYUKA_INTERNAL_OK__ = true;
  assert.equal(w.ctx.getDb_(), w.db);
  w.ctx.__KYUKA_INTERNAL_OK__ = false;
  assert.throws(() => w.ctx.getDb_(), MOVED);
});

test('移転モード: トリガーの 5 関数は何もせずに戻る（DB・メール・ロックに触れない）', () => {
  const w = makeWorld();
  w.migrate(NEW);
  w.calls.length = 0;
  for (const fn of TRIGGER_FNS) assert.equal(w.ctx[fn](), undefined, fn);
  assert.equal(w.calls.length, 0, JSON.stringify(w.calls));
  assert.equal(w.mails.length, 0);
  assert.equal(w.writes.length, 0);
});

test('移転モード: onOpen は「移転済み（新しいアドレス）」の 1 項目だけ', () => {
  const w = makeWorld();
  w.migrate(NEW);
  w.ctx.onOpen();
  assert.equal(w.menus.length, 1);
  assert.deepEqual(w.menus[0].items, ['showMigratedAddress']);
  assert.match(w.ctx.showMigratedAddress(), /https:\/\/leave\.lineworks-local\.net\/ を開いてください/);
});

test('移転モード: google.script.run で呼べる公開関数の全部が、書かずに止まる', () => {
  const fns = publicFunctions();
  const apis = fns.filter(f => /^api_/.test(f.name));
  assert.equal(apis.length, 30, 'api_* は 30 関数');
  const checked = [];
  for (const { name } of fns) {
    if (ALLOWED_IN_MIGRATION[name]) continue;
    const w = makeWorld();
    w.migrate(NEW);
    w.users.active = 'someone@lineworks-local.info'; // 画面から呼ぶ人（持ち主ではない）
    w.calls.length = 0;
    const args = ['LV-20261120-001', { leaveDate: '2026-12-20' }, 'x@lineworks-local.info'];
    if (TRIGGER_FNS.includes(name)) {
      assert.equal(w.ctx[name](), undefined, name);
    } else if (OPERATOR_FNS.includes(name)) {
      assert.throws(() => w.ctx[name](NEW), /スクリプトの持ち主/, name);
    } else {
      assert.throws(() => w.ctx[name].apply(null, args), MOVED, name);
    }
    assert.equal(w.writes.length, 0, name + ' wrote: ' + JSON.stringify(w.writes));
    assert.equal(w.mails.length, 0, name);
    assert.equal(w.count('newTrigger') + w.count('deleteTrigger') + w.count('setProperty') + w.count('deleteProperty'), 0, name);
    assert.equal(w.count('createFile') + w.count('setTrashed') + w.count('SpreadsheetApp.create'), 0, name);
    checked.push(name);
  }
  // 7-5 の表の関数がすべて含まれている
  for (const must of ['setupAllSheets', 'setupAllTriggers', 'syncCalendarMaster', 'syncStampMaster', 'syncDeptApprovers',
    'rebuildLeaveSummary', 'checkAndSendWarnMails', 'debug_checkLeaveRequestSheet', 'test_generatePdf', 'test_stampLookup',
    'api_getApproverList', 'api_getLeaveRequestById', 'api_approveLeaveRequest2Batch', 'api_canAccessAdmin']) {
    assert.ok(checked.includes(must), must);
  }
});

test('移転モード: getDb_ を通らない末尾 _ の入口（setupAllSheets 以外のトリガー作成・ヘッダ追加）も止まる', () => {
  const w = makeWorld();
  w.migrate(NEW);
  for (const fn of ['setupSyncTrigger_', 'setupWarnMailTrigger_', 'setupNotifyTriggers_', 'migrateLeaveRequestHeaders_']) {
    assert.throws(() => w.ctx[fn](), MOVED, fn);
  }
  assert.equal(w.count('newTrigger'), 0);
  assert.equal(w.writes.length, 0);
});

test('移転モードでも内部用の入口は動く（1 次承認の後処理・ping）', () => {
  const w = makeWorld();
  w.migrate(NEW);
  const r = w.post({ internal: true, token: TOKEN, action: 'internal.afterApprove1', jobId: 'm1', reqId: 'LV-20261120-001',
    expect: { status: '承認済' }, args: { signBase64: '', leaveDate: '2026-12-01' } });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.mail.status, 'sent');
  assert.equal(r.ledger.status, 'upserted');
  assert.equal(w.mails.length, 1);
  const p = w.post({ internal: true, token: TOKEN, action: 'internal.ping' });
  assert.equal(p.migrated, true);
  assert.equal(p.migratedTo, NEW);
  // 入口の後は、また止まっている
  assert.throws(() => w.ctx.getDb_(), MOVED);
  assert.throws(() => w.ctx.api_getLeaveRequestById('LV-20261120-001'), MOVED);
});

// ---------------------------------------------------------------- 運用関数
test('運用関数は持ち主以外（画面から呼ぶ人）では拒む', () => {
  for (const fn of OPERATOR_FNS) {
    const w = makeWorld();
    w.users.active = 'someone@lineworks-local.info';
    assert.throws(() => w.ctx[fn](NEW), /スクリプトの持ち主/, fn);
    w.users.active = '';
    assert.throws(() => w.ctx[fn](NEW), /スクリプトの持ち主/, fn);
    assert.equal(w.count('setProperty') + w.count('deleteProperty') + w.count('deleteTrigger') + w.count('newTrigger'), 0, fn);
  }
});

test('運用関数は持ち主のメール（imaizumi@lineworks-local.info）だけ。active === effective でも別の人なら拒む', () => {
  for (const fn of OPERATOR_FNS) {
    const w = makeWorld();
    w.users.active = 'other@lineworks-local.info';
    w.users.effective = 'other@lineworks-local.info'; // スクリプトの写しを別の人がエディタから実行
    assert.throws(() => w.ctx[fn](NEW), /スクリプトの持ち主/, fn);
    w.users.active = 'imaizumi@lineworks-local.info'; // 持ち主が見ても、実行が別の人なら拒む
    assert.throws(() => w.ctx[fn](NEW), /スクリプトの持ち主/, fn);
    assert.equal(w.count('setProperty') + w.count('deleteProperty') + w.count('deleteTrigger') + w.count('newTrigger'), 0, fn);
  }
  const w2 = makeWorld();
  w2.users.active = ' Imaizumi@LINEWORKS-local.info ';
  w2.users.effective = 'imaizumi@lineworks-local.info';
  assert.equal(w2.ctx.assertCutoverOperator_(), 'imaizumi@lineworks-local.info', '大文字・空白は同じとみなす');
});

test('移転モードの止めの文: Error: で始まらず、新しいアドレスを含む（画面の withFailureHandler の err.message）', () => {
  const w = makeWorld();
  w.migrate(NEW);
  w.users.active = 'someone@lineworks-local.info';
  for (const fn of ['api_getLeaveRequestById', 'api_approveLeaveRequest', 'api_getApproverList']) {
    let err = null;
    try { w.ctx[fn]('LV-20261120-001'); } catch (e) { err = e; }
    assert.ok(err, fn);
    assert.ok(!/^Error: /.test(err.message), err.message);
    assert.equal(err.message, '休暇届は新しいアドレスへ移りました。https://leave.lineworks-local.net/ を開いてください。（この画面からは操作できません）');
  }
});

test('enterMigrationMode: 引数なし・違うアドレスは止まる。休暇届のアドレスだけ受ける', () => {
  const w = makeWorld();
  assert.throws(() => w.ctx.enterMigrationMode(), /引数に渡してください/);
  assert.throws(() => w.ctx.enterMigrationMode('https://overtime.lineworks-local.net'), /引数に渡してください/);
  assert.throws(() => w.ctx.enterMigrationMode('http://leave.lineworks-local.net'), /引数に渡してください/);
  assert.ok(!('MIGRATED_TO_URL' in w.props));
  assert.equal(w.ctx.enterMigrationMode(NEW + '/'), NEW);
  assert.equal(w.props.MIGRATED_TO_URL, NEW);
  assert.equal(w.ctx.isMigrationMode_(), true);
  assert.equal(w.ctx.exitMigrationMode(), true);
  assert.ok(!('MIGRATED_TO_URL' in w.props));
  assert.equal(w.ctx.isMigrationMode_(), false);
  assert.equal(w.ctx.doGet({ parameter: {} }).html, 'TEMPLATE:top', '外せば今の画面に戻る');
  assert.equal(w.ctx.enterMigrationModeForKyuka(), NEW);
});

test('removeAllTriggersForMigration は全部を控えてから全部消し、戻す関数で作り直せる', () => {
  const w = makeWorld({ triggers: ['syncAllMasters', 'notifyTodayLeaves_', 'checkAndSendWarnMails_', 'someOldFn'] });
  w.ctx.enterMigrationModeForKyuka();
  const r = w.ctx.removeAllTriggersForCutover();
  assert.equal(r.removed.length, 4);
  assert.equal(w.triggers.length, 0);
  const rec = JSON.parse(w.props.MIGRATION_REMOVED_TRIGGERS);
  assert.deepEqual(rec.map(x => x.fn), ['syncAllMasters', 'notifyTodayLeaves_', 'checkAndSendWarnMails_', 'someOldFn']);
  const iRec = w.calls.findIndex(c => c[0] === 'setProperty' && c[1] === 'MIGRATION_REMOVED_TRIGGERS');
  const iDel = w.calls.findIndex(c => c[0] === 'deleteTrigger');
  assert.ok(iRec < iDel, '控えてから消す');
  // 移転モードのままでは戻さない
  assert.throws(() => w.ctx.restoreTriggersAfterRollback(), /移転モードのままです/);
  assert.throws(() => w.ctx.restoreAllTriggersAfterRollback(), /移転モードのままです/);
  w.ctx.exitMigrationMode();
  const all = w.ctx.restoreAllTriggersAfterRollback();
  assert.deepEqual(Array.from(all.triggers).sort(), ['checkAndSendWarnMails_', 'notifyTodayLeaves_', 'syncAllMasters']);
  assert.deepEqual(Array.from(all.unknown), ['someOldFn (CLOCK)']);
  assert.ok('MIGRATION_REMOVED_TRIGGERS' in w.props, '作り直せないものがあれば控えを残す');
  const sync = w.triggers.find(t => t.fn === 'syncAllMasters');
  assert.deepEqual(sync.s, { hour: 6, minute: 0 });
  assert.deepEqual(w.triggers.find(t => t.fn === 'notifyTodayLeaves_').s, { hour: 7, minute: 0 });
  // もう一度でも二重に作らない
  w.ctx.restoreTriggersAfterRollback();
  assert.equal(w.triggers.filter(t => t.fn === 'syncAllMasters').length, 1);
});

test('restoreTriggersAfterRollback は syncAllMasters の 1 本だけ作る', () => {
  const w = makeWorld({ triggers: ['syncAllMasters', 'notifyTodayLeaves_'] });
  w.ctx.removeAllTriggersForMigration();
  assert.deepEqual(w.ctx.restoreTriggersAfterRollback(), ['syncAllMasters']);
});

test('setInternalToken は値を出さずに保存し、短い値を拒む。状態の関数は何も変えない', () => {
  const w = makeWorld();
  delete w.props.VPS_INTERNAL_TOKEN;
  const r = w.ctx.setInternalToken();
  assert.deepEqual(JSON.parse(JSON.stringify(r)), { saved: true, length: 64 });
  assert.equal(w.props.VPS_INTERNAL_TOKEN.length, 64);
  assert.ok(!w.logs.some(l => l.includes(w.props.VPS_INTERNAL_TOKEN)));
  assert.throws(() => w.ctx.setInternalToken('short'), /32 字以上/);
  w.calls.length = 0;
  w.migrate(NEW);
  const s = w.ctx.cutoverStatus();
  assert.equal(s.migratedTo, NEW);
  assert.equal(s.internalTokenSet, true);
  assert.equal(s.apiVersion, 1);
  assert.deepEqual(Array.from(s.triggers), ['syncAllMasters']);
  assert.equal(w.count('getActive'), 0, 'DB を開かない');
  assert.equal(w.count('setProperty') + w.count('deleteProperty'), 0);
  assert.equal(JSON.stringify(w.ctx.showMigrationStatus()), JSON.stringify(s));
});
