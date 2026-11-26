// 内部用の入口（src/Internal.gs）の試験。node --test gas-tests/
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { makeWorld, TOKEN, REQ_HEADER, LEDGER_HEADER } = require('./harness');

const PNG_MAGIC = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
const SIGN = 'data:image/png;base64,' + Buffer.concat([PNG_MAGIC, Buffer.from('PNGDATA')]).toString('base64');
const A1_EXPECT = { status: '承認済', approvedAt: '2026-11-20 10:00:00', approvedBy1Email: 'boss@lineworks-local.info' };
const A2_EXPECT = { status: '承認済', approvedAt2: '2026-11-21 08:00:00', approvedBy2Email: 'somu@lineworks-local.info' };

function approve1(w, extra) {
  return w.call('internal.afterApprove1', Object.assign({
    jobId: 'j-a1', reqId: 'LV-20261120-001', expect: A1_EXPECT,
    args: { signBase64: SIGN, leaveDate: '2026-12-01' },
  }, extra || {}));
}
function approve2(w, extra, args) {
  return w.call('internal.afterApprove2', Object.assign({
    jobId: 'j-a2', reqId: 'LV-20261120-002', expect: A2_EXPECT,
    args: Object.assign({ approver1Email: 'boss@lineworks-local.info', approver2Email: 'somu@lineworks-local.info',
      oldPdfUrl: 'https://drive.google.com/file/d/OLDPDF/view', oldPdfFileId: 'OLDPDF', cleanupMode: 'asis' }, args || {}),
  }, extra || {}));
}
function unapprove(w, args, extra) {
  return w.call('internal.afterUnapprove', Object.assign({
    jobId: 'j-un', reqId: 'LV-20261120-003', expect: { status: '申請中' },
    args: Object.assign({ reason: 'edit', oldLeaveDate: '2026-12-03', clearSign: true,
      oldPdfUrl: 'https://drive.google.com/file/d/OLDPDF3/view', oldPdfFileId: 'OLDPDF3', cleanupMode: 'asis' }, args || {}),
  }, extra || {}));
}
const signFiles = w => Object.values(w.files).filter(f => /^sign_/.test(f.name));

// ---------------------------------------------------------------- 認証と本文
test('token が違えば何もせず unauthorized', () => {
  const w = makeWorld();
  const r = approve1(w, { token: 'x'.repeat(40) });
  assert.equal(r.ok, false); assert.equal(r.error, 'unauthorized'); assert.equal(r.retryable, false);
  assert.equal(w.calls.length, 0); assert.equal(w.writes.length, 0);
});

test('token が未設定なら全部拒む（本文に token が無くても）', () => {
  for (const props of [{ VPS_INTERNAL_TOKEN: null }, { VPS_INTERNAL_TOKEN: 'short' }]) {
    const w = makeWorld();
    if (props.VPS_INTERNAL_TOKEN === null) delete w.props.VPS_INTERNAL_TOKEN; else w.props.VPS_INTERNAL_TOKEN = props.VPS_INTERNAL_TOKEN;
    assert.equal(approve1(w).error, 'unauthorized');
    assert.equal(approve1(w, { token: '' }).error, 'unauthorized');
    assert.equal(approve1(w, { token: 'short' }).error, 'unauthorized');
    assert.equal(w.mails.length, 0);
  }
});

test('internal:true が無い本文・JSON でない本文・知らない action は bad_request', () => {
  const w = makeWorld();
  assert.equal(w.post({ token: TOKEN, action: 'internal.ping' }).error, 'bad_request');
  assert.equal(w.post('not json').error, 'bad_request');
  assert.equal(w.call('internal.nope').error, 'bad_request');
  assert.equal(w.call('internal.afterApprove1', { jobId: '', reqId: 'LV-20261120-001' }).error, 'bad_request');
  assert.equal(w.call('internal.afterApprove1', { jobId: 'j', reqId: '' }).error, 'bad_request');
  assert.equal(w.writes.length, 0);
});

test('実行の印は入口の中だけ立ち、終わると戻る', () => {
  const w = makeWorld();
  approve1(w);
  assert.equal(w.ctx.__KYUKA_INTERNAL_OK__, false);
  assert.equal(w.call('internal.ping').ok, true);
  assert.equal(w.ctx.__KYUKA_INTERNAL_OK__, false);
});

// ---------------------------------------------------------------- ping
test('ping は書かずに scriptId・DB・設定・移転モード・時間帯を返し、古い段の記録だけ消す', () => {
  const w = makeWorld();
  w.props['ij:old:mail'] = JSON.stringify({ t: Date.now() - 8 * 86400000, a: 'internal.afterApprove1', r: { status: 'sent' } });
  w.props['ij:new:mail'] = JSON.stringify({ t: Date.now(), a: 'internal.afterApprove1', r: { status: 'sent' } });
  const r = w.call('internal.ping', { jobId: undefined });
  assert.equal(r.ok, true);
  assert.equal(r.apiVersion, 1);
  assert.equal(r.scriptId, 'SCRIPT1');
  assert.equal(r.spreadsheetId, 'DB1');
  assert.equal(r.spreadsheetTimeZone, '');
  assert.equal(r.scriptTimeZone, 'Asia/Tokyo');
  assert.equal(r.migrated, false);
  assert.equal(r.settings.PDF_FOLDER_ID, 'PDFROOT');
  assert.equal(r.settings.LEDGER_SSID, 'LEDGER1');
  assert.equal(r.settings.MASTER_SOURCE_SSID, 'MASTERSRC');
  assert.equal(r.settings.CALENDAR_SOURCE_SSID, 'CALSRC');
  assert.equal(r.settings.MAIL_TO_SET, true);
  assert.equal(JSON.stringify(r).indexOf('somu@lineworks-local.info'), -1, 'MAIL_TO の値は返さない');
  assert.equal(r.ledgerKeyRows, 1);
  assert.equal(r.ledgerOpenable, true);
  assert.equal(r.pdfFolderOpenable, true);
  assert.equal(r.auth, 'bearer');
  assert.equal(r.action, 'internal.ping');
  assert.equal(w.writes.length, 0);
  assert.ok(!('ij:old:mail' in w.props));
  assert.ok('ij:new:mail' in w.props);
  assert.equal(w.count('tryLock'), 0);
});

test('ping は移転モードを毎回読み直して返す', () => {
  const w = makeWorld();
  assert.equal(w.call('internal.ping').migrated, false);
  w.props.MIGRATED_TO_URL = 'https://leave.lineworks-local.net';
  const r = w.call('internal.ping');
  assert.equal(r.migrated, true);
  assert.equal(r.spreadsheetId, 'DB1', '移転モードでも DB を読める');
});

// ---------------------------------------------------------------- afterApprove1
test('afterApprove1（手書き）: サイン → メール → 台帳の順で、今の関数を 1 回ずつ', () => {
  const w = makeWorld();
  const r = approve1(w);
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.action, 'internal.afterApprove1');
  assert.equal(r.jobId, 'j-a1');
  assert.equal(r.reqId, 'LV-20261120-001');
  assert.equal(r.sign.status, 'saved');
  assert.match(r.sign.signImageUrl, /^https:\/\/drive\.google\.com\/file\/d\/FILE\d+\/view$/);
  assert.equal(w.cell('LV-20261120-001', 'SIGN_IMAGE_URL'), r.sign.signImageUrl);
  const files = signFiles(w);
  assert.equal(files.length, 1);
  assert.equal(files[0].name, 'sign_LV-20261120-001.png');
  assert.equal(files[0].parent.name, '2026.12.01', '休暇日のフォルダ');
  assert.equal(r.mail.status, 'sent');
  assert.equal(w.mails.length, 1);
  assert.equal(w.mails[0].to, 'somu@lineworks-local.info');
  assert.match(w.mails[0].subject, /^【休暇届 承認完了】山田 有給消化（私用）$/);
  assert.equal(r.ledger.status, 'upserted');
  assert.equal(r.ledger.sheetName, '2026年度');
  assert.ok(w.ledgerIds('2026年度').includes('LV-20261120-001'));
  // 順番
  const names = w.calls.map(c => c[0]);
  const iSign = names.indexOf('createFile');
  const iMail = names.indexOf('sendEmail');
  const iLedger = w.calls.findIndex(c => c[0] === 'openById' && c[1] === 'LEDGER1');
  assert.ok(iSign >= 0 && iSign < iMail && iMail < iLedger, 'サイン → メール → 台帳');
  assert.equal(w.count('rebuildLeaveSummary_'), 0, '集計は呼ばない');
  assert.equal(w.count('SpreadsheetApp.create'), 0);
  assert.equal(w.count('tryLock'), 1);
  assert.equal(w.count('releaseLock'), 1);
});

test('afterApprove1（電子印）: サインは保存せず、列を空で上書きする（今のまま）', () => {
  const w = makeWorld();
  const r = approve1(w, { args: { signBase64: '', leaveDate: '2026-12-01' } });
  assert.equal(r.ok, true);
  assert.equal(r.sign.status, 'cleared');
  assert.equal(r.sign.signImageUrl, '');
  assert.equal(w.cell('LV-20261120-001', 'SIGN_IMAGE_URL'), '');
  assert.equal(signFiles(w).length, 0);
  assert.equal(w.mails.length, 1);
});

test('afterApprove1: MAIL_TO が空なら送らず skipped', () => {
  const w = makeWorld({ mailTo: '' });
  const r = approve1(w);
  assert.equal(r.ok, true);
  assert.equal(r.mail.status, 'skipped');
  assert.equal(w.mails.length, 0);
});

test('同じ jobId の再試行でメール 1 通・サイン 1 つ（台帳だけ失敗 → ok:true のまま、送り直すと台帳だけやり直す）', () => {
  const w = makeWorld();
  w.ledgerDown = true;
  const r1 = approve1(w);
  assert.equal(r1.ok, true, '台帳だけの失敗で全体を失敗にしない');
  assert.equal(r1.ledger.status, 'error');
  assert.equal(r1.ledger.error, 'ledger_unavailable');
  assert.equal(r1.ledger.retryable, true);
  assert.ok(r1.ledger.message);
  assert.equal(r1.sign.status, 'saved', '済んだ段の結果を返す');
  assert.equal(r1.mail.status, 'sent');
  assert.equal(w.count('SpreadsheetApp.create'), 0, '台帳を作らない');
  assert.equal(w.db.sheets.M_SYSTEM_SETTING.getLastRow(), 11, 'Settings に行を足さない');
  assert.equal(w.dbWrites().filter(x => x[1] === 'M_SYSTEM_SETTING').length, 0);
  w.ledgerDown = false;
  const r2 = approve1(w);
  assert.equal(r2.ok, true, JSON.stringify(r2));
  assert.equal(r2.sign.signImageUrl, r1.sign.signImageUrl);
  assert.equal(signFiles(w).length, 1);
  assert.equal(w.mails.length, 1);
  assert.equal(r2.ledger.status, 'upserted');
  // 3 回目は処理せずに同じ返事
  const writes = w.writes.length;
  const r3 = approve1(w);
  assert.equal(r3.replayed, true);
  assert.equal(r3.ok, true);
  assert.equal(w.writes.length, writes);
  assert.equal(w.mails.length, 1);
});

test('キャッシュが壊れていても、段の記録で二重にならない（罠 36）', () => {
  const w = makeWorld();
  approve1(w);
  w.cacheThrows = true;
  const r = approve1(w);
  assert.equal(r.ok, true);
  assert.equal(w.mails.length, 1);
  assert.equal(signFiles(w).length, 1);
});

test('同じ jobId を別の action に使うと bad_request', () => {
  const w = makeWorld();
  approve1(w, { jobId: 'same' });
  const r = approve2(w, { jobId: 'same' });
  assert.equal(r.error, 'bad_request');
  assert.equal(w.count('generateLeavePdf_'), 0);
});

test('expect が今の行と違えば何もせず row_not_synced', () => {
  const w = makeWorld();
  for (const ex of [{ status: '取消' }, { approvedAt: '2026-11-20 10:00:02' }, { approvedAt: '2026-11-20T10:00:00' }, { approvedBy1Email: 'x@lineworks-local.info' }, { leaveDate: '2026-12-02' }]) {
    const r = approve1(w, { jobId: 'j' + JSON.stringify(ex).length, expect: Object.assign({}, A1_EXPECT, ex) });
    assert.equal(r.ok, false); assert.equal(r.error, 'row_not_synced'); assert.equal(r.retryable, true);
  }
  assert.equal(w.mails.length, 0);
  assert.equal(signFiles(w).length, 0);
  assert.equal(w.dbWrites().length, 0);
  assert.equal(w.writes.length, 0);
  // 休暇日（日付だけ）も照合できる
  assert.equal(approve1(w, { jobId: 'jd', expect: { leaveDate: '2026-12-01', status: '承認済' } }).ok, true);
});

test('シートに REQ_ID が無ければ row_not_found（6 秒待たない）', () => {
  const w = makeWorld();
  const r = w.call('internal.afterApprove1', { reqId: 'LV-20991231-001', expect: {}, args: {} });
  assert.equal(r.error, 'row_not_found');
  assert.equal(r.retryable, true);
  assert.ok(r.durationMs < 1000);
});

test('ロックが取れなければ busy（何もしない）', () => {
  const w = makeWorld();
  w.lockBusy = true;
  const r = approve1(w);
  assert.equal(r.error, 'busy');
  assert.equal(r.retryable, true);
  assert.equal(w.mails.length, 0);
  assert.equal(w.writes.length, 0);
});

test('台帳の見出しが違えば書かずに ledger.status:error（ledger_header_mismatch・retryable:false）。メールは送る', () => {
  const w = makeWorld({ badLedgerHeader: true });
  const before = JSON.stringify(w.ledger.sheets['2026年度'].data);
  const r = approve1(w);
  assert.equal(r.ok, true);
  assert.equal(r.ledger.status, 'error');
  assert.equal(r.ledger.error, 'ledger_header_mismatch');
  assert.equal(r.ledger.retryable, false);
  assert.equal(r.mail.status, 'sent');
  assert.equal(JSON.stringify(w.ledger.sheets['2026年度'].data), before);
});

test('台帳に書けたか読み直す: 黙って失敗したら ledger.status:error（ledger_unconfirmed）', () => {
  const w = makeWorld();
  w.ctx.updateLedger_ = () => { w.calls.push(['updateLedger_']); };
  const r = approve1(w);
  assert.equal(r.ok, true);
  assert.equal(r.ledger.status, 'error');
  assert.equal(r.ledger.error, 'ledger_unconfirmed');
  assert.equal(r.ledger.retryable, true);
  assert.equal(w.count('updateLedger_'), 1);
});

test('台帳の段の中の例外も ledger.status:error（全体は ok:true）', () => {
  const w = makeWorld();
  w.ctx.updateLedger_ = () => { throw new Error('boom'); };
  const r = approve1(w);
  assert.equal(r.ok, true);
  assert.equal(r.ledger.status, 'error');
  assert.equal(r.ledger.error, 'ledger_error');
  assert.equal(r.ledger.message, 'boom');
  assert.equal(r.mail.status, 'sent');
});

test('LEDGER_SSID が空・開けないとき台帳を作らない（getLedgerSs_ の自動作成を動かさない）', () => {
  const w = makeWorld();
  w.db.sheets.M_SYSTEM_SETTING.data.find(r => r[0] === 'LEDGER_SSID')[1] = '';
  const r = approve1(w);
  assert.equal(r.ok, true);
  assert.equal(r.ledger.error, 'ledger_unavailable');
  assert.equal(w.count('SpreadsheetApp.create'), 0);
  // 入口の後は、画面から呼ぶ getLedgerSs_ は今のまま（差し替えが戻っている）
  assert.match(String(w.ctx.getLedgerSs_), /SpreadsheetApp\.create/);
});

test('台帳の年度タブが無ければ今の関数が作る（見出し 20 列）', () => {
  const w = makeWorld();
  delete w.ledger.sheets['2026年度'];
  w.ledger.order = [];
  const r = approve1(w);
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(w.ledger.sheets['2026年度'].data[0].slice(0, 20), LEDGER_HEADER);
});

// ---------------------------------------------------------------- afterApprove2
test('afterApprove2: PDF を作って PDF_URL に書き、台帳を書く。PDF_FILE_ID 列は足さない', () => {
  const w = makeWorld();
  const r = approve2(w);
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.pdf.status, 'created');
  assert.match(r.pdf.pdfFileId, /^FILE\d+$/);
  assert.equal(r.pdf.pdfUrl, 'https://drive.google.com/file/d/' + r.pdf.pdfFileId + '/view');
  assert.equal(r.pdf.pdfName, '20261202_加工課_山田_休暇届.pdf');
  assert.equal(r.pdf.folderId, 'FOLDER-2026.12.02');
  assert.equal(w.cell('LV-20261120-002', 'PDF_URL'), r.pdf.pdfUrl);
  assert.deepEqual(w.reqSheet.data[0], REQ_HEADER, '見出しは英語 26 列のまま');
  assert.equal(w.reqSheet.getLastColumn(), 26);
  const gen = w.calls.find(c => c[0] === 'generateLeavePdf_');
  assert.deepEqual(gen.slice(1), ['LV-20261120-002', 'boss@lineworks-local.info', 'somu@lineworks-local.info']);
  assert.equal(r.trash.status, 'skipped', 'asis では捨てない');
  assert.equal(w.files.OLDPDF.trashed, false);
  assert.equal(r.ledger.status, 'upserted');
  const lrow = w.ledger.sheets['2026年度'].data.find(x => x[0] === 'LV-20261120-002');
  assert.ok(String(lrow[19]).includes(r.pdf.pdfUrl), '台帳の PDF は新しい URL');
  // 順番: PDF の後に台帳
  const iPdf = w.calls.findIndex(c => c[0] === 'generateLeavePdf_');
  const iLedger = w.calls.findIndex((c, i) => i > iPdf && c[0] === 'openById');
  assert.ok(iLedger > iPdf);
  assert.equal(w.mails.length, 0);
});

test('afterApprove2: 再試行で PDF は 1 つ（台帳で失敗 → 送り直し）', () => {
  const w = makeWorld();
  w.ledgerDown = true;
  const r1 = approve2(w);
  assert.equal(r1.ok, true);
  assert.equal(r1.ledger.status, 'error');
  assert.equal(r1.ledger.error, 'ledger_unavailable');
  assert.equal(r1.pdf.status, 'created');
  w.ledgerDown = false;
  const r2 = approve2(w);
  assert.equal(r2.ok, true);
  assert.equal(r2.pdf.pdfUrl, r1.pdf.pdfUrl);
  assert.equal(w.count('generateLeavePdf_'), 1);
});

test('afterApprove2: シートの PDF_URL が oldPdfUrl と違えば作らずに already', () => {
  const w = makeWorld();
  const r = approve2(w, { jobId: 'other-tab' }, { oldPdfUrl: 'https://drive.google.com/file/d/SOMETHINGELSE/view' });
  assert.equal(r.ok, true);
  assert.equal(r.pdf.status, 'already');
  assert.equal(r.pdf.pdfUrl, 'https://drive.google.com/file/d/OLDPDF/view');
  assert.equal(r.pdf.pdfFileId, 'OLDPDF');
  assert.equal(w.count('generateLeavePdf_'), 0);
  assert.equal(r.ledger.status, 'upserted');
});

test('afterApprove2: PDF が例外でも台帳は書く。返事は retryable の失敗で、次は PDF をやり直す', () => {
  const w = makeWorld();
  w.pdfFail = true;
  const r1 = approve2(w);
  assert.equal(r1.ok, false);
  assert.equal(r1.error, 'internal');
  assert.equal(r1.failedStage, 'pdf');
  assert.equal(r1.retryable, true);
  assert.equal(r1.pdf.status, 'error');
  assert.equal(r1.ledger.status, 'upserted');
  assert.equal(w.cell('LV-20261120-002', 'PDF_URL'), 'https://drive.google.com/file/d/OLDPDF/view');
  w.pdfFail = false;
  const r2 = approve2(w);
  assert.equal(r2.ok, true);
  assert.equal(r2.pdf.status, 'created');
  assert.equal(w.count('generateLeavePdf_'), 2);
});

test("afterApprove2: 'trash' なら新しい PDF ができてから前の PDF を捨てる。PDF が失敗なら捨てない", () => {
  const w = makeWorld();
  w.pdfFail = true;
  const r0 = approve2(w, {}, { cleanupMode: 'trash' });
  assert.equal(r0.trash.status, 'skipped');
  assert.equal(w.files.OLDPDF.trashed, false);
  w.pdfFail = false;
  const r = approve2(w, {}, { cleanupMode: 'trash' });
  assert.equal(r.ok, true);
  assert.equal(r.trash.status, 'trashed');
  assert.equal(w.files.OLDPDF.trashed, true);
  const iGen = w.calls.map(c => c[0]).lastIndexOf('generateLeavePdf_');
  const iTrash = w.calls.findIndex(c => c[0] === 'setTrashed' && c[1] === 'OLDPDF');
  assert.ok(iTrash > iGen, '作ってから捨てる');
});

test("afterApprove2: 'trash' でも PDF フォルダの外のファイルは捨てない", () => {
  const w = makeWorld();
  const r = approve2(w, {}, { cleanupMode: 'trash', oldPdfFileId: 'FOREIGNPDF' });
  assert.equal(r.trash.status, 'skipped');
  assert.equal(r.trash.reason, 'outside_pdf_folder');
  assert.equal(w.files.FOREIGNPDF.trashed, false);
});

test('cleanupMode が知らない値なら bad_request', () => {
  const w = makeWorld();
  const r = approve2(w, {}, { cleanupMode: 'delete' });
  assert.equal(r.error, 'bad_request');
  assert.equal(r.retryable, false);
  assert.equal(w.count('generateLeavePdf_'), 0);
});

// ---------------------------------------------------------------- afterUnapprove
test('afterUnapprove（編集）: サインの列を空に、引数の休暇日の年度タブから行を消す。asis は PDF を残す', () => {
  const w = makeWorld();
  w.ledger.insertSheet('2027年度', 0).data.push(LEDGER_HEADER.slice());
  const r = unapprove(w);
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.reason, 'edit');
  assert.equal(r.sign.status, 'cleared');
  assert.equal(w.cell('LV-20261120-003', 'SIGN_IMAGE_URL'), '');
  assert.equal(w.files.OLDPDF3.trashed, false);
  assert.equal(w.cell('LV-20261120-003', 'PDF_URL'), 'https://drive.google.com/file/d/OLDPDF3/view');
  assert.equal(r.trash.status, 'skipped');
  assert.equal(r.ledger.status, 'removed');
  assert.equal(r.ledger.sheetName, '2026年度', 'シートの新しい休暇日（2027 年度）ではなく、引数の休暇日');
  assert.ok(!w.ledgerIds('2026年度').includes('LV-20261120-003'));
  assert.equal(w.count('rebuildLeaveSummary_'), 0);
});

test('afterUnapprove（取消）: サインは残す。台帳に行が無くても成功（申請中の取消）', () => {
  const w = makeWorld();
  const r = w.call('internal.afterUnapprove', { jobId: 'j-c', reqId: 'LV-20261126-001', expect: { status: '申請中' },
    args: { reason: 'cancel', oldLeaveDate: '2026-12-10', clearSign: false, oldPdfUrl: '', oldPdfFileId: '', cleanupMode: 'asis' } });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.sign.status, 'skipped');
  assert.equal(r.ledger.status, 'removed');
});

test("afterUnapprove: 'trash' はシートの PDF_URL が oldPdfUrl と同じときだけ捨てて列を空に", () => {
  const w = makeWorld();
  const r = unapprove(w, { cleanupMode: 'trash' });
  assert.equal(r.trash.status, 'trashed');
  assert.equal(w.files.OLDPDF3.trashed, true);
  assert.equal(w.cell('LV-20261120-003', 'PDF_URL'), '');
  const w2 = makeWorld();
  const r2 = unapprove(w2, { cleanupMode: 'trash', oldPdfUrl: 'https://drive.google.com/file/d/OTHER/view' });
  assert.equal(r2.trash.status, 'skipped');
  assert.equal(w2.files.OLDPDF3.trashed, false);
});

test('afterUnapprove: reason が違えば bad_request（何も書かない）', () => {
  const w = makeWorld();
  assert.equal(unapprove(w, { reason: 'other' }).error, 'bad_request');
  assert.equal(w.writes.length, 0);
});

test('afterUnapprove: oldLeaveDate の形が違っても失敗にせず、台帳の行は REQ_ID だけで探して消す', () => {
  for (const bad of ['', '2026/12/03', null]) {
    const w = makeWorld();
    w.ledger.insertSheet('2027年度', 0).data.push(LEDGER_HEADER.slice());
    const r = unapprove(w, { oldLeaveDate: bad });
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(r.sign.status, 'cleared');
    assert.equal(r.ledger.status, 'removed');
    assert.equal(r.ledger.searchedBy, 'reqId');
    assert.equal(r.ledger.sheetName, '2026年度');
    assert.ok(!w.ledgerIds('2026年度').includes('LV-20261120-003'));
  }
  // 台帳が開けなければ台帳の段だけ error
  const w2 = makeWorld();
  w2.ledgerDown = true;
  const r2 = unapprove(w2, { oldLeaveDate: 'x' });
  assert.equal(r2.ok, true);
  assert.equal(r2.sign.status, 'cleared');
  assert.equal(r2.ledger.status, 'error');
  assert.equal(r2.ledger.error, 'ledger_unavailable');
  // 見出しの違う年度タブにあれば、その段だけ error（書かない）
  const w3 = makeWorld({ badLedgerHeader: true });
  const r3 = unapprove(w3, { oldLeaveDate: '' });
  assert.equal(r3.ok, true);
  assert.equal(r3.ledger.error, 'ledger_header_mismatch');
  assert.ok(w3.ledgerIds('2026年度').includes('LV-20261120-003'));
});

test('afterUnapprove: 台帳から消えたか読み直す（黙って失敗したら ledger.status:error・ledger_unconfirmed）', () => {
  const w = makeWorld();
  w.ctx.removeLedgerRow_ = () => {};
  const r = unapprove(w);
  assert.equal(r.ok, true);
  assert.equal(r.ledger.status, 'error');
  assert.equal(r.ledger.error, 'ledger_unconfirmed');
});

// ---------------------------------------------------------------- sendWarnMail
test('sendWarnMail: 集計の書き直しを呼ばず、送信日に未達者へ 1 通。再試行で 2 通にしない', () => {
  const w = makeWorld({ now: new Date('2026-07-01T08:00:00+09:00').getTime() });
  const r = w.call('internal.sendWarnMail', { jobId: 'warn-1', args: { fiscalYear: 2026 } });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.mail.status, 'sent');
  assert.equal(r.mail.count, 1);
  assert.equal(w.mails.length, 1);
  assert.match(w.mails[0].subject, /^【有給取得警告】有給取得義務未達者一覧 2026\/07\/01$/);
  assert.match(w.mails[0].body, /山田  有給: 0回  （警告）/);
  assert.equal(w.count('rebuildLeaveSummary_'), 0);
  w.cache = {};
  const r2 = w.call('internal.sendWarnMail', { jobId: 'warn-1', args: { fiscalYear: 2026 } });
  assert.equal(r2.ok, true);
  assert.equal(w.mails.length, 1);
});

test('sendWarnMail: 送信日でなければ not_warn_day、MAIL_TO が空なら mail_to_empty', () => {
  const w = makeWorld({ now: new Date('2026-11-26T08:00:00+09:00').getTime() });
  assert.equal(w.call('internal.sendWarnMail', { jobId: 'w1' }).mail.status, 'not_warn_day');
  const w2 = makeWorld({ now: new Date('2027-01-01T08:00:00+09:00').getTime(), mailTo: '' });
  const r2 = w2.call('internal.sendWarnMail', { jobId: 'w2' });
  assert.equal(r2.mail.status, 'mail_to_empty', '1/1 は 2026 年度で、未達者がいる');
  const w3 = makeWorld({ now: new Date('2027-07-01T08:00:00+09:00').getTime() });
  assert.equal(w3.call('internal.sendWarnMail', { jobId: 'w4' }).mail.status, 'no_targets', '2027 年度の行は無い');
  assert.equal(w3.mails.length, 0);
  assert.equal(w.mails.length + w2.mails.length, 0);
  assert.equal(w.count('rebuildLeaveSummary_') + w2.count('rebuildLeaveSummary_'), 0);
});

test('sendWarnMail: MAIL_TO が空なら送らない', () => {
  const w = makeWorld({ now: new Date('2026-09-01T08:00:00+09:00').getTime(), mailTo: '' });
  const r = w.call('internal.sendWarnMail', { jobId: 'w3' });
  assert.equal(r.mail.status, 'mail_to_empty');
  assert.equal(w.mails.length, 0);
});

// ---------------------------------------------------------------- R2: 台帳へ書く文字（式を動かさない）
test('台帳へ書く文字: 入口からは先頭 = + - @ の文字に \' を付ける（数は除く）。画面からの updateLedger_ は今のまま', () => {
  const w = makeWorld();
  const row = w.reqSheet.data.find(x => x[0] === 'LV-20261120-001');
  row[REQ_HEADER.indexOf('WORKER_NAME')] = '=IMPORTXML("https://evil.example.com","//a")';
  row[REQ_HEADER.indexOf('PAIDLEAVE_DETAIL')] = '+SUM(1,2)';
  row[REQ_HEADER.indexOf('DEPT_NAME')] = '-5';
  row[REQ_HEADER.indexOf('APPROVED_BY1_EMAIL')] = '@boss';
  const r = approve1(w, { expect: { status: '承認済' } });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.ledger.status, 'upserted');
  const lrow = w.ledger.sheets['2026年度'].data.find(x => x[0] === 'LV-20261120-001');
  assert.equal(lrow[3], '\'=IMPORTXML("https://evil.example.com","//a")', '氏名');
  assert.equal(lrow[10], "'+SUM(1,2)", '理由・詳細');
  assert.equal(lrow[2], '-5', '数として読めるものはそのまま');
  assert.equal(lrow[14], "'@boss", '1次承認者');
  assert.equal(lrow[0], 'LV-20261120-001');
  // 入口の後は api_getLeaveRequestById が元に戻っている
  assert.equal(w.ctx.api_getLeaveRequestById('LV-20261120-001').workerName, '=IMPORTXML("https://evil.example.com","//a")');
  // 画面から呼ぶ updateLedger_ は今のまま（文字をそのまま書く）
  w.ctx.updateLedger_('LV-20261120-001');
  const lrow2 = w.ledger.sheets['2026年度'].data.find(x => x[0] === 'LV-20261120-001');
  assert.equal(lrow2[3], '=IMPORTXML("https://evil.example.com","//a")');
});

test('台帳へ書く文字: PDF の列は今の HYPERLINK の式のまま', () => {
  const w = makeWorld();
  const r = approve2(w);
  assert.equal(r.ok, true);
  const lrow = w.ledger.sheets['2026年度'].data.find(x => x[0] === 'LV-20261120-002');
  assert.equal(lrow[19], '=HYPERLINK("' + r.pdf.pdfUrl + '","PDF")');
});

test('vpsLedgerSafeText_: 文字だけを直す', () => {
  const w = makeWorld();
  const f = w.ctx.vpsLedgerSafeText_;
  assert.equal(f('=1+1'), "'=1+1");
  assert.equal(f('+81-90'), "'+81-90");
  assert.equal(f('-'), "'-");
  assert.equal(f('@x'), "'@x");
  assert.equal(f('-1.5'), '-1.5');
  assert.equal(f('+3'), '+3');
  assert.equal(f('1e3'), '1e3');
  assert.equal(f('山田'), '山田');
  assert.equal(f(''), '');
  assert.equal(f(12), 12);
});

// ---------------------------------------------------------------- R2: 段の記録と reqId
test('同じ jobId で reqId が違えば bad_request（返事の再生・段の記録の両方）', () => {
  const w = makeWorld();
  assert.equal(approve1(w).ok, true);
  const r1 = approve1(w, { reqId: 'LV-20261120-002', expect: {} });
  assert.equal(r1.ok, false);
  assert.equal(r1.error, 'bad_request');
  assert.equal(r1.replayed, undefined);
  w.cache = {}; // 返事の再生が無くても、段の記録で止まる
  const r2 = approve1(w, { reqId: 'LV-20261120-002', expect: {} });
  assert.equal(r2.error, 'bad_request');
  assert.match(r2.message, /another reqId/);
  assert.equal(w.mails.length, 1);
  assert.equal(signFiles(w).length, 1);
  assert.equal(w.cell('LV-20261120-002', 'SIGN_IMAGE_URL'), '');
  // 記録には reqId が入っている
  assert.equal(JSON.parse(w.props['ij:j-a1:mail']).q, 'LV-20261120-001');
  // 同じ reqId なら今まで通り
  assert.equal(approve1(w).ok, true);
  assert.equal(w.mails.length, 1);
});

// ---------------------------------------------------------------- R2: 承認日時の照合
test('承認日時: セルを秒で切り捨てて比べ、±1 秒は同じとみなす', () => {
  const cases = [
    ['2026-11-20T10:00:00.700', true],
    ['2026-11-20T09:59:59.999', true],
    ['2026-11-20T10:00:01.500', true],
    ['2026-11-20T09:59:59.000', true],
    ['2026-11-20T10:00:02.000', false],
    ['2026-11-20T09:59:58.900', false],
  ];
  for (const [cell, same] of cases) {
    const w = makeWorld();
    const row = w.reqSheet.data.find(x => x[0] === 'LV-20261120-001');
    row[REQ_HEADER.indexOf('APPROVED_AT')] = w.jst(cell);
    const r = approve1(w, { expect: { status: '承認済', approvedAt: '2026-11-20 10:00:00' } });
    assert.equal(r.ok, same, cell + ' ' + JSON.stringify(r));
    if (!same) assert.equal(r.error, 'row_not_synced');
  }
  // 2 次承認日時も同じ比べ方
  const w2 = makeWorld();
  w2.reqSheet.data.find(x => x[0] === 'LV-20261120-002')[REQ_HEADER.indexOf('APPROVED_AT2')] = w2.jst('2026-11-21T08:00:00.400');
  assert.equal(approve2(w2).ok, true);
  // セルが空で expect が空なら合う。セルに日時があって expect が空なら合わない
  const w3 = makeWorld();
  assert.equal(approve1(w3, { expect: { approvedAt2: '' } }).ok, true);
  assert.equal(approve1(w3, { jobId: 'j-x', expect: { approvedAt: '' } }).error, 'row_not_synced');
});

// ---------------------------------------------------------------- R2: 手書きサインの確かめ
test('手書きサイン: 2MB を超える・PNG でなければ保存せず sign.status:error、メールと台帳は続ける', () => {
  const big = 'data:image/png;base64,' + Buffer.concat([PNG_MAGIC, Buffer.alloc(2 * 1024 * 1024)]).toString('base64');
  const notPng = 'data:image/png;base64,' + Buffer.from('GIF89a-not-png').toString('base64');
  for (const [sign, msg] of [[big, /2MB/], [notPng, /not PNG/], ['data:image/png;base64,AA', /not PNG/]]) {
    const w = makeWorld();
    const r = approve1(w, { args: { signBase64: sign, leaveDate: '2026-12-01' } });
    assert.equal(r.ok, true, JSON.stringify(r).slice(0, 300));
    assert.equal(r.sign.status, 'error');
    assert.match(r.sign.error, msg);
    assert.equal(r.sign.signImageUrl, '');
    assert.equal(signFiles(w).length, 0);
    assert.equal(w.cell('LV-20261120-001', 'SIGN_IMAGE_URL'), '', '今の失敗時と同じく列は空');
    assert.equal(r.mail.status, 'sent');
    assert.equal(r.ledger.status, 'upserted');
  }
});

test('手書きサイン: GAS の base64Decode が符号付きのバイトを返しても PNG と分かる。ちょうど 2MB は通す', () => {
  const w = makeWorld();
  w.ctx.Utilities.base64Decode = s => Array.from(new Int8Array(Buffer.from(s, 'base64')));
  const exact = 'data:image/png;base64,' + Buffer.concat([PNG_MAGIC, Buffer.alloc(2 * 1024 * 1024 - 8)]).toString('base64');
  const r = approve1(w, { args: { signBase64: exact, leaveDate: '2026-12-01' } });
  assert.equal(r.sign.status, 'saved', JSON.stringify(r.sign));
  assert.equal(signFiles(w).length, 1);
});
