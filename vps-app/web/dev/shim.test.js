// gas-shim.js の試験。外部の npm パッケージを使わない。
//   node --test vps-app/web/dev/
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { webcrypto } = require('crypto');

const SHIM = fs.readFileSync(path.join(__dirname, '..', 'static', 'js', 'gas-shim.js'), 'utf8');
const EMAIL = 'imaizumi@lineworks-local.info';

// ---- 偽物のブラウザ ----
function makeEl(tag, attrs, parent) {
  const el = {
    nodeType: 1, tagName: tag, attrs: Object.assign({}, attrs || {}), parentNode: parent || null,
    disabled: false, innerHTML: '', value: '', connected: true,
    getAttribute(k) { return Object.prototype.hasOwnProperty.call(this.attrs, k) ? this.attrs[k] : null; },
    hasAttribute(k) { return Object.prototype.hasOwnProperty.call(this.attrs, k); },
    contains(o) { for (let n = o; n; n = n.parentNode) { if (n === this) { return true; } } return false; }
  };
  return el;
}

function makeEnv(opts) {
  opts = opts || {};
  const listeners = {};
  const store = new Map(Object.entries(opts.session || {}));
  const env = {
    alerts: [], fetches: [], buttons: [], errors: [],
    me: opts.me === undefined ? { email: EMAIL } : opts.me,
    queue: [],        // RPC の返事（先頭から 1 つずつ使う）。無ければ hold（返事しない）
    held: []
  };
  const location = { pathname: '/', search: opts.search || '?page=top', href: '' };
  const document = {
    readyState: 'complete', visibilityState: 'visible',
    addEventListener(t, f) { (listeners[t] = listeners[t] || []).push(f); },
    getElementById(id) {
      if (id === 'gas-template-data') { return { textContent: JSON.stringify({ PAGE: opts.page || 'top' }) }; }
      return null;
    },
    querySelectorAll() { return env.buttons; },
    contains(el) { return el.connected !== false; }
  };
  function response(status, body) {
    const text = typeof body === 'string' ? body : JSON.stringify(body);
    return { status: status, text: () => Promise.resolve(text), json: () => Promise.resolve(JSON.parse(text)) };
  }
  const win = {
    document: document, location: location,
    alert: (m) => env.alerts.push(m),
    sessionStorage: {
      getItem: (k) => (store.has(k) ? store.get(k) : null),
      setItem: (k, v) => store.set(k, String(v)),
      removeItem: (k) => store.delete(k)
    },
    setTimeout: (f, ms) => setTimeout(() => { try { f(); } catch (e) { env.errors.push(e); } }, ms),
    crypto: webcrypto,
    addEventListener() {}, removeEventListener() {},
    fetch: (url, init) => {
      env.fetches.push({ url: url, init: init || {} });
      if (url === '/auth/me') {
        return Promise.resolve(env.me === 401 ? response(401, { ok: false }) : response(200, env.me));
      }
      const next = env.queue.shift();
      if (!next) { return new Promise((resolve) => env.held.push((s, b) => resolve(response(s, b)))); }
      if (next === 'network') { return Promise.reject(new TypeError('Failed to fetch')); }
      return Promise.resolve(response(next[0], next[1]));
    }
  };
  win.window = win;
  if (opts.retryMs !== undefined) { win.__VPS_SHIM_RETRY_MS = opts.retryMs; }
  vm.createContext(win);
  vm.runInContext(SHIM, win, { filename: 'gas-shim.js' });
  env.win = win;
  env.global = vm.runInContext('this', win);
  env.store = store;
  env.location = location;
  env.fire = (type, ev) => (listeners[type] || []).forEach((f) => f(ev));
  env.rpcFetches = () => env.fetches.filter((f) => f.url.startsWith('/api/rpc/'));
  return env;
}

const flush = (n) => new Promise((r) => setTimeout(r, n || 20));

// ---- 呼び出しの形 ----

test('ハンドラの付け方 4 通りが動き、呼ぶたびに別の runner になる', async () => {
  const env = makeEnv();
  const got = [];
  const run = env.win.google.script.run;
  env.queue.push([200, { ok: true, result: 1 }], [200, { ok: true, result: 2 }], [400, { ok: false, error: 'だめ' }], [200, { ok: true, result: 4 }]);
  run.withSuccessHandler((r) => got.push(['s1', r])).api_getDeptList();
  run.withFailureHandler((e) => got.push(['f2', e.message])).withSuccessHandler((r) => got.push(['s2', r])).api_getLookupData();
  run.withSuccessHandler((r) => got.push(['s3', r])).withFailureHandler((e) => got.push(['f3', e.message])).api_getAllPendingRequests();
  run.api_getDeptList();  // どちらも無し（何も起きない）
  await flush();
  assert.deepEqual(got.sort(), [['f3', 'だめ'], ['s1', 1], ['s2', 2]].sort());
  assert.equal(typeof run.withSuccessHandler(() => {}).api_getDeptList, 'function');
  assert.notEqual(run.withSuccessHandler(() => {}), run);
  assert.equal(run.api_getFormUrl, undefined, '閉じたリストに無い関数は生えない');
});

test('省略した引数は送らない・配列とオブジェクトはそのまま送る', async () => {
  const env = makeEnv();
  env.queue.push([200, { ok: true, result: {} }], [200, { ok: true, result: {} }], [200, { ok: true, result: {} }]);
  env.win.google.script.run.api_editLeaveRequest('id-1');
  env.win.google.script.run.api_editLeaveRequest('id-1', { leaveDate: '2026-11-30' });
  env.win.google.script.run.api_getDeptAdminView(2026, ['製造1課', '品質管理課']);
  await flush();
  const f = env.rpcFetches();
  assert.equal(f[0].url, '/api/rpc/api_editLeaveRequest');
  assert.equal(f[0].init.body, '["id-1"]');
  assert.equal(f[1].init.body, '["id-1",{"leaveDate":"2026-11-30"}]');
  assert.equal(f[2].init.body, '[2026,["製造1課","品質管理課"]]');
  assert.equal(f[0].init.method, 'POST');
  assert.equal(f[0].init.credentials, 'same-origin');
  assert.equal(f[0].init.headers['Content-Type'], 'application/json');
});

test('200 の {error}・{ok:false}・null は成功ハンドラに届く', async () => {
  const env = makeEnv();
  const got = [];
  env.queue.push([200, { ok: true, result: { error: 'x', items: [] } }], [200, { ok: true, result: { ok: false, message: 'm' } }], [200, { ok: true, result: null }]);
  const r = env.win.google.script.run.withSuccessHandler((v) => got.push(v)).withFailureHandler(() => got.push('FAIL'));
  r.api_getAllPendingRequests();
  r.api_submitLeaveRequest({ a: 1 });
  r.api_getWorkerInfo();
  await flush();
  assert.deepEqual(JSON.parse(JSON.stringify(got)), [{ error: 'x', items: [] }, { ok: false, message: 'm' }, null]);
});

test('400 は失敗ハンドラに err.message と String(err) の両方で文が出る', async () => {
  const env = makeEnv();
  let err = null;
  env.queue.push([400, { ok: false, error: '既にキャンセル済みです。' }]);
  env.win.google.script.run.withFailureHandler((e) => { err = e; }).api_deleteLeaveRequest('a');
  await flush();
  assert.equal(err.message, '既にキャンセル済みです。');
  assert.equal(String(err), '既にキャンセル済みです。');
  assert.equal('エラー: ' + (err.message || err), 'エラー: 既にキャンセル済みです。');
});

test('5xx・通信エラーは決まった文で失敗ハンドラへ', async () => {
  const env = makeEnv();
  const msgs = [];
  env.queue.push([500, { ok: false, error: '内部の詳細' }], 'network', [502, '<html>bad gateway</html>']);
  const r = env.win.google.script.run.withFailureHandler((e) => msgs.push(e.message));
  r.api_getDeptList(); r.api_getLookupData(); r.api_getAllPendingRequests();
  await flush();
  const M = env.win.__vpsShim.INTERNAL_MSG;
  assert.equal(M, 'サーバーでエラーが起きました。時間をおいてもう一度お試しください。');
  assert.deepEqual(msgs, [M, M, M]);
});

// ---- 冪等キー ----

test('書き込みにだけ冪等キー。5xx の後は同じキー、200 の後は新しいキー', async () => {
  const env = makeEnv();
  const r = env.win.google.script.run.withFailureHandler(() => {}).withSuccessHandler(() => {});
  env.queue.push([200, { ok: true, result: [] }]);
  r.api_getDeptList();
  await flush();
  assert.equal(env.rpcFetches()[0].init.headers['X-Idempotency-Key'], undefined);

  env.queue.push([500, {}], [200, { ok: true, result: { ok: true } }], [200, { ok: true, result: { ok: true } }]);
  r.api_approveLeaveRequest('req-1');
  await flush();
  r.api_approveLeaveRequest('req-1');
  await flush();
  r.api_approveLeaveRequest('req-1');
  await flush();
  const keys = env.rpcFetches().slice(1).map((f) => f.init.headers['X-Idempotency-Key']);
  assert.ok(keys[0]);
  assert.equal(keys[1], keys[0], '5xx の後の押し直しは同じキー');
  assert.notEqual(keys[2], keys[0], '200 で捨てたので新しいキー');
});

test('400 でキーを捨て、引数が違えば別のキー', async () => {
  const env = makeEnv();
  const r = env.win.google.script.run.withFailureHandler(() => {});
  env.queue.push([400, { ok: false, error: 'x' }], [400, { ok: false, error: 'x' }], [400, { ok: false, error: 'x' }]);
  r.api_deleteLeaveRequest('a'); await flush();
  r.api_deleteLeaveRequest('a'); await flush();
  r.api_deleteLeaveRequest('b'); await flush();
  const keys = env.rpcFetches().map((f) => f.init.headers['X-Idempotency-Key']);
  assert.notEqual(keys[0], keys[1]);
  assert.notEqual(keys[1], keys[2]);
});

test('同じ書き込みの二度押しは 1 回だけ送り、両方のハンドラに同じ返事', async () => {
  const env = makeEnv();
  const got = [];
  const payload = { leaveType: '有給消化（私用）', leaveDate: '2026-11-30' };
  env.win.google.script.run.withSuccessHandler((v) => got.push(['a', v])).api_submitLeaveRequest(payload);
  env.win.google.script.run.withSuccessHandler((v) => got.push(['b', v])).api_submitLeaveRequest(payload);
  await flush();
  assert.equal(env.rpcFetches().length, 1);
  env.held.shift()(200, { ok: true, result: { ok: true, reqId: 'LV-20261130-001' } });
  await flush();
  assert.deepEqual(JSON.parse(JSON.stringify(got)), [['a', { ok: true, reqId: 'LV-20261130-001' }], ['b', { ok: true, reqId: 'LV-20261130-001' }]]);
  // 読む関数は二度押しをまとめない
  env.win.google.script.run.api_getAllPendingRequests();
  env.win.google.script.run.api_getAllPendingRequests();
  await flush();
  assert.equal(env.rpcFetches().length, 3);
});

test('409 は同じキーで送り直し、通れば成功', async () => {
  const env = makeEnv({ retryMs: 1 });
  let ok = null;
  env.queue.push([409, { ok: false, error: '処理中です' }], [409, { ok: false, error: '処理中です' }], [200, { ok: true, result: { ok: true } }]);
  env.win.google.script.run.withSuccessHandler((v) => { ok = v; }).api_approveLeaveRequest2('r');
  await flush(60);
  const f = env.rpcFetches();
  assert.equal(f.length, 3);
  assert.equal(new Set(f.map((x) => x.init.headers['X-Idempotency-Key'])).size, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(ok)), { ok: true });
});

test('409 が 10 回続いたらサーバーの文で失敗', async () => {
  const env = makeEnv({ retryMs: 1 });
  let err = null;
  for (let i = 0; i < 11; i++) { env.queue.push([409, { ok: false, error: '同じ操作を処理中です。' }]); }
  env.win.google.script.run.withFailureHandler((e) => { err = e; }).api_approveLeaveRequest('r');
  await flush(150);
  assert.equal(env.rpcFetches().length, 11);
  assert.equal(err.message, '同じ操作を処理中です。');
});

// ---- 401（罠 20） ----

test('401 は失敗ハンドラを呼ばず、書きかけを退避し、ボタンを戻して、ログインへ', async () => {
  const env = makeEnv({ search: '?page=sign&reqId=LV-20261126-001', page: 'sign' });
  await flush();  // 読み込み直後の /auth/me でメールが分かる
  env.win.__vpsDraft.register({ name: 'p', collect: () => ({ typed: '書きかけ' }), restore() {} });
  env.win.__unloadRemoved = 0;
  env.win.onbeforeunload = () => 'x';

  // 受け口から押したボタンが、処理中に無効になったまま 401 になる
  const btn = makeEl('BUTTON', { 'data-on-click': 'doApprove', 'data-args': '["R1","$this"]' });
  btn.innerHTML = '承認する';
  env.buttons = [btn];
  let failed = false;
  env.win.doApprove = function (id, b) {
    b.disabled = true; b.innerHTML = '処理中...';
    env.win.google.script.run.withFailureHandler(() => { failed = true; }).api_approveLeaveRequest2(id);
  };
  env.queue.push([401, { ok: false, error: 'ログインが必要です。', login: '/auth/login' }]);
  env.fire('click', { type: 'click', target: btn });
  assert.equal(btn.disabled, true);
  await flush();

  assert.equal(failed, false, '失敗ハンドラは呼ばない');
  assert.deepEqual(env.alerts, ['ログインが切れました。ログイン後にもう一度押してください。']);
  assert.equal(env.location.href, '/auth/login?next=' + encodeURIComponent('/?page=sign&reqId=LV-20261126-001'));
  assert.equal(btn.disabled, false);
  assert.equal(btn.innerHTML, '承認する');
  assert.equal(env.win.onbeforeunload, null);
  const saved = JSON.parse(env.store.get('kyuka-draft:sign'));
  assert.equal(saved.email, EMAIL);
  assert.deepEqual(saved.data, { p: { typed: '書きかけ' } });

  // 2 つ目の 401 で alert を重ねない。冪等キーは捨てていない
  env.queue.push([401, {}]);
  env.win.google.script.run.api_approveLeaveRequest2('R1');
  await flush();
  assert.equal(env.alerts.length, 1);
  const keys = env.rpcFetches().map((f) => f.init.headers['X-Idempotency-Key']);
  assert.equal(keys[0], keys[1]);
});

test('ログインから戻ったら、同じメールで 24 時間以内の書きかけだけ戻す', async () => {
  const draft = (email, age) => JSON.stringify({ savedAt: Date.now() - age, email: email, data: { top: { v: 1 } } });
  const cases = [
    [draft(EMAIL, 1000), true],
    [draft('other@lineworks-local.info', 1000), false],
    [draft(EMAIL, 25 * 3600 * 1000), false]
  ];
  for (const [saved, expected] of cases) {
    const env = makeEnv({ session: { 'kyuka-draft:top': saved } });
    let restored = null;
    env.win.__vpsDraft.register({ name: 'top', collect: () => null, restore: (d) => { restored = d; } });
    await flush();
    assert.equal(restored !== null, expected, saved);
    assert.equal(env.store.has('kyuka-draft:top'), false, '読んだら消す');
  }
});

test('画面に戻ったとき /auth/me が 401 ならその場でログインへ', async () => {
  const env = makeEnv();
  await flush();
  env.me = 401;
  env.fire('visibilitychange', {});
  await flush();
  assert.equal(env.location.href, '/auth/login?next=' + encodeURIComponent('/?page=top'));
  assert.equal(env.alerts.length, 1);
});

// ---- on 属性の代わりの受け口 ----

test('受け口は data-on-* の関数を、内側から順に、$this・$value・数・真偽値の引数で呼ぶ', () => {
  const env = makeEnv();
  const calls = [];
  env.win.loadData = function () { calls.push(['outer', this === env.global]); };
  env.win.switchTab = function (a, b, c, d) { calls.push(['inner', a, b === inner, c, d]); };
  env.win.onApproverChange = function (v) { calls.push(['val', v]); };
  const div = makeEl('DIV', { 'data-on-click': 'loadData' });
  const inner = makeEl('BUTTON', { 'data-on-click': 'switchTab', 'data-args': '["x","$this",3,true]' }, div);
  const text = { nodeType: 3, parentNode: inner };
  env.fire('click', { type: 'click', target: text });
  const sel = makeEl('SELECT', { 'data-on-change': 'onApproverChange', 'data-args': '["$value"]' });
  sel.value = '1.5';
  env.fire('change', { type: 'change', target: sel });
  // 無効なボタンの click は呼ばない
  inner.disabled = true;
  env.fire('click', { type: 'click', target: inner });
  assert.deepEqual(calls, [['inner', 'x', true, 3, true], ['outer', true], ['val', '1.5'], ['outer', true]]);
});

test('受け口は一覧（ON_FNS）に無い関数名を呼ばない（data-on-click="eval" などを入れられても）', async () => {
  const env = makeEnv();
  const calls = [];
  for (const name of ['evil', 'eval', 'alert', 'print', 'open']) { env.win[name] = () => calls.push(name); }
  for (const name of ['evil', 'eval', 'alert', 'print', 'open', 'constructor', '__proto__', 'toString', 'hasOwnProperty']) {
    const btn = makeEl('BUTTON', { 'data-on-click': name, 'data-args': '["x"]' });
    env.fire('click', { type: 'click', target: btn });
  }
  await flush();
  assert.deepEqual(calls, []);
  assert.equal(env.errors.length, 9);
  assert.ok(env.errors.every((e) => /呼べない画面の関数です/.test(e.message)));
  // 一覧の名前は呼ぶ
  env.win.loadData = () => calls.push('loadData');
  env.fire('click', { type: 'click', target: makeEl('BUTTON', { 'data-on-click': 'loadData' }) });
  assert.deepEqual(calls, ['loadData']);
});

test('mouseenter・mouseleave は要素の外から入ったとき・外へ出たときだけ', () => {
  const env = makeEnv();
  const calls = [];
  env.win.opsLoadStatus = (el, t, s) => calls.push([t, s]);
  const btn = makeEl('BUTTON', {
    'data-on-click': 'x',
    'data-on-mouseenter': 'opsLoadStatus', 'data-args-mouseenter': '["$this","translateY(-1px)","s"]',
    'data-on-mouseleave': 'opsLoadStatus', 'data-args-mouseleave': '["$this","",""]'
  });
  const span = makeEl('SPAN', {}, btn);
  const outside = makeEl('DIV', {});
  env.fire('mouseover', { type: 'mouseover', target: span, relatedTarget: outside });
  env.fire('mouseover', { type: 'mouseover', target: btn, relatedTarget: span });   // 中で動いただけ
  env.fire('mouseout', { type: 'mouseout', target: span, relatedTarget: btn });     // 中で動いただけ
  env.fire('mouseout', { type: 'mouseout', target: btn, relatedTarget: outside });
  assert.deepEqual(calls, [['translateY(-1px)', 's'], ['', '']]);
});

test('__vpsArgs は属性から抜けられない形で、元の値に戻る（XSS の直し）', () => {
  const env = makeEnv();
  const evil = ["x');alert(1)//", '"><img src=x onerror=alert(1)>', "a&b<c>'d"];
  const attr = env.win.__vpsArgs(evil[0], evil[1], evil[2], 3);
  assert.ok(!/["<>']/.test(attr));
  const decoded = attr.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  assert.deepEqual(JSON.parse(decoded), evil.concat([3]));
});

test('__escText は数の 0 を空にせず、HTML として効かない形にする（XSS の直し）', () => {
  const env = makeEnv();
  const esc = env.win.__escText;
  assert.equal(esc('"><img src=x onerror=alert(1)>'), '&quot;&gt;&lt;img src=x onerror=alert(1)&gt;');
  assert.equal(esc("x');alert(1)//"), 'x&#39;);alert(1)//');
  assert.equal(esc('a&b'), 'a&amp;b');
  // 数は今の連結と同じ文字、null・undefined は空（仕様 3-8）
  assert.equal(esc(0), '0');
  assert.equal(esc(2.5), '2.5');
  assert.equal(esc(''), '');
  assert.equal(esc(null), '');
  assert.equal(esc(undefined), '');
});

test('__vpsAttrSel は " や \\ があっても [属性="…"] を壊さない', () => {
  const env = makeEnv();
  const sel = env.win.__vpsAttrSel;
  assert.equal(sel('LV-20261126-001'), 'LV-20261126-001');
  assert.equal(sel('a"]b'), 'a\\"]b');
  assert.equal(sel('a\\b'), 'a\\\\b');
});

test('google.script.run の関数は画面の 27 関数と運用の 7 関数。冪等キーは書き込みの関数だけ', () => {
  const env = makeEnv();
  const SPEC = require('./rpc-spec');
  const shim = env.win.__vpsShim;
  const names = JSON.parse(JSON.stringify(shim.READ_FNS.concat(shim.WRITE_FNS))).sort();
  assert.deepEqual(names, SPEC.ALL.slice().sort());
  assert.equal(SPEC.SCREEN.length, 27);
  assert.equal(SPEC.OPS.length, 7);
  assert.deepEqual(JSON.parse(JSON.stringify(shim.WRITE_FNS)).sort(), SPEC.WRITE.slice().sort());
});

test('書き込みの 5 関数（申請・編集・取消・1 次承認・2 次承認）に冪等キーが付く', async () => {
  const env = makeEnv();
  const r = env.win.google.script.run.withSuccessHandler(() => {}).withFailureHandler(() => {});
  const calls = [
    () => r.api_submitLeaveRequest({ leaveDate: '2026-11-30' }),
    () => r.api_editLeaveRequest('LV-20261126-001', { leaveDate: '2026-11-30', leaveKubun: '全日休', halfDayType: '' }),
    () => r.api_deleteLeaveRequest('LV-20261126-001'),
    () => r.api_approveLeaveRequest('LV-20261126-001', '', 'bucho@lineworks-local.info'),
    () => r.api_approveLeaveRequest2('LV-20261126-001', 'imaizumi@lineworks-local.info'),
    () => r.api_getLeaveRequestById('LV-20261126-001')
  ];
  calls.forEach((c) => { env.queue.push([200, { ok: true, result: { ok: true } }]); c(); });
  await flush();
  const keys = env.rpcFetches().map((f) => f.init.headers['X-Idempotency-Key']);
  assert.equal(keys.slice(0, 5).filter(Boolean).length, 5);
  assert.equal(keys[5], undefined, '読む関数には付けない');
});

test('申請の成功の中の {ok:false, error} はそのまま成功ハンドラへ（画面が alert を出す）', async () => {
  const env = makeEnv();
  const got = [];
  env.queue.push([200, { ok: true, result: { ok: false, error: '休暇日の形式が不正です（yyyy-MM-dd）。' } }]);
  env.win.google.script.run.withSuccessHandler((v) => got.push(v)).withFailureHandler(() => got.push('FAIL'))
    .api_submitLeaveRequest({ leaveDate: 'x' });
  await flush();
  assert.deepEqual(JSON.parse(JSON.stringify(got)), [{ ok: false, error: '休暇日の形式が不正です（yyyy-MM-dd）。' }]);
});
