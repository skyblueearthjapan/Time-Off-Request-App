// 7 ページの JS の試験（on 属性の置き換え・呼び出しの行・XSS の直し）。jsdom は使わず、何でも受け付ける偽物の DOM で読み込む。
//   node --test vps-app/web/dev/
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { spawnSync } = require('child_process');
const { webcrypto } = require('crypto');
const SPEC = require('./rpc-spec');
const { RPC } = require('./mock-server');

const VPS = path.resolve(__dirname, '..', '..');
const ROOT = path.resolve(VPS, '..');
const JS_DIR = path.join(VPS, 'web', 'static', 'js');
const TPL_DIR = path.join(VPS, 'api', 'app', 'templates', 'pages');
const PAGES = ['top', 'preview', 'sign', 'somuAdmin', 'deptAdmin', 'admin', 'leaveStatus'];
const SHIM = fs.readFileSync(path.join(JS_DIR, 'gas-shim.js'), 'utf8');
const EVIL1 = '"><img src=x onerror=alert(1)>';
const EVIL2 = "x');alert(1)//";

function srcScript(page) {
  const src = fs.readFileSync(path.join(ROOT, 'src', page + '.html'), 'utf8').replace(/\r\n/g, '\n');
  const m = /<script>([\s\S]*?)<\/script>/.exec(src);
  return m ? m[1] : '';
}
// VPS 版で足した部分より前（GAS 版から写した部分）
function portedPart(page) {
  const js = fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8');
  const i = js.indexOf('VPS 版で足した部分（ここから下だけ');
  return i < 0 ? js : js.slice(0, i);
}

// 何を聞かれても「それらしい空の物」を返す偽物（関数としても呼べる）。
// textContent を入れると、ブラウザと同じく innerHTML は & < > を逃がした文字になる（escHtml・e() の形のため）
function browserEscape(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function makeAny(log) {
  const store = {};
  const target = function () {};
  const proxy = new Proxy(target, {
    get(t, k) {
      if (Object.prototype.hasOwnProperty.call(store, k)) { return store[k]; }
      if (k === Symbol.toPrimitive) { return (hint) => (hint === 'number' ? 0 : ''); }
      if (k === Symbol.iterator) { return function* () {}; }
      if (k === 'then') { return undefined; }
      if (k === 'length') { return 0; }
      if (k === 'options' || k === 'children' || k === 'childNodes') { return []; }
      if (['value', 'textContent', 'innerHTML', 'innerText', 'className', 'id'].includes(k)) { return ''; }
      if (k === 'checked' || k === 'disabled' || k === 'open') { return false; }
      return makeAny(log);
    },
    set(t, k, v) {
      store[k] = v;
      if (k === 'textContent') { store.innerHTML = browserEscape(v); }
      if (k === 'innerHTML' && log) { log.push(String(v)); }
      return true;
    },
    apply() { return makeAny(log); },
    construct() { return makeAny(log); }
  });
  return proxy;
}

// ページを読み込み、DOMContentLoaded を起こす。RPC は見本サーバーの関数（overrides で差し替え可）で答える
function loadPage(page, opts) {
  opts = opts || {};
  const calls = [];
  const html = [];        // innerHTML に入れた文字の全部
  const els = {};
  const listeners = [];
  const document = makeAny(html);
  const docStore = {
    readyState: 'loading',
    getElementById(id) {
      if (id === 'gas-template-data') { return { textContent: JSON.stringify(Object.assign({ PAGE: page, APP_URL: 'https://leave.example/', PORTAL_URL: '', REQ_ID: '' }, opts.data || {})) }; }
      if (!els[id]) { els[id] = makeAny(html); }
      return els[id];
    },
    addEventListener(t, f) { listeners.push([t, f]); },
    querySelector() { return makeAny(html); },
    querySelectorAll() { return []; },
    createElement() { return makeAny(html); },
    contains() { return true; }
  };
  const doc = new Proxy(document, {
    get(t, k) { return Object.prototype.hasOwnProperty.call(docStore, k) ? docStore[k] : t[k]; }
  });
  function response(status, body) {
    const text = JSON.stringify(body);
    return { status: status, text: () => Promise.resolve(text), json: () => Promise.resolve(JSON.parse(text)) };
  }
  const win = {
    document: doc, location: { pathname: '/', search: '?page=' + page, href: '' },
    alert() {}, confirm() { return false; }, prompt() { return null; },
    setTimeout: (f, ms) => setTimeout(f, Math.min(ms || 0, 5)), clearTimeout() {}, setInterval() { return 0; }, clearInterval() {},
    addEventListener(t, f) { listeners.push([t, f]); }, removeEventListener() {}, scrollTo() {},
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    crypto: webcrypto, devicePixelRatio: 1, console: { log() {}, error() {}, warn() {} },
    URL: { createObjectURL() { return ''; }, revokeObjectURL() {} }, Blob: function () {}, Image: function () {},
    fetch(url, init) {
      if (url === '/auth/me') { return Promise.resolve(response(200, { email: 'imaizumi@lineworks-local.info' })); }
      const name = url.slice('/api/rpc/'.length);
      const args = JSON.parse(init.body);
      calls.push({ name: name, args: args, key: init.headers['X-Idempotency-Key'] });
      const fn = (opts.overrides || {})[name] || RPC[name];
      if (!fn) { return Promise.resolve(response(501, { ok: false, error: '未登録の関数です: ' + name })); }
      try {
        const r = fn.apply(null, args);
        return Promise.resolve(response(200, { ok: true, result: r === undefined ? null : r }));
      } catch (e) {
        return Promise.resolve(response(400, { ok: false, error: e.message }));
      }
    }
  };
  win.window = win;
  vm.createContext(win);
  vm.runInContext(SHIM, win, { filename: 'gas-shim.js' });
  vm.runInContext(fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8'), win, { filename: page + '.js' });
  docStore.readyState = 'complete';
  listeners.filter((l) => l[0] === 'DOMContentLoaded').forEach((l) => l[1]({ type: 'DOMContentLoaded' }));
  return { win, calls, html, els };
}
const flush = (n) => new Promise((r) => setTimeout(r, n || 60));

test('画面は src/ から port-pages.js で作った物と同じ（手で直していない）', () => {
  const r = spawnSync(process.execPath, [path.join(__dirname, 'port-pages.js'), '--check'], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stderr);
  assert.match(r.stdout, /on 属性の置き換え 37 個/);
});

test('呼び出し 37 か所は GAS 版と 1 字も変わらない（画面ごとの数も同じ）', () => {
  let total = 0;
  for (const page of PAGES) {
    // 呼び出しの行＝「.api_xxx(引数)」を含む行。並びも含めて同じ
    const pick = (js) => js.split('\n').filter((l) => /\.api_\w+\(/.test(l));
    const before = pick(srcScript(page));
    const after = pick(portedPart(page));
    assert.deepEqual(after, before, page);
    const n = (srcScript(page).match(/\.api_\w+\(/g) || []).length;
    assert.equal(n, SPEC.CALLS_PER_PAGE[page], page);
    total += n;
  }
  assert.equal(total, 37);
  // 呼び出しの行は google.script.run のまま（シムが受ける）
  for (const page of PAGES) { assert.ok(!/fetch\(/.test(portedPart(page)), page); }
});

test('on 属性が HTML にも JS の文字にも 0 個、インラインの script は JSON の塊だけ', () => {
  let dataOn = 0;
  for (const page of PAGES.concat(['no_auth'])) {
    const html = fs.readFileSync(path.join(TPL_DIR, page + '.html'), 'utf8');
    assert.ok(!/<[A-Za-z][^>]*\son[a-z]+\s*=/i.test(html), page);
    const scripts = html.match(/<script\b[^>]*>[\s\S]*?<\/script>/g) || [];
    for (const s of scripts) {
      assert.ok(/^<script type="application\/json" id="gas-template-data">\{\{ data_json\|safe \}\}<\/script>$/.test(s) ||
        /^<script src="\/static\/js\/[\w-]+\.js"><\/script>$/.test(s), page + ': ' + s);
    }
    dataOn += (html.match(/data-on-(click|change)=/g) || []).length;
    if (page !== 'no_auth') {
      const js = fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8');
      assert.ok(!/\son(click|change|input|mouseenter|mouseleave)\s*=/.test(js), page);
      assert.ok(!/<\?/.test(js), page);
    }
  }
  // 静的な 37 個＋運用の欄の 6 個（somuAdmin）
  assert.equal(dataOn, 37 + 6);
});

for (const page of PAGES) {
  test(page + '.js: 構文の誤りが無く、読み込みで呼ぶ RPC はすべて閉じたリストの中', async () => {
    const code = fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8');
    assert.doesNotThrow(() => new vm.Script(code, { filename: page + '.js' }));
    const { calls } = loadPage(page, { data: { REQ_ID: 'LV-20261126-001' } });
    await flush();
    assert.ok(calls.length > 0, page);
    for (const c of calls) { assert.ok(SPEC.ALL.includes(c.name), c.name); }
  });

  test(page + '.js: data-on-* で呼ぶ関数がすべてある', () => {
    const { win } = loadPage(page);
    const html = fs.readFileSync(path.join(TPL_DIR, page + '.html'), 'utf8');
    const js = fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8');
    const names = new Set();
    for (const m of (html + '\n' + js).matchAll(/data-on-(?:click|input|change|mouseenter|mouseleave)=\\?"(\w+)/g)) { names.add(m[1]); }
    for (const n of names) { assert.equal(vm.runInContext('typeof ' + n, win), 'function', n); }
  });
}

test('テンプレートの値が今の変数名で受け取られる', () => {
  const data = { APP_URL: 'https://leave.example/', PORTAL_URL: 'https://portal.example/exec', REQ_ID: 'LV-20261126-002' };
  for (const page of ['top', 'somuAdmin', 'deptAdmin', 'admin']) {
    const { win } = loadPage(page, { data: data });
    assert.deepEqual(JSON.parse(JSON.stringify(vm.runInContext('[APP_URL, PORTAL_URL]', win))), [data.APP_URL, data.PORTAL_URL], page);
  }
  for (const page of ['preview', 'sign']) {
    const { win, calls } = loadPage(page, { data: data });
    assert.deepEqual(JSON.parse(JSON.stringify(vm.runInContext('[REQ_ID, APP_URL]', win))), [data.REQ_ID, data.APP_URL], page);
    if (page === 'preview') { assert.ok(calls.some((c) => c.name === 'api_getLeaveRequestById')); }
  }
});

test('preview は REQ_ID が空なら呼ばない（「REQ_IDが指定されていません。」のまま）', async () => {
  const { calls, els } = loadPage('preview', { data: { REQ_ID: '' } });
  await flush();
  assert.equal(calls.filter((c) => c.name === 'api_getLeaveRequestById').length, 0);
  assert.equal(els.loadingMsg.textContent, 'REQ_IDが指定されていません。');
});

test('preview は存在しない reqId で「申請データが見つかりません。（REQ_ID: …）」', async () => {
  const { els } = loadPage('preview', { data: { REQ_ID: 'LV-20000101-999' } });
  await flush();
  assert.equal(els.loadingMsg.textContent, '申請データが見つかりません。（REQ_ID: LV-20000101-999）');
});

// ---- XSS の直し: 氏名が 2 つの悪い文字の申請が、どの画面でも文字のまま見える ----

function assertSafe(all, label, expectEvil1, expectEvil2) {
  const text = all.join('\n');
  assert.ok(!/<img/i.test(text), label + ': <img が HTML として入った');
  assert.ok(!text.includes("x');alert"), label + ': 引用符が逃がされていない');
  assert.ok(!/\son[a-z]+=/i.test(text.replace(/&lt;img src=x onerror=alert\(1\)&gt;/g, '')), label + ': on 属性が入った');
  if (expectEvil1) { assert.ok(text.includes('&quot;&gt;&lt;img src=x onerror=alert(1)&gt;'), label + ': EVIL1 が文字のまま見えない'); }
  if (expectEvil2) { assert.ok(text.includes('x&#39;);alert(1)//'), label + ': EVIL2 が文字のまま見えない'); }
}

test('XSS: top の 3 つの一覧（承認待ち・本日の休暇者・今後の休暇予定）', async () => {
  const { els } = loadPage('top');
  await flush(120);
  for (const id of ['pendingWrap', 'todayLeaveWrap', 'leaveTableWrap']) {
    assertSafe([els[id].innerHTML], 'top ' + id, true, id !== 'todayLeaveWrap');
  }
});

test('XSS: top の申請 ID・日付がサーバーから悪い文字で来ても属性から抜けない', async () => {
  const evilItem = { reqId: "LV-1'\"><img src=x onerror=alert(1)>", deptName: EVIL1, workerName: EVIL2, leaveKubun: '半日休', halfDayType: EVIL1,
    leaveDate: '2999/01/01', leaveType: EVIL1, status: '申請中', isRetroactive: true, daysAgo: EVIL1, submittedAt: EVIL1 };
  const { els } = loadPage('top', { overrides: {
    api_getUpcomingLeaves: () => [evilItem], api_getAllPendingRequests: () => [Object.assign({}, evilItem, { leaveDate: EVIL1 })]
  } });
  await flush(120);
  assertSafe([els.leaveTableWrap.innerHTML, els.pendingWrap.innerHTML], 'top evil reqId', true, true);
  // 編集・削除のボタンは data-args（JSON を属性として逃がしたもの）で reqId を渡す
  assert.ok(els.leaveTableWrap.innerHTML.includes('data-on-click="openEditLeaveModal" data-args="[&quot;LV-1&#39;\\&quot;&gt;&lt;img'));
});

test('XSS: preview（悪い氏名の申請）', async () => {
  const evil = RPC.api_getUpcomingLeaves('製造2課', '2002')[0];
  const { html } = loadPage('preview', { data: { REQ_ID: evil.reqId } });
  await flush();
  assertSafe(html.filter((h) => h.includes('paper')), 'preview', true, false);
});

test('XSS: sign の承認者一覧の取得エラーの文', async () => {
  const { els } = loadPage('sign', { data: { REQ_ID: 'LV-20261126-001' }, overrides: { api_getApproverList: () => { throw new Error(EVIL1 + EVIL2); } } });
  await flush();
  assertSafe([els.approverSelect.innerHTML], 'sign', true, true);
  assert.ok(els.approverSelect.innerHTML.startsWith('<option value="">取得エラー: '));
});

test('XSS: somuAdmin の 2 次承認待ちと有給の表（最終有給日・回数も）', async () => {
  const { els } = loadPage('somuAdmin');
  await flush(120);
  assertSafe([els.approval2Wrap.innerHTML], 'somuAdmin 2次承認', true, true);
  assertSafe([els.tableWrap.innerHTML], 'somuAdmin 有給', true, true);
  assert.ok(els.approval2Wrap.innerHTML.includes('data-on-click="singleApprove2" data-args="[0,&quot;$this&quot;]"'));
  assert.ok(els.approval2Wrap.innerHTML.includes('data-on-change="updateSelectedCount"'));
  const { els: e2 } = loadPage('somuAdmin', { overrides: {
    api_getSomuAdminView: () => [{ deptName: 'a', workerName: 'b', paidCount: EVIL1, paidHalfCount: 1, substituteCount: EVIL1, specialCount: EVIL1,
      totalCount: EVIL1, lastPaidDate: EVIL1 + EVIL2, warnLevel: '危険', quarters: [] }]
  } });
  await flush(120);
  assertSafe([e2.tableWrap.innerHTML], 'somuAdmin 数と最終有給日', true, true);
});

test('XSS: deptAdmin の有給の表', async () => {
  const { win, els } = loadPage('deptAdmin');
  await flush();
  win.selectApprover('全部署', ['製造1課', '製造2課']);
  await flush();
  assertSafe([els.tableWrap.innerHTML], 'deptAdmin', true, true);
});

test('XSS: admin の表（最終有給日・警告レベル）とエラーの文', async () => {
  const { els } = loadPage('admin');
  await flush(120);
  assertSafe([els.tableWrap.innerHTML], 'admin', true, true);
  const row = { deptName: 'a', workerName: 'b', paidCount: 1, paidHalfCount: 0, substituteCount: 0, specialCount: 0, totalCount: 1,
    lastPaidDate: EVIL1 + EVIL2, warnLevel: '危険' };
  const { els: e2 } = loadPage('admin', { overrides: { api_getAdminBossView: () => [row] } });
  await flush(120);
  assertSafe([e2.tableWrap.innerHTML], 'admin 最終有給日', true, true);
  const { els: e3 } = loadPage('admin', { overrides: { api_getAdminBossView: () => { throw new Error(EVIL1 + EVIL2); } } });
  await flush(120);
  assertSafe([e3.tableWrap.innerHTML], 'admin エラー', true, true);
});

test('XSS: leaveStatus の一覧と本人ビュー', async () => {
  const { win, els } = loadPage('leaveStatus');
  await flush(120);
  assertSafe([els.body.innerHTML], 'leaveStatus 一覧', true, true);
  win.goToPerson('製造2課', '2002');
  await flush(120);
  assertSafe([els.body.innerHTML], 'leaveStatus 本人', true, false);
  assert.ok(els.body.innerHTML.includes('class="meHead"'));
});

test('sign: 承認の後の 1 秒の meta refresh は今のまま（仕様 3-10）', () => {
  const js = fs.readFileSync(path.join(JS_DIR, 'sign.js'), 'utf8');
  assert.ok(js.includes("meta.httpEquiv = 'refresh';\n        meta.content = '1;url=' + previewUrl;"));
});

test('画面が呼ぶ RPC の名前は閉じたリストの 34 関数だけで、全部どこかから呼ばれる', () => {
  const used = new Set();
  for (const page of PAGES) {
    const js = fs.readFileSync(path.join(JS_DIR, page + '.js'), 'utf8');
    for (const m of js.matchAll(/\.((?:api|ops)_\w+)\s*\(/g)) { used.add(m[1]); }
  }
  assert.deepEqual([...used].filter((n) => !SPEC.ALL.includes(n)), [], '閉じたリストに無い関数');
  assert.deepEqual(SPEC.ALL.filter((n) => !used.has(n)), [], '画面から呼ばれない関数');
});
