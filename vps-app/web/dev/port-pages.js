#!/usr/bin/env node
// GAS の画面（src/*.html）から VPS 版の画面を機械的に作る（docs/vps-implementation-spec.md の 3 章）。
//
//   node vps-app/web/dev/port-pages.js          作り直して書き出す
//   node vps-app/web/dev/port-pages.js --check  書き出した物が src/ から作った物と同じかだけ確かめる（違えば終了コード 1）
//
// 作る物:
//   vps-app/api/app/templates/pages/<page>.html  … HTML と CSS は元のまま。テンプレートの差し込みを Jinja2 の {{ }} に、
//                                                   on 属性を data-on-<event> に、<script> を外部の .js の読み込みにする
//   vps-app/web/static/js/<page>.js              … 元の <script> の中身。変えるのは下の PATCHES に書いた所だけ
// 手で直さないこと（直すときはこのファイルか port-additions/ を直して作り直す）。
'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..', '..');          // リポジトリの直下
const SRC = path.join(ROOT, 'src');
const TPL_DIR = path.join(ROOT, 'vps-app', 'api', 'app', 'templates', 'pages');
const JS_DIR = path.join(ROOT, 'vps-app', 'web', 'static', 'js');
const ADD_DIR = path.join(__dirname, 'port-additions');

const PAGES = ['top', 'preview', 'sign', 'somuAdmin', 'deptAdmin', 'admin', 'leaveStatus', 'no_auth'];
// GAS の setTitle（src/menu.gs の doGet）。no_auth は 2 種類あるので pages.py が {{ title }} で渡す
const TITLES = {
  top: '休暇届 申請TOP', preview: '休暇届 プレビュー', sign: '休暇届 サイン', somuAdmin: '休暇届 総務管理',
  deptAdmin: '休暇届 部署管理', admin: '休暇届 管理', leaveStatus: '休暇届 取得状況', no_auth: '{{ title }}'
};

// ---- 置き換えの道具（数が合わなければ止める） ----
function replaceExact(text, from, to, count, label) {
  const n = text.split(from).length - 1;
  if (n !== count) { throw new Error(label + ': 「' + from.slice(0, 80) + '」が ' + n + ' 件（' + count + ' 件のはず）'); }
  return text.split(from).join(to);
}
function replaceRegex(text, re, fn, count, label) {
  let n = 0;
  const out = text.replace(re, function () { n++; return fn.apply(null, arguments); });
  if (n !== count) { throw new Error(label + ': ' + re + ' が ' + n + ' 件（' + count + ' 件のはず）'); }
  return out;
}
// [from, to, count] の並びを順に当てる
function apply(list, label) {
  return list.map((p, i) => (t) => replaceExact(t, p[0], p[1], p[2], label + ' #' + (i + 1)));
}

// ---- HTML の on 属性 → data-on-<event>・data-args ----
const ON_ATTR = /\son(click|input|change|mouseenter|mouseleave)="([^"]*)"/g;

function parseHandler(code) {
  const m = /^\s*([A-Za-z_$][\w$]*)\((.*)\)\s*;?\s*$/.exec(code);
  if (!m || /\bthis\b|\bdocument\b|\bevent\b|;/.test(m[2])) { throw new Error('機械的に置き換えられない on 属性: ' + code); }
  // 引数は文字・数・真偽値の並びだけ（src は信頼できる元なので、ここで評価してよい）
  const args = new Function('return [' + m[2] + '];')();
  return [m[1], args];
}
function attrJson(v) {
  return JSON.stringify(v).replace(/&/g, '&amp;').replace(/'/g, '&#39;');
}
function convertOnAttributes(html, label) {
  let total = 0;
  const out = html.replace(/<[A-Za-z][^>]*>/g, function (tag) {
    const found = [];
    tag.replace(ON_ATTR, function (all, ev, code) { found.push([ev, code]); return all; });
    if (found.length === 0) { return tag; }
    total += found.length;
    const single = found.length === 1;
    return tag.replace(ON_ATTR, function (all, ev, code) {
      const h = parseHandler(code);
      let s = ' data-on-' + ev + '="' + h[0] + '"';
      if (h[1].length > 0) { s += ' ' + (single ? 'data-args' : 'data-args-' + ev) + "='" + attrJson(h[1]) + "'"; }
      return s;
    });
  });
  if (/<[A-Za-z][^>]*\son[a-z]+\s*=/.test(out)) { throw new Error(label + ': on 属性が残っている'); }
  return { html: out, count: total };
}

// ---- JS の文字列で作る onclick（top の 3 個は同じ形） ----
// '...onclick="FN(\'' + EXPR + '\')"' → data-on-click="FN" data-args="' + __vpsArgs(EXPR) + '"
const JS_ONCLICK = /onclick="(\w+)\(\\'' \+ ([\w.]+) \+ '\\'\)"/g;
function jsOnclick(all, fn, expr) {
  return 'data-on-click="' + fn + '" data-args="\' + __vpsArgs(' + expr + ') + \'"';
}

// escHtml・esc・e（DOM の textContent→innerHTML で逃がす形）に引用符を足す。表示は同じ
const DOM_ESC_OLD = '  d.textContent = String(s);\n  return d.innerHTML;\n}';
const DOM_ESC_NEW = '  d.textContent = String(s);\n  return d.innerHTML.replace(/"/g, \'&quot;\').replace(/\'/g, \'&#39;\');\n}';

// somuAdmin・admin の有給の表の数と最終有給日（XSS の直し。表示は同じ）
const PAID_CELLS = [
  [String.raw`'<td>' + d.paidCount + (d.paidHalfCount > 0 ? ' <span style="font-size:11px;color:var(--muted);">(' + d.paidHalfCount + '半)</span>' : '') + '</td>'`,
    String.raw`'<td>' + __escText(d.paidCount) + (d.paidHalfCount > 0 ? ' <span style="font-size:11px;color:var(--muted);">(' + __escText(d.paidHalfCount) + '半)</span>' : '') + '</td>'`, 1],
  [String.raw`'<td>' + d.substituteCount + '</td>'`, String.raw`'<td>' + __escText(d.substituteCount) + '</td>'`, 1],
  [String.raw`'<td>' + d.specialCount + '</td>'`, String.raw`'<td>' + __escText(d.specialCount) + '</td>'`, 1],
  [String.raw`'<td style="font-weight:700;">' + d.totalCount + '</td>'`, String.raw`'<td style="font-weight:700;">' + __escText(d.totalCount) + '</td>'`, 1],
  [String.raw`'<td>' + (d.lastPaidDate || '-') + '</td>'`, String.raw`'<td>' + __escText(d.lastPaidDate || '-') + '</td>'`, 1]
];

// ---- ページごとの直し ----
const PATCHES = {
  top: {
    html: [],
    js: apply([
      // 申請完了・本日の休暇者・今後の休暇予定・承認待ちの 3 つの一覧（XSS の直し。表示は同じ）
      [String.raw`'<td>' + it.deptName + '</td>'`, String.raw`'<td>' + __escText(it.deptName) + '</td>'`, 3],
      [String.raw`'<td>' + it.workerName + '</td>'`, String.raw`'<td>' + __escText(it.workerName) + '</td>'`, 2],
      [String.raw`'<td style="font-weight:700;">' + it.workerName + '</td>'`, String.raw`'<td style="font-weight:700;">' + __escText(it.workerName) + '</td>'`, 1],
      [String.raw`'<td>' + it.leaveKubun + (it.halfDayType ? '(' + it.halfDayType + ')' : '') + '</td>'`,
        String.raw`'<td>' + __escText(it.leaveKubun) + (it.halfDayType ? '(' + __escText(it.halfDayType) + ')' : '') + '</td>'`, 3],
      [String.raw`'<td>' + it.leaveType + '</td>'`, String.raw`'<td>' + __escText(it.leaveType) + '</td>'`, 3],
      [String.raw`'">' + it.status + '</span></td>'`, String.raw`'">' + __escText(it.status) + '</span></td>'`, 2],
      [String.raw`'<td>' + it.leaveDate + '</td>'`, String.raw`'<td>' + __escText(it.leaveDate) + '</td>'`, 1],
      [String.raw`'<td>' + dateDisp + retroBadge + '</td>'`, String.raw`'<td>' + __escText(dateDisp) + retroBadge + '</td>'`, 1],
      [String.raw`(' + it.daysAgo + '日前)`, String.raw`(' + __escText(it.daysAgo) + '日前)`, 1],
      [String.raw`'<td style="font-size:12px;color:var(--muted);">' + (it.submittedAt || '') + '</td>'`,
        String.raw`'<td style="font-size:12px;color:var(--muted);">' + __escText(it.submittedAt || '') + '</td>'`, 1],
      // href の reqId は encodeURIComponent（LV-yyyyMMdd-NNN は変わらない）
      [String.raw`'?page=preview&reqId=' + it.reqId + '"`, String.raw`'?page=preview&reqId=' + encodeURIComponent(it.reqId) + '"`, 2],
      [String.raw`var previewUrl = APP_URL + '?page=preview&reqId=' + result.reqId;`,
        String.raw`var previewUrl = APP_URL + '?page=preview&reqId=' + encodeURIComponent(result.reqId);`, 1],
      [String.raw`REQ_ID: ' + result.reqId + '</div>'`, String.raw`REQ_ID: ' + __escText(result.reqId) + '</div>'`, 1],
      // 振替の月次期間の検証の枠と、振替元の選択肢の先頭（休暇日・振替元はサーバーの値）
      [String.raw`'休暇日 <strong>' + leaveDate + '</strong> は期間「' + lp.start + '〜' + lp.end + '」に属します。<br>' +`,
        String.raw`'休暇日 <strong>' + __escText(leaveDate) + '</strong> は期間「' + __escText(lp.start) + '〜' + __escText(lp.end) + '」に属します。<br>' +`, 1],
      [String.raw`'振替元 <strong>' + subDate + '</strong> は期間「' + sp.start + '〜' + sp.end + '」に属します。<br>' +`,
        String.raw`'振替元 <strong>' + __escText(subDate) + '</strong> は期間「' + __escText(sp.start) + '〜' + __escText(sp.end) + '」に属します。<br>' +`, 1],
      [String.raw`'✅ 月次期間OK（' + lp.start + '〜' + lp.end + '）'`, String.raw`'✅ 月次期間OK（' + __escText(lp.start) + '〜' + __escText(lp.end) + '）'`, 1],
      [String.raw`-- 期間内(' + period.start + '〜' + period.end + ')に候補なし --`,
        String.raw`-- 期間内(' + __escText(period.start) + '〜' + __escText(period.end) + ')に候補なし --`, 1],
      [String.raw`-- 休日出勤日を選択（' + period.start + '〜' + period.end + '） --`,
        String.raw`-- 休日出勤日を選択（' + __escText(period.start) + '〜' + __escText(period.end) + '） --`, 1]
    ], 'top').concat([
      // JS の文字列で作る onclick 3 個（編集・削除・過去申請カレンダーの日）
      (t) => replaceRegex(t, JS_ONCLICK, jsOnclick, 3, 'top onclick')
    ])
  },
  preview: {
    html: [],
    js: apply([[DOM_ESC_OLD, DOM_ESC_NEW, 1]], 'preview')
  },
  sign: {
    html: [],
    js: apply([
      [String.raw`'<option value="">取得エラー: ' + err.message + '</option>'`,
        String.raw`'<option value="">取得エラー: ' + __escText(err.message) + '</option>'`, 1]
      // 承認の後の <meta http-equiv="refresh"> を JS で足す動きは今のまま（仕様 3-6・3-10。CSP は止めない）
    ], 'sign')
  },
  somuAdmin: {
    html: [
      // 運用（管理者）の欄を一番下（有給取得義務 管理一覧の後ろ）に足す。運用者のときだけ（pages.py の SHOW_OPS）
      (h) => replaceExact(h,
        '      <div class="table-scroll" id="tableWrap"><div class="loading">読込中...</div></div>\n    </div>\n',
        '      <div class="table-scroll" id="tableWrap"><div class="loading">読込中...</div></div>\n    </div>\n' +
          fs.readFileSync(path.join(ADD_DIR, 'somuAdmin-ops.html'), 'utf8').replace(/\r\n/g, '\n').replace(/^\n/, ''),
        1, 'somuAdmin ops panel')
    ],
    js: apply([
      [DOM_ESC_OLD, DOM_ESC_NEW, 1],
      [String.raw`(' + d.daysAgo + '日前)`, String.raw`(' + __escText(d.daysAgo) + '日前)`, 1],
      [String.raw`'<td><input type="checkbox" class="row-check" data-idx="' + i + '" onchange="updateSelectedCount()"></td>'`,
        String.raw`'<td><input type="checkbox" class="row-check" data-idx="' + i + '" data-on-change="updateSelectedCount"></td>'`, 1],
      [String.raw`'<td><button class="btn-sm" onclick="singleApprove2(' + i + ', this)">2次承認</button></td>'`,
        String.raw`'<td><button class="btn-sm" data-on-click="singleApprove2" data-args="' + __vpsArgs(i, '$this') + '">2次承認</button></td>'`, 1],
      [String.raw`document.querySelector('tr[data-reqid="' + reqId + '"]')`,
        String.raw`document.querySelector('tr[data-reqid="' + __vpsAttrSel(reqId) + '"]')`, 2]
    ].concat(PAID_CELLS), 'somuAdmin')
  },
  deptAdmin: {
    html: [],
    js: apply([
      [DOM_ESC_OLD, DOM_ESC_NEW, 1],
      [String.raw`'<td>' + (d.paidCount || 0) + (d.paidHalfCount > 0 ? ' <span style="font-size:11px;color:var(--muted);">(' + d.paidHalfCount + '半)</span>' : '') + '</td>'`,
        String.raw`'<td>' + __escText(d.paidCount || 0) + (d.paidHalfCount > 0 ? ' <span style="font-size:11px;color:var(--muted);">(' + __escText(d.paidHalfCount) + '半)</span>' : '') + '</td>'`, 1],
      [String.raw`'<td>' + (d.substituteCount || 0) + '</td>'`, String.raw`'<td>' + __escText(d.substituteCount || 0) + '</td>'`, 1],
      [String.raw`'<td>' + (d.specialCount || 0) + '</td>'`, String.raw`'<td>' + __escText(d.specialCount || 0) + '</td>'`, 1],
      [String.raw`'<td style="font-weight:700;">' + (d.totalCount || 0) + '</td>'`, String.raw`'<td style="font-weight:700;">' + __escText(d.totalCount || 0) + '</td>'`, 1]
    ], 'deptAdmin')
  },
  admin: {
    html: [],
    js: apply([
      [DOM_ESC_OLD, DOM_ESC_NEW, 1],
      [String.raw`エラー: ' + err.message + '</div>'`, String.raw`エラー: ' + __escText(err.message) + '</div>'`, 2],
      [String.raw`">' + d.warnLevel + '</span>'`, String.raw`">' + __escText(d.warnLevel) + '</span>'`, 3]
    ].concat(PAID_CELLS), 'admin')
  },
  leaveStatus: {
    html: [],
    js: apply([
      [String.raw`function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }`,
        String.raw`function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }`, 1],
      // 本人ビュー・一覧の数と class 名（XSS の直し。表示は同じ）
      [String.raw`'<div class="empty">選択したスタッフの' + res.fyYear + '年度`, String.raw`'<div class="empty">選択したスタッフの' + esc(res.fyYear) + '年度`, 1],
      [String.raw`('うち半休 ' + w.paidHalfCount + '回')`, String.raw`('うち半休 ' + esc(w.paidHalfCount) + '回')`, 1],
      [String.raw`'残'+m.remaining : '不足'+m.deficit`, String.raw`'残'+esc(m.remaining) : '不足'+esc(m.deficit)`, 1],
      // class に入れる m.state は決まった 5 つ以外なら空（仕様 3-8）
      [String.raw`'<div class="node '+m.state+'" style=`,
        String.raw`'<div class="node '+(['done','risk','warn','note','future'].indexOf(m.state) >= 0 ? m.state : '')+'" style=`, 1],
      [String.raw`'<div class="circle">'+m.required+'</div>'`, String.raw`'<div class="circle">'+esc(m.required)+'</div>'`, 1],
      [String.raw`<b>'+def+'回</b>`, String.raw`<b>'+esc(def)+'回</b>`, 1],
      [String.raw`'（'+nextFuture.required+'回）`, String.raw`'（'+esc(nextFuture.required)+'回）`, 1],
      [String.raw`<b>'+nextFuture.required+'回</b>（あと <b>'+need+'回</b>）`, String.raw`<b>'+esc(nextFuture.required)+'回</b>（あと <b>'+esc(need)+'回</b>）`, 1],
      [String.raw`<td class="num tabnum">'+b.n+'</td>`, String.raw`<td class="num tabnum">'+esc(b.n)+'</td>`, 1],
      [String.raw`'　/　'+res.fyYear+'年度（3/16〜翌3/15）</div>`, String.raw`'　/　'+esc(res.fyYear)+'年度（3/16〜翌3/15）</div>`, 1],
      [String.raw`<div class="k-val tabnum">'+count+'<span class="slash"> / 5回</span></div>`, String.raw`<div class="k-val tabnum">'+esc(count)+'<span class="slash"> / 5回</span></div>`, 1],
      [String.raw`年度目標5回・残 '+rest+'回 ／ '+halfNote+'</div>`, String.raw`年度目標5回・残 '+esc(rest)+'回 ／ '+halfNote+'</div>`, 1],
      [String.raw`<div class="k-val tabnum">'+w.substituteCount+'<span`, String.raw`<div class="k-val tabnum">'+esc(w.substituteCount)+'<span`, 1],
      [String.raw`<div class="k-val tabnum">'+w.specialCount+'<span`, String.raw`<div class="k-val tabnum">'+esc(w.specialCount)+'<span`, 1],
      [String.raw`<div class="k-val tabnum">'+total+'<span`, String.raw`<div class="k-val tabnum">'+esc(total)+'<span`, 1],
      [String.raw`<span>'+res.fyYear+'年度の取得内訳</span>`, String.raw`<span>'+esc(res.fyYear)+'年度の取得内訳</span>`, 1],
      [String.raw`'<td><span class="tabnum">'+o.paidCount+' / 5</span>`, String.raw`'<td><span class="tabnum">'+esc(o.paidCount)+' / 5</span>`, 1],
      [String.raw`'<td class="num tabnum">'+o.substituteCount+'</td>'`, String.raw`'<td class="num tabnum">'+esc(o.substituteCount)+'</td>'`, 1],
      [String.raw`'<td class="num tabnum">'+o.specialCount+'</td>'`, String.raw`'<td class="num tabnum">'+esc(o.specialCount)+'</td>'`, 1],
      [String.raw`'<td class="num tabnum">'+o.totalCount+'</td>'`, String.raw`'<td class="num tabnum">'+esc(o.totalCount)+'</td>'`, 1],
      [String.raw`'名 ・ '+res.fyYear+'年度</div>`, String.raw`'名 ・ '+esc(res.fyYear)+'年度</div>`, 1]
    ], 'leaveStatus')
  },
  no_auth: { html: [], js: [] }
};

// JS の中のテンプレートの差し込み（GAS の <?= ?>）→ JSON の塊 __T から今の変数名で受ける
const JS_TEMPLATE_VARS = ['APP_URL', 'PORTAL_URL', 'REQ_ID'];
const TEMPLATE_DATA = "var __T = JSON.parse(document.getElementById('gas-template-data').textContent);\n";

function countOf(text, s) { return text.split(s).length - 1; }

function buildPage(page) {
  const src = fs.readFileSync(path.join(SRC, page + '.html'), 'utf8').replace(/\r\n/g, '\n');
  const patch = PATCHES[page];

  // <script> は各ページに 1 つだけ（最後）。no_auth には無い
  const scripts = src.match(/<script>[\s\S]*?<\/script>/g) || [];
  if (scripts.length > 1) { throw new Error(page + ': <script> が ' + scripts.length + ' 個'); }
  let html = src;
  let js = null;
  if (scripts.length === 1) {
    js = scripts[0].slice('<script>'.length, -'</script>'.length);
    html = html.replace(scripts[0], '\u0000SCRIPT\u0000');
  }

  // ---- HTML ----
  for (const f of patch.html) { html = f(html); }
  for (const v of ['APP_URL', 'PORTAL_URL', 'message']) {
    const from = '<?= ' + v + ' ?>';
    if (countOf(html, from)) { html = replaceExact(html, from, '{{ ' + v + ' }}', countOf(html, from), page + ' ' + v); }
  }
  if (/<\?/.test(html)) { throw new Error(page + ': GAS のテンプレートの書き方が残っている'); }
  const conv = convertOnAttributes(html, page);
  html = conv.html;
  // 題名（GAS の setTitle の代わり）。leaveStatus は HTML に同じ題名がある
  if (/<title>/.test(html)) {
    if (!html.includes('<title>' + TITLES[page] + '</title>')) { throw new Error(page + ': 題名が setTitle と違う'); }
  } else {
    html = replaceRegex(html, /(<meta name="viewport"[^>]*>\n)/, (all, m) => m + '  <title>' + TITLES[page] + '</title>\n', 1, page + ' title');
  }
  // スマホ向けの CSS（すべて 820px 以下の @media の中）
  html = replaceExact(html, '</head>', '<link rel="stylesheet" href="/static/css/mobile.css">\n</head>', 1, page + ' mobile.css');
  html = replaceExact(html, '<body>', '<body data-page="' + page + '">', 1, page + ' body');
  if (js !== null) {
    html = html.replace('\u0000SCRIPT\u0000',
      '<script type="application/json" id="gas-template-data">{{ data_json|safe }}</script>\n' +
      '<script src="/static/js/gas-shim.js"></script>\n' +
      '<script src="/static/js/' + page + '.js"></script>');
  }

  // ---- JS ----
  let usesTemplate = false;
  if (js !== null) {
    for (const v of JS_TEMPLATE_VARS) {
      const from = "var " + v + " = '<?= " + v + " ?>';";
      if (countOf(js, from)) {
        js = replaceExact(js, from, 'var ' + v + ' = __T.' + v + ';', 1, page + ' JS ' + v);
        usesTemplate = true;
      }
    }
    for (const f of patch.js) { js = f(js); }
    if (/<\?/.test(js)) { throw new Error(page + ': JS に GAS のテンプレートの書き方が残っている'); }
    if (/\son(click|input|change|mouseenter|mouseleave)=/.test(js)) { throw new Error(page + ': JS の文字列に on 属性が残っている'); }
    const head = '/* ' + page + '.js — GAS 版 src/' + page + '.html の <script> をそのまま写したもの。\n' +
      ' * vps-app/web/dev/port-pages.js が作る。手で直さない。 */\n' +
      (usesTemplate ? TEMPLATE_DATA : '');
    const addFile = path.join(ADD_DIR, page + '.js');
    const tail = fs.existsSync(addFile) ? fs.readFileSync(addFile, 'utf8').replace(/\r\n/g, '\n') : '';
    js = head + js.replace(/^\n/, '') + tail;
  }
  return { html: html, js: js, onAttrs: conv.count };
}

function main() {
  const check = process.argv.includes('--check');
  const diffs = [];
  let totalOn = 0;
  for (const page of PAGES) {
    const out = buildPage(page);
    totalOn += out.onAttrs;
    const targets = [[path.join(TPL_DIR, page + '.html'), out.html]];
    if (out.js !== null) { targets.push([path.join(JS_DIR, page + '.js'), out.js]); }
    for (const [file, text] of targets) {
      if (check) {
        const cur = fs.existsSync(file) ? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n') : null;
        if (cur !== text) { diffs.push(path.relative(ROOT, file)); }
      } else {
        fs.mkdirSync(path.dirname(file), { recursive: true });
        fs.writeFileSync(file, text);
      }
    }
  }
  if (check) {
    if (diffs.length) { console.error('src/ から作り直した物と違う: ' + diffs.join(', ')); process.exit(1); }
    console.log('ok: 書き出した画面は src/ から作った物と同じ（HTML の on 属性の置き換え ' + totalOn + ' 個）');
  } else {
    console.log('書き出した（HTML の on 属性の置き換え ' + totalOn + ' 個）');
  }
}

module.exports = { buildPage, PAGES, TITLES };
if (require.main === module) { main(); }
