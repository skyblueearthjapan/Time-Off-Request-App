// gas-tests/harness.js - src/*.gs を Node の vm で読み、GAS のサービスの偽物で動かす。
// src/ の外に置くので clasp は送らない。動かし方: node --test gas-tests/
'use strict';

process.env.TZ = 'Asia/Tokyo'; // GAS のスクリプトの時間帯（appsscript.json）と同じにする

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = path.resolve(__dirname, '..', 'src');
const TOKEN = 'tok-0123456789abcdef0123456789abcdef-XYZ';
const OWNER = 'owner@lineworks-local.info';
const LEDGER_ID = 'LEDGER1';
const PDF_ROOT = 'PDFROOT';
const DB_ID = 'DB1';

// 本番の T_LEAVE_REQUEST の見出し（英語 26 列。PDF_FILE_ID は無い。06 の 2-5 節）
const REQ_HEADER = ['REQ_ID', 'FY_START_YEAR', 'STATUS', 'APPLIED_AT', 'DEPT_ID', 'DEPT_NAME', 'WORKER_ID', 'WORKER_NAME',
  'LEAVE_DATE', 'DAY_TYPE', 'HALF_TYPE', 'LEAVE_KIND', 'SUBSTITUTE_FOR_WORK_DATE', 'SPECIAL_REASON_REF',
  'SPECIAL_REASON_TEXT', 'PAIDLEAVE_DETAIL', 'DETAIL_FREE', 'APPROVED_AT', 'SIGN_IMAGE_URL', 'PDF_URL',
  'CREATED_BY_EMAIL', 'UPDATED_AT', 'APPROVED_BY1_EMAIL', 'APPROVED_BY2', 'APPROVED_BY2_EMAIL', 'APPROVED_AT2'];
const LEDGER_HEADER = ['管理番号', '作業員ID', '部署', '氏名', '休暇日', '曜日', '休暇区分', '半日区分', '勤務時間帯',
  '休暇種類', '理由・詳細', '振替元出勤日', '申請日', '1次承認日', '1次承認者', '2次承認日', '2次承認者',
  '承認状態', '有給累計取得回数', 'PDF'];
const SUMMARY_HEADER = ['年度', '部署ID', '部署名', '作業員ID', '作業員名', '有給回数', '有給半日回数', '振替回数',
  '特別回数', '合計回数', '最終有給日', '警告レベル'];

function col(name) { return REQ_HEADER.indexOf(name); }

// JST の書式（Asia/Tokyo は夏時間なし）
function formatDate(d, tz, pattern) {
  const x = new Date(d.getTime() + 9 * 3600 * 1000);
  const p2 = n => String(n).padStart(2, '0');
  const map = { yyyy: String(x.getUTCFullYear()), MM: p2(x.getUTCMonth() + 1), dd: p2(x.getUTCDate()),
    HH: p2(x.getUTCHours()), mm: p2(x.getUTCMinutes()), ss: p2(x.getUTCSeconds()) };
  return pattern.replace(/'([^']*)'|yyyy|MM|dd|HH|mm|ss/g, (m, lit) => lit !== undefined ? lit : map[m]);
}

function iter(arr) { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; }

function makeWorld(opts) {
  opts = opts || {};
  const world = {
    now: opts.now || null,
    writes: [],      // [ss, sheet, kind, ...]
    mails: [],
    calls: [],
    logs: [],
    files: {},       // Drive のファイル
    folders: {},
    props: Object.assign({ VPS_INTERNAL_TOKEN: TOKEN }, opts.props || {}),
    cache: {},
    cacheThrows: false,
    lockBusy: false,
    ledgerDown: false,
    pdfFail: false,
    users: { active: OWNER, effective: OWNER },
    triggers: (opts.triggers || ['syncAllMasters']).map((fn, i) => ({ fn, id: 'T' + i, type: 'CLOCK' })),
    menus: [],
  };

  // ---- Date（new Date() の「今」を決められる）
  class FakeDate extends Date {
    constructor(...a) { if (a.length === 0) super(world.now === null ? Date.now() : world.now); else super(...a); }
    static now() { return world.now === null ? Date.now() : world.now; }
  }
  const D = FakeDate;
  const jst = s => new D(s + '+09:00');
  world.D = D;
  world.jst = jst;

  // ---- Sheets
  class FakeRange {
    constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
    getValues() {
      const out = [];
      for (let i = 0; i < this.nr; i++) {
        const row = [];
        for (let j = 0; j < this.nc; j++) {
          const rr = this.sheet.data[this.r - 1 + i];
          const v = rr ? rr[this.c - 1 + j] : undefined;
          row.push(v === undefined || v === null ? '' : v);
        }
        out.push(row);
      }
      return out;
    }
    getValue() { return this.getValues()[0][0]; }
    setValues(vals) {
      for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.set(this.r + i, this.c + j, vals[i][j]);
      this.sheet.log('setValues', this.r, this.c);
      return this;
    }
    setValue(v) { this.sheet.set(this.r, this.c, v); this.sheet.log('setValue', this.r, this.c, v); return this; }
    setFormula(f) { this.sheet.set(this.r, this.c, f); this.sheet.log('setFormula', this.r, this.c); return this; }
  }
  ['setNumberFormat', 'setFontWeight', 'setBackground', 'setFontColor', 'setFontSize', 'setHorizontalAlignment',
    'setVerticalAlignment', 'setWrap', 'setBorder'].forEach(m => { FakeRange.prototype[m] = function() { return this; }; });

  class FakeSheet {
    constructor(parent, name, data) { this.parent = parent; this.name = name; this.data = (data || []).map(r => r.slice()); this.id = Math.floor(Math.random() * 1e6); }
    log(kind, ...a) { world.writes.push([this.parent.id, this.name, kind, ...a]); }
    getName() { return this.name; }
    getParent() { return this.parent; }
    getSheetId() { return this.id; }
    getLastRow() { let n = this.data.length; while (n > 0 && (this.data[n - 1] || []).every(v => v === '' || v == null)) n--; return n; }
    getLastColumn() { return this.data.reduce((m, r) => Math.max(m, r.length), 0); }
    getMaxColumns() { return Math.max(26, this.getLastColumn()); }
    getMaxRows() { return Math.max(1000, this.data.length); }
    set(r, c, v) { while (this.data.length < r) this.data.push([]); const row = this.data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
    getRange(r, c, nr, nc) { return new FakeRange(this, r, c, nr || 1, nc || 1); }
    getDataRange() { return new FakeRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
    appendRow(row) { this.data.splice(this.getLastRow(), 0, row.slice()); this.log('appendRow'); }
    deleteRow(r) { this.data.splice(r - 1, 1); this.log('deleteRow', r); }
    setFrozenRows() {} setRowHeight() {} setColumnWidth() {} setName(n) { this.name = n; }
  }
  class FakeSpreadsheet {
    constructor(id, sheets) { this.id = id; this.order = []; this.sheets = {}; for (const k of Object.keys(sheets || {})) { this.sheets[k] = new FakeSheet(this, k, sheets[k]); this.order.push(k); } }
    getId() { return this.id; }
    getSpreadsheetTimeZone() { return ''; } // GMT の DB は空を返す（罠 68）
    getSheetByName(n) { return this.sheets[n] || null; }
    getSheets() { return this.order.map(n => this.sheets[n]); }
    insertSheet(n, i) { this.sheets[n] = new FakeSheet(this, n, []); this.order.splice(i || 0, 0, n); world.writes.push([this.id, n, 'insertSheet']); return this.sheets[n]; }
    deleteSheet(sh) { delete this.sheets[sh.name]; this.order = this.order.filter(n => n !== sh.name); world.writes.push([this.id, sh.name, 'deleteSheet']); }
  }

  // ---- 申請の行（英語 26 列）
  function reqRow(o) {
    const r = REQ_HEADER.map(() => '');
    for (const k of Object.keys(o)) r[col(k)] = o[k];
    return r;
  }
  const rows = [
    // 1 次承認の直後（VPS が行を反映済み）
    reqRow({ REQ_ID: 'LV-20261120-001', FY_START_YEAR: 2026, STATUS: '承認済', APPLIED_AT: jst('2026-11-20T09:00:00'),
      DEPT_ID: '加工課', DEPT_NAME: '加工課', WORKER_ID: '101', WORKER_NAME: '山田', LEAVE_DATE: jst('2026-12-01T00:00:00'),
      DAY_TYPE: '全日休', LEAVE_KIND: '有給消化（私用）', APPROVED_AT: jst('2026-11-20T10:00:00'),
      SIGN_IMAGE_URL: 'https://drive.google.com/file/d/OLDSIGN/view', CREATED_BY_EMAIL: 'a@lineworks-local.info', APPROVED_BY1_EMAIL: 'boss@lineworks-local.info' }),
    // 2 次承認の直後
    reqRow({ REQ_ID: 'LV-20261120-002', FY_START_YEAR: 2026, STATUS: '承認済', APPLIED_AT: jst('2026-11-20T09:05:00'),
      DEPT_ID: '加工課', DEPT_NAME: '加工課', WORKER_ID: '101', WORKER_NAME: '山田', LEAVE_DATE: jst('2026-12-02T00:00:00'),
      DAY_TYPE: '半日休', HALF_TYPE: '午前', LEAVE_KIND: '有給消化（私用）', APPROVED_AT: jst('2026-11-20T10:05:00'),
      PDF_URL: 'https://drive.google.com/file/d/OLDPDF/view', APPROVED_BY1_EMAIL: 'boss@lineworks-local.info',
      APPROVED_BY2: '総務', APPROVED_BY2_EMAIL: 'somu@lineworks-local.info', APPROVED_AT2: jst('2026-11-21T08:00:00') }),
    // 承認済みから編集された直後（休暇日は新しい日＝翌年度、承認の列は VPS が空にした）
    reqRow({ REQ_ID: 'LV-20261120-003', FY_START_YEAR: 2027, STATUS: '申請中', APPLIED_AT: jst('2026-11-20T09:10:00'),
      DEPT_ID: '組立課', DEPT_NAME: '組立課', WORKER_ID: '202', WORKER_NAME: '佐藤', LEAVE_DATE: jst('2027-04-10T00:00:00'),
      DAY_TYPE: '全日休', LEAVE_KIND: '振替休暇', SIGN_IMAGE_URL: 'https://drive.google.com/file/d/SIGN3/view',
      PDF_URL: 'https://drive.google.com/file/d/OLDPDF3/view' }),
    reqRow({ REQ_ID: 'LV-20261126-001', FY_START_YEAR: 2026, STATUS: '申請中', APPLIED_AT: jst('2026-11-26T10:12:00'),
      DEPT_ID: '加工課', DEPT_NAME: '加工課', WORKER_ID: '103', WORKER_NAME: '鈴木', LEAVE_DATE: jst('2026-12-10T00:00:00'),
      DAY_TYPE: '全日休', LEAVE_KIND: '有給消化（私用）' }),
  ];
  const db = new FakeSpreadsheet(DB_ID, {
    M_SYSTEM_SETTING: [['KEY', 'VALUE', 'MEMO'],
      ['APP_URL', 'https://script.google.com/macros/s/APPDEPLOY/exec'], ['PORTAL_URL', ''],
      ['ADMIN_EMAILS', OWNER], ['SOMU_EMAILS', 'somu@lineworks-local.info'], ['ALLOW_ADMIN_PAGE_EMAILS', OWNER],
      ['MAIL_TO', opts.mailTo === undefined ? 'somu@lineworks-local.info' : opts.mailTo],
      ['PDF_FOLDER_ID', PDF_ROOT], ['MASTER_SOURCE_SSID', 'MASTERSRC'], ['CALENDAR_SOURCE_SSID', 'CALSRC'],
      ['LEDGER_SSID', LEDGER_ID]],
    T_LEAVE_REQUEST: [REQ_HEADER].concat(rows),
    V_LEAVE_SUMMARY: [SUMMARY_HEADER,
      [2026, '加工課', '加工課', '101', '山田', 0, 0, 0, 0, 0, '', '警告'],
      [2026, '組立課', '組立課', '202', '佐藤', 3, 0, 1, 0, 4, '2026-09-01', '']],
    M_STAMP: [['メール', 'stampFileId', '備考'], ['somu@lineworks-local.info', 'STAMP1', '総務']],
    M_LOOKUP: [['カテゴリ', '値', 'SORT'], ['特別休暇理由', 'その他', 99]],
    M_CALENDAR: [['日付', '区分', '曜日', '備考']],
    M_DEPT_APPROVERS: [['dept', 'approverEmails', 'approverName']],
    '作業員マスタ': [['作業員コード', '氏名', '部署'], ['101', '山田', '加工課']],
  });
  const ledgerHeader = opts.badLedgerHeader ? LEDGER_HEADER.map((h, i) => (i === 5 ? '曜' : h)) : LEDGER_HEADER;
  const ledger = new FakeSpreadsheet(LEDGER_ID, {
    '2026年度': [ledgerHeader, ['LV-20261120-003', '202', '組立課', '佐藤', '2026/12/03']],
  });
  world.db = db;
  world.ledger = ledger;
  world.reqSheet = db.sheets.T_LEAVE_REQUEST;
  world.cell = (reqId, name) => {
    const r = world.reqSheet.data.find(x => x[0] === reqId);
    return r ? r[col(name)] : undefined;
  };
  world.ledgerIds = tab => {
    const sh = ledger.sheets[tab];
    return sh ? sh.data.slice(1).map(r => r[0]) : null;
  };

  // ---- Drive
  function makeFolder(id, name, parent) {
    const f = { id, name, parent, children: [],
      getId: () => id, getName: () => name,
      getParents: () => iter(parent ? [parent] : []),
      getFoldersByName: n => iter(f.children.filter(c => c.name === n)),
      createFolder: n => { const c = makeFolder('FOLDER-' + n, n, f); f.children.push(c); world.calls.push(['createFolder', n]); return c; },
      createFile: blob => makeFile('FILE' + (Object.keys(world.files).length + 1), blob.name, blob.type || 'application/pdf', f),
      addFile: () => {},
    };
    world.folders[id] = f;
    return f;
  }
  function makeFile(id, name, mime, parent) {
    const file = { id, name, mime, parent, trashed: false,
      getId: () => id, getName: () => file.name, getUrl: () => 'https://drive.google.com/file/d/' + id + '/view',
      getMimeType: () => file.mime, getParents: () => iter([parent]),
      setName: n => { file.name = n; return file; },
      setTrashed: v => { file.trashed = v; world.calls.push(['setTrashed', id]); },
      getBlob: () => ({}),
    };
    world.files[id] = file;
    world.calls.push(['createFile', id, name]);
    return file;
  }
  const root = makeFolder(PDF_ROOT, 'PDF', null);
  const d1202 = root.createFolder('2026.12.02');
  makeFile('OLDPDF', 'old.pdf', 'application/pdf', d1202);
  makeFile('OLDPDF3', 'old3.pdf', 'application/pdf', d1202);
  const elsewhere = makeFolder('OTHERROOT', 'other', null);
  makeFile('FOREIGNPDF', 'prod.pdf', 'application/pdf', elsewhere);
  world.calls.length = 0;

  const ctx = {
    console: { log: (...a) => world.logs.push(a.join(' ')), warn: (...a) => world.logs.push(a.join(' ')), error: (...a) => world.logs.push(a.join(' ')) },
    Logger: { log: m => world.logs.push(String(m)) },
    JSON, Math, Date: D, Object, Array, String, Number, Boolean, RegExp, Error, Map, Set, isNaN, isFinite, parseInt, parseFloat, encodeURIComponent, decodeURIComponent,
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (k in world.props ? world.props[k] : null),
      setProperty: (k, v) => { world.props[k] = String(v); world.calls.push(['setProperty', k]); },
      deleteProperty: k => { delete world.props[k]; world.calls.push(['deleteProperty', k]); },
      getProperties: () => Object.assign({}, world.props),
    }) },
    CacheService: { getScriptCache: () => ({
      get: k => { if (world.cacheThrows) throw new Error('cache down'); return k in world.cache ? world.cache[k] : null; },
      put: (k, v) => { if (world.cacheThrows) throw new Error('cache down'); world.cache[k] = v; },
    }) },
    LockService: { getScriptLock: () => ({
      tryLock: () => { world.calls.push(['tryLock']); return !world.lockBusy; },
      waitLock: () => { world.calls.push(['waitLock']); if (world.lockBusy) throw new Error('Lock timeout'); },
      releaseLock: () => { world.calls.push(['releaseLock']); },
      hasLock: () => true,
    }) },
    SpreadsheetApp: {
      getActive: () => { world.calls.push(['getActive']); return db; },
      openById: id => {
        world.calls.push(['openById', id]);
        if (id === LEDGER_ID && !world.ledgerDown) return ledger;
        throw new Error('not found ' + id);
      },
      create: name => { world.calls.push(['SpreadsheetApp.create', name]); throw new Error('create must not be called in tests'); },
      flush: () => {},
      getUi: () => ({
        createMenu: title => { const m = { title, items: [], addItem: (a, b) => { m.items.push(b); return m; }, addSeparator: () => m, addToUi: () => { world.menus.push(m); } }; return m; },
        alert: () => {},
      }),
    },
    DriveApp: {
      getFolderById: id => { world.calls.push(['getFolderById', id]); if (!world.folders[id]) throw new Error('no folder ' + id); return world.folders[id]; },
      getFileById: id => { world.calls.push(['getFileById', id]); if (!world.files[id]) throw new Error('no file ' + id); return world.files[id]; },
      getRootFolder: () => ({ removeFile: () => {} }),
    },
    GmailApp: { sendEmail: (to, subject, body) => { world.mails.push({ to, subject, body }); world.calls.push(['sendEmail', subject]); } },
    MailApp: { sendEmail: () => { throw new Error('MailApp not used'); } },
    UrlFetchApp: { fetch: (url) => { world.calls.push(['fetch', url]); throw new Error('no network in tests'); } },
    Utilities: {
      formatDate, sleep: () => {}, getUuid: () => require('crypto').randomUUID(),
      base64Decode: s => Array.from(Buffer.from(s, 'base64')),
      newBlob: (bytes, type, name) => ({ bytes, type, name, getName: () => name }),
    },
    Session: {
      getActiveUser: () => ({ getEmail: () => world.users.active }),
      getEffectiveUser: () => ({ getEmail: () => world.users.effective }),
      getScriptTimeZone: () => 'Asia/Tokyo',
    },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ content: t, setMimeType() { return this; }, getContent() { return this.content; } }) },
    HtmlService: {
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
      createHtmlOutput: h => ({ html: h, setTitle(t) { this.title = t; return this; }, setXFrameOptionsMode() { return this; }, getContent() { return this.html; } }),
      createTemplateFromFile: name => ({ evaluate() { const self = this; return { html: 'TEMPLATE:' + name, vars: self, setTitle(t) { this.title = t; return this; }, setXFrameOptionsMode() { return this; }, getContent() { return this.html; } }; } }),
      createHtmlOutputFromFile: () => ({ getContent: () => '' }),
    },
    ScriptApp: {
      EventType: { CLOCK: 'CLOCK' },
      getScriptId: () => 'SCRIPT1',
      getOAuthToken: () => 'oauth',
      getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/' + (world.serviceDeploymentId || 'APPDEPLOY') + '/exec' }),
      getProjectTriggers: () => world.triggers.map(t => ({ getHandlerFunction: () => t.fn, getEventType: () => t.type, getUniqueId: () => t.id, _t: t })),
      deleteTrigger: tr => { world.triggers = world.triggers.filter(t => t !== tr._t); world.calls.push(['deleteTrigger', tr._t.fn]); },
      newTrigger: fn => {
        const b = { s: {}, timeBased: () => b, everyDays: () => b, atHour: h => { b.s.hour = h; return b; }, nearMinute: m => { b.s.minute = m; return b; },
          create: () => { world.triggers.push({ fn, id: 'N' + world.triggers.length, type: 'CLOCK', s: b.s }); world.calls.push(['newTrigger', fn, b.s]); } };
        return b;
      },
    },
  };
  vm.createContext(ctx);
  for (const f of fs.readdirSync(SRC).filter(f => f.endsWith('.gs')).sort()) {
    vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f });
  }

  // PDF は重い（一時シート・export）ので偽物にする。呼ばれた順を残す。
  ctx.generateLeavePdf_ = (reqId, a1, a2) => {
    world.calls.push(['generateLeavePdf_', reqId, a1, a2]);
    if (world.pdfFail) throw new Error('PDFエクスポートに失敗しました: HTTP 500');
    const folder = world.folders['FOLDER-2026.12.02'] || root.createFolder('2026.12.02');
    const name = '20261202_加工課_山田_休暇届.pdf';
    const f = folder.createFile({ name, type: 'application/pdf' });
    return { ok: true, pdfFileId: f.getId(), pdfUrl: f.getUrl(), pdfName: name, folderId: folder.getId() };
  };
  // 集計の書き直しは内部用の入口から呼ばれないこと
  const realRebuild = ctx.rebuildLeaveSummary_;
  ctx.rebuildLeaveSummary_ = () => { world.calls.push(['rebuildLeaveSummary_']); return realRebuild(); };

  world.ctx = ctx;
  world.post = body => JSON.parse(ctx.doPost({ postData: { contents: typeof body === 'string' ? body : JSON.stringify(body) } }).getContent());
  world.call = (action, extra) => world.post(Object.assign({ internal: true, token: TOKEN, action, jobId: 'job-' + action }, extra || {}));
  world.count = name => world.calls.filter(c => c[0] === name).length;
  world.migrate = url => { world.props.MIGRATED_TO_URL = url || 'https://leave.lineworks-local.net'; ctx.invalidateMigrationCache_(); };
  world.dbWrites = () => world.writes.filter(w => w[0] === DB_ID);
  return world;
}

// 公開関数（末尾が _ でない function）の一覧。google.script.run から呼べるもの
function publicFunctions() {
  const out = [];
  for (const f of fs.readdirSync(SRC).filter(f => f.endsWith('.gs')).sort()) {
    const text = fs.readFileSync(path.join(SRC, f), 'utf8');
    const re = /^function ([A-Za-z0-9_$]*[A-Za-z0-9$])\(/gm;
    let m;
    while ((m = re.exec(text))) out.push({ file: f, name: m[1] });
  }
  return out;
}

module.exports = { makeWorld, publicFunctions, TOKEN, OWNER, REQ_HEADER, LEDGER_HEADER, SRC, col };
