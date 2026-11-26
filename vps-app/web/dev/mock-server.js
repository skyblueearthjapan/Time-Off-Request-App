#!/usr/bin/env node
// 画面（7 ページ＋権限エラー）を api 無しで確かめるための見本サーバー。
// GAS 版と同じ形の固定データを返す（形は docs/vps-migration/inv-1-workflow.md の 1 節）。業務の決まりは簡単にしか再現しない。
//
//   node vps-app/web/dev/mock-server.js [port]        既定 8767 → http://localhost:8767/?page=top
//   MOCK_ROLE=admin（既定）  管理者（総務の画面・運用の欄・古い管理画面も開ける）
//   MOCK_ROLE=somu           総務（運用の欄は出ない。部署管理は権限エラー）
//   MOCK_ROLE=approver       部署承認者（部署管理だけ開ける）
//   MOCK_ROLE=staff          一般（総務管理・部署管理・管理は権限エラー）
//
// 呼び出しの記録: GET /__calls（関数名・引数・X-Idempotency-Key）。GET /__reset で記録と失敗の予約を消す。
// 失敗の見本: GET /__fail?status=500&count=1        次の書き込みを status で失敗させる（処理はしない）
//            GET /__fail?status=500&count=1&after=1 処理した後に status を返す（届いたのに応答が失敗）
//            GET /__fail?status=401&count=1        次の呼び出し（読む関数も）と /auth/me をログイン切れにする
//            GET /__fail?status=409&count=3        次の書き込みを 3 回「処理中」で返す
// 画面は nginx と同じ CSP（script-src 'self'）で返す（インラインの script や on 属性を使うと動かない）。
'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = Number(process.argv[2]) || 8767;
const VPS = path.resolve(__dirname, '..', '..');
const STATIC_DIR = path.join(VPS, 'web', 'static');
const TPL_DIR = path.join(VPS, 'api', 'app', 'templates', 'pages');
const ROLE = process.env.MOCK_ROLE || 'admin';
const IS_ADMIN = ROLE === 'admin';
const IS_SOMU = ROLE === 'admin' || ROLE === 'somu';
const IS_APPROVER = ROLE === 'approver';
const USER = { email: { admin: 'imaizumi', somu: 'somu', approver: 'shonin', staff: 'genba' }[ROLE] + '@lineworks-local.info' };
const PORTAL_URL = 'https://script.google.com/a/macros/lineworks-local.info/s/PORTAL_MOCK/exec';
const CSP = "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; " +
  "connect-src 'self'; frame-ancestors 'self'; base-uri 'self'; form-action 'self'";

// ---- 日付（JST） ----
function pad(n) { return String(n).padStart(2, '0'); }
function jst(d) { return new Date((d ? d.getTime() : Date.now()) + 9 * 3600 * 1000); }
function ymd(d) { const j = jst(d); return j.getUTCFullYear() + '-' + pad(j.getUTCMonth() + 1) + '-' + pad(j.getUTCDate()); }
function ymdhms(d) { const j = jst(d); return ymd(d) + ' ' + pad(j.getUTCHours()) + ':' + pad(j.getUTCMinutes()) + ':' + pad(j.getUTCSeconds()); }
function dateOf(s) { return new Date(s + 'T00:00:00+09:00'); }
function addDays(s, n) { return ymd(new Date(dateOf(s).getTime() + n * 86400000)); }
function dow(s) { return jst(dateOf(s)).getUTCDay(); }
function slash(s) { return s ? s.replace(/-/g, '/') : ''; }
const DOWS = ['日', '月', '火', '水', '木', '金', '土'];
const TODAY = ymd();
function fyOf(s) { const y = Number(s.slice(0, 4)), md = s.slice(5); return md < '03-16' ? y - 1 : y; }
const CUR_FY = fyOf(TODAY);
function inFy(s, fy) { return s >= fy + '-03-16' && s <= (fy + 1) + '-03-15'; }
// 月次期間（前月16日〜当月15日）
function period(s) {
  const y = Number(s.slice(0, 4)), m = Number(s.slice(5, 7)), d = Number(s.slice(8));
  const f = (yy, mm, dd) => { const t = new Date(Date.UTC(yy, mm - 1, dd)); return t.getUTCFullYear() + '-' + pad(t.getUTCMonth() + 1) + '-' + pad(t.getUTCDate()); };
  return d >= 16 ? { start: f(y, m, 16), end: f(y, m + 1, 15) } : { start: f(y, m - 1, 16), end: f(y, m, 15) };
}

// ---- 固定データ（XSS の確かめのため、氏名に 2 つの悪い文字を入れてある） ----
const EVIL1 = '"><img src=x onerror=alert(1)>';
const EVIL2 = "x');alert(1)//";
const DEPTS = ['品質管理課', '製造1課', '製造2課'];
const WORKERS = {
  '製造1課': [['1001', '山田 太郎'], ['1002', '佐藤 花子'], ['1003', EVIL2]],
  '製造2課': [['2001', '鈴木 一郎'], ['2002', EVIL1]],
  '品質管理課': [['3001', '高橋 次郎']]
};
const STAMPS = [{ email: 'bucho@lineworks-local.info', name: '部長 一郎' }, { email: 'kacho@lineworks-local.info', name: '課長 二郎' }];
const WARN = { '1001': '', '1002': '注意', '1003': '警告', '2001': '危険', '2002': '注意', '3001': '' };
function deptOfWorker(id) { return DEPTS.find((d) => WORKERS[d].some((w) => w[0] === id)); }
function nameOfWorker(id) { const d = deptOfWorker(id); return d ? WORKERS[d].find((w) => w[0] === id)[1] : ''; }

// 勤務日（平日。M_CALENDAR の代わり）と休日
function isWorkday(s) { const w = dow(s); return w !== 0 && w !== 6; }
function workingDays() {
  const out = [];
  for (let s = TODAY; s <= (CUR_FY + 1) + '-03-15'; s = addDays(s, 1)) { if (isWorkday(s)) { out.push(s); } }
  return out;
}

let seq = 0;
const reqs = [];
function mk(o) {
  const leaveDate = o.leaveDate || addDays(TODAY, 3);
  const r = Object.assign({
    reqId: 'LV-' + TODAY.replace(/-/g, '') + '-' + String(++seq).padStart(3, '0'),
    deptId: '製造1課', deptName: '製造1課', workerId: '1001', workerName: '山田 太郎',
    leaveKubun: '全日休', halfDayType: '', leaveDate: leaveDate, leaveType: '有給消化（私用）', substituteDate: '',
    specialReason: '', paidDetail: '', additionalDetail: '', submittedAt: TODAY + ' 09:00:00', approvedAt: '', status: '申請中',
    signImageUrl: '', pdfUrl: '', approvedBy1Email: '', approvedBy2: '', approvedBy2Email: '', approvedAt2: ''
  }, o);
  r.fiscalYear = String(fyOf(r.leaveDate));
  reqs.push(r);
  return r;
}
function nextWorkday(n) { let s = TODAY, k = 0; while (k < n) { s = addDays(s, 1); if (isWorkday(s)) { k++; } } return s; }
mk({ leaveDate: nextWorkday(2) });
mk({ workerId: '1003', workerName: EVIL2, leaveDate: nextWorkday(3), leaveKubun: '半日休', halfDayType: '午前', paidDetail: '通院のため' });
mk({ deptId: '製造2課', deptName: '製造2課', workerId: '2002', workerName: EVIL1, leaveDate: nextWorkday(4), leaveType: '特別休暇', specialReason: '忌引休暇' });
mk({ workerId: '1002', workerName: '佐藤 花子', leaveDate: TODAY, status: '承認済', approvedAt: TODAY + ' 08:30:00', approvedBy1Email: STAMPS[0].email });
mk({ deptId: '製造2課', deptName: '製造2課', workerId: '2002', workerName: EVIL1, leaveDate: TODAY, status: '承認済', approvedAt: TODAY + ' 08:40:00' });
mk({ deptId: '製造2課', deptName: '製造2課', workerId: '2001', workerName: '鈴木 一郎', leaveDate: nextWorkday(5), status: '承認済', approvedAt: TODAY + ' 10:00:00' });
mk({ workerId: '1003', workerName: EVIL2, leaveDate: addDays(TODAY, -10), submittedAt: TODAY + ' 07:50:00', status: '承認済', approvedAt: TODAY + ' 11:00:00' });
mk({ deptId: '品質管理課', deptName: '品質管理課', workerId: '3001', workerName: '高橋 次郎', leaveDate: addDays(TODAY, -3), submittedAt: TODAY + ' 08:10:00' });
mk({ leaveDate: nextWorkday(6), status: '承認済', approvedAt: TODAY + ' 09:30:00', approvedBy2: '部長 一郎', approvedBy2Email: STAMPS[0].email, approvedAt2: TODAY + ' 12:00:00' });
mk({ workerId: '1002', workerName: '佐藤 花子', leaveDate: nextWorkday(7), status: '取消' });

function find(reqId) { return reqs.find((r) => r.reqId === reqId) || null; }
function isRetro(r) { return r.submittedAt.slice(0, 10) > r.leaveDate; }
function daysAgo(r) { const d = Math.round((dateOf(TODAY) - dateOf(r.leaveDate)) / 86400000); return d > 0 ? d : 0; }
function stampName(email) { const s = STAMPS.find((x) => x.email === String(email).toLowerCase()); return s ? s.name : email; }

// 有給の集計（GAS の buildAdminViewData_ の形。数え方は見本）
const MILESTONES = [['6/30', '06-30'], ['8/31', '08-31'], ['12/31', '12-31'], ['2/15', '02-15'], ['3/15', '03-15']];
function summaryRows(fy, deptFilter) {
  const rows = [];
  DEPTS.forEach((dept) => {
    if (deptFilter && deptFilter !== 'ALL' && deptFilter !== dept) { return; }
    WORKERS[dept].forEach(([id, name]) => {
      const mine = reqs.filter((r) => r.workerId === id && r.status === '承認済' && inFy(r.leaveDate, fy));
      const paid = mine.filter((r) => r.leaveType === '有給消化（私用）');
      const paidCount = paid.reduce((s, r) => s + (r.leaveKubun === '半日休' ? 0.5 : 1), 0) + (id === '3001' ? 5 : 0);
      const paidHalfCount = paid.filter((r) => r.leaveKubun === '半日休').length;
      const substituteCount = mine.filter((r) => r.leaveType === '振替休暇').length;
      const specialCount = mine.filter((r) => r.leaveType === '特別休暇').length;
      const lastPaid = paid.map((r) => r.leaveDate).sort().pop() || '';
      let cum = 0;
      const quarters = MILESTONES.map((m, i) => {
        const end = (i < 3 ? fy : fy + 1) + '-' + m[1];
        cum = Math.min(paidCount, i + 1);
        const isFuture = end > TODAY && fy === CUR_FY ? true : fy > CUR_FY;
        const shortage = isFuture ? 0 : Math.max(0, i + 1 - paidCount);
        return { id: 'Q' + (i + 1), label: m[0], targetCount: i + 1, periodCount: cum, cumulativeCount: paidCount, shortage: shortage,
          level: shortage >= 2 ? 'danger' : (shortage >= 1 ? 'warn' : 'ok'), taken: !isFuture && shortage === 0, overdue: shortage > 0,
          isCurrent: false, isPast: !isFuture, isFuture: isFuture };
      });
      rows.push({ deptId: dept, deptName: dept, workerId: id, workerName: name, paidCount: paidCount, paidHalfCount: paidHalfCount,
        substituteCount: substituteCount, specialCount: specialCount, totalCount: paidCount + substituteCount + specialCount,
        lastPaidDate: lastPaid, warnLevel: WARN[id], quarters: quarters, currentShortage: 0 });
    });
  });
  return rows;
}
function plainRow(w) {
  return { deptId: w.deptId, deptName: w.deptName, workerId: w.workerId, workerName: w.workerName, paidCount: w.paidCount,
    paidHalfCount: w.paidHalfCount, substituteCount: w.substituteCount, specialCount: w.specialCount, totalCount: w.totalCount,
    lastPaidDate: w.lastPaidDate, warnLevel: w.warnLevel };
}
function needSomu() { if (!IS_SOMU && !IS_ADMIN) { throw new Error('権限がありません（総務または管理者のみ）'); } }

const RPC = {
  api_isDeptApprover: () => IS_APPROVER,
  api_isAdmin: () => IS_ADMIN,
  api_isSomu: () => IS_SOMU,
  api_getDeptList: () => DEPTS.map((d) => ({ deptId: d, deptName: d })),
  api_getWorkersByDept: (deptId) => (WORKERS[deptId] || []).map(([id, name]) => ({ workerId: id, workerName: name, deptId: deptId })),
  api_getWorkerInfo: () => null,  // 作業員マスタに Google アカウントの列が無いので、本番でもいつも null
  api_getLookupData: () => ({ leaveKubun: ['全日休', '半日休'], halfDayType: ['午前', '午後'],
    leaveTypeFull: ['有給消化（私用）', '振替休暇', '特別休暇'], leaveTypeHalf: ['有給消化（私用）', '特別休暇'],
    specialReasons: ['結婚休暇', '忌引休暇', '出産休暇', '生理休暇', '裁判員休暇', 'その他'] }),
  api_getWorkingDays: () => workingDays(),
  api_getHolidayDays: (workerId) => {
    const used = reqs.filter((r) => r.status !== '取消' && r.workerId === workerId && r.substituteDate).map((r) => r.substituteDate);
    const out = [];
    for (let s = addDays(TODAY, -45); s <= addDays(TODAY, 60); s = addDays(s, 1)) {
      if (!isWorkday(s) && used.indexOf(s) < 0) { out.push({ date: s, label: dow(s) === 0 ? '日曜' : '土曜' }); }
    }
    return out;
  },
  api_getPastWorkingDays: (ym) => {
    const m = /^(\d{4})-?(\d{2})$/.exec(ym || '') || [null, TODAY.slice(0, 4), TODAY.slice(5, 7)];
    const out = [];
    for (let d = 1; d <= 31; d++) {
      const s = m[1] + '-' + m[2] + '-' + pad(d);
      if (ymd(dateOf(s)) !== s || s > TODAY) { continue; }
      if (isWorkday(s)) { out.push({ date: s, dow: dow(s), label: s + '（' + DOWS[dow(s)] + '）' }); }
    }
    return out;
  },
  api_getUpcomingLeaves: (deptId, workerId) => reqs
    .filter((r) => r.status !== '取消' && (!deptId || r.deptId === deptId) && (!workerId || r.workerId === workerId) && r.leaveDate >= TODAY)
    .map((r) => ({ reqId: r.reqId, deptName: r.deptName, workerName: r.workerName, leaveKubun: r.leaveKubun, halfDayType: r.halfDayType,
      leaveDate: slash(r.leaveDate), leaveType: r.leaveType, status: r.status, isRetroactive: isRetro(r), daysAgo: daysAgo(r) }))
    .sort((a, b) => a.leaveDate.localeCompare(b.leaveDate)),
  api_getAllPendingRequests: () => reqs.filter((r) => r.status === '申請中')
    .map((r) => ({ reqId: r.reqId, deptName: r.deptName, workerName: r.workerName, leaveKubun: r.leaveKubun, halfDayType: r.halfDayType,
      leaveDate: slash(r.leaveDate), leaveType: r.leaveType, submittedAt: r.submittedAt.slice(5, 16).replace('-', '/'), status: r.status,
      isRetroactive: isRetro(r), daysAgo: daysAgo(r) }))
    .sort((a, b) => b.submittedAt.localeCompare(a.submittedAt)),
  api_submitLeaveRequest: (data) => {
    const d = data || {};
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.leaveDate || '')) { return { ok: false, error: '休暇日の形式が不正です（yyyy-MM-dd）。' }; }
    if (d.isRetroactive === true) {
      if (!(inFy(d.leaveDate, CUR_FY) && d.leaveDate <= TODAY)) { return { ok: false, error: '過去日申請は今年度内（年度開始の3/16から本日まで）の日付のみ指定できます。' }; }
    } else if (workingDays().indexOf(d.leaveDate) < 0) {
      return { ok: false, error: '指定された休暇日は勤務日ではありません。過去日の場合は「過去申請」から選択してください。' };
    }
    if (d.leaveType === '振替休暇') {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d.substituteDate || '')) { return { ok: false, error: '振替休暇の場合は振替元出勤日を選択してください。' }; }
      const lp = period(d.leaveDate), sp = period(d.substituteDate);
      if (lp.start !== sp.start) {
        return { ok: false, error: '振替休暇は同一月次期間（前月16日〜当月15日）内で消化する必要があります。\n休暇日: ' + d.leaveDate +
          '（期間: ' + lp.start + '〜' + lp.end + '）\n振替元: ' + d.substituteDate + '（期間: ' + sp.start + '〜' + sp.end + '）' };
      }
      if (reqs.some((r) => r.status !== '取消' && r.workerId === d.workerId && r.substituteDate === d.substituteDate)) {
        return { ok: false, error: '指定された振替元出勤日（' + d.substituteDate + '）は既に別の振替休暇申請で使用されています。' };
      }
    }
    const r = mk({ deptId: d.deptId, deptName: d.deptName, workerId: d.workerId, workerName: d.workerName, leaveKubun: d.leaveKubun,
      halfDayType: d.halfDayType, leaveDate: d.leaveDate, leaveType: d.leaveType, substituteDate: d.substituteDate || '',
      specialReason: d.specialReason, paidDetail: d.paidDetail, submittedAt: ymdhms() });
    return { ok: true, reqId: r.reqId };
  },
  api_editLeaveRequest: (reqId, patch) => {
    const r = find(reqId);
    if (!r) { throw new Error('申請が見つかりません。'); }
    if (r.status === '取消' || r.status === '却下') { throw new Error('取消/却下済みの申請は編集できません。'); }
    const p = patch || {};
    if (!/^\d{4}-\d{2}-\d{2}$/.test(p.leaveDate || '')) { throw new Error('休暇日の形式が不正です: ' + p.leaveDate); }
    if (p.leaveKubun !== '全日休' && p.leaveKubun !== '半日休') { throw new Error('休暇区分が不正です: ' + p.leaveKubun); }
    if (p.leaveKubun === '半日休' && !p.halfDayType) { throw new Error('半日区分を選択してください。'); }
    r.leaveDate = p.leaveDate; r.leaveKubun = p.leaveKubun; r.halfDayType = p.leaveKubun === '半日休' ? p.halfDayType : '';
    if (r.status === '承認済') {
      Object.assign(r, { status: '申請中', approvedAt: '', approvedBy1Email: '', signImageUrl: '', approvedBy2: '', approvedBy2Email: '', approvedAt2: '', pdfUrl: '' });
    }
    return { ok: true, newStatus: r.status };
  },
  api_deleteLeaveRequest: (reqId) => {
    const r = find(reqId);
    if (!r) { throw new Error('申請が見つかりません。'); }
    if (r.status === '取消') { throw new Error('既に取消済みです。'); }
    r.status = '取消'; r.pdfUrl = '';
    return { ok: true };
  },
  // 存在しない reqId はすぐ null（GAS は 6 秒待っていた）
  api_getLeaveRequestById: (reqId) => {
    const r = find(reqId);
    if (!r) { return null; }
    return { reqId: r.reqId, fiscalYear: r.fiscalYear, deptId: r.deptId, deptName: r.deptName, workerId: r.workerId, workerName: r.workerName,
      leaveKubun: r.leaveKubun, halfDayType: r.halfDayType, leaveDate: r.leaveDate, leaveType: r.leaveType, substituteDate: r.substituteDate,
      specialReason: r.specialReason, paidDetail: r.paidDetail, additionalDetail: r.additionalDetail, submittedAt: r.submittedAt,
      approvedAt: r.approvedAt, status: r.status, signImageUrl: r.signImageUrl, pdfUrl: r.pdfUrl, approvedBy1Email: r.approvedBy1Email,
      approvedBy2: r.approvedBy2, approvedBy2Email: r.approvedBy2Email, approvedAt2: r.approvedAt2, rowNo: reqs.indexOf(r) + 2 };
  },
  api_getApproverList: () => STAMPS.slice(),
  // 1 次承認は「申請中」だけ（オーナーの決定 K-4）
  api_approveLeaveRequest: (reqId, signBase64, approverEmail) => {
    if (!reqId) { throw new Error('REQ_IDが指定されていません。'); }
    const r = find(reqId);
    if (!r) { throw new Error('申請が見つかりません: ' + reqId); }
    if (r.status === '承認済') { throw new Error('この申請はすでに承認済みです。'); }
    if (r.status !== '申請中') { throw new Error('この申請はすでに取消済みです。'); }
    if (signBase64 && String(signBase64).indexOf('data:image/png;base64,') !== 0) { throw new Error('サインの画像が不正です。'); }
    Object.assign(r, { status: '承認済', approvedAt: ymdhms(), signImageUrl: signBase64 ? 'https://drive.google.com/mock/sign_' + reqId + '.png' : '',
      approvedBy1Email: approverEmail || '' });
    return { ok: true, reqId: reqId };
  },
  api_getCurrentApproverInfo: () => ({ email: USER.email, name: stampName(USER.email) }),
  api_getPending2ndApprovals: () => {
    needSomu();
    return reqs.filter((r) => r.status === '承認済' && !r.approvedAt2)
      .sort((a, b) => a.leaveDate.localeCompare(b.leaveDate))
      .map((r) => ({ reqId: r.reqId, deptName: r.deptName, workerName: r.workerName, leaveDate: slash(r.leaveDate), leaveKubun: r.leaveKubun,
        leaveType: r.leaveType, approvedAt: slash(r.approvedAt.slice(0, 16)), isRetroactive: isRetro(r), daysAgo: daysAgo(r) }));
  },
  api_approveLeaveRequest2: (reqId, approver2Email) => {
    if (!reqId) { throw new Error('REQ_IDが指定されていません。'); }
    needSomu();
    const r = find(reqId);
    if (!r) { throw new Error('申請が見つかりません: ' + reqId); }
    if (r.status !== '承認済') { throw new Error('1次承認されていない申請です。'); }
    Object.assign(r, { approvedBy2: stampName(approver2Email), approvedBy2Email: approver2Email, approvedAt2: ymdhms(),
      pdfUrl: 'https://drive.google.com/mock/' + reqId + '.pdf' });
    return { ok: true, reqId: reqId };
  },
  api_getSomuAdminView: (fy, deptFilter) => { needSomu(); return summaryRows(Number(fy), deptFilter); },
  api_getApproverProfiles: () => (IS_ADMIN
    ? { isAdmin: true, currentEmail: USER.email, profiles: [
      { email: 'shonin@lineworks-local.info', name: '承認 一郎', depts: ['製造1課', '製造2課'] },
      { email: 'qc@lineworks-local.info', name: '品管 二郎', depts: ['品質管理課'] }] }
    : { isAdmin: false, currentEmail: USER.email, profiles: IS_APPROVER ? [{ email: USER.email, name: '承認 一郎', depts: ['製造1課', '製造2課'] }] : [] }),
  api_getDeptAdminView: (fy, depts) => {
    if (!IS_ADMIN && !IS_APPROVER) { throw new Error('部署承認者の権限がありません。'); }
    const ds = Array.isArray(depts) ? depts : [];
    return { workers: summaryRows(Number(fy), 'ALL').filter((w) => ds.indexOf(w.deptName) >= 0), deptNames: ds };
  },
  api_getAdminBossView: (deptId, fy) => summaryRows(Number(fy), deptId || 'ALL').map(plainRow),
  api_getAdminSomuView: (fy, deptFilter) => summaryRows(Number(fy), deptFilter || 'ALL').map(plainRow),
  api_getLeaveStatusSearch: (fy, deptId, workerId) => {
    const fyYear = Number(fy);
    const rows = summaryRows(fyYear, deptId || 'ALL');
    if (!workerId) { return { mode: 'list', fyYear: fyYear, deptFilter: deptId || '', workers: rows }; }
    const w = rows.find((x) => x.workerId === workerId) || null;
    const milestones = MILESTONES.map((m, i) => {
      const end = (i < 3 ? fyYear : fyYear + 1) + '-' + m[1];
      const required = i + 1, count = w ? w.paidCount : 0, passed = end < TODAY, met = count >= required;
      const deficit = Math.max(0, required - count);
      return { label: m[0], month: Number(m[1].slice(0, 2)), day: Number(m[1].slice(3)), dateText: m[0], required: required, met: met,
        passed: passed, deficit: deficit, remaining: deficit, state: met ? 'done' : (!passed ? 'future' : (deficit >= 3 ? 'risk' : (deficit >= 2 ? 'warn' : 'note'))) };
    });
    const upcoming = fyYear === CUR_FY ? reqs.filter((r) => r.workerId === workerId && r.status !== '取消' && r.leaveDate >= TODAY)
      .map((r) => ({ leaveDate: slash(r.leaveDate), leaveKubun: r.leaveKubun, halfDayType: r.halfDayType, leaveType: r.leaveType, status: r.status })) : [];
    return { mode: 'personal', fyYear: fyYear, worker: w, milestones: milestones, upcoming: upcoming };
  },
  // 運用（管理者だけ）
  ops_getStatus: () => ({ rowSync: { pending: 1, oldestSeconds: 12, blocked: [] },
    gasJobs: { pending: 0, running: 0, failed: 1, failedJobs: [{ id: 41, request_id: reqs[3].reqId, action: 'afterApprove2', last_error: 'GAS が応答しませんでした', attempts: 8 }] },
    lastGasPing: { at: ymdhms(), ok: true }, masterSync: { at: TODAY + ' 00:30:05', ok: true }, configSync: { at: ymdhms(), ok: true } }),
  ops_retryJob: () => ({ ok: true }),
  ops_syncSheetConfigNow: () => ({ ok: true, counts: { settings: 15, dept_approvers: 13, stamps: 6 } }),
  ops_runMasterSync: () => ({ ok: true, tabs: { '作業員マスタ': { ok: true, rows: 50 } } }),
  ops_rebuildSummaryNow: () => ({ ok: true, rows: 37 }),
  ops_sendWarnMail: () => ({ ok: true, jobId: 103 }),
  ops_revokeSessions: () => ({ ok: true, revoked: 1 })
};
const WRITE = ['api_submitLeaveRequest', 'api_editLeaveRequest', 'api_deleteLeaveRequest', 'api_approveLeaveRequest', 'api_approveLeaveRequest2',
  'ops_retryJob', 'ops_syncSheetConfigNow', 'ops_runMasterSync', 'ops_rebuildSummaryNow', 'ops_sendWarnMail', 'ops_revokeSessions'];

// ---- テンプレート（pages.py と同じ値を差し込む。Jinja2 の代わりに {{ X }} と {% if SHOW_OPS %} だけを扱う） ----
function escHtml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&#34;').replace(/'/g, '&#39;'); }
function render(file, vars) {
  return fs.readFileSync(path.join(TPL_DIR, file), 'utf8')
    .replace(/\{% if SHOW_OPS %\}([\s\S]*?)\{% endif %\}\n?/g, (all, body) => (vars.SHOW_OPS ? body : ''))
    .replace(/\{\{ data_json\|safe \}\}/g, () => vars.data_json || '')
    .replace(/\{\{ (\w+) \}\}/g, (all, k) => escHtml(vars[k] === undefined ? '' : vars[k]));
}
function jsonForScript(o) { return JSON.stringify(o).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026'); }

const NO_AUTH = {
  admin: ['アクセス権限エラー', 'このページは管理者・総務のみアクセスできます。'],
  deptAdmin: ['アクセス権限なし', 'このページは部署承認者または管理者のみアクセスできます。'],
  somuAdmin: ['アクセス権限なし', 'このページは総務部・管理者のみアクセスできます。']
};
function allowed(page) {
  if (page === 'admin') { return IS_ADMIN || IS_SOMU; }
  if (page === 'deptAdmin') { return IS_ADMIN || IS_APPROVER; }
  if (page === 'somuAdmin') { return IS_SOMU || IS_ADMIN; }
  return true;
}

const calls = [];
const done = {};
let failNext = null;
function takeFail(isWrite) {
  if (!failNext || (!isWrite && failNext.status !== 401)) { return null; }
  const f = failNext;
  f.count--;
  if (f.count <= 0) { failNext = null; }
  return f;
}

function send(res, status, body, type, extra) {
  const text = (typeof body === 'string' || Buffer.isBuffer(body)) ? body : JSON.stringify(body);
  res.writeHead(status, Object.assign({ 'Content-Type': type || 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }, extra || {}));
  res.end(text);
}

const TYPES = { '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml' };
const PAGES = ['top', 'preview', 'sign', 'somuAdmin', 'deptAdmin', 'admin', 'leaveStatus'];

const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  const p = url.pathname;
  if (p === '/__calls') { return send(res, 200, calls); }
  if (p === '/__reset') { calls.length = 0; failNext = null; return send(res, 200, { ok: true }); }
  if (p === '/__fail') {
    failNext = { status: Number(url.searchParams.get('status')) || 500, count: Number(url.searchParams.get('count')) || 1, after: url.searchParams.get('after') === '1' };
    return send(res, 200, { ok: true, failNext: failNext });
  }
  if (p === '/auth/me') {
    if (failNext && failNext.status === 401) { return send(res, 401, { ok: false, error: 'ログインが必要です。' }); }
    return send(res, 200, { email: USER.email, name: USER.email.split('@')[0] });
  }
  if (p.startsWith('/auth/')) {
    const next = url.searchParams.get('next') || '/';
    return send(res, 200, '<!doctype html><meta charset="utf-8"><title>mock login</title><p>見本サーバーのログイン画面（' + escHtml(p) + '）</p>' +
      '<p><a href="' + escHtml(next.startsWith('/') && !next.startsWith('//') ? next : '/') + '">ログインしたことにして戻る</a></p>', 'text/html; charset=utf-8', { 'Content-Security-Policy': CSP });
  }
  if (p.startsWith('/static/')) {
    const file = path.normalize(path.join(STATIC_DIR, p.slice('/static/'.length)));
    if (!file.startsWith(STATIC_DIR) || !fs.existsSync(file) || !fs.statSync(file).isFile()) { return send(res, 404, 'not found', 'text/plain'); }
    return send(res, 200, fs.readFileSync(file), TYPES[path.extname(file)] || 'application/octet-stream');
  }
  if (p.startsWith('/api/rpc/')) {
    const name = decodeURIComponent(p.slice('/api/rpc/'.length));
    let raw = '';
    req.on('data', (c) => { raw += c; });
    req.on('end', () => {
      let args;
      try { args = raw ? JSON.parse(raw) : []; } catch (e) { return send(res, 400, { ok: false, error: '入力が不正です' }); }
      const key = req.headers['x-idempotency-key'] || '';
      calls.push({ name: name, args: args, idempotencyKey: key, at: ymdhms() });
      if (req.method !== 'POST') { return send(res, 405, { ok: false, error: 'POST だけです' }); }
      if (!Object.prototype.hasOwnProperty.call(RPC, name)) { return send(res, 501, { ok: false, error: '未登録の関数です: ' + name }); }
      if (!Array.isArray(args)) { return send(res, 400, { ok: false, error: '入力が不正です' }); }
      const write = WRITE.indexOf(name) >= 0;
      const fail = takeFail(write);
      if (fail && fail.status === 401) { return send(res, 401, { ok: false, error: 'ログインが必要です。', login: '/auth/login' }); }
      if (!IS_ADMIN && /^ops_/.test(name)) { return send(res, 403, { ok: false, error: '管理者権限がありません。' }); }
      if (write && key && done[key]) { return send(res, 200, done[key]); }
      if (fail && fail.status === 409) { return send(res, 409, { ok: false, error: '同じ操作を処理中です。' }); }
      if (fail && !fail.after) { return send(res, fail.status, { ok: false, error: 'サーバーでエラーが起きました。時間をおいてもう一度お試しください。' }); }
      let body;
      try {
        const result = RPC[name].apply(null, args);
        body = { ok: true, result: result === undefined ? null : result };
      } catch (e) {
        return send(res, 400, { ok: false, error: e.message });
      }
      if (write && key) { done[key] = body; }
      if (fail) { return send(res, fail.status, { ok: false, error: 'サーバーでエラーが起きました。時間をおいてもう一度お試しください。' }); }
      // 実際のサーバーの遅さに近づける
      setTimeout(() => send(res, 200, JSON.parse(JSON.stringify(body))), write ? 300 : 120);
    });
    return undefined;
  }
  if (p === '/') {
    const pageParam = url.searchParams.get('page') || 'top';
    const page = PAGES.includes(pageParam) ? pageParam : 'top';
    const APP_URL = 'http://localhost:' + PORT + '/';
    let html;
    if (!allowed(page)) {
      html = render('no_auth.html', { title: NO_AUTH[page][0], message: NO_AUTH[page][1], APP_URL: APP_URL });
    } else {
      const reqId = url.searchParams.get('reqId') || '';
      const data = { PAGE: page, APP_URL: APP_URL, PORTAL_URL: PORTAL_URL };
      if (page === 'preview' || page === 'sign') { data.REQ_ID = /^LV-[0-9]{8}-[0-9]{3}$/.test(reqId) ? reqId : ''; }
      html = render(page + '.html', { data_json: jsonForScript(data), APP_URL: APP_URL, PORTAL_URL: PORTAL_URL, SHOW_OPS: IS_ADMIN });
    }
    return send(res, 200, html, 'text/html; charset=utf-8', { 'Content-Security-Policy': CSP });
  }
  return send(res, 404, 'not found', 'text/plain');
});

module.exports = { RPC, WRITE, server };
if (require.main === module) {
  server.listen(PORT, () => {
    console.log('見本サーバー: http://localhost:' + PORT + '/?page=top  （' + ROLE + ' として）');
  });
}
