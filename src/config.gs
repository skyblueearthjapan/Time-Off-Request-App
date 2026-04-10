// ====== CONFIG ======
const TZ = 'Asia/Tokyo';

const SHEET = {
  SETTINGS: 'M_SYSTEM_SETTING',
  WORKER: '作業員マスタ',
  LOOKUP: 'M_LOOKUP',
  LEAVE_REQUEST: 'T_LEAVE_REQUEST',
  LEAVE_SUMMARY: 'V_LEAVE_SUMMARY',
  BATCH_LOGS: 'BatchLogs',
  CALENDAR: 'M_CALENDAR',
  STAMP: 'M_STAMP',
  DEPT_APPROVERS: 'M_DEPT_APPROVERS',
};

const STATUS = {
  SUBMITTED: '申請中',
  APPROVED: '承認済',
  CANCELED: '取消',
  REJECTED: '却下',
};

// 休暇区分
const LEAVE_KUBUN = {
  FULL_DAY: '全日休',
  HALF_DAY: '半日休',
};

// 半日区分
const HALF_DAY_TYPE = {
  AM: '午前',
  PM: '午後',
};

// 休暇種類
const LEAVE_TYPE = {
  PAID: '有給消化（私用）',
  SUBSTITUTE: '振替休暇',
  SPECIAL: '特別休暇',
};

// 有給警告基準（年度内）
const PAID_LEAVE_WARN = [
  { deadline: { month: 6, day: 30 }, required: 1, label: '6月末' },
  { deadline: { month: 8, day: 31 }, required: 2, label: '8月末' },
  { deadline: { month: 12, day: 31 }, required: 3, label: '12月末' },
  { deadline: { month: 2, day: 15 }, required: 4, label: '2/15' },
  { deadline: { month: 3, day: 15 }, required: 5, label: '年度末' },
];

// 有給取得義務期間（年5回の計画取得）
var PAID_LEAVE_QUARTERS = [
  { id: 'Q1',      label: '春期(3/16-6月)', startMonth: 3, startDay: 16, endMonth: 6,  endDay: 30, targetCount: 1 },
  { id: 'SUMMER',  label: '夏期(7-8月)',   startMonth: 7, startDay: 1,  endMonth: 8,  endDay: 31, targetCount: 2 },
  { id: 'Q3',      label: '秋期(9-11月)',  startMonth: 9, startDay: 1,  endMonth: 11, endDay: 30, targetCount: 3 },
  { id: 'WINTER',  label: '冬期(12-2/15)',  startMonth: 12, startDay: 1, endMonth: 2,  endDay: 15, targetCount: 4 },
  { id: 'YEAREND', label: '年度末(2/16-3/15)', startMonth: 2, startDay: 16, endMonth: 3,  endDay: 15, targetCount: 5 },
];

// 警告メール送信日（月初 or 期限翌日）
const WARN_MAIL_DATES = [
  { month: 7, day: 1 },
  { month: 9, day: 1 },
  { month: 1, day: 1 },
  { month: 2, day: 16 },
];

// ====== ヘッダエイリアス（英語ヘッダ → 日本語キー） ======
// T_LEAVE_REQUESTシートが英語ヘッダの場合に対応
var HEADER_ALIAS_ = {
  'FY_START_YEAR': '年度',
  'STATUS': '承認状態',
  'APPLIED_AT': '申請日時',
  'DEPT_ID': '部署ID',
  'DEPT_NAME': '部署名',
  'WORKER_ID': '作業員ID',
  'WORKER_NAME': '作業員名',
  'LEAVE_DATE': '休暇日',
  'DAY_TYPE': '休暇区分',
  'HALF_TYPE': '半日区分',
  'LEAVE_KIND': '休暇種類',
  'SUBSTITUTE_FOR_WORK_DATE': '振替元出勤日',
  'SPECIAL_REASON_REF': '特別理由参照',
  'SPECIAL_REASON_TEXT': '特別理由',
  'PAIDLEAVE_DETAIL': '有給詳細',
  'DETAIL_FREE': '追加詳細',
  'APPROVED_AT': '承認日時',
  'SIGN_IMAGE_URL': 'サイン画像URL',
  'CREATED_BY_EMAIL': '作成者メール',
  'UPDATED_AT': '更新日時',
  'APPROVED_BY1_EMAIL': '1次承認者メール',
  'APPROVED_BY2': '2次承認者名',
  'APPROVED_BY2_EMAIL': '2次承認者メール',
  'APPROVED_AT2': '2次承認日時',
};

// ====== UTIL ======
function fmtDate_(d, pattern) {
  pattern = pattern || 'yyyy-MM-dd';
  return Utilities.formatDate(d, TZ, pattern);
}

function getDb_() {
  return SpreadsheetApp.getActive();
}

function getSettings_() {
  var ss = getDb_();
  var sh = ss.getSheetByName(SHEET.SETTINGS);
  if (!sh) throw new Error('Sheet not found: ' + SHEET.SETTINGS);
  var values = sh.getDataRange().getValues();
  var map = {};
  for (var r = 1; r < values.length; r++) {
    var key = String(values[r][0] || '').trim();
    var val = values[r][1];
    if (key) map[key] = val;
  }
  return map;
}

function requireSheet_(name) {
  var ss = getDb_();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}

function normalize_(s) {
  return String(s == null ? '' : s).trim()
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    });
}

function buildHeaderIndex_(header) {
  var idx = {};
  // 逆引きマップ（日本語→英語）を構築
  var reverseAlias = {};
  for (var key in HEADER_ALIAS_) {
    if (HEADER_ALIAS_.hasOwnProperty(key)) {
      reverseAlias[HEADER_ALIAS_[key]] = key;
    }
  }
  for (var i = 0; i < header.length; i++) {
    var h = header[i];
    idx[h] = i;
    // 英語ヘッダー → 日本語エイリアスも登録
    if (HEADER_ALIAS_[h] && idx[HEADER_ALIAS_[h]] === undefined) {
      idx[HEADER_ALIAS_[h]] = i;
    }
    // 日本語ヘッダー → 英語エイリアスも登録
    if (reverseAlias[h] && idx[reverseAlias[h]] === undefined) {
      idx[reverseAlias[h]] = i;
    }
  }
  return idx;
}

function getSheetHeaderIndex_(sheetName, headerRowNo) {
  headerRowNo = headerRowNo || 1;
  var sh = requireSheet_(sheetName);
  var maxCol = sh.getMaxColumns();
  if (maxCol === 0) return { sh: sh, header: [], idx: {} };
  // ヘッダ行全体を読み取り、実際の最終列を検出
  var raw = sh.getRange(headerRowNo, 1, 1, maxCol).getValues()[0];
  var lastH = 0;
  for (var c = 0; c < raw.length; c++) {
    if (String(raw[c] || '').trim() !== '') lastH = c + 1;
  }
  if (lastH === 0) return { sh: sh, header: [], idx: {} };
  var header = [];
  for (var c = 0; c < lastH; c++) {
    header.push(normalize_(raw[c]));
  }
  return { sh: sh, header: header, idx: buildHeaderIndex_(header) };
}

// ====== 過去日遡及申請ヘルパー ======

/**
 * 当該日が属する年度の開始日（3/16）を返す
 * computeFiscalYear_ (leaveRequest.gs) を再利用し、年度算出ロジックを一元化
 */
function getFiscalYearStartDate_(dateObj) {
  var fy = computeFiscalYear_(dateObj || new Date());
  return new Date(fy, 2, 16, 0, 0, 0); // 3月は index=2
}

/**
 * 文字列（yyyy-MM-dd）が今年度内（3/16〜今日）の範囲に収まっているか判定
 * タイムゾーン安全化のため、全て 'yyyy-MM-dd' 文字列化して比較する
 */
function isWithinCurrentFiscalYear_(dateStr) {
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return false;
  var today = new Date();
  var fyStart = getFiscalYearStartDate_(today);
  var fyStartStr = Utilities.formatDate(fyStart, TZ, 'yyyy-MM-dd');
  var todayStr = Utilities.formatDate(today, TZ, 'yyyy-MM-dd');
  return dateStr >= fyStartStr && dateStr <= todayStr;
}

/**
 * 申請日時と休暇日から「過去日申請か」を算出
 * appliedAt の日付 > leaveDate の場合 true
 */
function computeIsRetroactive_(appliedAt, leaveDateStr) {
  if (!appliedAt || !leaveDateStr) return false;
  var applied = (appliedAt instanceof Date) ? appliedAt : new Date(appliedAt);
  if (isNaN(applied.getTime())) return false;
  var appliedDayStr = Utilities.formatDate(applied, TZ, 'yyyy-MM-dd');
  var ldStr = (leaveDateStr instanceof Date) ? Utilities.formatDate(leaveDateStr, TZ, 'yyyy-MM-dd') : String(leaveDateStr).substring(0, 10);
  return appliedDayStr > ldStr;
}

/**
 * 休暇日から「何日前の申請か」を算出
 * 今日 - leaveDate、過去日でなければ 0
 */
function computeDaysAgo_(leaveDateStr, today) {
  if (!leaveDateStr) return 0;
  var t = today || new Date();
  var todayStr = Utilities.formatDate(t, TZ, 'yyyy-MM-dd');
  var t0 = new Date(todayStr + 'T00:00:00+09:00');
  var ldStr = (leaveDateStr instanceof Date) ? Utilities.formatDate(leaveDateStr, TZ, 'yyyy-MM-dd') : String(leaveDateStr).substring(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ldStr)) return 0;
  var ld = new Date(ldStr + 'T00:00:00+09:00');
  if (isNaN(ld.getTime())) return 0;
  var diffMs = t0.getTime() - ld.getTime();
  var diffDays = Math.floor(diffMs / 86400000);
  return diffDays > 0 ? diffDays : 0;
}

function getLookupValues_(category) {
  var sh = requireSheet_(SHEET.LOOKUP);
  var values = sh.getDataRange().getValues();
  var H = values[0].map(function(h) { return normalize_(h); });
  var catIdx = H.indexOf('カテゴリ');
  var valIdx = H.indexOf('値');
  var sortIdx = H.indexOf('SORT');
  if (catIdx < 0 || valIdx < 0) return [];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    if (normalize_(values[r][catIdx]) === category) {
      out.push({
        value: normalize_(values[r][valIdx]),
        sort: sortIdx >= 0 ? Number(values[r][sortIdx] || 0) : r,
      });
    }
  }
  out.sort(function(a, b) { return a.sort - b.sort; });
  return out.map(function(o) { return o.value; });
}
