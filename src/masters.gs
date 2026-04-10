// ====== マスタ読込（部署/作業員/LOOKUP） ======
// 実データ列: 作業員コード | 氏名 | 部署 | 担当業務 | (Googleアカウント)
// 部署一覧は作業員マスタの「部署」列からユニーク取得（M_DEPTシート不要）

/**
 * 部署一覧を取得（M_WORKERの「部署」列からユニーク抽出）
 */
function loadDeptList_() {
  var workersByDept = loadWorkersByDeptMap_();
  var depts = Array.from(workersByDept.keys()).sort();
  return depts.map(function(d) {
    return { deptId: d, deptName: d };
  });
}

/**
 * 作業員マスタを部署ごとにMapで返す（内部用）
 */
function loadWorkersByDeptMap_() {
  var sh = requireSheet_(SHEET.WORKER);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return new Map();
  var H = values[0].map(function(h) { return normalize_(h); });
  var idx = {
    code: H.indexOf('作業員コード'),
    name: H.indexOf('氏名'),
    dept: H.indexOf('部署'),
    active: H.indexOf('在籍フラグ'),
    email: H.findIndex(function(h) { return h.indexOf('Googleアカウント') >= 0; }),
  };
  // ヘッダが見つからない場合のフォールバック
  if (idx.code < 0) idx.code = H.indexOf('作業員ID');
  if (idx.dept < 0) idx.dept = H.indexOf('部署ID');

  if (idx.code < 0 || idx.name < 0 || idx.dept < 0) {
    throw new Error('作業員マスタのヘッダが想定と違います（作業員コード/氏名/部署）。実際: ' + H.join(', '));
  }

  var map = new Map();
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var dept = normalize_(row[idx.dept]);
    var code = normalize_(row[idx.code]);
    var name = normalize_(row[idx.name]);
    if (!dept || !code || !name) continue;

    // 在籍フラグ列がある場合のみフィルタ
    if (idx.active >= 0) {
      var active = normalize_(row[idx.active]);
      if (active && active !== '1' && active.toLowerCase() !== 'true') continue;
    }

    if (!map.has(dept)) map.set(dept, []);
    map.get(dept).push({
      workerId: code,
      workerName: name,
      deptId: dept,
    });
  }
  // コード順ソート
  for (var entry of map.entries()) {
    entry[1].sort(function(a, b) { return a.workerId.localeCompare(b.workerId); });
  }
  return map;
}

/**
 * 指定部署の作業員一覧を取得
 */
function loadWorkersByDept_(deptId) {
  var map = loadWorkersByDeptMap_();
  return map.get(deptId) || [];
}

// ====== ログインユーザーの作業員情報取得 ======

function api_getWorkerInfo() {
  var email = Session.getActiveUser().getEmail();
  if (!email) return null;
  email = email.toLowerCase();

  var sh = requireSheet_(SHEET.WORKER);
  var values = sh.getDataRange().getValues();
  var H = values[0].map(function(h) { return normalize_(h); });
  var idx = {
    code: H.indexOf('作業員コード'),
    name: H.indexOf('氏名'),
    dept: H.indexOf('部署'),
    email: H.findIndex(function(h) { return h.indexOf('Googleアカウント') >= 0; }),
    active: H.indexOf('在籍フラグ'),
  };
  if (idx.code < 0) idx.code = H.indexOf('作業員ID');
  if (idx.dept < 0) idx.dept = H.indexOf('部署ID');
  if (idx.email < 0) return null;

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (idx.active >= 0) {
      var active = normalize_(row[idx.active]);
      if (active && active !== '1' && active.toLowerCase() !== 'true') continue;
    }
    if (normalize_(row[idx.email]).toLowerCase() === email) {
      var dept = normalize_(row[idx.dept]);
      return {
        workerId: normalize_(row[idx.code]),
        workerName: normalize_(row[idx.name]),
        deptId: dept,
        deptName: dept,
        email: email,
      };
    }
  }
  return null;
}

// ====== 電子印マスタ（StampMap） ======

/**
 * メールアドレスから電子印のDriveファイルIDを取得
 * ローカル M_STAMP シートを参照（syncStampMaster で同期済み）
 */
function getStampFileId_(email) {
  if (!email) return '';
  email = email.toLowerCase().trim();
  console.log('getStampFileId_: 検索メール=' + email);

  try {
    var ss = getDb_();
    var sh = ss.getSheetByName(SHEET.STAMP);
    if (!sh) {
      console.warn('M_STAMPシートが存在しません。setupAllSheets → syncStampMaster を実行してください。');
      return '';
    }
    var lastRow = sh.getLastRow();
    console.log('M_STAMP lastRow=' + lastRow);
    if (lastRow < 2) {
      console.warn('M_STAMPシートが空です（ヘッダのみ）。syncStampMaster を実行してください。');
      return '';
    }

    var data = sh.getDataRange().getValues();
    // ヘッダ: メール | stampFileId | 備考
    console.log('M_STAMPヘッダ: ' + data[0].join(' | '));
    for (var r = 1; r < data.length; r++) {
      var rowEmail = String(data[r][0] || '').toLowerCase().trim();
      var rowStampId = String(data[r][1] || '').trim();
      if (rowEmail === email) {
        console.log('M_STAMP一致: row=' + (r+1) + ', stampFileId=' + rowStampId);
        return rowStampId;
      }
    }
    console.warn('M_STAMPに該当メールなし: ' + email + ' (全' + (data.length - 1) + '件検索済)');
  } catch (e) {
    console.error('M_STAMP読取エラー: ' + e.message);
  }
  return '';
}

/**
 * 承認者一覧を返す（ローカル M_STAMP から取得、sign.htmlのドロップダウン用）
 * syncStampMaster で同期済みのローカルシートを参照
 */
function api_getApproverList() {
  try {
    var ss = getDb_();
    var sh = ss.getSheetByName(SHEET.STAMP);
    if (!sh || sh.getLastRow() < 2) {
      console.warn('M_STAMPシートが空です。syncStampMaster を実行してください。');
      return [];
    }

    var data = sh.getDataRange().getValues();
    var list = [];
    for (var r = 1; r < data.length; r++) {
      var email = String(data[r][0] || '').trim();
      var name = String(data[r][2] || '').trim();
      if (email) {
        list.push({ email: email, name: name });
      }
    }
    return list;
  } catch (e) {
    console.error('承認者一覧取得エラー: ' + e.message);
    return [];
  }
}

// ====== API（google.script.run用） ======

function api_getDeptList() {
  return loadDeptList_();
}

function api_getWorkersByDept(deptId) {
  if (!deptId) return [];
  return loadWorkersByDept_(deptId);
}

/**
 * LOOKUPデータを一括取得（申請モーダル用）
 */
function api_getLookupData() {
  return {
    leaveKubun: ['全日休', '半日休'],
    halfDayType: ['午前', '午後'],
    leaveTypeFull: ['有給消化（私用）', '振替休暇', '特別休暇'],
    leaveTypeHalf: ['有給消化（私用）', '特別休暇'],
    specialReasons: getLookupValues_('特別休暇理由'),
  };
}

/**
 * 休日出勤日（振替元）の候補を返す（年度内の休日・祝日・土日）
 * M_CALENDARにデータがある期間はそちらを優先、
 * データがない期間は一般的な土日・祝日（日本の国民の祝日）で自動補完
 */
function api_getHolidayDays() {
  // M_CALENDARからデータ取得
  var calMap = {};
  var calCoveredDates = {};
  var sh = getDb_().getSheetByName(SHEET.CALENDAR);
  if (sh && sh.getLastRow() >= 2) {
    var values = sh.getDataRange().getValues();
    var H = values[0].map(function(h) { return normalize_(h); });
    var dateIdx = H.indexOf('日付');
    var kubunIdx = H.indexOf('区分');
    if (dateIdx >= 0 && kubunIdx >= 0) {
      for (var r = 1; r < values.length; r++) {
        var d = values[r][dateIdx];
        if (!d) continue;
        if (!(d instanceof Date)) d = new Date(d);
        if (isNaN(d.getTime())) continue;
        var key = fmtDate_(d);
        calMap[key] = normalize_(values[r][kubunIdx]);
        calCoveredDates[key] = true;
      }
    }
  }

  // 日本の祝日マップ（一般カレンダー補完用）
  var publicHolidays = getJapanesePublicHolidays_();

  // 年度範囲を算出（年度初め〜今日）
  var today = new Date();
  var _m = today.getMonth() + 1, _d = today.getDate();
  var fy = (_m < 3 || (_m === 3 && _d < 16)) ? today.getFullYear() - 1 : today.getFullYear();
  var startDate = new Date(fy, 2, 16); // 3/16

  var result = [];
  var cursor = new Date(startDate);
  while (cursor <= today) {
    var key = fmtDate_(cursor);
    var dow = cursor.getDay();

    if (calCoveredDates[key]) {
      // M_CALENDARにデータあり → カレンダーの区分に従う
      var kubun = calMap[key];
      if (kubun === '休日' || kubun === '祝日') {
        result.push({ date: key, label: kubun });
      } else if (kubun === '出勤土曜') {
        // 出勤土曜は勤務日 → 候補に含めない
      } else if (dow === 0 || dow === 6) {
        result.push({ date: key, label: dow === 0 ? '日曜' : '土曜' });
      }
    } else {
      // M_CALENDARにデータなし → 一般カレンダーで補完
      if (publicHolidays[key]) {
        result.push({ date: key, label: '祝日（' + publicHolidays[key] + '）' });
      } else if (dow === 0) {
        result.push({ date: key, label: '日曜' });
      } else if (dow === 6) {
        result.push({ date: key, label: '土曜' });
      }
      // 振替休日チェック（祝日が日曜の場合、月曜が振替休日）
      // → publicHolidaysに含めて対応済み
    }

    cursor.setDate(cursor.getDate() + 1);
  }
  return result;
}

/**
 * 指定年の日本の国民の祝日マップを返す（日付文字列 → 祝日名）
 * M_CALENDARにデータがない期間の補完用
 */
function getJapanesePublicHolidays_() {
  var today = new Date();
  var _m2 = today.getMonth() + 1, _d2 = today.getDate();
  var fy = (_m2 < 3 || (_m2 === 3 && _d2 < 16)) ? today.getFullYear() - 1 : today.getFullYear();
  // 年度内に含まれる年（3/16〜12月=fy, 1月〜3/15=fy+1）
  var years = [fy, fy + 1];

  var holidays = {};
  for (var yi = 0; yi < years.length; yi++) {
    var y = years[yi];
    var h = getHolidaysForYear_(y);
    for (var key in h) {
      holidays[key] = h[key];
    }
  }

  // 振替休日を追加（祝日が日曜の場合、翌日以降の最初の平日が振替休日）
  var keys = Object.keys(holidays).sort();
  for (var i = 0; i < keys.length; i++) {
    var dt = new Date(keys[i] + 'T00:00:00');
    if (dt.getDay() === 0) { // 日曜
      var sub = new Date(dt);
      sub.setDate(sub.getDate() + 1);
      while (holidays[fmtDate_(sub)]) {
        sub.setDate(sub.getDate() + 1);
      }
      holidays[fmtDate_(sub)] = '振替休日';
    }
  }

  // 国民の休日（祝日と祝日に挟まれた平日）
  for (var i = 0; i < keys.length - 1; i++) {
    var d1 = new Date(keys[i] + 'T00:00:00');
    var d2 = new Date(keys[i + 1] + 'T00:00:00');
    var diff = (d2 - d1) / 86400000;
    if (diff === 2) {
      var between = new Date(d1);
      between.setDate(between.getDate() + 1);
      var bKey = fmtDate_(between);
      if (!holidays[bKey] && between.getDay() !== 0 && between.getDay() !== 6) {
        holidays[bKey] = '国民の休日';
      }
    }
  }

  return holidays;
}

/**
 * 指定年の固定祝日＋ハッピーマンデー＋春分・秋分を返す
 */
function getHolidaysForYear_(y) {
  var h = {};
  function add(m, d, name) { h[y + '-' + z2(m) + '-' + z2(d)] = name; }
  function z2(n) { return n < 10 ? '0' + n : '' + n; }
  // n番目の月曜日を求める
  function nthMonday(month, n) {
    var d = new Date(y, month - 1, 1);
    var count = 0;
    while (count < n) {
      if (d.getDay() === 1) count++;
      if (count < n) d.setDate(d.getDate() + 1);
    }
    return d.getDate();
  }

  // 固定祝日
  add(1, 1, '元日');
  add(2, 11, '建国記念の日');
  add(2, 23, '天皇誕生日');
  add(4, 29, '昭和の日');
  add(5, 3, '憲法記念日');
  add(5, 4, 'みどりの日');
  add(5, 5, 'こどもの日');
  add(8, 11, '山の日');
  add(11, 3, '文化の日');
  add(11, 23, '勤労感謝の日');

  // ハッピーマンデー
  add(1, nthMonday(1, 2), '成人の日');
  add(7, nthMonday(7, 3), '海の日');
  add(9, nthMonday(9, 3), '敬老の日');
  add(10, nthMonday(10, 2), 'スポーツの日');

  // 春分の日（概算: 20 or 21）
  var shunbun = Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
  add(3, shunbun, '春分の日');

  // 秋分の日（概算: 22 or 23）
  var shubun = Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
  add(9, shubun, '秋分の日');

  return h;
}

/**
 * M_CALENDARから勤務日一覧を返す（今日〜年度末）
 * 勤務日 = 平日(月〜金)で祝日でない日 + 出勤土曜
 */
function api_getWorkingDays() {
  var sh = getDb_().getSheetByName(SHEET.CALENDAR);
  if (!sh || sh.getLastRow() < 2) return [];

  var values = sh.getDataRange().getValues();
  var H = values[0].map(function(h) { return normalize_(h); });
  var dateIdx = H.indexOf('日付');
  var kubunIdx = H.indexOf('区分');
  if (dateIdx < 0 || kubunIdx < 0) return [];

  // カレンダーデータをMapに格納（日付文字列 → 区分）
  var calMap = {};
  for (var r = 1; r < values.length; r++) {
    var d = values[r][dateIdx];
    if (!d) continue;
    if (!(d instanceof Date)) d = new Date(d);
    if (isNaN(d.getTime())) continue;
    var key = fmtDate_(d);
    calMap[key] = normalize_(values[r][kubunIdx]);
  }

  // 今日〜年度末の範囲で勤務日を算出
  var today = new Date();
  var _m3 = today.getMonth() + 1, _d3 = today.getDate();
  var fy = (_m3 < 3 || (_m3 === 3 && _d3 < 16)) ? today.getFullYear() - 1 : today.getFullYear();
  var endDate = new Date(fy + 1, 2, 15); // 年度末 3/15

  var result = [];
  var cursor = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  while (cursor <= endDate) {
    var key = fmtDate_(cursor);
    var dow = cursor.getDay(); // 0=日, 6=土
    var kubun = calMap[key] || '';

    if (kubun === '休日' || kubun === '祝日') {
      // 休日・祝日 → スキップ
    } else if (kubun === '出勤土曜') {
      // 出勤土曜 → 勤務日
      result.push(key);
    } else if (dow >= 1 && dow <= 5) {
      // 平日（月〜金）でカレンダーに載っていない → 勤務日
      result.push(key);
    }
    // それ以外（通常の土日でカレンダーに未登録） → スキップ

    cursor.setDate(cursor.getDate() + 1);
  }
  return result;
}

/**
 * 過去日（3/16〜今日）の勤務日一覧を月単位で返す
 * 休日/祝日を除外（api_getWorkingDays と同じロジックで判定）
 * @param {string|number} yearMonth 'yyyy-MM' 文字列または 'yyyyMM' 数値文字列（省略時は今月）
 * @return {Array<{date:string, dow:number, label:string}>}
 */
function api_getPastWorkingDays(yearMonth) {
  var today = new Date();
  var todayKey = fmtDate_(today);

  // 対象月の決定: 'yyyy-MM' / 'yyyyMM' / 数値 いずれも許容
  var ym = normalize_(yearMonth == null ? '' : String(yearMonth));
  var y, m;
  if (/^\d{4}-\d{2}$/.test(ym)) {
    y = parseInt(ym.substring(0, 4), 10);
    m = parseInt(ym.substring(5, 7), 10);
  } else if (/^\d{6}$/.test(ym)) {
    // 将来互換: yyyyMM 形式（数値文字列）
    y = parseInt(ym.substring(0, 4), 10);
    m = parseInt(ym.substring(4, 6), 10);
  } else {
    y = today.getFullYear();
    m = today.getMonth() + 1;
  }

  // 範囲: 今年度の3/16 および 月初 のうち新しい方 〜 今日 および 月末 のうち古い方
  var fyStart = getFiscalYearStartDate_(today);
  var monthStart = new Date(y, m - 1, 1, 0, 0, 0);
  var monthEnd = new Date(y, m, 0, 23, 59, 59);
  var rangeStart = monthStart > fyStart ? monthStart : fyStart;
  var rangeEndDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var rangeEnd = monthEnd < rangeEndDate ? monthEnd : rangeEndDate;
  if (rangeStart > rangeEnd) return [];

  // M_CALENDAR 読込
  var calMap = {};
  var sh = getDb_().getSheetByName(SHEET.CALENDAR);
  if (sh && sh.getLastRow() >= 2) {
    var values = sh.getDataRange().getValues();
    var H = values[0].map(function(h) { return normalize_(h); });
    var dateIdx = H.indexOf('日付');
    var kubunIdx = H.indexOf('区分');
    if (dateIdx >= 0 && kubunIdx >= 0) {
      for (var r = 1; r < values.length; r++) {
        var d = values[r][dateIdx];
        if (!d) continue;
        if (!(d instanceof Date)) d = new Date(d);
        if (isNaN(d.getTime())) continue;
        calMap[fmtDate_(d)] = normalize_(values[r][kubunIdx]);
      }
    }
  }

  // 祝日補完マップ（M_CALENDAR未登録期間向け）
  var publicHolidays = getJapanesePublicHolidays_();

  var dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  var result = [];
  var cursor = new Date(rangeStart);
  while (cursor <= rangeEnd) {
    var key = fmtDate_(cursor);
    if (key > todayKey) break;
    var dow = cursor.getDay();
    var kubun = calMap[key] || '';
    var isWorkDay = false;
    if (kubun === '休日' || kubun === '祝日') {
      isWorkDay = false;
    } else if (kubun === '出勤土曜') {
      isWorkDay = true;
    } else if (dow >= 1 && dow <= 5) {
      // 平日: M_CALENDAR未登録でも祝日マップに該当すれば除外
      if (publicHolidays[key]) isWorkDay = false;
      else isWorkDay = true;
    } else {
      isWorkDay = false;
    }
    if (isWorkDay) {
      result.push({
        date: key,
        dow: dow,
        label: key + '（' + dowNames[dow] + '）',
      });
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return result;
}
