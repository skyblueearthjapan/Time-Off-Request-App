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
 * 外部SS（残業・休日出勤申請app）の StampMap シートを参照
 */
function getStampFileId_(email) {
  if (!email) return '';
  email = email.toLowerCase().trim();

  var settings = getSettings_();
  var sourceId = normalize_(settings['CALENDAR_SOURCE_SSID']);
  if (!sourceId) sourceId = '1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ';

  try {
    var ss = SpreadsheetApp.openById(sourceId);
    var sh = ss.getSheetByName('StampMap');
    if (!sh || sh.getLastRow() < 2) return '';

    var data = sh.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var rowEmail = String(data[r][0] || '').toLowerCase().trim();
      if (rowEmail === email) {
        return String(data[r][1] || '').trim();
      }
    }
  } catch (e) {
    console.error('StampMap読取エラー: ' + e.message);
  }
  return '';
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
  var fy = today.getMonth() < 3 ? today.getFullYear() - 1 : today.getFullYear();
  var endDate = new Date(fy + 1, 2, 31); // 年度末 3/31

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
