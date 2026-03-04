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
