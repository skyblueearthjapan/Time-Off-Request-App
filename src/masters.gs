// ====== マスタ読込（部署/作業員/LOOKUP） ======

/**
 * 部署一覧を取得（M_DEPTから、IS_ACTIVE=1のみ、SORT順）
 */
function loadDeptList_() {
  var sh = requireSheet_(SHEET.DEPT);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var H = values[0].map(function(h) { return normalize_(h); });
  var idx = {
    deptId: H.indexOf('部署ID'),
    deptName: H.indexOf('部署名'),
    active: H.indexOf('IS_ACTIVE'),
    sort: H.indexOf('SORT'),
  };
  if (idx.deptId < 0 || idx.deptName < 0) {
    throw new Error('M_DEPTのヘッダが想定と違います（部署ID/部署名）。');
  }

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    // IS_ACTIVE フィルタ（列がない場合は全員有効）
    if (idx.active >= 0) {
      var active = normalize_(row[idx.active]);
      if (active !== '' && active !== '1' && active.toLowerCase() !== 'true') continue;
    }
    out.push({
      deptId: normalize_(row[idx.deptId]),
      deptName: normalize_(row[idx.deptName]),
      sort: idx.sort >= 0 ? Number(row[idx.sort] || 9999) : r,
    });
  }
  out.sort(function(a, b) { return a.sort - b.sort; });
  return out;
}

/**
 * 指定部署の作業員一覧を取得（M_WORKERから）
 */
function loadWorkersByDept_(deptId) {
  var sh = requireSheet_(SHEET.WORKER);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var H = values[0].map(function(h) { return normalize_(h); });
  var idx = {
    workerId: H.indexOf('作業員ID'),
    workerName: H.indexOf('氏名'),
    dept: H.indexOf('部署ID'),
    active: H.indexOf('IS_ACTIVE'),
    sort: H.indexOf('SORT'),
    email: H.findIndex(function(h) { return h.indexOf('Googleアカウント') >= 0; }),
  };
  if (idx.workerId < 0 || idx.workerName < 0 || idx.dept < 0) {
    throw new Error('M_WORKERのヘッダが想定と違います（作業員ID/氏名/部署ID）。');
  }

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var rowDeptId = normalize_(row[idx.dept]);
    if (deptId && rowDeptId !== deptId) continue;

    if (idx.active >= 0) {
      var active = normalize_(row[idx.active]);
      if (active !== '' && active !== '1' && active.toLowerCase() !== 'true') continue;
    }

    out.push({
      workerId: normalize_(row[idx.workerId]),
      workerName: normalize_(row[idx.workerName]),
      deptId: rowDeptId,
      sort: idx.sort >= 0 ? Number(row[idx.sort] || 9999) : r,
    });
  }
  out.sort(function(a, b) { return a.sort - b.sort; });
  return out;
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
    workerId: H.indexOf('作業員ID'),
    workerName: H.indexOf('氏名'),
    dept: H.indexOf('部署ID'),
    email: H.findIndex(function(h) { return h.indexOf('Googleアカウント') >= 0; }),
    active: H.indexOf('IS_ACTIVE'),
  };
  if (idx.email < 0) return null;

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (idx.active >= 0) {
      var active = normalize_(row[idx.active]);
      if (active !== '' && active !== '1' && active.toLowerCase() !== 'true') continue;
    }
    if (normalize_(row[idx.email]).toLowerCase() === email) {
      // 部署名も取得
      var deptId = normalize_(row[idx.dept]);
      var depts = loadDeptList_();
      var deptObj = depts.find(function(d) { return d.deptId === deptId; });
      return {
        workerId: normalize_(row[idx.workerId]),
        workerName: normalize_(row[idx.workerName]),
        deptId: deptId,
        deptName: deptObj ? deptObj.deptName : deptId,
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
 * 戻り値: { leaveKubun: [...], leaveTypeFull: [...], leaveTypeHalf: [...], specialReasons: [...] }
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
