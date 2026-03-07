// ====== 休暇申請CRUD ======

/**
 * REQ_ID採番: LV-YYYYMMDD-NNN
 */
function generateReqId_() {
  var now = new Date();
  var ymd = fmtDate_(now, 'yyyyMMdd');
  var prefix = 'LV-' + ymd + '-';

  var sh = requireSheet_(SHEET.LEAVE_REQUEST);
  var lastRow = sh.getLastRow();
  var maxSeq = 0;

  if (lastRow >= 2) {
    var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var id = String(ids[i][0] || '');
      if (id.indexOf(prefix) === 0) {
        var seq = parseInt(id.substring(prefix.length), 10);
        if (seq > maxSeq) maxSeq = seq;
      }
    }
  }

  var nextSeq = String(maxSeq + 1);
  while (nextSeq.length < 3) nextSeq = '0' + nextSeq;
  return prefix + nextSeq;
}

/**
 * 年度算出: FY_START_MONTH(デフォルト4)基準
 */
function computeFiscalYear_(dateObj) {
  var d = dateObj || new Date();
  var year = d.getFullYear();
  var month = d.getMonth() + 1; // 1-12
  // 4月始まり: 1-3月は前年度
  if (month < 4) {
    return year - 1;
  }
  return year;
}

/**
 * 申請登録
 * data: { deptId, deptName, workerId, workerName, leaveKubun, halfDayType,
 *         leaveDate, leaveType, substituteDate, specialReason, paidDetail, additionalDetail }
 */
function api_submitLeaveRequest(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var reqId = generateReqId_();
    var now = new Date();
    var leaveDate = new Date(data.leaveDate);
    var fy = computeFiscalYear_(leaveDate);

    var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
    var sh = info.sh;
    var idx = info.idx;

    // ヘッダが無い場合は作成
    if (info.header.length === 0) {
      var headers = [
        'REQ_ID', '年度', '部署ID', '部署名', '作業員ID', '作業員名',
        '休暇区分', '半日区分', '休暇日', '休暇種類',
        '振替元出勤日', '特別理由', '有給詳細', '追加詳細',
        '申請日時', '承認日時', '承認状態', 'サイン画像URL', 'PDF_URL', 'PDF_FILE_ID'
      ];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      SpreadsheetApp.flush();
      // ヘッダインデックス再取得
      info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
      idx = info.idx;
    }

    // ヘッダ位置ベースで行データを構築（列の順序に依存しない）
    var row = new Array(info.header.length);
    for (var c = 0; c < row.length; c++) row[c] = '';

    function setCol(name, value) {
      if (idx[name] !== undefined) row[idx[name]] = value;
    }

    setCol('REQ_ID', reqId);
    setCol('年度', fy);
    setCol('部署ID', data.deptId || '');
    setCol('部署名', data.deptName || '');
    setCol('作業員ID', data.workerId || '');
    setCol('作業員名', data.workerName || '');
    setCol('休暇区分', data.leaveKubun || '');
    setCol('半日区分', data.halfDayType || '');
    setCol('休暇日', leaveDate);
    setCol('休暇種類', data.leaveType || '');
    setCol('振替元出勤日', data.substituteDate ? new Date(data.substituteDate) : '');
    setCol('特別理由', data.specialReason || '');
    setCol('有給詳細', data.paidDetail || '');
    setCol('追加詳細', data.additionalDetail || '');
    setCol('申請日時', now);
    setCol('承認状態', STATUS.SUBMITTED);
    setCol('作成者メール', Session.getActiveUser().getEmail() || '');

    // appendRowではなくgetRange().setValues()で確実に書込み
    // （空行が大量にある場合のappendRow位置ずれ防止）
    var writeRow = sh.getLastRow() + 1;
    sh.getRange(writeRow, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();

    Logger.log('申請書込み完了: ' + reqId + ' at row ' + writeRow);
    return { ok: true, reqId: reqId };
  } finally {
    lock.releaseLock();
  }
}

/**
 * REQ_IDで申請データ取得
 */
function api_getLeaveRequestById(reqId) {
  if (!reqId) return null;

  // リトライ付き（書込み直後の読取り遅延対策）
  var MAX_RETRY = 5;
  for (var attempt = 0; attempt < MAX_RETRY; attempt++) {
    if (attempt > 0) {
      SpreadsheetApp.flush();
      Utilities.sleep(1500);
    }

    var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
    var sh = info.sh;
    var idx = info.idx;
    var lastRow = sh.getLastRow();

    console.log('attempt ' + (attempt + 1) + ': lastRow=' + lastRow + ', header=' + JSON.stringify(info.header));

    if (lastRow < 2) continue;

    // REQ_IDカラムを検出（直接 or ヘッダ部分一致）
    var reqIdCol = idx['REQ_ID'];
    if (reqIdCol === undefined) {
      for (var h = 0; h < info.header.length; h++) {
        if (info.header[h].toUpperCase().indexOf('REQ') >= 0) {
          reqIdCol = h; break;
        }
      }
    }
    if (reqIdCol === undefined) {
      console.log('REQ_IDカラムなし。header=' + info.header.join(','));
      continue;
    }

    var readCols = Math.max(info.header.length, sh.getLastColumn());
    var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();

    // 全REQ_IDをログ出力（デバッグ用）
    var allIds = [];
    for (var i = 0; i < values.length; i++) {
      var cellVal = values[i][reqIdCol];
      allIds.push(String(cellVal));
    }
    console.log('検索対象reqId=[' + reqId + '], シート内REQ_ID一覧=' + JSON.stringify(allIds) + ', reqIdCol=' + reqIdCol);

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var cellReqId = normalize_(row[reqIdCol]);
      if (cellReqId === reqId) {
        console.log('FOUND at row ' + (i + 2));
        var result = {
          reqId: cellReqId,
          fiscalYear: String(row[idx['年度']] || ''),
          deptId: normalize_(row[idx['部署ID']]),
          deptName: normalize_(row[idx['部署名']]),
          workerId: normalize_(row[idx['作業員ID']]),
          workerName: normalize_(row[idx['作業員名']]),
          leaveKubun: normalize_(row[idx['休暇区分']]),
          halfDayType: normalize_(row[idx['半日区分']]),
          leaveDate: row[idx['休暇日']] instanceof Date ? fmtDate_(row[idx['休暇日']]) : String(row[idx['休暇日']] || ''),
          leaveType: normalize_(row[idx['休暇種類']]),
          substituteDate: row[idx['振替元出勤日']] instanceof Date ? fmtDate_(row[idx['振替元出勤日']]) : String(row[idx['振替元出勤日']] || ''),
          specialReason: normalize_(row[idx['特別理由']]),
          paidDetail: normalize_(row[idx['有給詳細']]),
          additionalDetail: normalize_(row[idx['追加詳細']]),
          submittedAt: row[idx['申請日時']] instanceof Date ? fmtDate_(row[idx['申請日時']], 'yyyy-MM-dd HH:mm:ss') : String(row[idx['申請日時']] || ''),
          approvedAt: row[idx['承認日時']] instanceof Date ? fmtDate_(row[idx['承認日時']], 'yyyy-MM-dd HH:mm:ss') : String(row[idx['承認日時']] || ''),
          status: normalize_(row[idx['承認状態']]),
          signImageUrl: normalize_(row[idx['サイン画像URL']]),
          pdfUrl: normalize_(row[idx['PDF_URL']]),
          rowNo: i + 2,
        };
        console.log('返却データ: ' + JSON.stringify(result));
        return result;
      }
    }
  }

  console.log('全リトライ失敗。reqId=' + reqId);
  return null;
}

/**
 * デバッグ用：T_LEAVE_REQUESTの状態を確認（GASエディタから直接実行）
 */
function debug_checkLeaveRequestSheet() {
  var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var sh = info.sh;
  var lastRow = sh.getLastRow();
  var msg = '=== T_LEAVE_REQUEST デバッグ ===\n';
  msg += 'シート名: ' + sh.getName() + '\n';
  msg += 'ヘッダ: ' + JSON.stringify(info.header) + '\n';
  msg += 'idx keys: ' + JSON.stringify(Object.keys(info.idx)) + '\n';
  msg += 'lastRow: ' + lastRow + '\n';

  if (lastRow >= 2) {
    var reqIdCol = info.idx['REQ_ID'];
    msg += 'REQ_ID col index: ' + reqIdCol + '\n';
    var vals = sh.getRange(2, 1, lastRow - 1, info.header.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      var raw = vals[i][reqIdCol !== undefined ? reqIdCol : 0];
      msg += 'Row ' + (i + 2) + ': REQ_ID=[' + raw + '] type=' + typeof raw + '\n';
    }
  }

  console.log(msg);
  Logger.log(msg);
  return msg;
}

/**
 * 休暇予定一覧（本日以降 & leaveDate >= 今日0:00）
 */
function api_getUpcomingLeaves(deptId, workerId) {
  var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var readCols = Math.max(info.header.length, sh.getLastColumn());
  var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    // 空行スキップ（REQ_IDが空の行は無視）
    var rowReqId = normalize_(row[idx['REQ_ID']]);
    if (!rowReqId) continue;

    var status = normalize_(row[idx['承認状態']]);
    if (status === STATUS.CANCELED) continue;

    var rowDeptId = normalize_(row[idx['部署ID']]);
    var rowWorkerId = normalize_(row[idx['作業員ID']]);

    // フィルタ
    if (deptId && rowDeptId !== deptId) continue;
    if (workerId && rowWorkerId !== workerId) continue;

    var leaveDate = row[idx['休暇日']];
    if (leaveDate instanceof Date) {
      if (leaveDate < today) continue;
    }

    out.push({
      reqId: normalize_(row[idx['REQ_ID']]),
      deptName: normalize_(row[idx['部署名']]),
      workerName: normalize_(row[idx['作業員名']]),
      leaveKubun: normalize_(row[idx['休暇区分']]),
      halfDayType: normalize_(row[idx['半日区分']]),
      leaveDate: leaveDate instanceof Date ? fmtDate_(leaveDate, 'yyyy/MM/dd') : '',
      leaveType: normalize_(row[idx['休暇種類']]),
      status: status,
    });
  }

  // 日付昇順
  out.sort(function(a, b) {
    return (a.leaveDate || '').localeCompare(b.leaveDate || '');
  });

  return out;
}

/**
 * 承認処理（sign.htmlから呼ばれる）
 * @param {string} reqId - 申請ID
 * @param {string} signBase64 - 手書きサインBase64（電子印がない場合のフォールバック）
 * @param {string} approverEmail - 承認者メールアドレス（StampMap検索用）
 */
function api_approveLeaveRequest(reqId, signBase64, approverEmail) {
  if (!reqId) throw new Error('REQ_IDが指定されていません。');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
    var sh = info.sh;
    var idx = info.idx;
    var lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('申請データがありません。');

    var readCols = Math.max(info.header.length, sh.getLastColumn());
    var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();
    var targetRow = -1;

    for (var i = 0; i < values.length; i++) {
      if (normalize_(values[i][idx['REQ_ID']]) === reqId) {
        targetRow = i + 2;
        break;
      }
    }
    if (targetRow < 0) throw new Error('申請が見つかりません: ' + reqId);

    var now = new Date();

    // ヘッダインデックス検証（エイリアス解決済み）
    var requiredCols = ['REQ_ID', '承認状態', '承認日時', 'サイン画像URL'];
    for (var rc = 0; rc < requiredCols.length; rc++) {
      if (idx[requiredCols[rc]] === undefined) {
        throw new Error('T_LEAVE_REQUEST のヘッダに "' + requiredCols[rc] + '" が見つかりません。ヘッダ: ' + info.header.join(', '));
      }
    }

    // 1. サイン画像をDriveに保存（PDFと同じ日付フォルダ）
    var leaveDate = values[targetRow - 2][idx['休暇日']];
    var signImageUrl = '';
    if (signBase64) {
      try {
        signImageUrl = saveSignImage_(reqId, signBase64, leaveDate);
      } catch (signErr) {
        console.error('サイン画像保存エラー（続行）: ' + signErr.message);
      }
    }

    // 2. ステータス更新
    sh.getRange(targetRow, idx['承認状態'] + 1).setValue(STATUS.APPROVED);
    sh.getRange(targetRow, idx['承認日時'] + 1).setValue(now);
    sh.getRange(targetRow, idx['サイン画像URL'] + 1).setValue(signImageUrl);
    SpreadsheetApp.flush();

    // 3. PDF生成（承認者メールを渡す）
    var pdfResult = null;
    try {
      pdfResult = generateLeavePdf_(reqId, approverEmail);
      if (pdfResult && pdfResult.ok) {
        if (idx['PDF_URL'] !== undefined) {
          sh.getRange(targetRow, idx['PDF_URL'] + 1).setValue(pdfResult.pdfUrl || '');
        }
        if (idx['PDF_FILE_ID'] !== undefined) {
          sh.getRange(targetRow, idx['PDF_FILE_ID'] + 1).setValue(pdfResult.pdfFileId || '');
        }
      }
    } catch (pdfErr) {
      console.error('PDF生成エラー（続行）: ' + pdfErr.message);
    }

    // 4. サマリー再構築
    try {
      rebuildLeaveSummary_();
    } catch (sumErr) {
      console.error('サマリー再構築エラー: ' + sumErr.message);
    }

    // 5. メール通知
    try {
      sendApprovalNotification_(reqId);
    } catch (mailErr) {
      console.error('メール通知エラー: ' + mailErr.message);
    }

    SpreadsheetApp.flush();
    return { ok: true, reqId: reqId };
  } finally {
    lock.releaseLock();
  }
}
