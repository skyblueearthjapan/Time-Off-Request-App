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
  var day = d.getDate();
  // 3/16始まり: 1月〜3/15は前年度
  if (month < 3 || (month === 3 && day < 16)) {
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
        '申請日時', '承認日時', '承認状態', 'サイン画像URL', 'PDF_URL', 'PDF_FILE_ID',
        'APPROVED_BY1_EMAIL', 'APPROVED_BY2', 'APPROVED_BY2_EMAIL', 'APPROVED_AT2'
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
          approvedBy1Email: idx['APPROVED_BY1_EMAIL'] !== undefined ? normalize_(row[idx['APPROVED_BY1_EMAIL']]) : '',
          approvedBy2: idx['APPROVED_BY2'] !== undefined ? normalize_(row[idx['APPROVED_BY2']]) : '',
          approvedBy2Email: idx['APPROVED_BY2_EMAIL'] !== undefined ? normalize_(row[idx['APPROVED_BY2_EMAIL']]) : '',
          approvedAt2: idx['APPROVED_AT2'] !== undefined && row[idx['APPROVED_AT2']] instanceof Date ? fmtDate_(row[idx['APPROVED_AT2']], 'yyyy-MM-dd HH:mm:ss') : '',
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
 * 全員の申請中一覧（日付制限なし・フィルタなし）
 * TOP画面の「承認待ち一覧」セクション用
 */
function api_getAllPendingRequests() {
  var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var readCols = Math.max(info.header.length, sh.getLastColumn());
  var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowReqId = normalize_(row[idx['REQ_ID']]);
    if (!rowReqId) continue;

    var status = normalize_(row[idx['承認状態']]);
    if (status !== STATUS.SUBMITTED) continue;

    var leaveDate = row[idx['休暇日']];
    var submittedAt = row[idx['申請日時']];

    out.push({
      reqId: rowReqId,
      deptName: normalize_(row[idx['部署名']]),
      workerName: normalize_(row[idx['作業員名']]),
      leaveKubun: normalize_(row[idx['休暇区分']]),
      halfDayType: normalize_(row[idx['半日区分']]),
      leaveDate: leaveDate instanceof Date ? fmtDate_(leaveDate, 'yyyy/MM/dd') : '',
      leaveType: normalize_(row[idx['休暇種類']]),
      submittedAt: submittedAt instanceof Date ? fmtDate_(submittedAt, 'MM/dd HH:mm') : '',
      status: status,
    });
  }

  // 申請日時の新しい順（最新が上）
  out.sort(function(a, b) {
    return (b.submittedAt || '').localeCompare(a.submittedAt || '');
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
  lock.waitLock(60000);
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
    // 1次承認者メールを保存（2次承認PDF再生成用）
    if (idx['APPROVED_BY1_EMAIL'] !== undefined || idx['1次承認者メール'] !== undefined) {
      var col1 = idx['APPROVED_BY1_EMAIL'] !== undefined ? idx['APPROVED_BY1_EMAIL'] : idx['1次承認者メール'];
      sh.getRange(targetRow, col1 + 1).setValue(approverEmail || '');
    }
    SpreadsheetApp.flush();

    // 3. PDF生成（承認者メールを渡す）
    console.log('=== PDF生成開始 === reqId=' + reqId + ', approverEmail=' + approverEmail + ' (type=' + typeof approverEmail + ')');
    var pdfResult = null;
    try {
      pdfResult = generateLeavePdf_(reqId, approverEmail);
      console.log('PDF生成結果: ' + JSON.stringify(pdfResult));
      if (pdfResult && pdfResult.ok) {
        if (idx['PDF_URL'] !== undefined) {
          sh.getRange(targetRow, idx['PDF_URL'] + 1).setValue(pdfResult.pdfUrl || '');
        }
        if (idx['PDF_FILE_ID'] !== undefined) {
          sh.getRange(targetRow, idx['PDF_FILE_ID'] + 1).setValue(pdfResult.pdfFileId || '');
        }
      }
    } catch (pdfErr) {
      console.error('PDF生成エラー（続行）: ' + pdfErr.message + '\n' + pdfErr.stack);
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

    // 6. 台帳更新
    try {
      updateLedger_(reqId);
    } catch (ledgerErr) {
      console.error('台帳更新エラー: ' + ledgerErr.message);
    }

    return { ok: true, reqId: reqId };
  } finally {
    lock.releaseLock();
  }
}

// ====== 2次承認 ======

/**
 * 2次承認待ち一覧取得（総務管理画面用）
 * @return {Object[]} 1次承認済み・2次未承認の申請一覧
 */
function api_getPending2ndApprovals() {
  if (!isSomu_() && !isAdmin_()) throw new Error('権限がありません（総務または管理者のみ）');

  var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var readCols = Math.max(info.header.length, sh.getLastColumn());
  var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();
  var out = [];

  // APPROVED_AT2 列のインデックス取得
  var at2Col = idx['APPROVED_AT2'] !== undefined ? idx['APPROVED_AT2'] : idx['2次承認日時'];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var reqId = normalize_(row[idx['REQ_ID']]);
    if (!reqId) continue;

    var status = normalize_(row[idx['承認状態']]);
    if (status !== STATUS.APPROVED) continue;

    // 2次承認済みはスキップ
    if (at2Col !== undefined && row[at2Col] && String(row[at2Col]).trim() !== '') continue;

    var leaveDate = row[idx['休暇日']];
    out.push({
      reqId: reqId,
      deptName: normalize_(row[idx['部署名']]),
      workerName: normalize_(row[idx['作業員名']]),
      leaveDate: leaveDate instanceof Date ? fmtDate_(leaveDate, 'yyyy/MM/dd') : String(leaveDate || ''),
      leaveKubun: normalize_(row[idx['休暇区分']]),
      leaveType: normalize_(row[idx['休暇種類']]),
      approvedAt: row[idx['承認日時']] instanceof Date ? fmtDate_(row[idx['承認日時']], 'yyyy/MM/dd HH:mm') : '',
    });
  }

  // 休暇日昇順
  out.sort(function(a, b) { return (a.leaveDate || '').localeCompare(b.leaveDate || ''); });
  return out;
}

/**
 * 2次承認実行（総務管理画面から呼ばれる）
 * @param {string} reqId - 申請ID
 * @param {string} approver2Email - 2次承認者メール（電子印取得用）
 * @return {Object} { ok: boolean }
 */
function api_approveLeaveRequest2(reqId, approver2Email) {
  if (!reqId) throw new Error('REQ_IDが指定されていません。');
  if (!isSomu_() && !isAdmin_()) throw new Error('権限がありません（総務または管理者のみ）');

  var lock = LockService.getScriptLock();
  lock.waitLock(60000);
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

    var status = normalize_(values[targetRow - 2][idx['承認状態']]);
    if (status !== STATUS.APPROVED) throw new Error('1次承認されていない申請です。');

    var now = new Date();

    // 2次承認者名を取得（M_STAMPの備考列から）
    var approver2Name = approver2Email || '';
    if (approver2Email) {
      try {
        var stampSh = getDb_().getSheetByName(SHEET.STAMP);
        if (stampSh) {
          var stampData = stampSh.getDataRange().getValues();
          for (var s = 1; s < stampData.length; s++) {
            if (normalize_(String(stampData[s][0] || '')).toLowerCase() === approver2Email.toLowerCase()) {
              approver2Name = normalize_(String(stampData[s][2] || '')) || approver2Email;
              break;
            }
          }
        }
      } catch (e) { /* ignore */ }
    }

    // 2次承認情報を書き込み（英語キーのみ。HEADER_ALIAS_でエイリアス解決済み）
    function setCol2(name, value) {
      var colIdx = idx[name];
      if (colIdx !== undefined) {
        sh.getRange(targetRow, colIdx + 1).setValue(value);
      }
    }
    setCol2('APPROVED_BY2', approver2Name);
    setCol2('APPROVED_BY2_EMAIL', approver2Email);
    setCol2('APPROVED_AT2', now);
    SpreadsheetApp.flush();

    // PDF再生成（2次承認者の電子印を追加）
    try {
      var approver1Email = '';
      var a1Col = idx['APPROVED_BY1_EMAIL'] !== undefined ? idx['APPROVED_BY1_EMAIL'] : idx['1次承認者メール'];
      if (a1Col !== undefined) {
        approver1Email = normalize_(values[targetRow - 2][a1Col]);
      }
      console.log('2次承認PDF再生成: reqId=' + reqId + ', approver1=' + approver1Email + ', approver2=' + approver2Email);

      // 既存PDFをゴミ箱へ
      var pdfFileIdCol = idx['PDF_FILE_ID'];
      if (pdfFileIdCol !== undefined) {
        var oldPdfId = normalize_(values[targetRow - 2][pdfFileIdCol]);
        if (oldPdfId) {
          try { DriveApp.getFileById(oldPdfId).setTrashed(true); } catch (e) { console.warn('旧PDF削除スキップ: ' + e.message); }
        }
      }

      // PDF再生成（2次承認者印付き）
      var pdfResult = generateLeavePdf_(reqId, approver1Email, approver2Email);
      if (pdfResult && pdfResult.ok) {
        if (idx['PDF_URL'] !== undefined) sh.getRange(targetRow, idx['PDF_URL'] + 1).setValue(pdfResult.pdfUrl || '');
        if (idx['PDF_FILE_ID'] !== undefined) sh.getRange(targetRow, idx['PDF_FILE_ID'] + 1).setValue(pdfResult.pdfFileId || '');
      }
    } catch (pdfErr) {
      console.error('2次承認PDF再生成エラー（続行）: ' + pdfErr.message);
    }

    SpreadsheetApp.flush();

    // 台帳更新（2次承認者情報を反映）
    try {
      updateLedger_(reqId);
    } catch (ledgerErr) {
      console.error('台帳更新エラー: ' + ledgerErr.message);
    }

    return { ok: true, reqId: reqId };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 2次承認一括実行（シート読込・ロック・スタンプ名を1回で処理）
 * PDF再生成はロック解放後に行い、タイムアウトを防止
 * @param {string[]} reqIds - 申請IDの配列
 * @param {string} approver2Email - 2次承認者メール
 * @return {Object} { ok: boolean, results: [...] }
 */
function api_approveLeaveRequest2Batch(reqIds, approver2Email) {
  if (!isSomu_() && !isAdmin_()) throw new Error('権限がありません');
  if (!reqIds || !reqIds.length) throw new Error('申請IDが指定されていません');
  if (reqIds.length > 20) throw new Error('一度に承認できるのは20件までです');

  // 2次承認者名を1回だけ取得
  var approver2Name = approver2Email || '';
  if (approver2Email) {
    try {
      var stampSh = getDb_().getSheetByName(SHEET.STAMP);
      if (stampSh) {
        var stampData = stampSh.getDataRange().getValues();
        for (var s = 1; s < stampData.length; s++) {
          if (normalize_(String(stampData[s][0] || '')).toLowerCase() === approver2Email.toLowerCase()) {
            approver2Name = normalize_(String(stampData[s][2] || '')) || approver2Email;
            break;
          }
        }
      }
    } catch (e) { /* ignore */ }
  }

  var now = new Date();
  var pdfTargets = []; // PDF再生成用の情報を収集
  var results = [];

  // ロック内でシート書き込みのみ（高速）
  var lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
    var sh = info.sh;
    var idx = info.idx;
    var lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('申請データがありません。');

    var readCols = Math.max(info.header.length, sh.getLastColumn());
    var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();

    // reqId→行番号マップ作成
    var reqIdToRow = {};
    for (var i = 0; i < values.length; i++) {
      var rid = normalize_(values[i][idx['REQ_ID']]);
      if (rid) reqIdToRow[rid] = i;
    }

    for (var j = 0; j < reqIds.length; j++) {
      var reqId = reqIds[j];
      try {
        var rowIdx = reqIdToRow[reqId];
        if (rowIdx === undefined) throw new Error('申請が見つかりません');
        var targetRow = rowIdx + 2;
        var status = normalize_(values[rowIdx][idx['承認状態']]);
        if (status !== STATUS.APPROVED) throw new Error('1次承認されていない申請です');

        // 2次承認情報を書き込み
        if (idx['APPROVED_BY2'] !== undefined) sh.getRange(targetRow, idx['APPROVED_BY2'] + 1).setValue(approver2Name);
        if (idx['APPROVED_BY2_EMAIL'] !== undefined) sh.getRange(targetRow, idx['APPROVED_BY2_EMAIL'] + 1).setValue(approver2Email);
        if (idx['APPROVED_AT2'] !== undefined) sh.getRange(targetRow, idx['APPROVED_AT2'] + 1).setValue(now);

        // PDF再生成用の情報を収集
        var approver1Email = '';
        var a1Col = idx['APPROVED_BY1_EMAIL'];
        if (a1Col !== undefined) approver1Email = normalize_(values[rowIdx][a1Col]);
        var oldPdfId = idx['PDF_FILE_ID'] !== undefined ? normalize_(values[rowIdx][idx['PDF_FILE_ID']]) : '';

        pdfTargets.push({ reqId: reqId, targetRow: targetRow, approver1Email: approver1Email, oldPdfId: oldPdfId });
        results.push({ reqId: reqId, ok: true });
      } catch (e) {
        results.push({ reqId: reqId, ok: false, error: e.message });
      }
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // ロック解放後にPDF再生成（時間がかかる処理）
  for (var p = 0; p < pdfTargets.length; p++) {
    var t = pdfTargets[p];
    try {
      // 既存PDFをゴミ箱へ
      if (t.oldPdfId) {
        try { DriveApp.getFileById(t.oldPdfId).setTrashed(true); } catch (e) { /* skip */ }
      }
      var pdfResult = generateLeavePdf_(t.reqId, t.approver1Email, approver2Email);
      if (pdfResult && pdfResult.ok) {
        var info2 = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
        if (info2.idx['PDF_URL'] !== undefined) info2.sh.getRange(t.targetRow, info2.idx['PDF_URL'] + 1).setValue(pdfResult.pdfUrl || '');
        if (info2.idx['PDF_FILE_ID'] !== undefined) info2.sh.getRange(t.targetRow, info2.idx['PDF_FILE_ID'] + 1).setValue(pdfResult.pdfFileId || '');
      }
    } catch (pdfErr) {
      console.error('バッチPDF再生成エラー（続行）: reqId=' + t.reqId + ' ' + pdfErr.message);
    }

    // 台帳更新
    try {
      updateLedger_(t.reqId);
    } catch (ledgerErr) {
      console.error('バッチ台帳更新エラー（続行）: reqId=' + t.reqId + ' ' + ledgerErr.message);
    }
  }

  return { ok: true, results: results };
}

// ====== 申請編集 API ======
// TOP画面から呼ばれる。休暇日・休暇区分（全日/半日）・半日区分（午前/午後）を変更可能。
// 承認済みの場合は申請中に戻す（再承認必要）。PDF削除・台帳削除も行う。

function api_editLeaveRequest(reqId, patch) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var req = api_getLeaveRequestById(reqId);
    if (!req) throw new Error('申請が見つかりません。');
    if (req.status === STATUS.CANCELED || req.status === STATUS.REJECTED) {
      throw new Error('取消/却下済みの申請は編集できません。');
    }

    var leaveDate = normalize_(patch.leaveDate || '');
    var leaveKubun = normalize_(patch.leaveKubun || '');
    var halfDayType = normalize_(patch.halfDayType || '');

    // バリデーション
    if (leaveDate && !/^\d{4}-\d{2}-\d{2}$/.test(leaveDate)) {
      throw new Error('休暇日の形式が不正です: ' + leaveDate);
    }
    if (leaveKubun && leaveKubun !== LEAVE_KUBUN.FULL_DAY && leaveKubun !== LEAVE_KUBUN.HALF_DAY) {
      throw new Error('休暇区分が不正です: ' + leaveKubun);
    }
    if (leaveKubun === LEAVE_KUBUN.HALF_DAY && halfDayType !== HALF_DAY_TYPE.AM && halfDayType !== HALF_DAY_TYPE.PM) {
      throw new Error('半日区分を選択してください。');
    }

    var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
    var sh = info.sh;
    var idx = info.idx;
    var rowNo = req.rowNo;

    // 休暇日の更新
    if (leaveDate) {
      sh.getRange(rowNo, idx['休暇日'] + 1).setValue(leaveDate);
      // 年度も再計算
      var newFy = computeFiscalYear_(new Date(leaveDate + 'T00:00:00+09:00'));
      if (idx['年度'] !== undefined) sh.getRange(rowNo, idx['年度'] + 1).setValue(newFy);
    }

    // 休暇区分の更新
    if (leaveKubun) {
      sh.getRange(rowNo, idx['休暇区分'] + 1).setValue(leaveKubun);
      if (leaveKubun === LEAVE_KUBUN.FULL_DAY) {
        // 全日休の場合は半日区分をクリア
        sh.getRange(rowNo, idx['半日区分'] + 1).setValue('');
      } else {
        sh.getRange(rowNo, idx['半日区分'] + 1).setValue(halfDayType);
      }
    }

    // 承認済みの場合は申請中に戻す
    var newStatus = req.status;
    if (req.status === STATUS.APPROVED) {
      newStatus = STATUS.SUBMITTED;
      sh.getRange(rowNo, idx['承認状態'] + 1).setValue(STATUS.SUBMITTED);
      // 承認関連フィールドをクリア
      if (idx['承認日時'] !== undefined) sh.getRange(rowNo, idx['承認日時'] + 1).setValue('');
      if (idx['APPROVED_BY1_EMAIL'] !== undefined) sh.getRange(rowNo, idx['APPROVED_BY1_EMAIL'] + 1).setValue('');
      if (idx['サイン画像URL'] !== undefined) sh.getRange(rowNo, idx['サイン画像URL'] + 1).setValue('');
      if (idx['APPROVED_BY2'] !== undefined) sh.getRange(rowNo, idx['APPROVED_BY2'] + 1).setValue('');
      if (idx['APPROVED_BY2_EMAIL'] !== undefined) sh.getRange(rowNo, idx['APPROVED_BY2_EMAIL'] + 1).setValue('');
      if (idx['APPROVED_AT2'] !== undefined) sh.getRange(rowNo, idx['APPROVED_AT2'] + 1).setValue('');

      // PDF削除
      var pdfFileId = idx['PDF_FILE_ID'] !== undefined ? normalize_(sh.getRange(rowNo, idx['PDF_FILE_ID'] + 1).getValue()) : '';
      if (pdfFileId) {
        try { DriveApp.getFileById(pdfFileId).setTrashed(true); } catch (e) { console.warn('旧PDF削除スキップ: ' + e.message); }
        if (idx['PDF_URL'] !== undefined) sh.getRange(rowNo, idx['PDF_URL'] + 1).setValue('');
        if (idx['PDF_FILE_ID'] !== undefined) sh.getRange(rowNo, idx['PDF_FILE_ID'] + 1).setValue('');
      }

      // 台帳から行削除
      try { removeLedgerRow_(reqId, req.leaveDate); } catch (e) { console.warn('台帳削除スキップ: ' + e.message); }
    }

    SpreadsheetApp.flush();

    // サマリー再構築（有給カウントが変わる可能性）
    try { rebuildLeaveSummary_(); } catch (e) { console.warn('サマリー再構築スキップ: ' + e.message); }

    return { ok: true, newStatus: newStatus };
  } finally {
    lock.releaseLock();
  }
}

// ====== 申請削除（キャンセル）API ======

function api_deleteLeaveRequest(reqId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var req = api_getLeaveRequestById(reqId);
    if (!req) throw new Error('申請が見つかりません。');
    if (req.status === STATUS.CANCELED) throw new Error('既に取消済みです。');

    var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
    var sh = info.sh;
    var idx = info.idx;
    var rowNo = req.rowNo;

    // ステータスを取消に変更
    sh.getRange(rowNo, idx['承認状態'] + 1).setValue(STATUS.CANCELED);

    // PDF削除
    var pdfFileId = idx['PDF_FILE_ID'] !== undefined ? normalize_(sh.getRange(rowNo, idx['PDF_FILE_ID'] + 1).getValue()) : '';
    if (pdfFileId) {
      try { DriveApp.getFileById(pdfFileId).setTrashed(true); } catch (e) { console.warn('旧PDF削除スキップ: ' + e.message); }
      if (idx['PDF_URL'] !== undefined) sh.getRange(rowNo, idx['PDF_URL'] + 1).setValue('');
      if (idx['PDF_FILE_ID'] !== undefined) sh.getRange(rowNo, idx['PDF_FILE_ID'] + 1).setValue('');
    }

    SpreadsheetApp.flush();

    // 台帳から行削除
    try { removeLedgerRow_(reqId, req.leaveDate); } catch (e) { console.warn('台帳削除スキップ: ' + e.message); }

    // サマリー再構築
    try { rebuildLeaveSummary_(); } catch (e) { console.warn('サマリー再構築スキップ: ' + e.message); }

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
