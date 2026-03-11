// ====== 総務確認用 休暇届台帳（外部スプレッドシート） ======

/**
 * 台帳スプレッドシートを取得（なければPDFフォルダ内に自動作成）
 * @return {Spreadsheet}
 */
function getLedgerSs_() {
  var settings = getSettings_();
  var ledgerSsId = normalize_(settings['LEDGER_SSID']);

  // 既存IDがあればそれを使う
  if (ledgerSsId) {
    try {
      return SpreadsheetApp.openById(ledgerSsId);
    } catch (e) {
      console.warn('LEDGER_SSID のSSが開けません。再作成します: ' + e.message);
    }
  }

  // PDFフォルダ内に新規作成
  var rootFolderId = normalize_(settings['PDF_FOLDER_ID']);
  if (!rootFolderId) {
    console.warn('PDF_FOLDER_ID未設定のため台帳作成スキップ');
    return null;
  }

  var ss = SpreadsheetApp.create('休暇届台帳（総務確認用）');
  var file = DriveApp.getFileById(ss.getId());
  var folder = DriveApp.getFolderById(rootFolderId);
  folder.addFile(file);
  // ルートから除去（フォルダ内にのみ残す）
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) { /* ignore */ }

  // M_SYSTEM_SETTINGに保存
  var settingSh = getDb_().getSheetByName(SHEET.SETTINGS);
  if (settingSh) {
    settingSh.appendRow(['LEDGER_SSID', ss.getId(), '総務確認用台帳スプレッドシートID（自動生成）']);
  }

  // デフォルトのSheet1を削除用に名前変更（年度シート作成後に削除）
  ss.getSheets()[0].setName('_tmp');

  console.log('台帳SS新規作成: ' + ss.getId());
  return ss;
}

/**
 * 年度シートを取得（なければ作成＋ヘッダ書式設定）
 * @param {Spreadsheet} ss
 * @param {number} fy - 年度（例: 2025）
 * @return {Sheet}
 */
function getLedgerFySheet_(ss, fy) {
  var sheetName = fy + '年度';
  var sh = ss.getSheetByName(sheetName);
  if (sh) return sh;

  // 新規作成
  sh = ss.insertSheet(sheetName, 0); // 先頭に挿入

  // ヘッダ
  var headers = [
    '管理番号', '作業員ID', '部署', '氏名',
    '休暇日', '曜日', '休暇区分', '半日区分', '勤務時間帯',
    '休暇種類', '理由・詳細', '振替元出勤日',
    '申請日', '1次承認日', '1次承認者', '2次承認日', '2次承認者',
    '承認状態', '有給累計取得回数', 'PDF'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダ書式
  var headerRange = sh.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a365d');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);

  // 列幅設定
  var widths = [
    140, // 管理番号
    90,  // 作業員ID
    100, // 部署
    100, // 氏名
    100, // 休暇日
    50,  // 曜日
    80,  // 休暇区分
    70,  // 半日区分
    130, // 勤務時間帯
    130, // 休暇種類
    200, // 理由・詳細
    100, // 振替元出勤日
    100, // 申請日
    100, // 1次承認日
    120, // 1次承認者
    100, // 2次承認日
    120, // 2次承認者
    80,  // 承認状態
    80,  // 有給累計
    80   // PDF
  ];
  for (var c = 0; c < widths.length; c++) {
    sh.setColumnWidth(c + 1, widths[c]);
  }

  // _tmpシートがあれば削除
  var tmp = ss.getSheetByName('_tmp');
  if (tmp) {
    try { ss.deleteSheet(tmp); } catch (e) { /* ignore */ }
  }

  console.log('台帳年度シート作成: ' + sheetName);
  return sh;
}

// ====== 台帳更新メイン ======

/**
 * 承認時に台帳を更新（upsert: REQ_IDが既存なら更新、なければ追記）
 * @param {string} reqId
 */
function updateLedger_(reqId) {
  try {
    var reqData = api_getLeaveRequestById(reqId);
    if (!reqData) {
      console.warn('台帳更新: 申請データなし reqId=' + reqId);
      return;
    }

    var ss = getLedgerSs_();
    if (!ss) return;

    // 年度判定
    var leaveDate = new Date(reqData.leaveDate);
    var fy = computeFiscalYear_(leaveDate);
    var sh = getLedgerFySheet_(ss, fy);

    var DOWS = ['日', '月', '火', '水', '木', '金', '土'];

    // 休暇日・曜日
    var leaveDateStr = fmtDate_(leaveDate, 'yyyy/MM/dd');
    var dowStr = DOWS[leaveDate.getDay()];

    // 勤務時間帯
    var timeStr = '8:05 〜 17:10（終日）';
    if (reqData.leaveKubun === '半日休') {
      if (reqData.halfDayType === '午前') {
        timeStr = '8:05 〜 12:15（午前半休）';
      } else {
        timeStr = '13:00 〜 17:10（午後半休）';
      }
    }

    // 理由
    var reason = '';
    if (reqData.leaveType === '特別休暇' && reqData.specialReason) reason = reqData.specialReason;
    else if (reqData.leaveType === '有給消化（私用）') reason = reqData.paidDetail || '私用の為';
    else if (reqData.leaveType === '振替休暇') reason = '振替休暇';
    else reason = reqData.leaveType || '';

    // 振替元出勤日
    var subDateStr = '―';
    if (reqData.leaveType === '振替休暇' && reqData.substituteDate) {
      var subD = new Date(reqData.substituteDate);
      subDateStr = fmtDate_(subD, 'yyyy/MM/dd');
    }

    // 申請日
    var submittedStr = '';
    if (reqData.submittedAt) {
      var sd = new Date(reqData.submittedAt);
      submittedStr = fmtDate_(sd, 'yyyy/MM/dd');
    }

    // 1次承認日・承認者
    var approved1Str = '';
    if (reqData.approvedAt) {
      var a1 = new Date(reqData.approvedAt);
      approved1Str = fmtDate_(a1, 'yyyy/MM/dd');
    }
    var approver1Name = reqData.approvedBy1Email || '';

    // 2次承認日・承認者
    var approved2Str = '';
    if (reqData.approvedAt2) {
      var a2 = new Date(reqData.approvedAt2);
      approved2Str = fmtDate_(a2, 'yyyy/MM/dd');
    }
    var approver2Name = reqData.approvedBy2 || '';

    // 有給累計取得回数（当該年度・当該作業員の承認済有給回数）
    var paidCumulative = countPaidLeaves_(fy, reqData.workerId);

    // PDF URL
    var pdfUrl = reqData.pdfUrl || '';

    // 行データ構築
    var rowData = [
      reqData.reqId,
      reqData.workerId,
      reqData.deptName,
      reqData.workerName,
      leaveDateStr,
      dowStr,
      reqData.leaveKubun,
      reqData.halfDayType || '―',
      timeStr,
      reqData.leaveType,
      reason,
      subDateStr,
      submittedStr,
      approved1Str,
      approver1Name,
      approved2Str,
      approver2Name,
      reqData.status,
      paidCumulative,
      pdfUrl
    ];

    // upsert: 既存行を検索
    var lastRow = sh.getLastRow();
    var targetRow = -1;
    if (lastRow >= 2) {
      var reqIds = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < reqIds.length; i++) {
        if (normalize_(reqIds[i][0]) === reqId) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow > 0) {
      // 更新
      sh.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // 追記
      targetRow = lastRow + 1;
      sh.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    }

    // 行の書式設定（偶数行は薄い背景色）
    var dataRange = sh.getRange(targetRow, 1, 1, rowData.length);
    dataRange.setFontSize(10);
    dataRange.setVerticalAlignment('middle');
    if ((targetRow % 2) === 0) {
      dataRange.setBackground('#f8fafc');
    } else {
      dataRange.setBackground('#ffffff');
    }

    // PDF列にハイパーリンク設定
    if (pdfUrl) {
      sh.getRange(targetRow, 20).setFormula('=HYPERLINK("' + pdfUrl + '","PDF")');
    }

    SpreadsheetApp.flush();
    console.log('台帳更新完了: reqId=' + reqId + ', row=' + targetRow + ', fy=' + fy);

  } catch (e) {
    console.error('台帳更新エラー（続行）: ' + e.message + '\n' + e.stack);
  }
}

/**
 * 当該年度・当該作業員の承認済有給取得回数をカウント
 * @param {number} fy - 年度
 * @param {string} workerId - 作業員ID
 * @return {number}
 */
function countPaidLeaves_(fy, workerId) {
  var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  var readCols = Math.max(info.header.length, sh.getLastColumn());
  var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();
  var count = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowReqId = normalize_(row[idx['REQ_ID']]);
    if (!rowReqId) continue;

    var status = normalize_(row[idx['承認状態']]);
    if (status !== STATUS.APPROVED) continue;

    var rowWorkerId = normalize_(row[idx['作業員ID']]);
    if (rowWorkerId !== workerId) continue;

    var leaveType = normalize_(row[idx['休暇種類']]);
    if (leaveType !== '有給消化（私用）') continue;

    var leaveDate = row[idx['休暇日']];
    if (leaveDate instanceof Date) {
      var rowFy = computeFiscalYear_(leaveDate);
      if (rowFy === fy) count++;
    }
  }

  return count;
}
