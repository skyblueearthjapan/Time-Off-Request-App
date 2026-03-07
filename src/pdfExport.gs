// ====== PDF生成 ======

/**
 * 休暇届PDFのセル位置マッピング（PDF_MAPシート準拠）
 * docs/休暇届アプリ_最強DBマスターガイド.md §4.1 参照
 */
var LEAVE_PDF_MAP = {
  LEAVE_START_REIWA_YEAR: 'B6',  // 開始：令和年（数値）
  LEAVE_START_MONTH:      'E6',  // 開始：月
  LEAVE_START_DAY:        'H6',  // 開始：日
  LEAVE_START_WDAY:       'K6',  // 開始：曜日（例：月）
  LEAVE_END_REIWA_YEAR:   'B7',  // 終了：令和年
  LEAVE_END_MONTH:        'E7',  // 終了：月
  LEAVE_END_DAY:          'H7',  // 終了：日
  LEAVE_END_WDAY:         'K7',  // 終了：曜日
  LEAVE_END_HOUR:         'M7',  // 終了：時
  LEAVE_END_MIN:          'B8',  // 終了：分
  LEAVE_REASON_SHORT:     'A14', // 理由（短文）
  SUBMIT_REIWA_YEAR:      'J30', // 届出日：令和年
  SUBMIT_MONTH:           'M30', // 届出日：月
  SUBMIT_DAY:             'H31', // 届出日：日
  WORKER_NAME:            'J34', // 氏名
};

/**
 * 令和変換
 */
function toReiwa_(dateObj) {
  if (!dateObj) return '';
  var d = dateObj instanceof Date ? dateObj : new Date(dateObj);
  var year = d.getFullYear();
  var reiwa = year - 2018;
  return '令和' + reiwa + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日';
}

/**
 * Drive日付フォルダを取得または作成（yyyy.MM.dd）
 */
function getOrCreateDateFolder_(rootFolderId, dateObj) {
  var root = DriveApp.getFolderById(rootFolderId);
  var folderName = fmtDate_(dateObj, 'yyyy.MM.dd');
  var it = root.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return root.createFolder(folderName);
}

/**
 * サイン画像をDriveに保存
 */
function saveSignImage_(reqId, base64Data, leaveDate) {
  var settings = getSettings_();
  var rootFolderId = normalize_(settings['PDF_FOLDER_ID']);
  if (!rootFolderId) {
    console.warn('PDF_FOLDER_ID未設定のためサイン画像保存スキップ');
    return '';
  }

  // Base64デコード（data:image/png;base64, 部分を除去）
  var base64 = base64Data.replace(/^data:image\/\w+;base64,/, '');
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', 'sign_' + reqId + '.png');

  // PDFと同じ日付フォルダに保存
  var dateObj = leaveDate instanceof Date ? leaveDate : (leaveDate ? new Date(leaveDate) : new Date());
  var folder = getOrCreateDateFolder_(rootFolderId, dateObj);
  var file = folder.createFile(blob);
  return file.getUrl();
}

/**
 * 休暇届PDF生成（1件）
 */
function generateLeavePdf_(reqId) {
  var reqData = api_getLeaveRequestById(reqId);
  if (!reqData) throw new Error('申請が見つかりません: ' + reqId);

  var settings = getSettings_();
  var rootFolderId = normalize_(settings['PDF_FOLDER_ID']);
  var templateSsid = normalize_(settings['PDF_TEMPLATE_SSID']);

  if (!rootFolderId) throw new Error('M_SYSTEM_SETTINGに PDF_FOLDER_ID が未設定です。');
  if (!templateSsid) throw new Error('M_SYSTEM_SETTINGに PDF_TEMPLATE_SSID が未設定です。');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // テンプレSSをコピー
    var templateFile = DriveApp.getFileById(templateSsid);
    var tmpName = 'TMP_' + reqId + '_' + fmtDate_(new Date(), 'yyyyMMdd_HHmmss');
    var tmpFile = templateFile.makeCopy(tmpName);
    var tmpSs = SpreadsheetApp.openById(tmpFile.getId());

    // PDF_TEMPLATEシートを名前で取得
    var formSheet = tmpSs.getSheetByName('PDF_TEMPLATE');
    if (!formSheet) {
      // フォールバック: シート名が異なる場合は最初のシートを試行
      formSheet = tmpSs.getSheets()[0];
      console.warn('PDF_TEMPLATEシートが見つからないため最初のシートを使用: ' + formSheet.getName());
    }
    fillLeaveTemplate_(formSheet, reqData);
    SpreadsheetApp.flush();

    // 保存先フォルダ
    var leaveDate = reqData.leaveDate instanceof Date ? reqData.leaveDate : new Date(reqData.leaveDate);
    var dateFolder = getOrCreateDateFolder_(rootFolderId, leaveDate);

    // ファイル名
    var ymd = fmtDate_(leaveDate, 'yyyyMMdd');
    var safeDept = (reqData.deptName || '').replace(/[\\\/\:\*\?\"\<\>\|]/g, '_');
    var safeName = (reqData.workerName || '').replace(/[\\\/\:\*\?\"\<\>\|]/g, '_');
    var pdfName = ymd + '_' + safeDept + '_' + safeName + '_休暇届.pdf';

    // PDFエクスポート
    var pdfBlob = exportSheetToPdfBlob_(tmpSs.getId(), formSheet.getSheetId(), pdfName);

    // Drive保存
    var pdfFile = dateFolder.createFile(pdfBlob).setName(pdfName);

    // 一時コピー削除
    tmpFile.setTrashed(true);

    return {
      ok: true,
      pdfFileId: pdfFile.getId(),
      pdfUrl: pdfFile.getUrl(),
      pdfName: pdfName,
      folderId: dateFolder.getId(),
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * テンプレートシートにデータ書込み（PDF_MAP準拠）
 */
function fillLeaveTemplate_(sheet, data) {
  var DOWS = ['日','月','火','水','木','金','土'];

  // --- 休暇日（開始・終了） ---
  if (data.leaveDate) {
    var ld = data.leaveDate instanceof Date ? data.leaveDate : new Date(data.leaveDate);
    var reiwaYear = ld.getFullYear() - 2018;
    var month = ld.getMonth() + 1;
    var day = ld.getDate();
    var wday = DOWS[ld.getDay()];

    // 開始日（Row 6）
    sheet.getRange(LEAVE_PDF_MAP.LEAVE_START_REIWA_YEAR).setValue(reiwaYear);
    sheet.getRange(LEAVE_PDF_MAP.LEAVE_START_MONTH).setValue(month);
    sheet.getRange(LEAVE_PDF_MAP.LEAVE_START_DAY).setValue(day);
    sheet.getRange(LEAVE_PDF_MAP.LEAVE_START_WDAY).setValue(wday);

    // 終了日（Row 7）- 全日休は同日、半日休も同日
    sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_REIWA_YEAR).setValue(reiwaYear);
    sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_MONTH).setValue(month);
    sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_DAY).setValue(day);
    sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_WDAY).setValue(wday);

    // 終了時刻（半日区分に応じて設定）
    // 午前: 8:05～12:15、午後: 13:00～17:10、全日: 8:05～17:10
    if (data.leaveKubun === '半日休') {
      if (data.halfDayType === '午前') {
        sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_HOUR).setValue(12);
        sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_MIN).setValue(15);
      } else {
        // 午後半休
        sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_HOUR).setValue(17);
        sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_MIN).setValue(10);
      }
    } else {
      // 全日休
      sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_HOUR).setValue(17);
      sheet.getRange(LEAVE_PDF_MAP.LEAVE_END_MIN).setValue(10);
    }
  }

  // --- 理由（短文） ---
  var reason = '';
  if (data.leaveType === '特別休暇' && data.specialReason) {
    reason = data.specialReason;
  } else if (data.leaveType === '有給消化（私用）' && data.paidDetail) {
    reason = data.paidDetail;
  } else if (data.leaveType === '振替休暇' && data.substituteDate) {
    var subD = data.substituteDate instanceof Date ? data.substituteDate : new Date(data.substituteDate);
    reason = '振替（出勤日: ' + fmtDate_(subD, 'yyyy/MM/dd') + '）';
  } else if (data.leaveType) {
    reason = data.leaveType;
  }
  if (data.additionalDetail) {
    reason += (reason ? ' ' : '') + data.additionalDetail;
  }
  if (reason) {
    sheet.getRange(LEAVE_PDF_MAP.LEAVE_REASON_SHORT).setValue(reason);
  }

  // --- 届出日（令和） ---
  if (data.submittedAt) {
    var sd = data.submittedAt instanceof Date ? data.submittedAt : new Date(data.submittedAt);
    sheet.getRange(LEAVE_PDF_MAP.SUBMIT_REIWA_YEAR).setValue(sd.getFullYear() - 2018);
    sheet.getRange(LEAVE_PDF_MAP.SUBMIT_MONTH).setValue(sd.getMonth() + 1);
    sheet.getRange(LEAVE_PDF_MAP.SUBMIT_DAY).setValue(sd.getDate());
  }

  // --- 氏名 ---
  sheet.getRange(LEAVE_PDF_MAP.WORKER_NAME).setValue(data.workerName || '');
}

/**
 * シート1枚をPDF化するユーティリティ
 */
function exportSheetToPdfBlob_(spreadsheetId, sheetId, filename) {
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export'
    + '?format=pdf'
    + '&gid=' + sheetId
    + '&portrait=true'
    + '&size=A4'
    + '&fitw=true'
    + '&sheetnames=false&printtitle=false'
    + '&pagenumbers=false'
    + '&gridlines=false'
    + '&fzr=false';

  var token = ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });

  var code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('PDFエクスポートに失敗しました: HTTP ' + code);
  }

  return res.getBlob().setName(filename);
}
