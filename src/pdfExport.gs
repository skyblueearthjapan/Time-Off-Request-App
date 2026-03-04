// ====== PDF生成 ======

/**
 * 休暇届PDFのセル位置マッピング（テンプレートSS内のシート）
 */
var LEAVE_PDF_MAP = {
  title: 'C1',          // タイトル「休 暇 届」
  submittedDate: 'F2',  // 申請日
  deptName: 'C4',       // 部署名
  workerName: 'E4',     // 氏名
  leaveKubun: 'C6',     // 休暇区分（全日休/半日休）
  halfDayType: 'E6',    // 半日区分（午前/午後）
  leaveDate: 'C8',      // 休暇日
  leaveDateReiwa: 'E8', // 休暇日（令和）
  leaveType: 'C10',     // 休暇種類
  substituteDate: 'C12',// 振替元出勤日
  reason: 'C14',        // 理由・詳細
  signArea: 'F18',      // 承認印エリア
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
function saveSignImage_(reqId, base64Data) {
  var settings = getSettings_();
  var rootFolderId = normalize_(settings['PDF_FOLDER_ID']);
  if (!rootFolderId) {
    console.warn('PDF_FOLDER_ID未設定のためサイン画像保存スキップ');
    return '';
  }

  // Base64デコード（data:image/png;base64, 部分を除去）
  var base64 = base64Data.replace(/^data:image\/\w+;base64,/, '');
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', 'sign_' + reqId + '.png');

  var folder = DriveApp.getFolderById(rootFolderId);
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

    // 申請書シートに直接書込み
    var formSheet = tmpSs.getSheets()[0]; // 最初のシート
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
 * テンプレートシートにデータ書込み
 */
function fillLeaveTemplate_(sheet, data) {
  // 申請日
  if (data.submittedAt) {
    var sd = data.submittedAt instanceof Date ? data.submittedAt : new Date(data.submittedAt);
    sheet.getRange(LEAVE_PDF_MAP.submittedDate).setValue(fmtDate_(sd, 'yyyy/MM/dd'));
  }
  // 部署名
  sheet.getRange(LEAVE_PDF_MAP.deptName).setValue(data.deptName || '');
  // 氏名
  sheet.getRange(LEAVE_PDF_MAP.workerName).setValue(data.workerName || '');
  // 休暇区分
  sheet.getRange(LEAVE_PDF_MAP.leaveKubun).setValue(data.leaveKubun || '');
  // 半日区分
  if (data.halfDayType) {
    sheet.getRange(LEAVE_PDF_MAP.halfDayType).setValue(data.halfDayType);
  }
  // 休暇日
  if (data.leaveDate) {
    var ld = data.leaveDate instanceof Date ? data.leaveDate : new Date(data.leaveDate);
    sheet.getRange(LEAVE_PDF_MAP.leaveDate).setValue(fmtDate_(ld, 'yyyy/MM/dd'));
    sheet.getRange(LEAVE_PDF_MAP.leaveDateReiwa).setValue(toReiwa_(ld));
  }
  // 休暇種類
  sheet.getRange(LEAVE_PDF_MAP.leaveType).setValue(data.leaveType || '');
  // 振替元出勤日
  if (data.substituteDate && data.leaveType === '振替休暇') {
    var subD = data.substituteDate instanceof Date ? data.substituteDate : new Date(data.substituteDate);
    sheet.getRange(LEAVE_PDF_MAP.substituteDate).setValue(fmtDate_(subD, 'yyyy/MM/dd'));
  }
  // 理由・詳細
  var reasons = [];
  if (data.specialReason) reasons.push(data.specialReason);
  if (data.paidDetail) reasons.push(data.paidDetail);
  if (data.additionalDetail) reasons.push(data.additionalDetail);
  if (reasons.length) {
    sheet.getRange(LEAVE_PDF_MAP.reason).setValue(reasons.join('\n'));
  }
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
