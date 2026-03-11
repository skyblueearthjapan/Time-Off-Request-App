// ====== PDF生成（テンプレートレス・動的構築方式） ======

/**
 * テスト用：GASエディタから直接実行してPDF生成を確認
 * 実行前に REQ_ID を実在のものに書き換えてください
 */
function test_generatePdf() {
  var reqId = 'LV-20260307-001';  // ← 実在のREQ_IDに変更
  var approverEmail = 'imaizumi@lineworks-local.info'; // ← 承認者メール
  try {
    var result = generateLeavePdf_(reqId, approverEmail);
    console.log('PDF生成成功: ' + JSON.stringify(result));
    Logger.log('PDF URL: ' + result.pdfUrl);
  } catch (e) {
    console.error('PDF生成エラー: ' + e.message);
    console.error(e.stack);
  }
}

/**
 * 電子印取得テスト（GASエディタから実行）
 * M_STAMPの同期状態とスタンプファイルの存在確認
 */
function test_stampLookup() {
  var testEmail = 'imaizumi@lineworks-local.info'; // ← テスト対象メール
  console.log('=== 電子印テスト開始 ===');

  // 1. M_STAMPシート確認
  var ss = getDb_();
  var sh = ss.getSheetByName(SHEET.STAMP);
  if (!sh) {
    console.error('M_STAMPシートが存在しません！ setupAllSheets を実行してください。');
    return;
  }
  var lastRow = sh.getLastRow();
  console.log('M_STAMP: ' + lastRow + ' 行（ヘッダ含む）');
  if (lastRow >= 2) {
    var data = sh.getDataRange().getValues();
    console.log('ヘッダ: ' + data[0].join(' | '));
    for (var r = 1; r < data.length; r++) {
      console.log('  行' + (r+1) + ': ' + data[r][0] + ' → ' + data[r][1] + ' (' + data[r][2] + ')');
    }
  }

  // 2. スタンプFileId取得
  var stampFileId = getStampFileId_(testEmail);
  console.log('取得結果: stampFileId=' + (stampFileId || '(空)'));

  // 3. Driveファイル確認
  if (stampFileId) {
    try {
      var file = DriveApp.getFileById(stampFileId);
      console.log('Driveファイル: 名前=' + file.getName() + ', MIME=' + file.getMimeType() + ', サイズ=' + file.getSize());
      var blob = file.getBlob();
      console.log('Blob取得: type=' + blob.getContentType() + ', size=' + blob.getBytes().length);
      console.log('=== 電子印テスト: OK ===');
    } catch (e) {
      console.error('Driveファイル取得エラー: ' + e.message);
    }
  } else {
    console.warn('=== 電子印テスト: スタンプ未登録 ===');
  }
}

/**
 * 令和変換
 */
function toReiwa_(dateObj) {
  if (!dateObj) return '';
  var d = dateObj instanceof Date ? dateObj : new Date(dateObj);
  var reiwa = d.getFullYear() - 2018;
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
 * サイン画像をDriveに保存（PDFと同じ日付フォルダ）
 */
function saveSignImage_(reqId, base64Data, leaveDate) {
  var settings = getSettings_();
  var rootFolderId = normalize_(settings['PDF_FOLDER_ID']);
  if (!rootFolderId) {
    console.warn('PDF_FOLDER_ID未設定のためサイン画像保存スキップ');
    return '';
  }

  var base64 = base64Data.replace(/^data:image\/\w+;base64,/, '');
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', 'sign_' + reqId + '.png');

  var dateObj = leaveDate instanceof Date ? leaveDate : (leaveDate ? new Date(leaveDate) : new Date());
  var folder = getOrCreateDateFolder_(rootFolderId, dateObj);
  var file = folder.createFile(blob);
  return file.getUrl();
}

// ====== 色定義 ======
var PDF_COLORS_ = {
  TITLE_BG:   '#1a365d',
  TITLE_FG:   '#ffffff',
  LABEL_BG:   '#e8edf3',
  LABEL_FG:   '#1a365d',
  BORDER:     '#94a3b8',
  SECTION_BG: '#f1f5f9',
  TEXT:        '#1a202c',
  MUTED:      '#718096',
  ACCENT:     '#cc0000',
};

// ====== PDF生成メイン ======

/**
 * 休暇届PDF生成（GASでシートを動的構築）
 */
function generateLeavePdf_(reqId, approverEmail) {
  var reqData = api_getLeaveRequestById(reqId);
  if (!reqData) throw new Error('申請が見つかりません: ' + reqId);

  var settings = getSettings_();
  var rootFolderId = normalize_(settings['PDF_FOLDER_ID']);
  if (!rootFolderId) throw new Error('PDF_FOLDER_IDが未設定です。');

  // ※ロックは呼び出し元（api_approveLeaveRequest）が保持済み。
  //   GASスクリプトロックは再入不可のため、ここでは取得しない。

  var tmpSs = null;
  try {
    // 一時スプレッドシートを新規作成（テンプレートコピー不要）
    tmpSs = SpreadsheetApp.create('TMP_PDF_' + reqId + '_' + fmtDate_(new Date(), 'yyyyMMdd_HHmmss'));
    var sheet = tmpSs.getSheets()[0];
    sheet.setName('休暇届');

    // シート構築＋データ書込み（IMAGE数式での印影挿入含む）
    buildLeavePdfSheet_(sheet, reqData, approverEmail);
    SpreadsheetApp.flush();
    // IMAGE数式がレンダリングされるまで待機
    Utilities.sleep(3000);
    SpreadsheetApp.flush();

    // 保存先フォルダ
    var leaveDate = new Date(reqData.leaveDate);
    var dateFolder = getOrCreateDateFolder_(rootFolderId, leaveDate);

    // ファイル名
    var ymd = fmtDate_(leaveDate, 'yyyyMMdd');
    var safeDept = (reqData.deptName || '').replace(/[\\\/\:\*\?\"\<\>\|]/g, '_');
    var safeName = (reqData.workerName || '').replace(/[\\\/\:\*\?\"\<\>\|]/g, '_');
    var pdfName = ymd + '_' + safeDept + '_' + safeName + '_休暇届.pdf';

    // PDFエクスポート
    var pdfBlob = exportSheetToPdfBlob_(tmpSs.getId(), sheet.getSheetId(), pdfName);

    // Drive保存
    var pdfFile = dateFolder.createFile(pdfBlob).setName(pdfName);

    return {
      ok: true,
      pdfFileId: pdfFile.getId(),
      pdfUrl: pdfFile.getUrl(),
      pdfName: pdfName,
      folderId: dateFolder.getId(),
    };
  } finally {
    // 一時SSは成功・失敗問わず必ず削除
    if (tmpSs) {
      try { DriveApp.getFileById(tmpSs.getId()).setTrashed(true); }
      catch (e) { console.error('一時SS削除エラー: ' + e.message); }
    }
  }
}

// ====== シート構築 ======

/**
 * 休暇届シートをゼロから構築しデータを書き込む
 */
function buildLeavePdfSheet_(sheet, data, approverEmail) {
  var C = PDF_COLORS_;
  var DOWS = ['日','月','火','水','木','金','土'];

  // --- 列幅設定（A4縦に最適化） ---
  sheet.setColumnWidth(1, 12);   // A: 左余白
  sheet.setColumnWidth(2, 80);   // B: ラベル左
  sheet.setColumnWidth(3, 80);   // C: ラベル右
  sheet.setColumnWidth(4, 68);   // D: 値
  sheet.setColumnWidth(5, 68);   // E: 値
  sheet.setColumnWidth(6, 68);   // F: 値
  sheet.setColumnWidth(7, 68);   // G: 値/ラベル
  sheet.setColumnWidth(8, 68);   // H: 値
  sheet.setColumnWidth(9, 68);   // I: 値
  sheet.setColumnWidth(10, 12);  // J: 右余白

  // --- 全体フォント ---
  sheet.getRange('A1:J35').setFontFamily('Noto Sans JP, sans-serif')
    .setFontColor(C.TEXT);

  // ============================================================
  //  ROW 1: 上余白
  // ============================================================
  sheet.setRowHeight(1, 12);

  // ============================================================
  //  ROW 2: タイトル「休 暇 届」
  // ============================================================
  sheet.setRowHeight(2, 52);
  var titleR = sheet.getRange('B2:I2');
  titleR.merge().setValue('休 暇 届')
    .setFontSize(24).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C.TITLE_BG).setFontColor(C.TITLE_FG);

  // ============================================================
  //  ROW 3: スペーサー
  // ============================================================
  sheet.setRowHeight(3, 6);

  // ============================================================
  //  ROW 4: 届出日（右寄せ）
  // ============================================================
  sheet.setRowHeight(4, 22);
  var sd = data.submittedAt ? new Date(data.submittedAt) : new Date();
  var submitStr = '令和' + (sd.getFullYear() - 2018) + '年'
    + (sd.getMonth() + 1) + '月' + sd.getDate() + '日 届出';
  sheet.getRange('F4:I4').merge().setValue(submitStr)
    .setFontSize(10).setHorizontalAlignment('right');

  // ============================================================
  //  ROW 5: スペーサー
  // ============================================================
  sheet.setRowHeight(5, 6);

  // ============================================================
  //  ROW 6: 所属部署
  // ============================================================
  sheet.setRowHeight(6, 32);
  pdfLabel_(sheet, 'B6:C6', '所 属 部 署', C);
  pdfValue_(sheet, 'D6:I6', data.deptName || '');
  pdfBorder_(sheet, 'B6:I6', C);

  // ============================================================
  //  ROW 7: 氏名
  // ============================================================
  sheet.setRowHeight(7, 32);
  pdfLabel_(sheet, 'B7:C7', '氏　　　名', C);
  pdfValue_(sheet, 'D7:I7', data.workerName || '');
  pdfBorder_(sheet, 'B7:I7', C);

  // ============================================================
  //  ROW 8: スペーサー + セクションライン
  // ============================================================
  sheet.setRowHeight(8, 14);
  sheet.getRange('B8:I8').merge().setValue('届 出 内 容')
    .setFontSize(9).setFontWeight('bold')
    .setFontColor(C.LABEL_FG).setHorizontalAlignment('center')
    .setVerticalAlignment('bottom');

  // ============================================================
  //  ROW 9: 休暇日
  // ============================================================
  sheet.setRowHeight(9, 34);
  pdfLabel_(sheet, 'B9:C9', '休 暇 日', C);
  var ld = new Date(data.leaveDate);
  var leaveDateStr = '令和' + (ld.getFullYear() - 2018) + '年'
    + (ld.getMonth() + 1) + '月' + ld.getDate() + '日'
    + '（' + DOWS[ld.getDay()] + '曜日）';
  pdfValue_(sheet, 'D9:I9', leaveDateStr);
  pdfBorder_(sheet, 'B9:I9', C);

  // ============================================================
  //  ROW 10: 休暇時間帯
  // ============================================================
  sheet.setRowHeight(10, 34);
  pdfLabel_(sheet, 'B10:C10', '休 暇 時 間', C);
  var timeStr = '';
  if (data.leaveKubun === '半日休') {
    if (data.halfDayType === '午前') {
      timeStr = '午前半休　8:05 ～ 12:15';
    } else {
      timeStr = '午後半休　13:00 ～ 17:10';
    }
  } else {
    timeStr = '終日　8:05 ～ 17:10';
  }
  pdfValue_(sheet, 'D10:I10', timeStr);
  pdfBorder_(sheet, 'B10:I10', C);

  // ============================================================
  //  ROW 11: 休暇区分
  // ============================================================
  sheet.setRowHeight(11, 32);
  pdfLabel_(sheet, 'B11:C11', '休 暇 区 分', C);
  var kubunStr = data.leaveKubun || '';
  if (data.halfDayType) kubunStr += '（' + data.halfDayType + '）';
  pdfValue_(sheet, 'D11:I11', kubunStr);
  pdfBorder_(sheet, 'B11:I11', C);

  // ============================================================
  //  ROW 12: 休暇種類
  // ============================================================
  sheet.setRowHeight(12, 32);
  pdfLabel_(sheet, 'B12:C12', '休 暇 種 類', C);
  pdfValue_(sheet, 'D12:I12', data.leaveType || '');
  pdfBorder_(sheet, 'B12:I12', C);

  // ============================================================
  //  ROW 13: 振替元出勤日
  // ============================================================
  sheet.setRowHeight(13, 32);
  pdfLabel_(sheet, 'B13:C13', '振替元出勤日', C);
  var subStr = '―';
  if (data.leaveType === '振替休暇' && data.substituteDate) {
    var subD = new Date(data.substituteDate);
    subStr = '令和' + (subD.getFullYear() - 2018) + '年'
      + (subD.getMonth() + 1) + '月' + subD.getDate() + '日'
      + '（' + DOWS[subD.getDay()] + '）';
  }
  pdfValue_(sheet, 'D13:I13', subStr);
  pdfBorder_(sheet, 'B13:I13', C);

  // ============================================================
  //  ROW 14-15: 理由・詳細（2行分）
  // ============================================================
  sheet.setRowHeight(14, 25);
  sheet.setRowHeight(15, 25);
  pdfLabel_(sheet, 'B14:C15', '理由・詳細', C);
  var reason = pdfBuildReason_(data);
  sheet.getRange('D14:I15').merge().setValue(reason)
    .setFontSize(11).setVerticalAlignment('top').setWrap(true);
  pdfBorder_(sheet, 'B14:I15', C);

  // ============================================================
  //  ROW 16: スペーサー
  // ============================================================
  sheet.setRowHeight(16, 10);

  // ============================================================
  //  ROW 17: 宣言文
  // ============================================================
  sheet.setRowHeight(17, 30);
  sheet.getRange('B17:I17').merge()
    .setValue('上記のとおり休暇を届け出いたします。')
    .setFontSize(11).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // ============================================================
  //  ROW 18: スペーサー
  // ============================================================
  sheet.setRowHeight(18, 8);

  // ============================================================
  //  ROW 19-20: 届出者（印なし）
  // ============================================================
  sheet.setRowHeight(19, 28);
  sheet.setRowHeight(20, 28);
  pdfLabel_(sheet, 'B19:C20', '届 出 者', C);
  sheet.getRange('D19:D20').merge().setValue('氏名')
    .setFontSize(9).setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setFontColor(C.MUTED);
  sheet.getRange('E19:I20').merge().setValue(data.workerName || '')
    .setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  pdfBorder_(sheet, 'B19:I20', C);

  // ============================================================
  //  ROW 21: スペーサー
  // ============================================================
  sheet.setRowHeight(21, 12);

  // ============================================================
  //  ROW 22: 承認欄ヘッダー
  // ============================================================
  sheet.setRowHeight(22, 26);
  sheet.getRange('B22:I22').merge()
    .setValue('承 認 欄')
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C.SECTION_BG).setFontColor(C.LABEL_FG);
  pdfBorder_(sheet, 'B22:I22', C);

  // ============================================================
  //  ROW 23: 承認者ラベル
  // ============================================================
  sheet.setRowHeight(23, 24);
  sheet.getRange('B23:C23').merge().setValue('')
    .setBackground(C.LABEL_BG);
  sheet.getRange('D23:F23').merge().setValue('部　長')
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C.SECTION_BG);
  sheet.getRange('G23:I23').merge().setValue('所 属 長')
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C.SECTION_BG);
  pdfBorder_(sheet, 'B23:I23', C);

  // ============================================================
  //  ROW 24-27: 承認印エリア
  // ============================================================
  sheet.setRowHeight(24, 20);
  sheet.setRowHeight(25, 20);
  sheet.setRowHeight(26, 20);
  sheet.setRowHeight(27, 20);
  sheet.getRange('B24:C27').merge().setValue('承認印')
    .setFontSize(9).setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(C.LABEL_BG).setFontColor(C.LABEL_FG);
  sheet.getRange('D24:F27').merge().setValue('');
  sheet.getRange('G24:I27').merge().setValue('');
  pdfBorder_(sheet, 'B24:I27', C);

  // ============================================================
  //  ROW 28: 承認日
  // ============================================================
  sheet.setRowHeight(28, 24);
  sheet.getRange('B28:C28').merge().setValue('承認日')
    .setFontSize(9).setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(C.LABEL_BG).setFontColor(C.LABEL_FG);
  var approvedStr = '';
  if (data.approvedAt) {
    var ad = new Date(data.approvedAt);
    approvedStr = (ad.getFullYear() - 2018) + '/' + (ad.getMonth() + 1) + '/' + ad.getDate();
  }
  sheet.getRange('D28:F28').merge().setValue(approvedStr)
    .setFontSize(9).setHorizontalAlignment('center');
  sheet.getRange('G28:I28').merge().setValue('')
    .setFontSize(9).setHorizontalAlignment('center');
  pdfBorder_(sheet, 'B28:I28', C);

  // ============================================================
  //  ROW 29: フッター注記
  // ============================================================
  sheet.setRowHeight(29, 6);
  sheet.setRowHeight(30, 18);
  sheet.getRange('B30:I30').merge()
    .setValue('※後報の場合は、理由を詳しく記載すること。')
    .setFontSize(8).setFontColor(C.MUTED);

  // ============================================================
  //  承認印挿入（所属長欄: G24、IMAGE数式方式）
  //  ※ insertImage(blob) はPDFエクスポート時に反映されない
  //    GAS既知問題のため、IMAGE()セル数式で確実にPDFに含める
  // ============================================================
  var tmpStampFileIds = [];
  var approverStampInserted = false;
  try {
    var stampEmail = approverEmail || '';
    if (!stampEmail) {
      try { stampEmail = Session.getActiveUser().getEmail() || ''; } catch (e2) { /* ignore */ }
    }
    console.log('電子印検索: approverEmail=' + approverEmail + ', stampEmail=' + stampEmail);

    if (stampEmail) {
      var stampFileId = getStampFileId_(stampEmail);
      console.log('電子印FileId: ' + stampFileId);

      if (stampFileId) {
        // 元ファイルを共有設定してIMAGE数式で直接参照（blob操作なし）
        var readyId = prepareStampForImageFormula_(stampFileId);
        if (readyId) {
          var imageUrl = 'https://drive.google.com/uc?export=view&id=' + readyId;
          sheet.getRange('G24').setFormula('=IMAGE("' + imageUrl + '")');
          approverStampInserted = true;
          console.log('電子印IMAGE数式挿入成功: G24, fileId=' + readyId);
        }
      } else {
        console.warn('電子印FileIdが空です。M_STAMPにメール=' + stampEmail + 'の登録があるか確認してください。');
      }
    }
  } catch (e) {
    console.error('電子印挿入エラー: ' + e.message + '\n' + e.stack);
  }

  // 2. 手書きサイン（電子印がない場合のフォールバック）
  if (!approverStampInserted && data.signImageUrl) {
    try {
      console.log('手書きサイン挿入: URL=' + data.signImageUrl);
      var signTmpId = pdfInsertSignAsFormula_(sheet, data.signImageUrl, 'G24');
      if (signTmpId) tmpStampFileIds.push(signTmpId);
    } catch (e) {
      console.error('サイン画像挿入エラー: ' + e.message);
    }
  }

  if (!approverStampInserted && !data.signImageUrl) {
    console.warn('承認印なし: 電子印もサイン画像もありません');
  }

  // --- 印刷範囲外の列を非表示 ---
  if (sheet.getMaxColumns() > 10) {
    sheet.hideColumns(11, sheet.getMaxColumns() - 10);
  }
  if (sheet.getMaxRows() > 30) {
    sheet.hideRows(31, sheet.getMaxRows() - 30);
  }

}

// ====== ヘルパー関数 ======

/** ラベルセル（背景色+太字+中央揃え） */
function pdfLabel_(sheet, rangeStr, text, C) {
  sheet.getRange(rangeStr).merge().setValue(text)
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground(C.LABEL_BG).setFontColor(C.LABEL_FG);
}

/** 値セル（左揃え） */
function pdfValue_(sheet, rangeStr, text) {
  sheet.getRange(rangeStr).merge().setValue(text)
    .setFontSize(11).setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
}

/** 外枠+内部罫線 */
function pdfBorder_(sheet, rangeStr, C) {
  sheet.getRange(rangeStr).setBorder(
    true, true, true, true, true, true,
    C.BORDER, SpreadsheetApp.BorderStyle.SOLID
  );
}

/** 理由テキスト構築 */
function pdfBuildReason_(data) {
  if (data.leaveType === '特別休暇' && data.specialReason) {
    return data.specialReason;
  } else if (data.leaveType === '有給消化（私用）') {
    return data.paidDetail || '私用のため';
  } else if (data.leaveType === '振替休暇') {
    return '振替休暇';
  } else if (data.leaveType) {
    return data.leaveType;
  }
  return '―';
}

/**
 * 大きな電子印画像を縮小して返す（2MB/100万ピクセル制限対策）
 * 複数のサムネイルAPI URLを試行し、全て失敗したら一時Docを経由して縮小
 */
function getResizedStampBlob_(fileId, originalBlob) {
  var token = ScriptApp.getOAuthToken();
  var headers = { Authorization: 'Bearer ' + token };

  // 方法1: Drive thumbnail API
  var urls = [
    'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w200',
    'https://lh3.googleusercontent.com/d/' + fileId + '=s200',
  ];

  for (var i = 0; i < urls.length; i++) {
    try {
      var res = UrlFetchApp.fetch(urls[i], {
        headers: headers,
        muteHttpExceptions: true,
        followRedirects: true,
      });
      console.log('サムネイル試行' + (i+1) + ': HTTP ' + res.getResponseCode() + ', size=' + res.getBlob().getBytes().length + ', url=' + urls[i]);
      if (res.getResponseCode() === 200) {
        var blob = res.getBlob();
        if (blob.getBytes().length > 1000 && blob.getBytes().length < 1500000) {
          return blob.setName('stamp.png');
        }
      }
    } catch (e) {
      console.warn('サムネイル試行' + (i+1) + 'エラー: ' + e.message);
    }
  }

  // 方法2: 一時Googleスライドに挿入→サムネイル取得→削除
  try {
    console.log('一時スライド方式でリサイズ試行');
    var slide = SlidesApp.create('TMP_STAMP_RESIZE');
    var page = slide.getSlides()[0];
    var img = page.insertImage(originalBlob);
    // アスペクト比を維持して200x200以内に
    var w = img.getWidth();
    var h = img.getHeight();
    var scale = Math.min(200 / w, 200 / h, 1);
    img.setWidth(w * scale).setHeight(h * scale);
    slide.saveAndClose();

    // スライドのサムネイルを取得
    var slideFile = DriveApp.getFileById(slide.getId());
    var thumbBlob = slideFile.getThumbnail();
    slideFile.setTrashed(true);

    if (thumbBlob && thumbBlob.getBytes().length > 500) {
      console.log('スライドサムネイル取得成功: size=' + thumbBlob.getBytes().length);
      return thumbBlob.setName('stamp.png');
    }
  } catch (slideErr) {
    console.error('スライドリサイズエラー: ' + slideErr.message);
    // 一時ファイルクリーンアップ
    try {
      var files = DriveApp.getFilesByName('TMP_STAMP_RESIZE');
      while (files.hasNext()) files.next().setTrashed(true);
    } catch (e) { /* ignore */ }
  }

  // 全方法失敗 → 元の画像をそのまま返す（insertImageでエラーになる可能性あり）
  console.warn('全リサイズ方法失敗、元画像を使用');
  return originalBlob;
}

/**
 * 電子印のDriveファイルから一時公開ファイルを作成（IMAGE数式用）
 * @param {string} stampFileId - 元の電子印DriveファイルID
 * @return {string|null} 一時ファイルID（クリーンアップ用）、失敗時null
 */
/**
 * スタンプファイルをIMAGE数式で使えるように共有設定する
 * blob操作を一切行わず、元ファイルの共有設定のみ変更
 * @param {string} stampFileId - DriveファイルID
 * @return {string|null} 使用するファイルID（成功時）、失敗時null
 */
function prepareStampForImageFormula_(stampFileId) {
  try {
    var stampFile = DriveApp.getFileById(stampFileId);
    console.log('電子印ファイル取得成功: ' + stampFile.getName() + ' (' + stampFile.getMimeType() + ', size=' + stampFile.getSize() + ')');

    // 元ファイルを「リンクを知っている全員」に共有設定
    // blob操作なし → 2MB制限の影響を完全に回避
    stampFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    console.log('電子印共有設定完了: id=' + stampFileId);
    return stampFileId;
  } catch (e) {
    console.error('電子印共有設定エラー: ' + e.message + '\n' + e.stack);
    return null;
  }
}

/**
 * サイン画像をIMAGE数式でセルに挿入（PDF出力対応）
 * @param {Sheet} sheet - 対象シート
 * @param {string} imageUrl - Drive URL
 * @param {string} cellRange - セル範囲（例: 'G24'）
 * @return {string|null} 一時ファイルID（クリーンアップ用）
 */
function pdfInsertSignAsFormula_(sheet, imageUrl, cellRange) {
  var fileId = '';
  var match = imageUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) fileId = match[1];
  if (!fileId) {
    match = imageUrl.match(/id=([a-zA-Z0-9_-]+)/);
    if (match) fileId = match[1];
  }
  if (!fileId) {
    console.warn('サイン画像URLからファイルIDを抽出できません: ' + imageUrl);
    return null;
  }

  // サイン画像を共有設定してIMAGE数式で参照（blob操作なし）
  var readyId = prepareStampForImageFormula_(fileId);
  if (readyId) {
    var url = 'https://drive.google.com/uc?export=view&id=' + readyId;
    sheet.getRange(cellRange).setFormula('=IMAGE("' + url + '")');
    console.log('サイン画像IMAGE数式挿入成功: ' + cellRange);
  }
  return null; // 元ファイル使用のため一時ファイルクリーンアップ不要
}

// ====== PDFエクスポート ======

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
    + '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4'
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
