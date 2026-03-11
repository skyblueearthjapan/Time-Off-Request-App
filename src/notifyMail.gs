// ====== スケジュール通知メール ======

/**
 * 総務メールアドレス一覧を取得
 * @return {string[]}
 */
function getSomuEmails_() {
  var settings = getSettings_();
  var raw = normalize_(settings['SOMU_EMAILS'] || '');
  return raw.split(',').map(function(s) { return s.trim().toLowerCase(); }).filter(Boolean);
}

/**
 * 指定部署の承認者メール一覧を取得
 * @param {string} deptName 部署名
 * @return {string[]}
 */
function getDeptApproverEmailsForDept_(deptName) {
  try {
    var sh = requireSheet_(SHEET.DEPT_APPROVERS);
    var data = sh.getDataRange().getValues();
    var emails = [];
    for (var r = 1; r < data.length; r++) {
      var dept = normalize_(String(data[r][0] || '')).trim();
      if (dept !== deptName) continue;
      var rawEmails = normalize_(String(data[r][1] || ''));
      var list = rawEmails.split(',').map(function(s) { return s.trim().toLowerCase(); }).filter(Boolean);
      for (var i = 0; i < list.length; i++) {
        if (emails.indexOf(list[i]) < 0) emails.push(list[i]);
      }
    }
    return emails;
  } catch (e) {
    console.error('getDeptApproverEmailsForDept_ エラー: ' + e.message);
    return [];
  }
}

// ====== 17:15頃 1次未承認通知 ======

/**
 * 申請済み（submitted）で1次承認されていない申請がある場合、
 * 該当部署の1次承認者 + 総務にメール送信
 */
function notifyPendingApproval1_() {
  var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var readCols = Math.max(info.header.length, sh.getLastColumn());
  var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();

  // 申請中の申請を部署別にグループ化（休暇日が今日以降のみ）
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var pendingByDept = {}; // { deptName: [{ workerName, leaveDate, leaveType }] }
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var reqId = normalize_(row[idx['REQ_ID']]);
    if (!reqId) continue;
    var status = normalize_(row[idx['承認状態']]);
    if (status !== STATUS.SUBMITTED) continue;

    var leaveDate = row[idx['休暇日']];
    // 過去の申請はスキップ
    if (leaveDate instanceof Date && leaveDate < today) continue;

    var deptName = normalize_(row[idx['部署名']]);
    var leaveDateStr = leaveDate instanceof Date ? fmtDate_(leaveDate, 'yyyy/MM/dd') : String(leaveDate || '');

    if (!pendingByDept[deptName]) pendingByDept[deptName] = [];
    pendingByDept[deptName].push({
      workerName: normalize_(row[idx['作業員名']]),
      leaveDate: leaveDateStr,
      leaveType: normalize_(row[idx['休暇種類']]),
      leaveKubun: normalize_(row[idx['休暇区分']]),
    });
  }

  var deptNames = Object.keys(pendingByDept);
  if (deptNames.length === 0) {
    Logger.log('1次未承認の申請なし');
    return;
  }

  var settings = getSettings_();
  var appUrl = normalize_(settings['APP_URL'] || '');
  var somuEmails = getSomuEmails_();
  var totalCount = 0;

  // 部署ごとにメール送信
  for (var d = 0; d < deptNames.length; d++) {
    var dept = deptNames[d];
    var items = pendingByDept[dept];
    totalCount += items.length;

    // 宛先: 該当部署の承認者 + 総務
    var approverEmails = getDeptApproverEmailsForDept_(dept);
    var allTo = approverEmails.concat(somuEmails);
    // 重複除去
    var uniqueTo = [];
    for (var u = 0; u < allTo.length; u++) {
      if (uniqueTo.indexOf(allTo[u]) < 0) uniqueTo.push(allTo[u]);
    }
    if (uniqueTo.length === 0) continue;

    var subject = '【休暇届 未承認通知】' + dept + ' ' + items.length + '件が未承認です';
    var lines = [];
    lines.push('【休暇届 1次承認未完了通知】');
    lines.push('');
    lines.push('以下の休暇届が申請されていますが、まだ1次承認されていません。');
    lines.push('ご確認の上、承認をお願いいたします。');
    lines.push('');
    lines.push('■ ' + dept + '（' + items.length + '件）');
    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      lines.push('  - ' + it.workerName + '　' + it.leaveDate + '　' + it.leaveType + '（' + it.leaveKubun + '）');
    }
    lines.push('');
    if (appUrl) {
      lines.push('承認はこちらから（申請TOP）:');
      lines.push(appUrl + '?page=top');
      lines.push('');
    }
    lines.push('※本メールは自動送信です。');

    GmailApp.sendEmail(uniqueTo.join(','), subject, lines.join('\n'));
  }

  Logger.log('1次未承認通知送信完了: ' + deptNames.length + '部署, ' + totalCount + '件');
}

// ====== 18:00頃 2次未承認通知 ======

/**
 * 1次承認済み（approved）で2次承認されていない申請がある場合、
 * 総務部にメール送信
 */
function notifyPendingApproval2_() {
  var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var readCols = Math.max(info.header.length, sh.getLastColumn());
  var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();

  var at2Col = idx['APPROVED_AT2'] !== undefined ? idx['APPROVED_AT2'] : idx['2次承認日時'];

  var pendingItems = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var reqId = normalize_(row[idx['REQ_ID']]);
    if (!reqId) continue;
    var status = normalize_(row[idx['承認状態']]);
    if (status !== STATUS.APPROVED) continue;

    // 2次承認済みはスキップ
    if (at2Col !== undefined && row[at2Col] && String(row[at2Col]).trim() !== '') continue;

    var leaveDate = row[idx['休暇日']];
    pendingItems.push({
      deptName: normalize_(row[idx['部署名']]),
      workerName: normalize_(row[idx['作業員名']]),
      leaveDate: leaveDate instanceof Date ? fmtDate_(leaveDate, 'yyyy/MM/dd') : String(leaveDate || ''),
      leaveType: normalize_(row[idx['休暇種類']]),
      leaveKubun: normalize_(row[idx['休暇区分']]),
      approvedAt: row[idx['承認日時']] instanceof Date ? fmtDate_(row[idx['承認日時']], 'yyyy/MM/dd HH:mm') : '',
    });
  }

  if (pendingItems.length === 0) {
    Logger.log('2次未承認の申請なし');
    return;
  }

  var somuEmails = getSomuEmails_();
  if (somuEmails.length === 0) {
    Logger.log('SOMU_EMAILS未設定のため2次未承認通知スキップ');
    return;
  }

  var settings = getSettings_();
  var appUrl = normalize_(settings['APP_URL'] || '');

  var subject = '【休暇届 2次承認待ち】' + pendingItems.length + '件が2次承認待ちです';
  var lines = [];
  lines.push('【休暇届 2次承認未完了通知】');
  lines.push('');
  lines.push('以下の休暇届が1次承認済みですが、2次承認がまだ完了していません。');
  lines.push('ご確認の上、2次承認をお願いいたします。');
  lines.push('');

  // 部署別にグループ化
  var byDept = {};
  for (var j = 0; j < pendingItems.length; j++) {
    var it = pendingItems[j];
    if (!byDept[it.deptName]) byDept[it.deptName] = [];
    byDept[it.deptName].push(it);
  }
  var depts = Object.keys(byDept).sort();
  for (var d = 0; d < depts.length; d++) {
    var dept = depts[d];
    var members = byDept[dept];
    lines.push('■ ' + dept + '（' + members.length + '件）');
    for (var m = 0; m < members.length; m++) {
      var mem = members[m];
      lines.push('  - ' + mem.workerName + '　' + mem.leaveDate + '　' + mem.leaveType
        + '（' + mem.leaveKubun + '）　1次承認: ' + mem.approvedAt);
    }
    lines.push('');
  }

  lines.push('合計: ' + pendingItems.length + '件');
  lines.push('');
  if (appUrl) {
    lines.push('2次承認はこちらから（総務管理画面）:');
    lines.push(appUrl + '?page=somuAdmin');
    lines.push('');
  }
  lines.push('※本メールは自動送信です。');

  GmailApp.sendEmail(somuEmails.join(','), subject, lines.join('\n'));
  Logger.log('2次未承認通知送信完了: ' + pendingItems.length + '件');
}

// ====== 7:00 本日休暇者通知 ======

/**
 * 本日が休暇日の承認済み申請を総務にメール送信
 */
function notifyTodayLeaves_() {
  var info = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var today = new Date();
  var todayStr = fmtDate_(today, 'yyyy/MM/dd');
  var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
  var todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);

  var readCols = Math.max(info.header.length, sh.getLastColumn());
  var values = sh.getRange(2, 1, lastRow - 1, readCols).getValues();

  var todayLeaves = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var reqId = normalize_(row[idx['REQ_ID']]);
    if (!reqId) continue;
    var status = normalize_(row[idx['承認状態']]);
    // 承認済み（1次 or 2次）のみ
    if (status !== STATUS.APPROVED) continue;

    var leaveDate = row[idx['休暇日']];
    if (!(leaveDate instanceof Date)) continue;
    if (leaveDate < todayStart || leaveDate > todayEnd) continue;

    todayLeaves.push({
      deptName: normalize_(row[idx['部署名']]),
      workerName: normalize_(row[idx['作業員名']]),
      leaveType: normalize_(row[idx['休暇種類']]),
      leaveKubun: normalize_(row[idx['休暇区分']]),
      halfDayType: normalize_(row[idx['半日区分']]),
    });
  }

  if (todayLeaves.length === 0) {
    Logger.log('本日の休暇者なし: ' + todayStr);
    return;
  }

  var somuEmails = getSomuEmails_();
  if (somuEmails.length === 0) {
    Logger.log('SOMU_EMAILS未設定のため本日休暇者通知スキップ');
    return;
  }

  var settings = getSettings_();
  var appUrl = normalize_(settings['APP_URL'] || '');
  var pdfFolderId = normalize_(settings['PDF_FOLDER_ID'] || '');

  var subject = '【本日の休暇者】' + todayStr + ' ' + todayLeaves.length + '名';
  var lines = [];
  lines.push('【本日の休暇者一覧】');
  lines.push('');
  lines.push('日付: ' + todayStr);
  lines.push('');

  // 部署別にグループ化
  var byDept = {};
  for (var j = 0; j < todayLeaves.length; j++) {
    var it = todayLeaves[j];
    if (!byDept[it.deptName]) byDept[it.deptName] = [];
    byDept[it.deptName].push(it);
  }
  var depts = Object.keys(byDept).sort();
  for (var d = 0; d < depts.length; d++) {
    var dept = depts[d];
    var members = byDept[dept];
    lines.push('■ ' + dept);
    for (var m = 0; m < members.length; m++) {
      var mem = members[m];
      var timeInfo = mem.leaveKubun;
      if (mem.halfDayType) timeInfo += '（' + mem.halfDayType + '）';
      lines.push('  - ' + mem.workerName + '　' + mem.leaveType + '　' + timeInfo);
    }
    lines.push('');
  }

  lines.push('合計: ' + todayLeaves.length + '名');
  lines.push('');
  if (pdfFolderId) {
    lines.push('PDFフォルダ:');
    lines.push('https://drive.google.com/drive/folders/' + pdfFolderId);
    lines.push('');
  }
  if (appUrl) {
    lines.push('総務管理画面:');
    lines.push(appUrl + '?page=somuAdmin');
    lines.push('');
  }
  lines.push('※本メールは自動送信です。');

  GmailApp.sendEmail(somuEmails.join(','), subject, lines.join('\n'));
  Logger.log('本日休暇者通知送信完了: ' + todayLeaves.length + '名');
}
