// ====== 有給警告メール ======

/** メニューから呼べる公開ラッパー（日付チェックをスキップして強制送信） */
function checkAndSendWarnMails() {
  checkAndSendWarnMails_();
}

/**
 * 毎日08:00に実行。当日がWARN_MAIL_DATESに該当すれば警告メール送信。
 */
function checkAndSendWarnMails_() {
  var now = new Date();
  var month = now.getMonth() + 1; // 1-12
  var day = now.getDate();

  // 当日が警告送信日か判定
  var isWarnDay = false;
  for (var i = 0; i < WARN_MAIL_DATES.length; i++) {
    if (WARN_MAIL_DATES[i].month === month && WARN_MAIL_DATES[i].day === day) {
      isWarnDay = true;
      break;
    }
  }

  if (!isWarnDay) {
    Logger.log('本日は有給警告送信日ではありません: ' + fmtDate_(now, 'yyyy/MM/dd'));
    return;
  }

  Logger.log('有給警告メール送信開始: ' + fmtDate_(now, 'yyyy/MM/dd'));

  // サマリー再構築
  rebuildLeaveSummary_();

  // 未達者抽出
  var fy = computeFiscalYear_(now);
  var underWorkers = getUnderCountWorkers_(fy);

  if (underWorkers.length === 0) {
    Logger.log('有給未達者なし');
    return;
  }

  // メール送信
  var settings = getSettings_();
  var mailTo = normalize_(settings['MAIL_TO']);
  if (!mailTo) {
    Logger.log('MAIL_TO未設定のため警告メール送信スキップ');
    return;
  }

  var subject = '【有給取得警告】有給取得義務未達者一覧 ' + fmtDate_(now, 'yyyy/MM/dd');
  var body = buildWarnMailBody_(underWorkers, fy, now);

  GmailApp.sendEmail(mailTo, subject, body);
  Logger.log('有給警告メール送信完了: ' + underWorkers.length + '名');
}

/**
 * 有給未達者を抽出
 */
function getUnderCountWorkers_(fyYear) {
  var info = getSheetHeaderIndex_(SHEET.LEAVE_SUMMARY, 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (Number(row[idx['年度']]) !== Number(fyYear)) continue;

    var warnLevel = normalize_(row[idx['警告レベル']]);
    if (warnLevel) {
      out.push({
        deptId: normalize_(row[idx['部署ID']]),
        deptName: normalize_(row[idx['部署名']]),
        workerId: normalize_(row[idx['作業員ID']]),
        workerName: normalize_(row[idx['作業員名']]),
        paidCount: Number(row[idx['有給回数']] || 0),
        warnLevel: warnLevel,
        lastPaidDate: normalize_(row[idx['最終有給日']]),
      });
    }
  }

  // 部署別→氏名順
  out.sort(function(a, b) {
    return (a.deptName + a.workerName).localeCompare(b.deptName + b.workerName, 'ja');
  });

  return out;
}

/**
 * 警告メール本文生成
 */
function buildWarnMailBody_(workers, fyYear, dateObj) {
  var settings = getSettings_();
  var appUrl = normalize_(settings['APP_URL'] || '');

  var lines = [];
  lines.push('【有給取得義務 未達者一覧】');
  lines.push('');
  lines.push('年度: ' + fyYear + '年度（' + fyYear + '/4/1 - ' + (Number(fyYear) + 1) + '/3/31）');
  lines.push('確認日: ' + fmtDate_(dateObj, 'yyyy/MM/dd'));
  lines.push('');

  // 現在の必要回数を表示
  var now = dateObj || new Date();
  var currentRequired = 0;
  var currentLabel = '';
  for (var i = 0; i < PAID_LEAVE_WARN.length; i++) {
    var w = PAID_LEAVE_WARN[i];
    var deadlineMonth = w.deadline.month;
    var deadlineYear = deadlineMonth >= 4 ? Number(fyYear) : Number(fyYear) + 1;
    var deadlineDate = new Date(deadlineYear, deadlineMonth - 1, w.deadline.day, 23, 59, 59);
    if (now > deadlineDate) {
      currentRequired = w.required;
      currentLabel = w.label;
    }
  }
  lines.push('現在の必要有給回数: ' + currentRequired + '回（' + currentLabel + '基準）');
  lines.push('');

  // 部署別にグルーピング
  var deptGroups = {};
  for (var j = 0; j < workers.length; j++) {
    var w = workers[j];
    if (!deptGroups[w.deptName]) deptGroups[w.deptName] = [];
    deptGroups[w.deptName].push(w);
  }

  var deptNames = Object.keys(deptGroups).sort();
  for (var d = 0; d < deptNames.length; d++) {
    var dept = deptNames[d];
    lines.push('■ ' + dept);
    var members = deptGroups[dept];
    for (var m = 0; m < members.length; m++) {
      var member = members[m];
      var levelMark = member.warnLevel === '危険' ? '!!!' : member.warnLevel === '警告' ? '!!' : '!';
      lines.push('  ' + levelMark + ' ' + member.workerName
        + '  有給: ' + member.paidCount + '回'
        + '  （' + member.warnLevel + '）'
        + (member.lastPaidDate ? '  最終: ' + member.lastPaidDate : ''));
    }
    lines.push('');
  }

  lines.push('対象者数: ' + workers.length + '名');
  lines.push('');

  if (appUrl) {
    lines.push('管理画面: ' + appUrl + '?page=admin');
    lines.push('');
  }
  lines.push('※本メールは自動送信です。');

  return lines.join('\n');
}
