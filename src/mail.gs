// ====== 承認通知メール ======

/**
 * 承認完了時にメール通知を送信
 */
function sendApprovalNotification_(reqId) {
  var settings = getSettings_();
  var mailTo = normalize_(settings['MAIL_TO']);
  if (!mailTo) {
    console.warn('MAIL_TO未設定のためメール通知スキップ');
    return;
  }

  var data = api_getLeaveRequestById(reqId);
  if (!data) {
    console.warn('申請データなし: ' + reqId);
    return;
  }

  var subject = '【休暇届 承認完了】' + data.workerName + ' ' + data.leaveType;
  var body = buildApprovalMailBody_(data);

  GmailApp.sendEmail(mailTo, subject, body);
}

/**
 * 承認通知メール本文生成
 */
function buildApprovalMailBody_(data) {
  var leaveDate = '';
  if (data.leaveDate) {
    var d = data.leaveDate instanceof Date ? data.leaveDate : new Date(data.leaveDate);
    leaveDate = fmtDate_(d, 'yyyy/MM/dd');
  }

  var settings = getSettings_();
  var appUrl = normalize_(settings['APP_URL'] || '');

  var lines = [];
  lines.push('【休暇届 承認完了通知】');
  lines.push('');
  lines.push('以下の休暇届が承認されました。');
  lines.push('');
  lines.push('■ 申請情報');
  lines.push('  REQ_ID: ' + data.reqId);
  lines.push('  部署: ' + data.deptName);
  lines.push('  氏名: ' + data.workerName);
  lines.push('  休暇区分: ' + data.leaveKubun + (data.halfDayType ? '（' + data.halfDayType + '）' : ''));
  lines.push('  休暇日: ' + leaveDate);
  lines.push('  休暇種類: ' + data.leaveType);

  if (data.leaveType === '振替休暇' && data.substituteDate) {
    var subDate = data.substituteDate instanceof Date ? data.substituteDate : new Date(data.substituteDate);
    lines.push('  振替元出勤日: ' + fmtDate_(subDate, 'yyyy/MM/dd'));
  }

  if (data.specialReason) {
    lines.push('  特別理由: ' + data.specialReason);
  }

  lines.push('');
  if (appUrl) {
    lines.push('管理画面: ' + appUrl + '?page=admin');
    lines.push('');
  }
  lines.push('※本メールは自動送信です。');

  return lines.join('\n');
}
