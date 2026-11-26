/* preview.js — GAS 版 src/preview.html の <script> をそのまま写したもの。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
var __T = JSON.parse(document.getElementById('gas-template-data').textContent);
var REQ_ID = __T.REQ_ID;
var APP_URL = __T.APP_URL;
var reqData = null;
var DOWS = ['日','月','火','水','木','金','土'];

window.addEventListener('DOMContentLoaded', function() {
  if (!REQ_ID) {
    document.getElementById('loadingMsg').textContent = 'REQ_IDが指定されていません。';
    return;
  }
  google.script.run
    .withSuccessHandler(function(data) {
      if (!data) {
        document.getElementById('loadingMsg').textContent = '申請データが見つかりません。（REQ_ID: ' + REQ_ID + '）';
        return;
      }
      reqData = data;
      try { renderPreview(data); } catch (err) {
        console.error('renderPreview error:', err);
        document.getElementById('loadingMsg').textContent = '表示エラー: ' + err.message;
      }
    })
    .withFailureHandler(function(err) {
      document.getElementById('loadingMsg').textContent = 'エラー: ' + err.message;
    })
    .api_getLeaveRequestById(REQ_ID);
});

function renderPreview(d) {
  var ld = d.leaveDate ? new Date(d.leaveDate) : new Date();
  var ry = ld.getFullYear() - 2018;

  // 届出日
  var sd = d.submittedAt ? new Date(d.submittedAt) : new Date();
  var submitStr = '令和' + (sd.getFullYear() - 2018) + '年'
    + (sd.getMonth() + 1) + '月' + sd.getDate() + '日 届出';

  // 休暇日
  var leaveDateStr = '令和' + ry + '年' + (ld.getMonth() + 1) + '月'
    + ld.getDate() + '日（' + DOWS[ld.getDay()] + '曜日）';

  // 休暇時間
  var timeStr = '終日　8:05 〜 17:10';
  if (d.leaveKubun === '半日休') {
    if (d.halfDayType === '午前') {
      timeStr = '午前半休　8:05 〜 12:15';
    } else {
      timeStr = '午後半休　13:00 〜 17:10';
    }
  }

  // 休暇区分
  var kubunStr = d.leaveKubun || '';
  if (d.halfDayType) kubunStr += '（' + d.halfDayType + '）';

  // 振替元出勤日
  var subStr = '―';
  if (d.leaveType === '振替休暇' && d.substituteDate) {
    var subD = new Date(d.substituteDate);
    subStr = '令和' + (subD.getFullYear() - 2018) + '年'
      + (subD.getMonth() + 1) + '月' + subD.getDate() + '日（' + DOWS[subD.getDay()] + '）';
  }

  // 理由
  var reason = '';
  if (d.leaveType === '特別休暇' && d.specialReason) reason = d.specialReason;
  else if (d.leaveType === '有給消化（私用）') reason = d.paidDetail || '私用の為';
  else if (d.leaveType === '振替休暇') reason = '振替休暇';
  else reason = d.leaveType || '―';

  // 承認
  var isApproved = (d.status === '承認済');
  var at1 = '';
  if (d.approvedAt) {
    var a1 = new Date(d.approvedAt);
    at1 = (a1.getFullYear() - 2018) + '/' + (a1.getMonth() + 1) + '/' + a1.getDate();
  }
  var at2 = '';
  if (d.approvedAt2) {
    var a2 = new Date(d.approvedAt2);
    at2 = (a2.getFullYear() - 2018) + '/' + (a2.getMonth() + 1) + '/' + a2.getDate();
  }

  var h = '<div class="paper">';

  // タイトル帯
  h += '<div class="p-title">休 暇 届</div>';

  h += '<div class="p-body">';

  // 過去申請バッジ
  if (d.isRetroactive === true) {
    var daysAgoTxt = (typeof d.daysAgo === 'number') ? '（' + d.daysAgo + '日前）' : '';
    h += '<div style="margin:12px 0 0;padding:8px 12px;background:#fffbeb;border:1px solid #fcd34d;border-radius:8px;color:#7c2d12;font-size:12px;font-weight:700;">'
      + '&#128281; 過去申請' + e(daysAgoTxt) + ' ／ 事前に承認者・総務へ連絡済みであることをご確認ください。'
      + '</div>';
  }

  // 届出日
  h += '<div class="p-date">' + e(submitStr) + '</div>';

  // 所属部署・氏名
  h += '<table class="t">';
  h += '<tr><th>所 属 部 署</th><td>' + e(d.deptName) + '</td></tr>';
  h += '<tr><th>氏　　　名</th><td>' + e(d.workerName) + '</td></tr>';
  h += '</table>';

  // 届出内容
  h += '<div class="sec-head">届 出 内 容</div>';
  h += '<table class="t">';
  h += '<tr><th>休 暇 日</th><td style="font-size:16px;font-weight:900;color:#1a365d;">' + e(leaveDateStr) + '</td></tr>';
  h += '<tr><th>休 暇 時 間</th><td>' + e(timeStr) + '</td></tr>';
  h += '<tr><th>休 暇 区 分</th><td style="font-size:16px;font-weight:900;color:#1a365d;">' + e(kubunStr) + '</td></tr>';
  h += '<tr><th>休 暇 種 類</th><td>' + e(d.leaveType) + '</td></tr>';
  h += '<tr><th>振替元出勤日</th><td>' + e(subStr) + '</td></tr>';
  h += '<tr><th>理由・詳細</th><td style="white-space:pre-wrap;">' + e(reason) + '</td></tr>';
  h += '</table>';

  // 宣言文
  h += '<div class="declaration">上記のとおり休暇を届け出いたします。</div>';

  // 届出者
  h += '<table class="t t-applicant">';
  h += '<tr>';
  h += '<th>届 出 者</th>';
  h += '<td class="sub-label">氏名</td>';
  h += '<td class="name-val">' + e(d.workerName) + '</td>';
  h += '</tr>';
  h += '</table>';

  // 承認欄
  h += '<table class="t t-approval">';
  h += '<tr><td class="approval-header" colspan="3">承 認 欄</td></tr>';
  h += '<tr>';
  h += '<td class="role-header" style="width:80px;"></td>';
  h += '<td class="role-header">部　長</td>';
  h += '<td class="role-header">所 属 長</td>';
  h += '</tr>';
  h += '<tr>';
  h += '<td class="seal-label">承認印</td>';
  // 部長セル
  h += '<td class="seal-cell">';
  if (d.approvedBy2) {
    h += '<div class="stamp-text">' + e(d.approvedBy2) + '</div>';
  }
  h += '</td>';
  // 所属長セル
  h += '<td class="seal-cell">';
  if (isApproved) {
    h += '<div class="stamp-circle">承認済</div>';
  }
  h += '</td>';
  h += '</tr>';
  h += '<tr>';
  h += '<td class="date-label">承認日</td>';
  h += '<td class="date-cell">' + e(at2) + '</td>';
  h += '<td class="date-cell">' + e(at1) + '</td>';
  h += '</tr>';
  h += '</table>';

  // 注記
  h += '<div class="note">※後報の場合は、理由を詳しく記載すること。</div>';

  // REQ_ID
  h += '<div class="req-id">REQ_ID: ' + e(d.reqId) + '</div>';

  h += '</div>'; // p-body
  h += '</div>'; // paper

  document.getElementById('contentWrap').innerHTML = h;
  document.getElementById('actionBar').style.display = '';

  if (isApproved) {
    document.getElementById('approveBtn').style.display = 'none';
    var bb = document.getElementById('backBtn');
    bb.textContent = 'TOPに戻る';
    bb.style.background = 'linear-gradient(135deg,#1f9d55,#1bd67f)';
    bb.style.color = '#fff';
    bb.style.fontWeight = '800';
    bb.style.fontSize = '16px';
    bb.style.padding = '14px 32px';
    bb.style.boxShadow = '0 10px 22px rgba(31,157,85,.20)';
    bb.style.border = '0';
  } else {
    document.getElementById('backBtn').textContent = 'TOPに戻る';
  }
}

document.getElementById('backBtn').addEventListener('click', function() {
  window.top.location.href = APP_URL + '?page=top';
});
document.getElementById('approveBtn').addEventListener('click', function() {
  window.top.location.href = APP_URL + '?page=sign&reqId=' + REQ_ID;
});

function e(s) {
  if (s === null || s === undefined) return '';
  var d = document.createElement('div');
  d.textContent = String(s);
  return d.innerHTML.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
