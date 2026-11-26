/* somuAdmin.js — GAS 版 src/somuAdmin.html の <script> をそのまま写したもの。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
var __T = JSON.parse(document.getElementById('gas-template-data').textContent);
var APP_URL = __T.APP_URL;
var PORTAL_URL = __T.PORTAL_URL;
var currentData = [];
var pending2Data = [];
var selectedApprover2 = '';

window.addEventListener('DOMContentLoaded', function() {
  if (PORTAL_URL) document.getElementById('portalLink').style.display = '';
  google.script.run.withSuccessHandler(function(ok) {
    if (ok) document.getElementById('navDept').style.display = '';
  }).api_isDeptApprover();
  google.script.run.withSuccessHandler(function(ok) {
    if (ok) document.getElementById('navDept').style.display = '';
  }).api_isAdmin();

  // 年度セレクト初期化
  var fySelect = document.getElementById('fySelect');
  var now = new Date();
  var _m = now.getMonth() + 1, _d = now.getDate();
  var currentFy = (_m < 3 || (_m === 3 && _d < 16)) ? now.getFullYear() - 1 : now.getFullYear();
  for (var y = currentFy + 2; y >= currentFy - 2; y--) {
    var opt = document.createElement('option');
    opt.value = y;
    opt.textContent = y + '年度';
    if (y === currentFy) opt.selected = true;
    fySelect.appendChild(opt);
  }

  // 部署フィルタ読込
  google.script.run.withSuccessHandler(function(depts) {
    var sel = document.getElementById('deptFilter');
    depts.forEach(function(d) {
      var opt = document.createElement('option');
      opt.value = d.deptId;
      opt.textContent = d.deptName;
      sel.appendChild(opt);
    });
  }).api_getDeptList();

  document.getElementById('deptFilter').addEventListener('change', function() { loadData(); });

  // 2次承認者（ログインユーザー）取得
  loadApprover2Info();

  // 2次承認待ち読込
  loadPending2();

  // 有給データ読込
  loadData();
});

// ====== 2次承認者（ログインユーザー自動取得） ======
function loadApprover2Info() {
  google.script.run.withSuccessHandler(function(info) {
    // info = { email, name }
    selectedApprover2 = info.email;
    var display = document.getElementById('approver2Name');
    display.innerHTML = '<strong>' + escHtml(info.name) + '</strong>（' + escHtml(info.email) + '）';
    updateBatchBtn();
  }).withFailureHandler(function(err) {
    document.getElementById('approver2Name').textContent = 'エラー: ' + err.message;
  }).api_getCurrentApproverInfo();
}

// ====== 2次承認待ち一覧 ======
function loadPending2() {
  var wrap = document.getElementById('approval2Wrap');
  wrap.innerHTML = '<div class="loading">読込中...</div>';
  google.script.run.withSuccessHandler(function(data) {
    pending2Data = data || [];
    document.getElementById('pending2Count').textContent = pending2Data.length + '件';
    renderPending2();
  }).withFailureHandler(function(err) {
    wrap.innerHTML = '<div class="loading" style="color:var(--danger);">エラー: ' + escHtml(err.message) + '</div>';
    document.getElementById('pending2Count').textContent = 'エラー';
  }).api_getPending2ndApprovals();
}

function renderPending2() {
  var wrap = document.getElementById('approval2Wrap');
  if (!pending2Data.length) {
    wrap.innerHTML = '<div class="loading" style="color:var(--accent);">2次承認待ちの申請はありません</div>';
    document.getElementById('batchBar').style.display = 'none';
    return;
  }
  document.getElementById('batchBar').style.display = 'flex';

  var html = '<table><thead><tr>'
    + '<th style="width:30px;"></th>'
    + '<th>部署</th><th>氏名</th><th>休暇日</th><th>休暇種類</th><th>区分</th><th>1次承認日</th><th>操作</th>'
    + '</tr></thead><tbody>';

  for (var i = 0; i < pending2Data.length; i++) {
    var d = pending2Data[i];
    var retroBadge = '';
    var rowExtra = '';
    if (d.isRetroactive === true) {
      retroBadge = ' <span style="display:inline-block;padding:2px 8px;border-radius:999px;background:#fee2e2;color:#b91c1c;border:1px solid #fca5a5;font-size:11px;font-weight:800;">&#128281; 過去申請</span>'
        + (typeof d.daysAgo === 'number' ? ' <span style="font-size:11px;color:var(--muted);">(' + __escText(d.daysAgo) + '日前)</span>' : '');
      rowExtra = 'background:#fffbeb;';
    }
    html += '<tr data-reqid="' + escHtml(d.reqId) + '" style="transition:all .5s ease;' + rowExtra + '">'
      + '<td><input type="checkbox" class="row-check" data-idx="' + i + '" data-on-change="updateSelectedCount"></td>'
      + '<td>' + escHtml(d.deptName) + '</td>'
      + '<td>' + escHtml(d.workerName) + retroBadge + '</td>'
      + '<td>' + escHtml(d.leaveDate) + '</td>'
      + '<td>' + escHtml(d.leaveType) + '</td>'
      + '<td>' + escHtml(d.leaveKubun) + '</td>'
      + '<td>' + escHtml(d.approvedAt) + '</td>'
      + '<td><button class="btn-sm" data-on-click="singleApprove2" data-args="' + __vpsArgs(i, '$this') + '">2次承認</button></td>'
      + '</tr>';
  }
  html += '</tbody></table>';
  wrap.innerHTML = html;
}

function fadeOutRow(reqId, callback) {
  var row = document.querySelector('tr[data-reqid="' + __vpsAttrSel(reqId) + '"]');
  if (row) {
    row.style.background = '#e6ffed';
    setTimeout(function() {
      row.style.opacity = '0';
      row.style.transform = 'translateX(40px)';
      row.style.maxHeight = row.offsetHeight + 'px';
      row.style.overflow = 'hidden';
      setTimeout(function() {
        row.style.maxHeight = '0';
        row.style.padding = '0';
        setTimeout(function() {
          row.remove();
          // カウント更新
          var remaining = document.querySelectorAll('tr[data-reqid]').length;
          document.getElementById('pending2Count').textContent = remaining + '件';
          if (remaining === 0) {
            document.getElementById('approval2Wrap').innerHTML = '<div class="loading" style="color:var(--accent);">2次承認待ちの申請はありません</div>';
            document.getElementById('batchBar').style.display = 'none';
          }
          if (callback) callback();
        }, 300);
      }, 300);
    }, 400);
  } else {
    if (callback) callback();
  }
}

function singleApprove2(idx, btn) {
  if (!selectedApprover2) {
    alert('2次承認者（電子印）を選択してください。');
    return;
  }
  var d = pending2Data[idx];
  if (!confirm(d.workerName + 'さんの休暇届（' + d.leaveDate + '）を2次承認しますか？')) return;

  btn.disabled = true;
  btn.textContent = '処理中...';
  google.script.run
    .withSuccessHandler(function(result) {
      if (result && result.ok) {
        btn.textContent = '✓';
        fadeOutRow(d.reqId);
      } else {
        btn.textContent = '失敗';
        btn.disabled = false;
      }
    })
    .withFailureHandler(function(err) {
      alert('エラー: ' + err.message);
      btn.textContent = '2次承認';
      btn.disabled = false;
    })
    .api_approveLeaveRequest2(d.reqId, selectedApprover2);
}

function toggleSelectAll() {
  var checked = document.getElementById('selectAll').checked;
  var boxes = document.querySelectorAll('.row-check');
  for (var i = 0; i < boxes.length; i++) boxes[i].checked = checked;
  updateSelectedCount();
}

function updateSelectedCount() {
  var boxes = document.querySelectorAll('.row-check:checked');
  document.getElementById('selectedCount').textContent = boxes.length + '件選択中';
  updateBatchBtn();
}

function updateBatchBtn() {
  var boxes = document.querySelectorAll('.row-check:checked');
  document.getElementById('batchApproveBtn').disabled = boxes.length === 0 || !selectedApprover2;
}

function batchApprove2() {
  if (!selectedApprover2) {
    alert('2次承認者（電子印）を選択してください。');
    return;
  }
  var boxes = document.querySelectorAll('.row-check:checked');
  var reqIds = [];
  for (var i = 0; i < boxes.length; i++) {
    var idx = Number(boxes[i].dataset.idx);
    reqIds.push(pending2Data[idx].reqId);
  }
  if (!reqIds.length) return;
  if (!confirm(reqIds.length + '件を一括2次承認しますか？\n（1件あたり約15秒かかります）')) return;

  var btn = document.getElementById('batchApproveBtn');
  btn.disabled = true;
  document.getElementById('selectAll').disabled = true;
  var boxes2 = document.querySelectorAll('.row-check');
  for (var b = 0; b < boxes2.length; b++) boxes2[b].disabled = true;

  var total = reqIds.length;
  var okCount = 0;
  var errCount = 0;
  var current = 0;

  function processNext() {
    if (current >= total) {
      btn.textContent = '一括2次承認';
      document.getElementById('selectAll').disabled = false;
      if (errCount > 0) {
        alert('完了: ' + okCount + '件成功, ' + errCount + '件失敗');
        loadPending2();
      }
      updateBatchBtn();
      return;
    }
    var reqId = reqIds[current];
    btn.textContent = '承認中... (' + (current + 1) + '/' + total + ')';

    google.script.run
      .withSuccessHandler(function() {
        okCount++;
        current++;
        fadeOutRow(reqId, function() { processNext(); });
      })
      .withFailureHandler(function(err) {
        console.error('2次承認エラー: ' + reqId + ' ' + err.message);
        errCount++;
        current++;
        // エラー行は赤くする
        var row = document.querySelector('tr[data-reqid="' + __vpsAttrSel(reqId) + '"]');
        if (row) { row.style.background = '#ffe0e0'; }
        processNext();
      })
      .api_approveLeaveRequest2(reqId, selectedApprover2);
  }
  processNext();
}

// ====== 有給管理 ======
function loadData() {
  var wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '<div class="loading">読込中...</div>';
  var fy = Number(document.getElementById('fySelect').value);
  var deptFilter = document.getElementById('deptFilter').value;

  google.script.run.withSuccessHandler(function(data) {
    currentData = data;
    renderKpi(data);
    renderTable(data);
  }).withFailureHandler(function(err) {
    wrap.innerHTML = '<div class="loading" style="color:var(--danger);">エラー: ' + escHtml(err.message) + '</div>';
  }).api_getSomuAdminView(fy, deptFilter);
}

function renderKpi(data) {
  var total = data.length;
  var warnCount = 0;
  var dangerCount = 0;
  var paidSum = 0;

  data.forEach(function(d) {
    paidSum += (d.paidCount || 0);
    if (d.warnLevel !== '') warnCount++;
    if (d.warnLevel === '危険') dangerCount++;
  });

  document.getElementById('kpiTotal').textContent = total;
  document.getElementById('kpiWarn').textContent = warnCount;
  document.getElementById('kpiDanger').textContent = dangerCount;
  document.getElementById('kpiAvg').textContent = total > 0 ? (paidSum / total).toFixed(1) : '-';
}

function renderTable(data) {
  var wrap = document.getElementById('tableWrap');
  if (!data.length) {
    wrap.innerHTML = '<div class="loading">データなし</div>';
    return;
  }

  var html = '<table><thead><tr>'
    + '<th>部署</th><th>氏名</th><th>有給</th>'
    + '<th>春期(3/16-6月)<br><span style="font-size:10px;color:var(--muted);">目標1回</span></th><th>夏期(7-8月)<br><span style="font-size:10px;color:var(--muted);">目標2回</span></th><th>秋期(9-11月)<br><span style="font-size:10px;color:var(--muted);">目標3回</span></th><th>冬期(12-2/15)<br><span style="font-size:10px;color:var(--muted);">目標4回</span></th><th>年度末(2/16-3/15)<br><span style="font-size:10px;color:var(--muted);">目標5回</span></th>'
    + '<th>振替</th><th>特別</th><th>合計</th><th>最終有給日</th><th>警告</th>'
    + '</tr></thead><tbody>';

  data.forEach(function(d) {
    var rowClass = '';
    if (d.warnLevel === '危険') rowClass = 'danger-row';
    else if (d.warnLevel === '警告' || d.warnLevel === '注意') rowClass = 'warn-row';

    var warnPill = '';
    if (d.warnLevel === '危険') warnPill = '<span class="pill danger">' + escHtml(d.warnLevel) + '</span>';
    else if (d.warnLevel === '警告') warnPill = '<span class="pill warn">' + escHtml(d.warnLevel) + '</span>';
    else if (d.warnLevel === '注意') warnPill = '<span class="pill info">' + escHtml(d.warnLevel) + '</span>';
    else warnPill = '<span class="pill ok">OK</span>';

    var quarters = d.quarters || [];

    html += '<tr class="' + rowClass + '">'
      + '<td>' + escHtml(d.deptName) + '</td>'
      + '<td>' + escHtml(d.workerName) + '</td>'
      + '<td>' + __escText(d.paidCount) + (d.paidHalfCount > 0 ? ' <span style="font-size:11px;color:var(--muted);">(' + __escText(d.paidHalfCount) + '半)</span>' : '') + '</td>'
      + renderQuarterCell(quarters[0])
      + renderQuarterCell(quarters[1])
      + renderQuarterCell(quarters[2])
      + renderQuarterCell(quarters[3])
      + renderQuarterCell(quarters[4])
      + '<td>' + __escText(d.substituteCount) + '</td>'
      + '<td>' + __escText(d.specialCount) + '</td>'
      + '<td style="font-weight:700;">' + __escText(d.totalCount) + '</td>'
      + '<td>' + __escText(d.lastPaidDate || '-') + '</td>'
      + '<td>' + warnPill + '</td>'
      + '</tr>';
  });

  html += '</tbody></table>';
  wrap.innerHTML = html;
}

function renderQuarterCell(q) {
  if (!q) return '<td><span class="pill">―</span></td>';
  var remaining = Math.max(0, 5 - Number(q.cumulativeCount || 0));
  var sh = Number(q.shortage || 0);
  if (q.isFuture) return '<td><span class="pill">残' + remaining + '</span></td>';
  if (remaining === 0) return '<td><span class="pill ok">達成</span></td>';
  if (sh >= 2) return '<td><span class="pill danger">残' + remaining + '</span></td>';
  if (sh >= 1) return '<td><span class="pill warn">残' + remaining + '</span></td>';
  return '<td><span class="pill ok">残' + remaining + '</span></td>';
}

function exportCsv() {
  if (!currentData.length) { alert('データがありません'); return; }
  var headers = ['部署','氏名','有給回数','有給半日回数','春期(3/16-6月)','夏期(7-8月)','秋期(9-11月)','冬期(12-2/15)','年度末(2/16-3/15)','振替回数','特別回数','合計回数','最終有給日','警告レベル'];
  var lines = [headers.join(',')];

  currentData.forEach(function(d) {
    var quarters = d.quarters || [];
    lines.push([
      csvEsc(d.deptName),
      csvEsc(d.workerName),
      d.paidCount,
      d.paidHalfCount,
      quarterText(quarters[0]),
      quarterText(quarters[1]),
      quarterText(quarters[2]),
      quarterText(quarters[3]),
      quarterText(quarters[4]),
      d.substituteCount,
      d.specialCount,
      d.totalCount,
      csvEsc(d.lastPaidDate || ''),
      csvEsc(d.warnLevel || 'OK')
    ].join(','));
  });

  var csv = lines.join('\n');
  var bom = '\uFEFF';
  var blob = new Blob([bom + csv], {type:'text/csv;charset=utf-8;'});
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = '有給取得管理_' + document.getElementById('fySelect').value + '年度.csv';
  a.click();
  URL.revokeObjectURL(url);
}

function quarterText(q) {
  if (!q) return '―';
  if (q.taken) return '✓';
  if (q.overdue) return '✗';
  return '―';
}

function csvEsc(s) {
  s = String(s || '');
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g,'""') + '"';
  return s;
}

function escHtml(s) {
  if (!s) return '';
  var d = document.createElement('div');
  d.textContent = String(s);
  return d.innerHTML.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/* ===================================================================
 * VPS 版で足した部分（ここから下だけ。上は GAS 版 src/somuAdmin.html のまま）
 * =================================================================== */

// ====== 運用（管理者）の欄（仕様 3-11・4-2） ======
// 関数名をそのまま書く（google.script.run に生えている名前と照らし合わせる試験のため）
var OPS_CALLS = {
  ops_syncSheetConfigNow: function(r){ r.ops_syncSheetConfigNow(); },
  ops_runMasterSync: function(r){ r.ops_runMasterSync(); },
  ops_rebuildSummaryNow: function(r){ r.ops_rebuildSummaryNow(); },
  ops_sendWarnMail: function(r){ r.ops_sendWarnMail(); },
  ops_revokeSessions: function(r, v){ r.ops_revokeSessions(v); },
  ops_retryJob: function(r, v){ r.ops_retryJob(v); }
};
// 押す前の確かめの文（有給警告メールだけ決まった文。仕様 3-11）
var OPS_CONFIRM = {
  ops_sendWarnMail: '有給警告メールを送信しますか？\n（警告送信日（7/1・9/1・1/1・2/16）以外は送信されません）'
};

function opsLoadStatus(){
  var pre = document.getElementById('opsStatus');
  pre.textContent = '読込中...';
  google.script.run.withSuccessHandler(function(st){
    pre.textContent = JSON.stringify(st, null, 2);
    opsRenderFailed(st);
  }).withFailureHandler(function(err){
    pre.textContent = 'エラー: ' + (err.message || err);
  }).ops_getStatus();
}

// 状態の中の「失敗した仕事」の一覧（id を持つ配列）を探して、行ごとに［再実行］を付ける
function opsFindFailedJobs(st){
  var found = [];
  (function walk(o, depth){
    if (!o || typeof o !== 'object' || depth > 3) return;
    Object.keys(o).forEach(function(k){
      var v = o[k];
      if (Array.isArray(v) && /fail/i.test(k)) {
        v.forEach(function(j){ if (j && typeof j === 'object' && j.id !== undefined && j.id !== null) found.push(j); });
      } else if (v && typeof v === 'object' && !Array.isArray(v)) {
        walk(v, depth + 1);
      }
    });
  })(st, 0);
  return found;
}

function opsRenderFailed(st){
  var box = document.getElementById('opsFailed');
  var jobs = opsFindFailedJobs(st);
  if (jobs.length === 0) { box.innerHTML = '<div class="sub">失敗した仕事はありません</div>'; return; }
  var html = '<table><thead><tr><th>id</th><th>申請 ID</th><th>仕事</th><th>理由</th><th>回数</th><th></th></tr></thead><tbody>';
  jobs.forEach(function(j){
    html += '<tr>' +
      '<td>' + __escText(j.id) + '</td>' +
      '<td>' + __escText(j.request_id || j.requestId || j.reqId || '') + '</td>' +
      '<td>' + __escText(j.action || j.kind || '') + '</td>' +
      '<td>' + __escText(j.last_error || j.lastError || j.reason || '') + '</td>' +
      '<td>' + __escText(j.attempts !== undefined ? j.attempts : '') + '</td>' +
      '<td><button class="btn-sm ghost" data-on-click="opsRetryJob" data-args="' + __vpsArgs(j.id, '$this') + '">失敗した仕事の再実行</button></td>' +
      '</tr>';
  });
  box.innerHTML = html + '</tbody></table>';
}

function opsCall(fn, value, label){
  if (!confirm(OPS_CONFIRM[fn] || ('「' + label + '」を実行しますか？'))) return;
  var r = google.script.run.withSuccessHandler(function(res){
    alert('「' + label + '」を受け付けました。\n' + JSON.stringify(res));
    opsLoadStatus();
  }).withFailureHandler(function(err){
    alert('エラー: ' + (err.message || err));
  });
  OPS_CALLS[fn](r, value);
}

function opsRun(fn, btn){
  opsCall(fn, undefined, btn.textContent);
}

function opsRunWithInput(fn, inputId, btn){
  var v = document.getElementById(inputId).value.trim();
  if (!v) { alert('値を入力してください。'); return; }
  opsCall(fn, v, btn.textContent + '（' + v + '）');
}

function opsRetryJob(jobId, btn){
  opsCall('ops_retryJob', jobId, '失敗した仕事の再実行（id ' + jobId + '）');
}

(function(){
  var panel = document.getElementById('opsPanel');
  if (panel) panel.addEventListener('toggle', function(){ if (panel.open) opsLoadStatus(); });
})();

// ====== 操作中にログインが切れたときの書きかけ（2 次承認待ちの選択中のチェック）の退避と復元（仕様 3-5） ======
(function(){
  if (!window.__vpsDraft) return;
  var when = window.__vpsDraft.when;
  window.__vpsDraft.register({
    name: 'somuAdmin',
    collect: function(){
      var ids = [];
      var boxes = document.querySelectorAll('.row-check:checked');
      for (var i = 0; i < boxes.length; i++) {
        var d = pending2Data[Number(boxes[i].dataset.idx)];
        if (d) ids.push(d.reqId);
      }
      return ids.length ? { reqIds: ids } : null;
    },
    restore: function(d){
      if (!d || !d.reqIds) return;
      when(function(){ return document.querySelectorAll('.row-check').length > 0; }, function(){
        var boxes = document.querySelectorAll('.row-check');
        for (var i = 0; i < boxes.length; i++) {
          var row = pending2Data[Number(boxes[i].dataset.idx)];
          boxes[i].checked = !!row && d.reqIds.indexOf(row.reqId) >= 0;
        }
        updateSelectedCount();
      });
    }
  });
})();
