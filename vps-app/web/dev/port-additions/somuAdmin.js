
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
