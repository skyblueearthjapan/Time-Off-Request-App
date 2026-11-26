/* deptAdmin.js — GAS 版 src/deptAdmin.html の <script> をそのまま写したもの。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
var __T = JSON.parse(document.getElementById('gas-template-data').textContent);
var APP_URL = __T.APP_URL;
var PORTAL_URL = __T.PORTAL_URL;
var currentData = [];
var SELECTED_DEPTS = [];
var SELECTED_NAME = '';
var IS_ADMIN = false;
var ALL_PROFILES = [];

window.addEventListener('DOMContentLoaded', function() {
  if (PORTAL_URL) document.getElementById('portalLink').style.display = '';
  google.script.run.withSuccessHandler(function(ok) {
    if (ok) document.getElementById('navSomu').style.display = '';
  }).api_isSomu();

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

  // 承認者プロファイル取得 → モーダル or 自動選択
  initApprover();
});

function initApprover() {
  document.getElementById('headerSub').textContent = '承認者情報を読込中...';
  google.script.run
    .withSuccessHandler(function(result) {
      IS_ADMIN = result.isAdmin;
      ALL_PROFILES = result.profiles || [];

      if (IS_ADMIN) {
        // 管理者 → 承認者選択モーダルを表示
        document.getElementById('btnChangeApprover').style.display = '';
        showApproverModal();
      } else if (ALL_PROFILES.length === 1) {
        // 一般承認者（1人）→ 自動選択してダッシュボード表示
        selectApprover(ALL_PROFILES[0].name, ALL_PROFILES[0].depts);
      } else if (ALL_PROFILES.length > 1) {
        // 複数プロファイル（通常はない）→ モーダル表示
        showApproverModal();
      } else {
        document.getElementById('headerSub').textContent = '担当部署が登録されていません';
        document.getElementById('tableWrap').innerHTML = '<div class="loading">M_DEPT_APPROVERSに登録がありません。</div>';
      }
    })
    .withFailureHandler(function(err) {
      document.getElementById('headerSub').textContent = 'エラー: ' + err.message;
    })
    .api_getApproverProfiles();
}

function showApproverModal() {
  var container = document.getElementById('approverList');
  container.innerHTML = '';

  if (IS_ADMIN) {
    // 「全部署」ボタン
    var allDepts = [];
    for (var i = 0; i < ALL_PROFILES.length; i++) {
      for (var j = 0; j < ALL_PROFILES[i].depts.length; j++) {
        if (allDepts.indexOf(ALL_PROFILES[i].depts[j]) < 0) {
          allDepts.push(ALL_PROFILES[i].depts[j]);
        }
      }
    }
    var allBtn = document.createElement('button');
    allBtn.className = 'approver-btn all-dept';
    allBtn.innerHTML = '全部署<div class="dept-list">' + escHtml(allDepts.join(', ')) + '</div>';
    allBtn.onclick = (function(d) { return function() { selectApprover('全部署', d); }; })(allDepts);
    container.appendChild(allBtn);
  }

  // 各承認者ボタン
  for (var i = 0; i < ALL_PROFILES.length; i++) {
    var p = ALL_PROFILES[i];
    var btn = document.createElement('button');
    btn.className = 'approver-btn';
    btn.innerHTML = escHtml(p.name) + '<div class="dept-list">' + escHtml(p.depts.join(', ')) + '</div>';
    btn.onclick = (function(name, depts) { return function() { selectApprover(name, depts); }; })(p.name, p.depts);
    container.appendChild(btn);
  }

  document.getElementById('approverModal').classList.remove('hidden');
}

function selectApprover(name, depts) {
  SELECTED_NAME = name;
  SELECTED_DEPTS = depts;
  document.getElementById('approverModal').classList.add('hidden');
  document.getElementById('headerSub').textContent = '担当: ' + name + '（' + depts.join(', ') + '）';
  loadData();
}

function loadData() {
  if (!SELECTED_DEPTS.length) return;
  var wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '<div class="loading">読込中...</div>';
  var fy = Number(document.getElementById('fySelect').value);

  google.script.run
    .withSuccessHandler(function(result) {
      currentData = result.workers || [];
      renderKpi(currentData);
      renderTable(currentData);
    })
    .withFailureHandler(function(err) {
      wrap.innerHTML = '<div class="loading" style="color:var(--danger);">エラー: ' + escHtml(err.message) + '</div>';
    })
    .api_getDeptAdminView(fy, SELECTED_DEPTS);
}

function renderKpi(data) {
  var total = data.length;
  var warnCount = 0;
  var dangerCount = 0;
  var paidSum = 0;

  for (var i = 0; i < data.length; i++) {
    var d = data[i];
    paidSum += (d.paidCount || 0);
    if (d.warnLevel && d.warnLevel !== '') warnCount++;
    if (d.warnLevel === '危険') dangerCount++;
  }

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
    + '<th>春期<br>(3/16-6月)<br><span style="font-size:10px;color:var(--muted);">目標1回</span></th><th>夏期<br>(7-8月)<br><span style="font-size:10px;color:var(--muted);">目標2回</span></th><th>秋期<br>(9-11月)<br><span style="font-size:10px;color:var(--muted);">目標3回</span></th><th>冬期<br>(12-2/15)<br><span style="font-size:10px;color:var(--muted);">目標4回</span></th><th>年度末<br>(2/16-3/15)<br><span style="font-size:10px;color:var(--muted);">目標5回</span></th>'
    + '<th>振替</th><th>特別</th><th>合計</th><th>最終有給日</th><th>警告</th>'
    + '</tr></thead><tbody>';

  for (var i = 0; i < data.length; i++) {
    var d = data[i];
    var rowClass = '';
    if (d.warnLevel === '危険') rowClass = 'danger-row';
    else if (d.warnLevel === '警告') rowClass = 'warn-row';
    else if (d.warnLevel === '注意') rowClass = 'caution-row';

    // Quarter cells - 「残X」表示（5 - 累計取得数）で色分け
    var quarters = d.quarters || [];
    var qHtml = '';
    for (var qi = 0; qi < 5; qi++) {
      var q = quarters[qi];
      if (!q) {
        qHtml += '<td class="qcell"><span class="pill">―</span></td>';
      } else {
        var remaining = Math.max(0, 5 - Number(q.cumulativeCount || 0));
        var sh = Number(q.shortage || 0);
        if (q.isFuture) {
          qHtml += '<td class="qcell"><span class="pill">残' + remaining + '</span></td>';
        } else if (remaining === 0) {
          qHtml += '<td class="qcell"><span class="pill ok">達成</span></td>';
        } else if (sh >= 2) {
          qHtml += '<td class="qcell"><span class="pill danger">残' + remaining + '</span></td>';
        } else if (sh >= 1) {
          qHtml += '<td class="qcell"><span class="pill warn">残' + remaining + '</span></td>';
        } else {
          qHtml += '<td class="qcell"><span class="pill ok">残' + remaining + '</span></td>';
        }
      }
    }

    // Warning pill
    var warnPill = '';
    if (d.warnLevel === '危険') warnPill = '<span class="pill danger">' + escHtml(d.warnLevel) + '</span>';
    else if (d.warnLevel === '警告') warnPill = '<span class="pill warn">' + escHtml(d.warnLevel) + '</span>';
    else if (d.warnLevel === '注意') warnPill = '<span class="pill caution">' + escHtml(d.warnLevel) + '</span>';
    else warnPill = '<span class="pill ok">OK</span>';

    html += '<tr class="' + rowClass + '">'
      + '<td>' + escHtml(d.deptName) + '</td>'
      + '<td>' + escHtml(d.workerName) + '</td>'
      + '<td>' + __escText(d.paidCount || 0) + (d.paidHalfCount > 0 ? ' <span style="font-size:11px;color:var(--muted);">(' + __escText(d.paidHalfCount) + '半)</span>' : '') + '</td>'
      + qHtml
      + '<td>' + __escText(d.substituteCount || 0) + '</td>'
      + '<td>' + __escText(d.specialCount || 0) + '</td>'
      + '<td style="font-weight:700;">' + __escText(d.totalCount || 0) + '</td>'
      + '<td>' + escHtml(d.lastPaidDate || '-') + '</td>'
      + '<td>' + warnPill + '</td>'
      + '</tr>';
  }

  html += '</tbody></table>';
  wrap.innerHTML = html;
}

function escHtml(s) {
  if (!s) return '';
  var d = document.createElement('div');
  d.textContent = String(s);
  return d.innerHTML.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
