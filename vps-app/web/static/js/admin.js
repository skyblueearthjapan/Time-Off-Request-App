/* admin.js — GAS 版 src/admin.html の <script> をそのまま写したもの。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
var __T = JSON.parse(document.getElementById('gas-template-data').textContent);
var APP_URL = __T.APP_URL;
var PORTAL_URL = __T.PORTAL_URL;
var currentTab = 'boss';
var currentData = [];
var isSomuUser = false;

window.addEventListener('DOMContentLoaded', function() {
  if (PORTAL_URL) document.getElementById('portalLink').style.display = '';

  // 年度セレクト初期化
  var fySelect = document.getElementById('fySelect');
  var now = new Date();
  var _m = now.getMonth() + 1, _d = now.getDate();
  var currentFy = (_m < 3 || (_m === 3 && _d < 16)) ? now.getFullYear() - 1 : now.getFullYear();
  for (var y = currentFy; y >= currentFy - 2; y--) {
    var opt = document.createElement('option');
    opt.value = y;
    opt.textContent = y + '年度';
    fySelect.appendChild(opt);
  }

  // 総務判定
  google.script.run.withSuccessHandler(function(isSomu) {
    isSomuUser = isSomu;
    if (isSomu) {
      document.getElementById('somuTab').style.display = '';
      document.getElementById('deptFilterWrap').style.display = '';
      document.getElementById('csvBtn').style.display = '';
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
    }
    loadData();
  }).api_isSomu();
});

function switchTab(tab) {
  currentTab = tab;
  document.querySelectorAll('.tab-btn').forEach(function(b) { b.classList.toggle('active', b.dataset.tab === tab); });
  if (tab === 'somu') {
    document.getElementById('deptFilterWrap').style.display = '';
    document.getElementById('csvBtn').style.display = '';
  } else {
    document.getElementById('deptFilterWrap').style.display = 'none';
    document.getElementById('csvBtn').style.display = 'none';
  }
  loadData();
}

function loadData() {
  var wrap = document.getElementById('tableWrap');
  wrap.innerHTML = '<div class="loading">読み込み中...</div>';
  var fy = document.getElementById('fySelect').value;
  var deptFilter = document.getElementById('deptFilter').value;

  if (currentTab === 'somu' && isSomuUser) {
    google.script.run.withSuccessHandler(function(data) {
      currentData = data;
      renderTable(data);
    }).withFailureHandler(function(err) {
      wrap.innerHTML = '<div class="loading" style="color:var(--danger);">エラー: ' + __escText(err.message) + '</div>';
    }).api_getAdminSomuView(Number(fy), deptFilter);
  } else {
    google.script.run.withSuccessHandler(function(data) {
      currentData = data;
      renderTable(data);
    }).withFailureHandler(function(err) {
      wrap.innerHTML = '<div class="loading" style="color:var(--danger);">エラー: ' + __escText(err.message) + '</div>';
    }).api_getAdminBossView('', Number(fy));
  }
}

function renderTable(data) {
  // KPI更新
  var total = data.length;
  var warnCount = 0;
  var dangerCount = 0;
  var paidSum = 0;

  data.forEach(function(d) {
    paidSum += d.paidCount;
    if (d.warnLevel === '注意' || d.warnLevel === '警告') warnCount++;
    if (d.warnLevel === '危険') { warnCount++; dangerCount++; }
  });

  document.getElementById('kpiTotal').textContent = total;
  document.getElementById('kpiWarn').textContent = warnCount;
  document.getElementById('kpiDanger').textContent = dangerCount;
  document.getElementById('kpiAvg').textContent = total > 0 ? (paidSum / total).toFixed(1) : '-';

  // テーブル
  var wrap = document.getElementById('tableWrap');
  if (!data.length) {
    wrap.innerHTML = '<div class="loading">データなし</div>';
    return;
  }

  var html = '<table><thead><tr><th>部署</th><th>氏名</th><th>有給</th><th>振替</th><th>特別</th><th>合計</th><th>最終有給日</th><th>警告</th></tr></thead><tbody>';

  data.forEach(function(d) {
    var rowClass = '';
    if (d.warnLevel === '危険') rowClass = 'danger-row';
    else if (d.warnLevel === '警告' || d.warnLevel === '注意') rowClass = 'warn-row';

    var warnPill = '';
    if (d.warnLevel === '危険') warnPill = '<span class="pill danger">' + __escText(d.warnLevel) + '</span>';
    else if (d.warnLevel === '警告') warnPill = '<span class="pill warn">' + __escText(d.warnLevel) + '</span>';
    else if (d.warnLevel === '注意') warnPill = '<span class="pill info">' + __escText(d.warnLevel) + '</span>';
    else warnPill = '<span class="pill ok">OK</span>';

    html += '<tr class="' + rowClass + '">'
      + '<td>' + esc(d.deptName) + '</td>'
      + '<td>' + esc(d.workerName) + '</td>'
      + '<td>' + __escText(d.paidCount) + (d.paidHalfCount > 0 ? ' <span style="font-size:11px;color:var(--muted);">(' + __escText(d.paidHalfCount) + '半)</span>' : '') + '</td>'
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

function exportCsv() {
  if (!currentData.length) { alert('データがありません'); return; }
  var lines = ['部署,氏名,有給回数,有給半日回数,振替回数,特別回数,合計回数,最終有給日,警告レベル'];
  currentData.forEach(function(d) {
    lines.push([
      csvEsc(d.deptName), csvEsc(d.workerName),
      d.paidCount, d.paidHalfCount, d.substituteCount, d.specialCount, d.totalCount,
      csvEsc(d.lastPaidDate || ''), csvEsc(d.warnLevel || 'OK'),
    ].join(','));
  });
  var csv = lines.join('\n');
  var bom = '\uFEFF';
  var blob = new Blob([bom + csv], {type:'text/csv;charset=utf-8;'});
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = '有給取得状況_' + document.getElementById('fySelect').value + '.csv';
  a.click();
  URL.revokeObjectURL(url);
}

function csvEsc(s) {
  s = String(s || '');
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g,'""') + '"';
  return s;
}

function esc(s) {
  if (!s) return '';
  var d = document.createElement('div');
  d.textContent = String(s);
  return d.innerHTML.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
