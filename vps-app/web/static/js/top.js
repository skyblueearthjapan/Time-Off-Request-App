/* top.js — GAS 版 src/top.html の <script> をそのまま写したもの。
 * vps-app/web/dev/port-pages.js が作る。手で直さない。 */
var __T = JSON.parse(document.getElementById('gas-template-data').textContent);
var APP_URL = __T.APP_URL;
var PORTAL_URL = __T.PORTAL_URL;
var selectedDeptId = '';
var selectedDeptName = '';
var selectedWorkerId = '';
var selectedWorkerName = '';
var lookupData = null;
var workingDays = [];
var currentStep = 1;
// ====== 過去申請カレンダー用ステート ======
// selectedRetroactiveLeaveDate は 'yyyy-MM-dd' (ISO) 形式で内部保持する
// （表示時のみスラッシュ形式に整形。バックエンドは ISO 形式を要求）
var selectedRetroactiveLeaveDate = '';
var retroCalYear = 0;
var retroCalMonth = 0;                    // 1-12
var retroWorkingDaysCache = {};           // { 'yyyy-MM': ['yyyy-MM-dd', ...] }

// 今年度の開始日（3/16）を動的算出
// 3/16 以降なら当年の 3/16、それ以前なら前年の 3/16
function getFyStart_() {
  // JST基準で年度開始日を算出（ブラウザTZに依存しない）
  var t = new Date(new Date().toLocaleString('en-US', {timeZone: 'Asia/Tokyo'}));
  var beforeMid = (t.getMonth() < 2) || (t.getMonth() === 2 && t.getDate() < 16);
  return new Date(beforeMid ? t.getFullYear() - 1 : t.getFullYear(), 2, 16);
}
var RETRO_MIN_DATE = getFyStart_();

// 'yyyy-MM-dd' → 'yyyy/MM/dd'（表示用）
function isoToSlash_(iso) {
  return iso ? iso.replace(/-/g, '/') : '';
}

// ====== 初期化 ======
window.addEventListener('DOMContentLoaded', function() {
  if (PORTAL_URL) document.getElementById('portalLink').style.display = '';

  // 部署管理リンク表示判定（管理者 or 部署承認者）
  google.script.run.withSuccessHandler(function(ok) {
    if (ok) document.getElementById('navDept').style.display = '';
  }).api_isDeptApprover();
  google.script.run.withSuccessHandler(function(ok) {
    if (ok) document.getElementById('navDept').style.display = '';
  }).api_isAdmin();

  // 総務管理リンク表示判定
  google.script.run.withSuccessHandler(function(ok) {
    if (ok) document.getElementById('navSomu').style.display = '';
  }).api_isSomu();

  // 部署一覧ロード
  google.script.run.withSuccessHandler(function(depts) {
    var sel = document.getElementById('deptSelect');
    sel.innerHTML = '<option value="">-- 部署を選択 --</option>';
    depts.forEach(function(d) {
      var opt = document.createElement('option');
      opt.value = d.deptId;
      opt.textContent = d.deptName;
      opt.dataset.name = d.deptName;
      sel.appendChild(opt);
    });

    // ログインユーザーの部署を自動選択
    google.script.run.withSuccessHandler(function(info) {
      if (info) {
        sel.value = info.deptId;
        selectedDeptId = info.deptId;
        selectedDeptName = info.deptName;
        loadWorkers(info.deptId, info.workerId);
      } else {
        // 自動検出できなかった場合も予定一覧を表示
        loadLeaveTable();
      }
    }).api_getWorkerInfo();
  }).withFailureHandler(function(err) {
    console.error('部署一覧取得エラー:', err);
    document.getElementById('deptSelect').innerHTML = '<option value="">-- 取得エラー --</option>';
    loadLeaveTable();
  }).api_getDeptList();

  // 承認待ち一覧ロード
  loadPendingRequests();

  // LOOKUPデータロード
  google.script.run.withSuccessHandler(function(data) {
    lookupData = data;
  }).api_getLookupData();

  // 勤務日一覧ロード
  google.script.run.withSuccessHandler(function(days) {
    workingDays = days || [];
    var sel = document.getElementById('leaveDate');
    sel.innerHTML = '<option value="">-- 休暇日を選択 --</option>';
    var dayNames = ['日','月','火','水','木','金','土'];
    workingDays.forEach(function(d) {
      var opt = document.createElement('option');
      opt.value = d;
      var dt = new Date(d + 'T00:00:00');
      var dow = dayNames[dt.getDay()];
      opt.textContent = d + '（' + dow + '）';
      if (dt.getDay() === 6) opt.style.color = '#0ea5e9'; // 土曜は青
      sel.appendChild(opt);
    });
  }).withFailureHandler(function(err) {
    console.error('勤務日取得エラー:', err);
    var sel = document.getElementById('leaveDate');
    sel.innerHTML = '<option value="">-- 取得エラー --</option>';
  }).api_getWorkingDays();

  // 通常プルダウンが選ばれたら過去申請カレンダー選択をクリア（排他）
  var leaveDateEl = document.getElementById('leaveDate');
  if (leaveDateEl) {
    leaveDateEl.addEventListener('change', function() {
      if (this.value && selectedRetroactiveLeaveDate) {
        selectedRetroactiveLeaveDate = '';
        var info = document.getElementById('retroSelectedInfo');
        if (info) info.textContent = '選択中: なし';
        var cells = document.querySelectorAll('#retroCalGrid .retro-cal-cell.selected');
        for (var i = 0; i < cells.length; i++) cells[i].classList.remove('selected');
      }
      // 休暇日が変わったら振替元プルダウンを月次期間で再フィルタ
      renderSubstitutePulldown();
      validateSubstituteSelection();
    });
  }

  // 初期ロード: 土日・祝日候補を取得（作業員選択後に使用済み除外で再ロード）
  loadHolidayDays('');
});

// ====== 月次期間ヘルパ（前月16日〜当月15日） ======
// dateStr: 'yyyy-MM-dd' → {start:'yyyy-MM-dd', end:'yyyy-MM-dd'}
function computeMonthlyPeriodJS(dateStr) {
  if (!dateStr) return null;
  var parts = dateStr.split('-');
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1; // 0-11
  var d = parseInt(parts[2], 10);
  var start, end;
  if (d >= 16) {
    start = new Date(y, m, 16);
    end = new Date(y, m + 1, 15);
  } else {
    start = new Date(y, m - 1, 16);
    end = new Date(y, m, 15);
  }
  function fmt(dt) {
    var yy = dt.getFullYear();
    var mm = ('0' + (dt.getMonth() + 1)).slice(-2);
    var dd = ('0' + dt.getDate()).slice(-2);
    return yy + '-' + mm + '-' + dd;
  }
  return { start: fmt(start), end: fmt(end), startDt: start, endDt: end };
}

// 現在選択中の休暇日を取得（通常/過去申請の両経路対応）
function getCurrentLeaveDate() {
  if (selectedRetroactiveLeaveDate) return selectedRetroactiveLeaveDate;
  var el = document.getElementById('leaveDate');
  return el ? el.value : '';
}

// 振替元選択時のリアルタイム月次期間バリデーション
function validateSubstituteSelection() {
  var sel = document.getElementById('substituteDate');
  var errBox = document.getElementById('substituteDateError');
  var okBox = document.getElementById('substituteDateOk');
  var nextBtn = document.getElementById('substituteNextBtn');
  if (!sel || !errBox || !okBox || !nextBtn) return;

  errBox.style.display = 'none';
  okBox.style.display = 'none';
  nextBtn.disabled = false;

  var subDate = sel.value;
  var leaveDate = getCurrentLeaveDate();
  if (!subDate) return;

  if (!leaveDate) {
    errBox.textContent = '先に休暇日を選択してください。';
    errBox.style.display = '';
    nextBtn.disabled = true;
    return;
  }

  var lp = computeMonthlyPeriodJS(leaveDate);
  var sp = computeMonthlyPeriodJS(subDate);
  if (!lp || !sp) return;

  if (lp.start !== sp.start) {
    errBox.innerHTML =
      '<strong>⚠ 月次期間違反</strong><br>' +
      '休暇日 <strong>' + __escText(leaveDate) + '</strong> は期間「' + __escText(lp.start) + '〜' + __escText(lp.end) + '」に属します。<br>' +
      '振替元 <strong>' + __escText(subDate) + '</strong> は期間「' + __escText(sp.start) + '〜' + __escText(sp.end) + '」に属します。<br>' +
      '振替休暇は同一月次期間（前月16日〜当月15日）で消化する必要があります。';
    errBox.style.display = '';
    nextBtn.disabled = true;
  } else {
    okBox.innerHTML = '✅ 月次期間OK（' + __escText(lp.start) + '〜' + __escText(lp.end) + '）';
    okBox.style.display = '';
    nextBtn.disabled = false;
  }
}

// 振替元プルダウン用の全候補キャッシュ（休暇日が決まり次第フィルタリング）
var ALL_HOLIDAY_CANDIDATES = [];

// 振替元プルダウンを再描画（休暇日の月次期間内の候補のみ表示）
function renderSubstitutePulldown() {
  var sel = document.getElementById('substituteDate');
  if (!sel) return;
  var prevValue = sel.value;
  var leaveDate = getCurrentLeaveDate();
  var period = leaveDate ? computeMonthlyPeriodJS(leaveDate) : null;

  var dayNames = ['日','月','火','水','木','金','土'];
  if (!leaveDate) {
    sel.innerHTML = '<option value="">-- 先に休暇日を選択してください --</option>';
    return;
  }

  var filtered = ALL_HOLIDAY_CANDIDATES.filter(function(h) {
    return h.date >= period.start && h.date <= period.end;
  });

  if (filtered.length === 0) {
    sel.innerHTML = '<option value="">-- 期間内(' + __escText(period.start) + '〜' + __escText(period.end) + ')に候補なし --</option>';
    return;
  }

  sel.innerHTML = '<option value="">-- 休日出勤日を選択（' + __escText(period.start) + '〜' + __escText(period.end) + '） --</option>';
  filtered.forEach(function(h) {
    var opt = document.createElement('option');
    opt.value = h.date;
    var dt = new Date(h.date + 'T00:00:00');
    var dow = dayNames[dt.getDay()];
    opt.textContent = h.date + '（' + dow + '・' + h.label + '）';
    if (dt.getDay() === 0) opt.style.color = '#ff5a5f';
    if (dt.getDay() === 6) opt.style.color = '#0ea5e9';
    sel.appendChild(opt);
  });

  // 前の選択値を復元（範囲内なら）
  if (prevValue && filtered.some(function(h) { return h.date === prevValue; })) {
    sel.value = prevValue;
  }
}

// 振替元休日出勤日プルダウンをロード（土日・祝日を候補に、使用済みは除外）
function loadHolidayDays(workerId) {
  var sel = document.getElementById('substituteDate');
  if (!sel) return;
  sel.innerHTML = '<option value="">読込中...</option>';
  google.script.run.withSuccessHandler(function(holidays) {
    ALL_HOLIDAY_CANDIDATES = holidays || [];
    renderSubstitutePulldown();
  }).withFailureHandler(function(err) {
    console.error('休日一覧取得エラー:', err);
    sel.innerHTML = '<option value="">-- 取得エラー --</option>';
  }).api_getHolidayDays(workerId || '');
}

// ====== 部署選択 ======
document.getElementById('deptSelect').addEventListener('change', function() {
  selectedDeptId = this.value;
  selectedDeptName = this.options[this.selectedIndex].dataset.name || '';
  if (selectedDeptId) {
    loadWorkers(selectedDeptId);
  } else {
    document.getElementById('workerSelect').innerHTML = '<option value="">-- 部署を先に選択 --</option>';
    document.getElementById('applyBtn').disabled = true;
  }
  loadLeaveTable();
});

function loadWorkers(deptId, autoSelectId) {
  var sel = document.getElementById('workerSelect');
  sel.innerHTML = '<option value="">読込中...</option>';
  google.script.run.withSuccessHandler(function(workers) {
    sel.innerHTML = '<option value="">-- 作業員を選択 --</option>';
    workers.forEach(function(w) {
      var opt = document.createElement('option');
      opt.value = w.workerId;
      opt.textContent = w.workerId + ' ' + w.workerName;
      opt.dataset.name = w.workerName;
      sel.appendChild(opt);
    });
    if (autoSelectId) {
      sel.value = autoSelectId;
      selectedWorkerId = autoSelectId;
      var selOpt = sel.options[sel.selectedIndex];
      selectedWorkerName = selOpt ? selOpt.dataset.name || '' : '';
      document.getElementById('applyBtn').disabled = false;
      loadHolidayDays(autoSelectId);
    } else {
      loadHolidayDays('');
    }
    loadLeaveTable();
  }).api_getWorkersByDept(deptId);
}

// ====== 作業員選択 ======
document.getElementById('workerSelect').addEventListener('change', function() {
  selectedWorkerId = this.value;
  selectedWorkerName = this.options[this.selectedIndex].dataset.name || '';
  document.getElementById('applyBtn').disabled = !selectedWorkerId;
  loadHolidayDays(selectedWorkerId);
  loadLeaveTable();
});

// ====== 休暇予定一覧 ======
function loadLeaveTable() {
  var wrap = document.getElementById('leaveTableWrap');
  wrap.innerHTML = '<div class="loading">読み込み中...</div>';
  google.script.run.withSuccessHandler(function(items) {
    if (!items) items = [];

    // 本日の日付文字列（yyyy/MM/dd）
    var now = new Date();
    var todayStr = now.getFullYear() + '/' + ('0'+(now.getMonth()+1)).slice(-2) + '/' + ('0'+now.getDate()).slice(-2);

    // 本日と今後に分離
    var todayItems = [];
    var futureItems = [];
    items.forEach(function(it) {
      if (it.leaveDate === todayStr) todayItems.push(it);
      else futureItems.push(it);
    });

    // 本日セクション
    var todayCard = document.getElementById('todaySectionCard');
    var todayWrap = document.getElementById('todayLeaveWrap');
    if (todayItems.length > 0) {
      todayCard.style.display = '';
      document.getElementById('todayLeaveCount').textContent = todayItems.length + '名';
      var tHtml = '<table><thead><tr><th>部署</th><th>氏名</th><th>区分</th><th>種類</th><th>状態</th></tr></thead><tbody>';
      todayItems.forEach(function(it) {
        var statusClass = it.status === '承認済' ? 'ok' : 'info';
        tHtml += '<tr style="font-weight:700;">'
          + '<td>' + __escText(it.deptName) + '</td>'
          + '<td>' + __escText(it.workerName) + '</td>'
          + '<td>' + __escText(it.leaveKubun) + (it.halfDayType ? '(' + __escText(it.halfDayType) + ')' : '') + '</td>'
          + '<td>' + __escText(it.leaveType) + '</td>'
          + '<td><span class="pill ' + statusClass + '">' + __escText(it.status) + '</span></td>'
          + '</tr>';
      });
      tHtml += '</tbody></table>';
      todayWrap.innerHTML = tHtml;
    } else {
      todayCard.style.display = 'none';
    }

    // 今後セクション
    document.getElementById('leaveCount').textContent = futureItems.length + '件';
    if (!futureItems.length) {
      wrap.innerHTML = '<div class="loading">予定なし</div>';
      return;
    }
    // 編集モーダル用にデータをキャッシュ
    window.LEAVE_ITEMS = items;

    var html = '<table><thead><tr><th>日付</th><th>部署</th><th>氏名</th><th>区分</th><th>種類</th><th>状態</th><th>操作</th></tr></thead><tbody>';
    futureItems.forEach(function(it) {
      var statusClass = it.status === '承認済' ? 'ok' : 'info';
      var actionBtns = '';
      if (it.status !== '取消' && it.status !== '却下') {
        actionBtns = '<button class="ghost btn-sm" data-on-click="openEditLeaveModal" data-args="' + __vpsArgs(it.reqId) + '" style="font-size:11px;">編集</button>'
          + ' <button class="ghost btn-sm" data-on-click="confirmDeleteLeave" data-args="' + __vpsArgs(it.reqId) + '" style="font-size:11px;color:var(--danger);">削除</button>';
      }
      if (it.status === '申請中') {
        actionBtns += ' <a href="' + APP_URL + '?page=preview&reqId=' + encodeURIComponent(it.reqId) + '" target="_top" class="btn-sm btn-blue" style="color:#fff;text-decoration:none;padding:4px 10px;border-radius:8px;">承認</a>';
      }
      html += '<tr>'
        + '<td>' + __escText(it.leaveDate) + '</td>'
        + '<td>' + __escText(it.deptName) + '</td>'
        + '<td>' + __escText(it.workerName) + '</td>'
        + '<td>' + __escText(it.leaveKubun) + (it.halfDayType ? '(' + __escText(it.halfDayType) + ')' : '') + '</td>'
        + '<td>' + __escText(it.leaveType) + '</td>'
        + '<td><span class="pill ' + statusClass + '">' + __escText(it.status) + '</span></td>'
        + '<td>' + actionBtns + '</td>'
        + '</tr>';
    });
    html += '</tbody></table>';
    wrap.innerHTML = html;
  }).withFailureHandler(function(err) {
    console.error('予定一覧取得エラー:', err);
    wrap.innerHTML = '<div class="loading">予定なし</div>';
    document.getElementById('leaveCount').textContent = '0件';
  }).api_getUpcomingLeaves(selectedDeptId, selectedWorkerId);
}

// ====== モーダル制御 ======
function openModal() {
  if (!selectedWorkerId) { alert('作業員を選択してください'); return; }
  resetModal();
  document.getElementById('modalOverlay').classList.add('active');
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('active');
}

function resetModal() {
  currentStep = 1;
  document.querySelectorAll('.step').forEach(function(s) { s.classList.remove('active'); });
  document.querySelector('.step[data-step="1"]').classList.add('active');
  document.querySelectorAll('input[type="radio"]').forEach(function(r) { r.checked = false; });
  document.getElementById('leaveDate').selectedIndex = 0;
  document.getElementById('substituteDate').selectedIndex = 0;
  document.getElementById('specialReason').value = '';
  document.getElementById('paidDetail').value = '';
  // 振替元バリデーション表示をリセット
  var errBox = document.getElementById('substituteDateError');
  var okBox = document.getElementById('substituteDateOk');
  var nextBtn = document.getElementById('substituteNextBtn');
  if (errBox) errBox.style.display = 'none';
  if (okBox) okBox.style.display = 'none';
  if (nextBtn) nextBtn.disabled = false;
  // 過去申請カレンダーもリセット
  selectedRetroactiveLeaveDate = '';
  document.getElementById('retroSelectedInfo').textContent = '選択中: なし';
  var sec = document.getElementById('retroSection');
  if (sec) sec.classList.remove('open');
  var icon = document.getElementById('retroToggleIcon');
  if (icon) icon.innerHTML = '&#9654;';
  // 今日基準月で初期化
  var now = new Date();
  retroCalYear = now.getFullYear();
  retroCalMonth = now.getMonth() + 1;
  renderRetroCalendar();
}

// ====== 過去申請カレンダー ======
function toggleRetroSection() {
  var sec = document.getElementById('retroSection');
  var icon = document.getElementById('retroToggleIcon');
  if (sec.classList.contains('open')) {
    sec.classList.remove('open');
    icon.innerHTML = '&#9654;';
  } else {
    sec.classList.add('open');
    icon.innerHTML = '&#9660;';
    // 開いた時に最新で再描画
    if (!retroCalYear) {
      var now = new Date();
      retroCalYear = now.getFullYear();
      retroCalMonth = now.getMonth() + 1;
    }
    renderRetroCalendar();
  }
}

function changeRetroMonth(delta) {
  var m = retroCalMonth + delta;
  var y = retroCalYear;
  if (m < 1) { m = 12; y--; }
  if (m > 12) { m = 1; y++; }
  // 範囲外のチェック
  var firstOfMonth = new Date(y, m - 1, 1);
  var lastOfMonth = new Date(y, m, 0);
  var today = new Date();
  today.setHours(0,0,0,0);
  // 月全体が許容範囲外なら切替不可
  if (lastOfMonth < RETRO_MIN_DATE) return;
  if (firstOfMonth > today) return;
  retroCalYear = y;
  retroCalMonth = m;
  renderRetroCalendar();
}

function renderRetroCalendar() {
  var title = document.getElementById('retroCalTitle');
  var grid = document.getElementById('retroCalGrid');
  if (!title || !grid) return;
  title.textContent = retroCalYear + '年 ' + retroCalMonth + '月';

  // 前後ボタンの活性制御
  var prevFirst = new Date(retroCalYear, retroCalMonth - 2, 1);
  var prevLast = new Date(retroCalYear, retroCalMonth - 1, 0);
  document.getElementById('retroPrevBtn').disabled = (prevLast < RETRO_MIN_DATE);
  var nextFirst = new Date(retroCalYear, retroCalMonth, 1);
  var today = new Date(); today.setHours(0,0,0,0);
  document.getElementById('retroNextBtn').disabled = (nextFirst > today);

  // ヘッダー（曜日）
  var html = '';
  var dows = ['日','月','火','水','木','金','土'];
  for (var i = 0; i < 7; i++) {
    var cls = 'dow';
    if (i === 0) cls += ' sun';
    if (i === 6) cls += ' sat';
    html += '<div class="' + cls + '">' + dows[i] + '</div>';
  }
  grid.innerHTML = html + '<div class="retro-cal-loading" style="grid-column:span 7;text-align:center;padding:14px;color:var(--muted);font-size:12px;">読込中...</div>';

  // バックエンドから勤務日取得（'yyyy-MM' 文字列で渡す）
  var ym = retroCalYear + '-' + ('0' + retroCalMonth).slice(-2);
  if (retroWorkingDaysCache[ym]) {
    drawRetroCalendarCells(retroWorkingDaysCache[ym]);
    return;
  }
  google.script.run
    .withSuccessHandler(function(days) {
      // days = [{date:'yyyy-MM-dd', dow, label}, ...]
      var list = (days || []).map(function(d) { return d.date; });
      retroWorkingDaysCache[ym] = list;
      // 描画前に月が切り替わっていたら無視
      var currentYm = retroCalYear + '-' + ('0' + retroCalMonth).slice(-2);
      if (currentYm !== ym) return;
      drawRetroCalendarCells(list);
    })
    .withFailureHandler(function(err) {
      console.error('過去勤務日取得エラー:', err);
      drawRetroCalendarCells([]);
    })
    .api_getPastWorkingDays(ym);
}

function drawRetroCalendarCells(workableList) {
  var grid = document.getElementById('retroCalGrid');
  var html = '';
  var dows = ['日','月','火','水','木','金','土'];
  for (var i = 0; i < 7; i++) {
    var cls = 'dow';
    if (i === 0) cls += ' sun';
    if (i === 6) cls += ' sat';
    html += '<div class="' + cls + '">' + dows[i] + '</div>';
  }

  var firstDay = new Date(retroCalYear, retroCalMonth - 1, 1);
  var startDow = firstDay.getDay();
  var lastDate = new Date(retroCalYear, retroCalMonth, 0).getDate();
  var today = new Date(); today.setHours(0,0,0,0);

  // 先頭の空セル
  for (var s = 0; s < startDow; s++) {
    html += '<div class="retro-cal-cell empty"></div>';
  }
  for (var d = 1; d <= lastDate; d++) {
    var dt = new Date(retroCalYear, retroCalMonth - 1, d);
    var iso = retroCalYear + '-' + ('0' + retroCalMonth).slice(-2) + '-' + ('0' + d).slice(-2);
    var slash = isoToSlash_(iso);
    var inRange = (dt >= RETRO_MIN_DATE && dt <= today);
    var isWorkable = inRange && workableList.indexOf(iso) >= 0;
    var cls = 'retro-cal-cell';
    if (!isWorkable) cls += ' disabled';
    else cls += ' workable';
    // 内部保持は ISO 形式のため ISO で比較
    if (selectedRetroactiveLeaveDate === iso) cls += ' selected';
    if (isWorkable) {
      // data-date には ISO 形式を格納し、onclick にも ISO を渡す（表示はスラッシュ）
      html += '<div class="' + cls + '" data-date="' + iso + '" data-on-click="selectRetroDate" data-args="' + __vpsArgs(iso) + '">' + d + '</div>';
    } else {
      html += '<div class="' + cls + '">' + d + '</div>';
    }
  }
  grid.innerHTML = html;
}

function selectRetroDate(isoDate) {
  // 内部保持は ISO 形式（'yyyy-MM-dd'）、表示はスラッシュ形式
  selectedRetroactiveLeaveDate = isoDate;
  var slashDate = isoToSlash_(isoDate);
  // 通常プルダウンの選択をクリア（排他）
  var ld = document.getElementById('leaveDate');
  if (ld) ld.selectedIndex = 0;
  document.getElementById('retroSelectedInfo').textContent = '選択中: ' + slashDate + '（過去申請）';
  // セルのselected再描画
  var cells = document.querySelectorAll('#retroCalGrid .retro-cal-cell');
  for (var i = 0; i < cells.length; i++) {
    cells[i].classList.remove('selected');
    if (cells[i].dataset && cells[i].dataset.date === isoDate) {
      cells[i].classList.add('selected');
    }
  }
  // 過去申請日が変わったら振替元プルダウンを再フィルタ＋再検証
  renderSubstitutePulldown();
  validateSubstituteSelection();
}

function showStep(n) {
  document.querySelectorAll('.step').forEach(function(s) { s.classList.remove('active'); });
  var target = document.querySelector('.step[data-step="' + n + '"]');
  if (target) target.classList.add('active');
  currentStep = n;
}

function nextStep(n) {
  // バリデーション
  if (currentStep === 1) {
    var kubun = document.querySelector('input[name="leaveKubun"]:checked');
    if (!kubun) { alert('休暇区分を選択してください'); return; }
  }
  if (currentStep === 2) {
    var hasNormal = !!document.getElementById('leaveDate').value;
    var hasRetro = !!selectedRetroactiveLeaveDate;
    if (!hasNormal && !hasRetro) { alert('休暇日を選択してください（過去申請カレンダーから選択することもできます）'); return; }
    var kubunVal = document.querySelector('input[name="leaveKubun"]:checked').value;
    if (kubunVal === '全日休') {
      // 半日区分スキップ → Step4へ
      buildLeaveTypeOptions(kubunVal);
      showStep(4);
      return;
    }
  }
  if (currentStep === 3) {
    var halfDay = document.querySelector('input[name="halfDayType"]:checked');
    if (!halfDay) { alert('午前・午後を選択してください'); return; }
    var kubunVal2 = document.querySelector('input[name="leaveKubun"]:checked').value;
    buildLeaveTypeOptions(kubunVal2);
  }
  if (currentStep === 9) return; // 確認画面の先はない

  // Step9: 確認画面を構築
  if (n === 9) {
    buildConfirmation();
  }

  showStep(n);
}

function prevStep(n) {
  showStep(n);
}

function prevStepFromType() {
  var kubun = document.querySelector('input[name="leaveKubun"]:checked').value;
  if (kubun === '全日休') {
    showStep(2);
  } else {
    showStep(3);
  }
}

function nextStepFromType() {
  var type = document.querySelector('input[name="leaveType"]:checked');
  if (!type) { alert('休暇種類を選択してください'); return; }
  var typeVal = type.value;
  if (typeVal === '振替休暇') { showStep(5); renderSubstitutePulldown(); validateSubstituteSelection(); }
  else if (typeVal === '特別休暇') { showStep(6); }
  else if (typeVal === '有給消化（私用）') { showStep(7); }
  else { showStep(9); buildConfirmation(); }
}

function prevStepFromConfirm() {
  var type = document.querySelector('input[name="leaveType"]:checked');
  if (!type) { showStep(4); return; }
  var typeVal = type.value;
  if (typeVal === '振替休暇') { showStep(5); renderSubstitutePulldown(); validateSubstituteSelection(); }
  else if (typeVal === '特別休暇') { showStep(6); }
  else if (typeVal === '有給消化（私用）') { showStep(7); }
  else { showStep(4); }
}

function buildLeaveTypeOptions(kubun) {
  var group = document.getElementById('leaveTypeGroup');
  group.innerHTML = '';
  var types = kubun === '全日休'
    ? ['有給消化（私用）', '振替休暇', '特別休暇']
    : ['有給消化（私用）', '特別休暇'];

  types.forEach(function(t) {
    var lbl = document.createElement('label');
    lbl.innerHTML = '<input type="radio" name="leaveType" value="' + t + '"><span>' + t + '</span>';
    group.appendChild(lbl);
  });
}

function buildConfirmation() {
  var kubun = document.querySelector('input[name="leaveKubun"]:checked');
  var halfDay = document.querySelector('input[name="halfDayType"]:checked');
  var leaveType = document.querySelector('input[name="leaveType"]:checked');

  document.getElementById('cfmDept').textContent = selectedDeptName;
  document.getElementById('cfmWorker').textContent = selectedWorkerId + ' ' + selectedWorkerName;
  document.getElementById('cfmKubun').textContent = kubun ? kubun.value : '';
  document.getElementById('cfmHalfDay').textContent = halfDay ? halfDay.value : '-';
  // 内部保持は ISO 形式のため、確認画面表示用にスラッシュ形式へ整形
  var normalVal = document.getElementById('leaveDate').value;
  var displayDate = selectedRetroactiveLeaveDate
    ? isoToSlash_(selectedRetroactiveLeaveDate) + '（過去申請）'
    : isoToSlash_(normalVal) || normalVal;
  document.getElementById('cfmDate').textContent = displayDate;
  document.getElementById('cfmType').textContent = leaveType ? leaveType.value : '';
  var subSel = document.getElementById('substituteDate');
  var subText = '-';
  if (subSel && subSel.value) {
    var selectedOpt = subSel.options[subSel.selectedIndex];
    subText = selectedOpt ? selectedOpt.textContent : subSel.value;
  }
  document.getElementById('cfmSubDate').textContent = subText;

  // 理由を統合表示
  var reason = '-';
  var typeVal = leaveType ? leaveType.value : '';
  if (typeVal === '特別休暇') {
    reason = document.getElementById('specialReason').value || '-';
  } else if (typeVal === '有給消化（私用）') {
    reason = document.getElementById('paidDetail').value || '私用のため';
  } else if (typeVal === '振替休暇') {
    reason = '振替休暇';
  }
  document.getElementById('cfmReason').textContent = reason;
}

function submitRequest() {
  var btn = document.getElementById('submitBtn');

  var kubun = document.querySelector('input[name="leaveKubun"]:checked');
  var halfDay = document.querySelector('input[name="halfDayType"]:checked');
  var leaveType = document.querySelector('input[name="leaveType"]:checked');

  // 過去申請かどうか判定
  var isRetro = !!selectedRetroactiveLeaveDate;
  var effectiveLeaveDate = isRetro
    ? selectedRetroactiveLeaveDate
    : document.getElementById('leaveDate').value;

  if (isRetro) {
    // 表示はスラッシュ形式、送信はそのまま ISO 形式
    if (!confirm('過去申請として送信します。\n対象日: ' + isoToSlash_(selectedRetroactiveLeaveDate) + '\n\n承認者および総務にすでにご連絡済みですか？\nよろしければ「OK」を押してください。')) {
      return;
    }
  }

  btn.disabled = true;
  btn.textContent = '送信中...';

  var data = {
    deptId: selectedDeptId,
    deptName: selectedDeptName,
    workerId: selectedWorkerId,
    workerName: selectedWorkerName,
    leaveKubun: kubun ? kubun.value : '',
    halfDayType: halfDay ? halfDay.value : '',
    leaveDate: effectiveLeaveDate,
    leaveType: leaveType ? leaveType.value : '',
    substituteDate: document.getElementById('substituteDate').value,
    specialReason: document.getElementById('specialReason').value,
    paidDetail: document.getElementById('paidDetail').value,
    isRetroactive: isRetro,
  };

  google.script.run
    .withSuccessHandler(function(result) {
      if (result && result.ok) {
        var previewUrl = APP_URL + '?page=preview&reqId=' + encodeURIComponent(result.reqId);
        // モーダル内に申請完了メッセージとプレビューリンクを表示
        document.querySelector('.step.active').innerHTML =
          '<div style="text-align:center;padding:20px 0;">'
          + '<div style="font-size:40px;margin-bottom:12px;">&#10004;</div>'
          + '<div style="font-weight:900;font-size:18px;margin-bottom:8px;">申請が完了しました</div>'
          + '<div style="color:var(--muted);font-size:13px;margin-bottom:20px;">REQ_ID: ' + __escText(result.reqId) + '</div>'
          + '<div style="display:flex;gap:12px;justify-content:center;flex-wrap:wrap;">'
          + '<a href="' + previewUrl + '" target="_top" style="display:inline-block;padding:14px 32px;border-radius:12px;background:linear-gradient(135deg,#0ea5e9,#38bdf8);color:#fff;font-weight:800;font-size:15px;text-decoration:none;box-shadow:0 10px 22px rgba(14,165,233,.20);">プレビュー画面へ（承認依頼）</a>'
          + '<a href="' + APP_URL + '?page=top" target="_top" style="display:inline-block;padding:14px 32px;border-radius:12px;background:#fff;color:var(--text);font-weight:800;font-size:15px;text-decoration:none;border:1px solid #d7e7f5;">TOPに戻る</a>'
          + '</div>'
          + '</div>';
      } else {
        var msg = (result && result.error) ? result.error : '申請に失敗しました。';
        alert('申請に失敗しました：\n' + msg);
        btn.disabled = false;
        btn.textContent = '申請する';
      }
    })
    .withFailureHandler(function(err) {
      alert('エラー: ' + err.message);
      btn.disabled = false;
      btn.textContent = '申請する';
    })
    .api_submitLeaveRequest(data);
}

// ====== 承認待ち一覧 ======
function loadPendingRequests() {
  var wrap = document.getElementById('pendingWrap');
  wrap.innerHTML = '<div class="loading">読み込み中...</div>';
  google.script.run
    .withSuccessHandler(function(items) {
      if (!items || !items.length) {
        document.getElementById('pendingCount').textContent = '0件';
        wrap.innerHTML = '<div class="loading">承認待ちの申請はありません</div>';
        return;
      }
      document.getElementById('pendingCount').textContent = items.length + '件';
      var dayNames = ['日','月','火','水','木','金','土'];
      var html = '<table><thead><tr><th>休暇日</th><th>部署</th><th>氏名</th><th>区分</th><th>種類</th><th>申請日時</th><th></th></tr></thead><tbody>';
      items.forEach(function(it) {
        // 曜日付き
        var dateDisp = it.leaveDate;
        if (it.leaveDate) {
          var parts = it.leaveDate.split('/');
          if (parts.length === 3) {
            var dt = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
            dateDisp = it.leaveDate + '(' + dayNames[dt.getDay()] + ')';
          }
        }
        // 過去申請バッジ
        var retroBadge = '';
        var rowStyle = '';
        if (it.isRetroactive === true) {
          retroBadge = ' <span style="display:inline-block;padding:2px 8px;border-radius:999px;background:#fee2e2;color:#b91c1c;border:1px solid #fca5a5;font-size:11px;font-weight:800;">&#128281; 過去申請</span>'
            + (typeof it.daysAgo === 'number' ? ' <span style="font-size:11px;color:var(--muted);">(' + __escText(it.daysAgo) + '日前)</span>' : '');
          rowStyle = ' style="background:#fffbeb;"';
        }
        html += '<tr' + rowStyle + '>'
          + '<td>' + __escText(dateDisp) + retroBadge + '</td>'
          + '<td>' + __escText(it.deptName) + '</td>'
          + '<td style="font-weight:700;">' + __escText(it.workerName) + '</td>'
          + '<td>' + __escText(it.leaveKubun) + (it.halfDayType ? '(' + __escText(it.halfDayType) + ')' : '') + '</td>'
          + '<td>' + __escText(it.leaveType) + '</td>'
          + '<td style="font-size:12px;color:var(--muted);">' + __escText(it.submittedAt || '') + '</td>'
          + '<td><a href="' + APP_URL + '?page=preview&reqId=' + encodeURIComponent(it.reqId) + '" target="_top" class="btn-sm btn-blue" style="color:#fff;text-decoration:none;padding:6px 14px;border-radius:8px;display:inline-block;">承認プレビュー</a></td>'
          + '</tr>';
      });
      html += '</tbody></table>';
      wrap.innerHTML = html;
    })
    .withFailureHandler(function(err) {
      console.error('承認待ち一覧取得エラー:', err);
      wrap.innerHTML = '<div class="loading">取得エラー</div>';
      document.getElementById('pendingCount').textContent = '';
    })
    .api_getAllPendingRequests();
}

// ====== 編集モーダル ======
var EDIT_LEAVE_CTX = {};

// 休暇区分ラジオ変更時に半日区分の表示切り替え
document.querySelectorAll('input[name="editLeaveKubun"]').forEach(function(r) {
  r.addEventListener('change', function() {
    document.getElementById('editHalfDayWrap').style.display =
      this.value === '半日休' ? '' : 'none';
  });
});

function openEditLeaveModal(reqId) {
  var item = null;
  (window.LEAVE_ITEMS || []).forEach(function(it) { if (it.reqId === reqId) item = it; });
  if (!item) { alert('申請データが見つかりません。画面を更新してください。'); return; }

  EDIT_LEAVE_CTX = { reqId: reqId, item: item };

  document.getElementById('editInfo').textContent = item.deptName + ' ／ ' + item.workerName;

  // 振替休暇の場合は月次制約の注記を表示
  var notice = document.getElementById('editSubstituteNotice');
  if (notice) notice.style.display = (item.leaveType === '振替休暇') ? '' : 'none';

  // 休暇日（yyyy/MM/dd → yyyy-MM-dd に変換）
  var dateVal = (item.leaveDate || '').replace(/\//g, '-');
  document.getElementById('editLeaveDate').value = dateVal;

  // 休暇区分
  document.querySelectorAll('input[name="editLeaveKubun"]').forEach(function(r) {
    r.checked = (r.value === item.leaveKubun);
  });

  // 半日区分
  document.getElementById('editHalfDayWrap').style.display =
    item.leaveKubun === '半日休' ? '' : 'none';
  document.querySelectorAll('input[name="editHalfDayType"]').forEach(function(r) {
    r.checked = (r.value === item.halfDayType);
  });

  var btn = document.getElementById('editSaveBtn');
  btn.disabled = false;
  btn.textContent = '保存';

  document.getElementById('editOverlay').classList.add('active');
}

function closeEditLeaveModal() {
  document.getElementById('editOverlay').classList.remove('active');
}

function saveEditLeave() {
  var leaveDate = document.getElementById('editLeaveDate').value;
  var kubunRadio = document.querySelector('input[name="editLeaveKubun"]:checked');
  var halfDayRadio = document.querySelector('input[name="editHalfDayType"]:checked');

  if (!leaveDate) { alert('休暇日を入力してください。'); return; }
  if (!kubunRadio) { alert('休暇区分を選択してください。'); return; }
  if (kubunRadio.value === '半日休' && !halfDayRadio) { alert('半日区分を選択してください。'); return; }

  var patch = {
    leaveDate: leaveDate,
    leaveKubun: kubunRadio.value,
    halfDayType: kubunRadio.value === '半日休' ? halfDayRadio.value : ''
  };

  var msg = '申請内容を変更しますか？\n休暇日: ' + leaveDate + '\n区分: ' + patch.leaveKubun;
  if (patch.halfDayType) msg += '(' + patch.halfDayType + ')';
  if (EDIT_LEAVE_CTX.item.status === '承認済') msg += '\n\n※ 承認済みのため、申請中に戻り再承認が必要になります。';
  if (!confirm(msg)) return;

  var btn = document.getElementById('editSaveBtn');
  btn.disabled = true;
  btn.textContent = '保存中...';

  google.script.run
    .withSuccessHandler(function(res) {
      closeEditLeaveModal();
      if (res && res.ok) {
        alert('申請を更新しました。');
        loadLeaveTable();
        loadPendingRequests();
      } else {
        alert('更新に失敗しました。');
      }
    })
    .withFailureHandler(function(err) {
      btn.disabled = false;
      btn.textContent = '保存';
      alert('エラー: ' + (err.message || err));
    })
    .api_editLeaveRequest(EDIT_LEAVE_CTX.reqId, patch);
}

function confirmDeleteLeave(reqId) {
  var item = null;
  (window.LEAVE_ITEMS || []).forEach(function(it) { if (it.reqId === reqId) item = it; });
  var label = item ? (item.workerName + ' の休暇届（' + item.leaveDate + '）') : '選択した申請';
  if (!confirm(label + 'を削除しますか？\n\nこの操作は取り消せません。')) return;

  google.script.run
    .withSuccessHandler(function(res) {
      if (res && res.ok) {
        alert('申請を削除しました。');
        loadLeaveTable();
        loadPendingRequests();
      } else {
        alert('削除に失敗しました。');
      }
    })
    .withFailureHandler(function(err) {
      alert('削除エラー: ' + (err.message || err));
    })
    .api_deleteLeaveRequest(reqId);
}

/* ===================================================================
 * VPS 版で足した部分（ここから下だけ。上は GAS 版 src/top.html のまま）
 * =================================================================== */

// 申請モーダルを開く前に、ログインが続いているかを確かめる（操作中のログイン切れ）
(function(){
  var orig = openModal;
  openModal = function(){
    if (window.__vpsCheckSession) window.__vpsCheckSession();
    return orig.apply(this, arguments);
  };
})();

// 操作中にログインが切れたときの書きかけ（申請モーダル・編集モーダル）の退避と復元
(function(){
  if (!window.__vpsDraft) return;
  var when = window.__vpsDraft.when;
  function byId(id){ return document.getElementById(id); }
  function isActive(id){ var el = byId(id); return !!el && el.classList.contains('active'); }
  function checkedValue(name){ var r = document.querySelector('input[name="' + name + '"]:checked'); return r ? r.value : ''; }
  function setChecked(name, value){
    document.querySelectorAll('input[name="' + name + '"]').forEach(function(r){ r.checked = (r.value === value); });
  }

  function restoreRequest(r){
    // 部署 → 作業員（部署一覧が読めてから。作業員を選ぶと申請ボタンが押せる）
    when(function(){ var s = byId('deptSelect'); return s.options.length > 1 && s.options[0].value === ''; }, function(){
      var sel = byId('deptSelect');
      sel.value = r.dept;
      if (sel.value !== r.dept) return;
      selectedDeptId = r.dept;
      selectedDeptName = sel.options[sel.selectedIndex].dataset.name || '';
      loadWorkers(r.dept, r.worker);
      // 作業員と休暇日の一覧が届いてからモーダルを開き直す
      when(function(){ return selectedWorkerId === r.worker && byId('leaveDate').options.length > 1; }, function(){
        openModal();
        setChecked('leaveKubun', r.kubun);
        setChecked('halfDayType', r.halfDay);
        if (r.retro) {
          var p = r.retro.split('-');
          retroCalYear = Number(p[0]);
          retroCalMonth = Number(p[1]);
          if (r.retroOpen) toggleRetroSection(); else renderRetroCalendar();
          selectRetroDate(r.retro);
        } else {
          byId('leaveDate').value = r.leaveDate || '';
          if (r.retroOpen) toggleRetroSection();
        }
        if (r.kubun) {
          buildLeaveTypeOptions(r.kubun);
          setChecked('leaveType', r.leaveType);
        }
        byId('specialReason').value = r.specialReason || '';
        byId('paidDetail').value = r.paidDetail || '';
        var step = Number(r.step) || 1;
        // 振替元は候補が届いてから選び直す
        when(function(){ return ALL_HOLIDAY_CANDIDATES.length > 0 || !r.substituteDate; }, function(){
          renderSubstitutePulldown();
          byId('substituteDate').value = r.substituteDate || '';
          validateSubstituteSelection();
          if (step === 9) buildConfirmation();
        });
        if (step === 9) buildConfirmation();
        showStep(step);
      });
    });
  }

  function restoreEdit(e){
    when(function(){
      return (window.LEAVE_ITEMS || []).some(function(it){ return it.reqId === e.reqId; });
    }, function(){
      openEditLeaveModal(e.reqId);
      byId('editLeaveDate').value = e.leaveDate || '';
      setChecked('editLeaveKubun', e.kubun);
      byId('editHalfDayWrap').style.display = e.kubun === '半日休' ? '' : 'none';
      setChecked('editHalfDayType', e.halfDay);
    });
  }

  window.__vpsDraft.register({
    name: 'top',
    collect: function(){
      var out = {};
      if (isActive('modalOverlay') && selectedWorkerId && byId('submitBtn')) {
        out.request = {
          dept: selectedDeptId, worker: selectedWorkerId, step: currentStep,
          kubun: checkedValue('leaveKubun'), halfDay: checkedValue('halfDayType'), leaveType: checkedValue('leaveType'),
          leaveDate: byId('leaveDate').value, retro: selectedRetroactiveLeaveDate,
          retroOpen: byId('retroSection').classList.contains('open'),
          substituteDate: byId('substituteDate').value,
          specialReason: byId('specialReason').value, paidDetail: byId('paidDetail').value
        };
      }
      if (isActive('editOverlay') && EDIT_LEAVE_CTX && EDIT_LEAVE_CTX.reqId) {
        out.edit = {
          reqId: EDIT_LEAVE_CTX.reqId, leaveDate: byId('editLeaveDate').value,
          kubun: checkedValue('editLeaveKubun'), halfDay: checkedValue('editHalfDayType')
        };
      }
      return (out.request || out.edit) ? out : null;
    },
    restore: function(d){
      if (d.request) restoreRequest(d.request);
      else if (d.edit) restoreEdit(d.edit);
    }
  });
})();
