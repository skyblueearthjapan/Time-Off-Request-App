// ====== スタッフ向け 休暇取得状況ダッシュボード API ======
// 本ページ専用。誰でも・どの部署/スタッフでも検索できる（権限ガードなし）。
// ※ 集計ロジックは既存 buildAdminViewData_ / PAID_LEAVE_WARN を再利用する。

/**
 * 休暇取得状況の検索
 * @param {number|string} fyYear 年度（省略時は現在年度）
 * @param {string} deptId 部署名（空=全部署）
 * @param {string} workerId 作業員ID（指定=本人ビュー / 空=一覧ビュー）
 * @return {Object} personalモード or listモード
 */
function api_getLeaveStatusSearch(fyYear, deptId, workerId) {
  var fy = fyYear ? Number(fyYear) : computeFiscalYear_(new Date());

  // 対象部署の決定（deptId は部署名そのもの）
  var depts;
  if (deptId) {
    depts = [deptId];
  } else {
    depts = loadDeptList_().map(function(d) { return d.deptName; });
  }

  // 集計（既存中核をそのまま利用）
  var workers = buildAdminViewData_(depts, fy);

  if (workerId) {
    var me = null;
    for (var i = 0; i < workers.length; i++) {
      if (workers[i].workerId === workerId) { me = workers[i]; break; }
    }
    var upcoming = (me && fy === computeFiscalYear_(new Date()))
      ? api_getUpcomingLeaves(me.deptName, workerId)
      : [];
    return {
      mode: 'personal',
      fyYear: fy,
      worker: me,                                   // null なら該当なし
      milestones: me ? buildMilestones_(me.paidCount, fy) : [],
      upcoming: upcoming
    };
  }

  return {
    mode: 'list',
    fyYear: fy,
    deptFilter: deptId || '',
    workers: workers
  };
}

/**
 * 5段階しきい値（PAID_LEAVE_WARN）を「今日時点」の達成状況に変換
 * @return {Object[]} タイムライン描画用の配列
 */
function buildMilestones_(paidCount, fyYear) {
  var now = new Date();
  var fy = Number(fyYear);
  return PAID_LEAVE_WARN.map(function(w) {
    var m = w.deadline.month, day = w.deadline.day;
    // 3/16以降は fy 年、1月〜3/15 は fy+1 年（computeFiscalYear_ と同一ロジック）
    var dyear = (m > 3 || (m === 3 && day >= 16)) ? fy : fy + 1;
    var deadline = new Date(dyear, m - 1, day, 23, 59, 59);
    var passed = now >= deadline;
    var met = paidCount >= w.required;
    var deficit = Math.max(0, w.required - paidCount);
    var state;
    if (met) state = 'done';
    else if (passed) state = (deficit >= 3) ? 'risk' : (deficit >= 2 ? 'warn' : 'note');
    else state = 'future';
    return {
      label: w.label, month: m, day: day, dateText: m + '/' + day,
      required: w.required, met: met, passed: passed,
      deficit: deficit, remaining: Math.max(0, w.required - paidCount),
      state: state
    };
  });
}
