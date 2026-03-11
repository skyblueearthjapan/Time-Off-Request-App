// ====== 有給取得集計 ======

/** メニューから呼べる公開ラッパー */
function rebuildLeaveSummary() {
  var count = rebuildLeaveSummary_();
  var msg = 'サマリー再構築が完了しました。(' + count + '件)';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

/**
 * V_LEAVE_SUMMARYを再構築
 * T_LEAVE_REQUESTの承認済データから年度×作業員で集計
 */
function rebuildLeaveSummary_() {
  var reqInfo = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var reqSh = reqInfo.sh;
  var reqIdx = reqInfo.idx;
  var lastRow = reqSh.getLastRow();

  // 集計データ構築
  var summaryMap = {}; // key = 年度_作業員ID

  if (lastRow >= 2) {
    var readCols = Math.max(reqInfo.header.length, reqSh.getLastColumn());
    var values = reqSh.getRange(2, 1, lastRow - 1, readCols).getValues();

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      // 空行スキップ（REQ_IDが空なら無視）
      var reqIdCol = reqIdx['REQ_ID'] !== undefined ? reqIdx['REQ_ID'] : 0;
      if (!normalize_(row[reqIdCol])) continue;

      var status = normalize_(row[reqIdx['承認状態']]);
      if (status !== STATUS.APPROVED) continue;

      var fy = row[reqIdx['年度']];
      var workerId = normalize_(row[reqIdx['作業員ID']]);
      var workerName = normalize_(row[reqIdx['作業員名']]);
      var deptId = normalize_(row[reqIdx['部署ID']]);
      var deptName = normalize_(row[reqIdx['部署名']]);
      var leaveType = normalize_(row[reqIdx['休暇種類']]);
      var leaveKubun = normalize_(row[reqIdx['休暇区分']]);
      var leaveDate = row[reqIdx['休暇日']];

      var key = fy + '_' + workerId;
      if (!summaryMap[key]) {
        summaryMap[key] = {
          fiscalYear: fy,
          deptId: deptId,
          deptName: deptName,
          workerId: workerId,
          workerName: workerName,
          paidCount: 0,
          paidHalfCount: 0,
          substituteCount: 0,
          specialCount: 0,
          totalCount: 0,
          lastPaidDate: '',
        };
      }

      var s = summaryMap[key];
      s.totalCount++;

      if (leaveType === LEAVE_TYPE.PAID) {
        if (leaveKubun === '半日休') {
          s.paidHalfCount++;
          s.paidCount += 0.5;
        } else {
          s.paidCount++;
        }
        // 最終有給日を更新
        if (leaveDate instanceof Date) {
          var ldStr = fmtDate_(leaveDate, 'yyyy-MM-dd');
          if (!s.lastPaidDate || ldStr > s.lastPaidDate) s.lastPaidDate = ldStr;
        }
      } else if (leaveType === LEAVE_TYPE.SUBSTITUTE) {
        s.substituteCount++;
      } else if (leaveType === LEAVE_TYPE.SPECIAL) {
        s.specialCount++;
      }
    }
  }

  // V_LEAVE_SUMMARYシートに書き込み
  var ss = getDb_();
  var sumSh = ss.getSheetByName(SHEET.LEAVE_SUMMARY);
  if (!sumSh) {
    sumSh = ss.insertSheet(SHEET.LEAVE_SUMMARY);
  }
  sumSh.clearContents();

  var header = ['年度', '部署ID', '部署名', '作業員ID', '作業員名',
                '有給回数', '有給半日回数', '振替回数', '特別回数', '合計回数',
                '最終有給日', '警告レベル'];
  sumSh.getRange(1, 1, 1, header.length).setValues([header]);

  var keys = Object.keys(summaryMap).sort();
  if (keys.length > 0) {
    var rows = [];
    for (var k = 0; k < keys.length; k++) {
      var s = summaryMap[keys[k]];
      var warnLevel = computeWarnLevel_(s.paidCount, s.fiscalYear);
      rows.push([
        s.fiscalYear, s.deptId, s.deptName, s.workerId, s.workerName,
        s.paidCount, s.paidHalfCount, s.substituteCount, s.specialCount, s.totalCount,
        s.lastPaidDate, warnLevel,
      ]);
    }
    sumSh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  SpreadsheetApp.flush();
  return keys.length;
}

/**
 * 有給警告レベル計算
 * 現在日時点で、期限内に必要回数に達しているかチェック
 * 戻り値: '' (問題なし), '注意', '警告', '危険'
 */
function computeWarnLevel_(paidCount, fiscalYear) {
  var now = new Date();
  var fy = Number(fiscalYear);

  for (var i = PAID_LEAVE_WARN.length - 1; i >= 0; i--) {
    var w = PAID_LEAVE_WARN[i];
    var deadlineMonth = w.deadline.month;
    var deadlineDay = w.deadline.day;

    // 年度内の日付を算出（4月〜12月はfy年、1月〜3月はfy+1年）
    var deadlineYear = deadlineMonth >= 4 ? fy : fy + 1;
    var deadlineDate = new Date(deadlineYear, deadlineMonth - 1, deadlineDay, 23, 59, 59);

    // まだ期限が来ていない場合はスキップ
    if (now < deadlineDate) continue;

    // 期限が過ぎているのに回数が足りない
    if (paidCount < w.required) {
      var deficit = w.required - paidCount;
      if (deficit >= 3) return '危険';
      if (deficit >= 2) return '警告';
      return '注意';
    }
    // この期限はクリア → 次の期限は未来のはずなのでOK
    break;
  }
  return '';
}

/**
 * 上長ビュー: 指定部署の年度別取得状況
 */
function api_getAdminBossView(deptId, fyYear) {
  if (!canAccessAdmin_()) throw new Error('権限がありません');
  if (!fyYear) {
    fyYear = computeFiscalYear_(new Date());
  }

  // サマリーを最新化
  rebuildLeaveSummary_();

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
    if (deptId && normalize_(row[idx['部署ID']]) !== deptId) continue;

    out.push({
      deptId: normalize_(row[idx['部署ID']]),
      deptName: normalize_(row[idx['部署名']]),
      workerId: normalize_(row[idx['作業員ID']]),
      workerName: normalize_(row[idx['作業員名']]),
      paidCount: Number(row[idx['有給回数']] || 0),
      paidHalfCount: Number(row[idx['有給半日回数']] || 0),
      substituteCount: Number(row[idx['振替回数']] || 0),
      specialCount: Number(row[idx['特別回数']] || 0),
      totalCount: Number(row[idx['合計回数']] || 0),
      lastPaidDate: normalize_(row[idx['最終有給日']]),
      warnLevel: normalize_(row[idx['警告レベル']]),
    });
  }

  out.sort(function(a, b) {
    return (a.deptName + a.workerName).localeCompare(b.deptName + b.workerName, 'ja');
  });

  return out;
}

/**
 * 総務ビュー: 全部署の年度別取得状況
 */
function api_getAdminSomuView(fyYear, deptFilter) {
  if (!canAccessAdmin_()) throw new Error('権限がありません');
  return api_getAdminBossView(deptFilter || '', fyYear);
}
// computeFiscalYear_ は leaveRequest.gs で定義済み

// ====== 期間別有給取得管理 ======

/**
 * 有給取得日の配列から期間別ステータスを算出
 * @param {Date[]} paidLeaveDates 承認済み有給の休暇日配列
 * @param {number} fiscalYear 年度
 * @return {Object[]} PAID_LEAVE_QUARTERS に対応する5要素の配列
 */
function computeQuarterStatus_(paidLeaveDates, fiscalYear) {
  var fy = Number(fiscalYear);
  var now = new Date();
  var result = [];
  var cumulativeCount = 0; // 年度通算の有給取得数

  for (var i = 0; i < PAID_LEAVE_QUARTERS.length; i++) {
    var q = PAID_LEAVE_QUARTERS[i];

    // 期間の開始日・終了日を算出（4-12月はfy年、1-3月はfy+1年）
    var startYear = q.startMonth >= 4 ? fy : fy + 1;
    var endYear = q.endMonth >= 4 ? fy : fy + 1;
    var periodStart = new Date(startYear, q.startMonth - 1, q.startDay, 0, 0, 0);
    var periodEnd = new Date(endYear, q.endMonth - 1, q.endDay, 23, 59, 59);

    // 期間内の有給日数をカウント
    var periodCount = 0;
    for (var j = 0; j < paidLeaveDates.length; j++) {
      var d = paidLeaveDates[j];
      if (d >= periodStart && d <= periodEnd) {
        periodCount++;
      }
    }
    cumulativeCount += periodCount;

    // 現在この期間内か、過ぎているか、まだ先か
    var isCurrent = now >= periodStart && now <= periodEnd;
    var isPast = now > periodEnd;
    var isFuture = now < periodStart;

    // 累計不足数（targetCount = この期間終了までに必要な累計回数）
    var shortage = Math.max(0, q.targetCount - cumulativeCount);

    // 危険度判定（現在 or 過去の期間のみ）
    // shortage >= 2 → danger（赤）, shortage === 1 → warn（黄）, 0 → ok
    var level = '';
    if (isCurrent || isPast) {
      if (shortage >= 2) level = 'danger';
      else if (shortage === 1) level = 'warn';
      else level = 'ok';
    }

    result.push({
      id: q.id,
      label: q.label,
      targetCount: q.targetCount,
      periodCount: periodCount,
      cumulativeCount: cumulativeCount,
      shortage: shortage,
      level: level,
      taken: periodCount > 0,
      overdue: shortage > 0 && isPast,
      isCurrent: isCurrent,
      isPast: isPast,
      isFuture: isFuture,
    });
  }

  return result;
}

/**
 * 部署管理ビュー用データ構築（共通処理）
 * @param {string[]} depts 対象部署名の配列
 * @param {number} fyYear 年度
 * @return {Object[]} 作業員ごとの取得状況配列
 */
function buildAdminViewData_(depts, fyYear) {
  var fy = Number(fyYear);

  // 1. サマリーを最新化
  rebuildLeaveSummary_();

  // 2. 作業員マスタから全アクティブ作業員を取得
  var workerMap = loadWorkersByDeptMap_();

  // 3. V_LEAVE_SUMMARYの読込
  var sumInfo = getSheetHeaderIndex_(SHEET.LEAVE_SUMMARY, 1);
  var sumSh = sumInfo.sh;
  var sumIdx = sumInfo.idx;
  var sumLastRow = sumSh.getLastRow();
  var summaryByWorker = {}; // workerId → サマリー行データ
  if (sumLastRow >= 2) {
    var sumValues = sumSh.getRange(2, 1, sumLastRow - 1, sumSh.getLastColumn()).getValues();
    for (var i = 0; i < sumValues.length; i++) {
      var row = sumValues[i];
      if (Number(row[sumIdx['年度']]) !== fy) continue;
      var wid = normalize_(row[sumIdx['作業員ID']]);
      summaryByWorker[wid] = {
        paidCount: Number(row[sumIdx['有給回数']] || 0),
        paidHalfCount: Number(row[sumIdx['有給半日回数']] || 0),
        substituteCount: Number(row[sumIdx['振替回数']] || 0),
        specialCount: Number(row[sumIdx['特別回数']] || 0),
        totalCount: Number(row[sumIdx['合計回数']] || 0),
        lastPaidDate: normalize_(row[sumIdx['最終有給日']]),
        warnLevel: normalize_(row[sumIdx['警告レベル']]),
      };
    }
  }

  // 4. T_LEAVE_REQUESTから承認済み有給の個別日付を取得
  var reqInfo = getSheetHeaderIndex_(SHEET.LEAVE_REQUEST, 1);
  var reqSh = reqInfo.sh;
  var reqIdx = reqInfo.idx;
  var reqLastRow = reqSh.getLastRow();
  var paidDatesByWorker = {}; // workerId → Date[]
  if (reqLastRow >= 2) {
    var readCols = Math.max(reqInfo.header.length, reqSh.getLastColumn());
    var reqValues = reqSh.getRange(2, 1, reqLastRow - 1, readCols).getValues();
    for (var i = 0; i < reqValues.length; i++) {
      var row = reqValues[i];
      var reqIdCol = reqIdx['REQ_ID'] !== undefined ? reqIdx['REQ_ID'] : 0;
      if (!normalize_(row[reqIdCol])) continue;

      var status = normalize_(row[reqIdx['承認状態']]);
      if (status !== STATUS.APPROVED) continue;

      var rowFy = Number(row[reqIdx['年度']]);
      if (rowFy !== fy) continue;

      var leaveType = normalize_(row[reqIdx['休暇種類']]);
      if (leaveType !== LEAVE_TYPE.PAID) continue;

      var workerId = normalize_(row[reqIdx['作業員ID']]);
      var leaveDate = row[reqIdx['休暇日']];
      if (!(leaveDate instanceof Date)) continue;

      if (!paidDatesByWorker[workerId]) paidDatesByWorker[workerId] = [];
      paidDatesByWorker[workerId].push(leaveDate);
    }
  }

  // 5. 対象部署の作業員ごとにデータ構築
  var deptsNorm = {};
  for (var d = 0; d < depts.length; d++) {
    deptsNorm[normalize_(depts[d])] = true;
  }

  var out = [];
  var entries = Array.from(workerMap.entries());
  for (var e = 0; e < entries.length; e++) {
    var deptName = entries[e][0];
    if (!deptsNorm[normalize_(deptName)]) continue;

    var workers = entries[e][1];
    for (var w = 0; w < workers.length; w++) {
      var worker = workers[w];
      var wid = worker.workerId;

      // サマリーデータ（なければゼロデフォルト）
      var sum = summaryByWorker[wid] || {
        paidCount: 0,
        paidHalfCount: 0,
        substituteCount: 0,
        specialCount: 0,
        totalCount: 0,
        lastPaidDate: '',
        warnLevel: '',
      };

      // 期間別ステータス算出
      var dates = paidDatesByWorker[wid] || [];
      var quarters = computeQuarterStatus_(dates, fy);

      out.push({
        deptName: deptName,
        workerId: wid,
        workerName: worker.workerName,
        paidCount: sum.paidCount,
        paidHalfCount: sum.paidHalfCount,
        substituteCount: sum.substituteCount,
        specialCount: sum.specialCount,
        totalCount: sum.totalCount,
        lastPaidDate: sum.lastPaidDate,
        warnLevel: computeWarnLevel_(sum.paidCount, fy),
        quarters: quarters,
      });
    }
  }

  // 6. 現在期間の累計不足数を算出し、不足が多い人を上位に
  for (var s = 0; s < out.length; s++) {
    var qs = out[s].quarters || [];
    var maxShortage = 0;
    for (var qi = 0; qi < qs.length; qi++) {
      if (qs[qi].isCurrent || qs[qi].isPast) {
        maxShortage = Math.max(maxShortage, qs[qi].shortage);
      }
    }
    out[s].currentShortage = maxShortage;
  }
  out.sort(function(a, b) {
    // 有給取得数昇順（少ない人が先頭）
    if (a.paidCount !== b.paidCount) return a.paidCount - b.paidCount;
    // 同数なら不足数降順
    if (a.currentShortage !== b.currentShortage) return b.currentShortage - a.currentShortage;
    var deptCmp = a.deptName.localeCompare(b.deptName, 'ja');
    if (deptCmp !== 0) return deptCmp;
    return a.workerName.localeCompare(b.workerName, 'ja');
  });

  return out;
}

/**
 * 部署承認者ビュー: 担当部署の年度別有給取得状況（期間別付き）
 * 管理者は任意の部署を指定可能（承認者選択モーダル用）
 * @param {number} fyYear 年度
 * @param {string[]} [selectedDepts] 管理者が指定する部署配列（省略時は自分の担当部署）
 * @return {Object} { workers: [...], deptNames: [...] }
 */
function api_getDeptAdminView(fyYear, selectedDepts) {
  var admin = isAdmin_();
  if (!admin && !isDeptApprover_()) throw new Error('部署承認者の権限がありません。');

  var depts;
  if (selectedDepts && selectedDepts.length > 0 && admin) {
    // 管理者が特定の承認者の部署を選択
    depts = selectedDepts;
  } else if (selectedDepts && selectedDepts.length > 0) {
    // 一般承認者 → 自分の担当部署のみ許可
    var myDepts = getDeptApproverDepts_();
    depts = selectedDepts.filter(function(d) { return myDepts.indexOf(d) >= 0; });
  } else {
    depts = getDeptApproverDepts_();
  }

  if (!depts || depts.length === 0) return { workers: [], deptNames: [] };
  var fy = fyYear || computeFiscalYear_(new Date());
  return { workers: buildAdminViewData_(depts, fy), deptNames: depts };
}

/**
 * 総務管理ビュー: 全部署または指定部署の年度別有給取得状況（期間別付き）
 * @param {number} fyYear 年度
 * @param {string} deptFilter 部署フィルタ（空/'ALL'で全部署）
 * @return {Object[]} 作業員ごとの取得状況配列
 */
function api_getSomuAdminView(fyYear, deptFilter) {
  if (!isSomu_() && !isAdmin_()) {
    throw new Error('権限がありません（総務または管理者のみ）');
  }

  if (!fyYear) {
    fyYear = computeFiscalYear_(new Date());
  }

  var depts;
  if (!deptFilter || normalize_(deptFilter) === 'ALL' || normalize_(deptFilter) === '') {
    // 全部署取得
    var deptList = loadDeptList_();
    depts = deptList.map(function(d) { return d.deptName; });
  } else {
    depts = [deptFilter];
  }

  return buildAdminViewData_(depts, fyYear);
}
