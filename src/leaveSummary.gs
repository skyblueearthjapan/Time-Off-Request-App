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
