// 画面から呼ぶ関数の閉じたリスト（docs/vps-implementation-spec.md 4-1 の RPC_FUNCTIONS、計 34）。試験が照らし合わせる元。
// SCREEN = 仕様 4-3。GAS の画面が呼ぶ 27 関数（inv-2-screens.md 4 節。呼び出し 37 か所）
// OPS    = 仕様 4-2。総務管理の一番下の「運用（管理者）」欄が呼ぶ 7 関数（VPS 版で足したもの）
// WRITE  = 書き込みの関数（シムが冪等キーを付ける）
'use strict';

const SCREEN = [
  'api_isDeptApprover', 'api_isAdmin', 'api_isSomu', 'api_getDeptList', 'api_getWorkerInfo',
  'api_getLookupData', 'api_getWorkingDays', 'api_getWorkersByDept', 'api_getUpcomingLeaves',
  'api_getHolidayDays', 'api_getPastWorkingDays', 'api_submitLeaveRequest', 'api_getAllPendingRequests',
  'api_editLeaveRequest', 'api_deleteLeaveRequest', 'api_getLeaveRequestById', 'api_getApproverList',
  'api_approveLeaveRequest', 'api_getCurrentApproverInfo', 'api_getPending2ndApprovals',
  'api_approveLeaveRequest2', 'api_getSomuAdminView', 'api_getApproverProfiles', 'api_getDeptAdminView',
  'api_getAdminSomuView', 'api_getAdminBossView', 'api_getLeaveStatusSearch'
];
const OPS = [
  'ops_getStatus', 'ops_retryJob', 'ops_syncSheetConfigNow', 'ops_runMasterSync',
  'ops_rebuildSummaryNow', 'ops_sendWarnMail', 'ops_revokeSessions'
];
const WRITE = [
  'api_submitLeaveRequest', 'api_editLeaveRequest', 'api_deleteLeaveRequest',
  'api_approveLeaveRequest', 'api_approveLeaveRequest2',
  'ops_retryJob', 'ops_syncSheetConfigNow', 'ops_runMasterSync', 'ops_rebuildSummaryNow',
  'ops_sendWarnMail', 'ops_revokeSessions'
];
// 画面ごとの呼び出しの数（inv-2 の 4-1 の #1〜#37。4 節の上の表は somuAdmin を 7 と書いているが、#19〜#26 の 8 か所が正）
const CALLS_PER_PAGE = { top: 15, somuAdmin: 8, admin: 4, leaveStatus: 4, deptAdmin: 3, sign: 2, preview: 1 };

module.exports = { SCREEN, OPS, WRITE, ALL: SCREEN.concat(OPS), CALLS_PER_PAGE };
