// ====== 権限判定 ======

/**
 * 管理ページアクセス可否（ALLOW_ADMIN_PAGE_EMAILS に含まれるか）
 */
function canAccessAdmin_() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var settings = getSettings_();
  var raw = normalize_(settings['ALLOW_ADMIN_PAGE_EMAILS'] || '');
  if (!raw) return false;
  var list = raw.split(',').map(function(s) { return s.trim().toLowerCase(); }).filter(Boolean);
  return list.indexOf(email) >= 0;
}

/**
 * 管理者か（ADMIN_EMAILS に含まれるか）
 */
function isAdmin_() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var settings = getSettings_();
  var raw = normalize_(settings['ADMIN_EMAILS'] || '');
  if (!raw) return false;
  var list = raw.split(',').map(function(s) { return s.trim().toLowerCase(); }).filter(Boolean);
  return list.indexOf(email) >= 0;
}

/**
 * 総務か（SOMU_EMAILS に含まれるか）
 */
function isSomu_() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var settings = getSettings_();
  var raw = normalize_(settings['SOMU_EMAILS'] || '');
  if (!raw) return false;
  var list = raw.split(',').map(function(s) { return s.trim().toLowerCase(); }).filter(Boolean);
  return list.indexOf(email) >= 0;
}

// ====== 部署承認者判定 ======

/**
 * 現在のユーザーが部署承認者かどうか判定
 * M_DEPT_APPROVERS シートの承認者メール列(B列)にメールが含まれていればtrue
 */
function isDeptApprover_() {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase();
    var depts = getDeptApproverDeptsForEmail_(email);
    return depts.length > 0;
  } catch (e) {
    console.log('isDeptApprover_ エラー: ' + e.message);
    return false;
  }
}

/**
 * 現在のユーザーが承認者として割り当てられている部署名の配列を返す
 * @return {string[]} 例: ['機械加工', '製缶・溶接']
 */
function getDeptApproverDepts_() {
  try {
    var email = Session.getActiveUser().getEmail().toLowerCase();
    return getDeptApproverDeptsForEmail_(email);
  } catch (e) {
    console.log('getDeptApproverDepts_ エラー: ' + e.message);
    return [];
  }
}

/**
 * 指定メールアドレスが承認者として割り当てられている部署名の配列を返す
 * サーバーサイドで既にメールを取得済みの場合に使用
 * @param {string} email メールアドレス
 * @return {string[]} 例: ['機械加工', '製缶・溶接']
 */
function getDeptApproverDeptsForEmail_(email) {
  try {
    if (!email) return [];
    var normalizedEmail = normalize_(email).toLowerCase();
    var sh = requireSheet_(SHEET.DEPT_APPROVERS);
    var data = sh.getDataRange().getValues();
    var depts = [];
    // 1行目はヘッダなのでスキップ
    for (var r = 1; r < data.length; r++) {
      var deptName = normalize_(String(data[r][0] || '')).trim();
      var approverEmails = normalize_(String(data[r][1] || ''));
      if (!deptName || !approverEmails) continue;
      var emailList = approverEmails.split(',').map(function(s) {
        return s.trim().toLowerCase();
      }).filter(Boolean);
      if (emailList.indexOf(normalizedEmail) >= 0) {
        depts.push(deptName);
      }
    }
    console.log('getDeptApproverDeptsForEmail_ email=' + normalizedEmail + ' depts=' + JSON.stringify(depts));
    return depts;
  } catch (e) {
    console.log('getDeptApproverDeptsForEmail_ エラー: ' + e.message);
    return [];
  }
}

// ====== API（google.script.run用） ======
function api_canAccessAdmin() {
  return canAccessAdmin_();
}

function api_isAdmin() {
  return isAdmin_();
}

function api_isSomu() {
  return isSomu_();
}

function api_isDeptApprover() {
  return isDeptApprover_();
}

function api_getDeptApproverDepts() {
  return getDeptApproverDepts_();
}

/**
 * 部署管理画面の初期化用: 承認者プロファイル一覧を返す
 * 管理者の場合 → 全承認者プロファイル + isAdmin=true
 * 一般承認者の場合 → 自分のプロファイルのみ + isAdmin=false
 * @return {Object} { isAdmin, currentEmail, profiles: [{ email, name, depts }] }
 */
function api_getApproverProfiles() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var admin = isAdmin_();

  // M_DEPT_APPROVERS を読んでプロファイル構築
  var sh = requireSheet_(SHEET.DEPT_APPROVERS);
  var data = sh.getDataRange().getValues();
  var profileMap = {}; // email → { email, name, depts: [] }

  for (var r = 1; r < data.length; r++) {
    var deptName = normalize_(String(data[r][0] || '')).trim();
    var approverEmails = normalize_(String(data[r][1] || ''));
    var approverName = normalize_(String(data[r][2] || ''));
    if (!deptName || !approverEmails) continue;

    var emailList = approverEmails.split(',').map(function(s) { return s.trim().toLowerCase(); }).filter(Boolean);
    for (var i = 0; i < emailList.length; i++) {
      var e = emailList[i];
      if (!profileMap[e]) {
        profileMap[e] = { email: e, name: approverName || e, depts: [] };
      }
      if (profileMap[e].depts.indexOf(deptName) < 0) {
        profileMap[e].depts.push(deptName);
      }
      // 名前が空だった場合、後の行で名前が見つかれば上書き
      if (approverName && profileMap[e].name === e) {
        profileMap[e].name = approverName;
      }
    }
  }

  var profiles = [];
  var keys = Object.keys(profileMap);
  for (var k = 0; k < keys.length; k++) {
    profiles.push(profileMap[keys[k]]);
  }
  // 名前でソート
  profiles.sort(function(a, b) { return a.name.localeCompare(b.name, 'ja'); });

  if (admin) {
    return { isAdmin: true, currentEmail: email, profiles: profiles };
  } else {
    // 自分のプロファイルだけ返す
    var myProfile = profileMap[email];
    return {
      isAdmin: false,
      currentEmail: email,
      profiles: myProfile ? [myProfile] : [],
    };
  }
}
