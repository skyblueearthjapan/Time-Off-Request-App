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

// ====== API（google.script.run用） ======
function api_canAccessAdmin() {
  return canAccessAdmin_();
}

function api_isSomu() {
  return isSomu_();
}
