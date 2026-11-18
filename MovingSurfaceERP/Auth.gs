function getAdminUser_(pin) {
  const credential = String(pin == null ? '' : pin).trim().split('|');
  const suppliedPin = credential[0];
  const selectedStaffId = String(credential[1] || '').trim();
  if (!isValidPin_(suppliedPin)) return {valid:false};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET);
  if (!sh || sh.getLastRow() < 2) return {valid:false};
  const rows = sh.getRange(2,1,sh.getLastRow()-1,Math.max(12,sh.getLastColumn())).getValues();
  const matched = findAdminCredentialMatch_(rows,suppliedPin,selectedStaffId);
  if (!matched) return {valid:false};
  const row = matched.row;
  const user = {
    valid:true, name:String(row[0]), role:String(row[3]), canView:String(row[4]).toUpperCase(),
    canOverride:String(row[5]).toUpperCase(), canReset:String(row[6]).toUpperCase(),
    canOnboard:String(row[7]).toUpperCase(), staffId:String(row[8] || ''),
    email:String(row[9] || '').trim(), accessPages:String(row[10] || '').trim(),
    writePages:String(row[11] || '').trim(), row:matched.index+2
  };
  user.access = getAdminAccess_(user);
  return user;
}

function findAdminCredentialMatch_(rows,suppliedPin,selectedStaffId) {
  const matches = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (normalizePin_(row[1]) === suppliedPin && String(row[2]).toUpperCase() === 'YES') {
      matches.push({row:row,index:i});
    }
  }
  const candidates = selectedStaffId
    ? matches.filter(item => String(item.row[8] || '').trim() === selectedStaffId)
    : matches;
  return candidates.length === 1 ? candidates[0] : null;
}

function getAdminLoginOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const activeStaffIds = Object.create(null);
  getActiveStaff_().forEach(staff => { activeStaffIds[staff.id] = staff.name; });
  const seen = Object.create(null);
  return sheet.getRange(2,1,sheet.getLastRow()-1,Math.max(12,sheet.getLastColumn())).getValues()
    .filter(row => String(row[2]).toUpperCase() === 'YES' && String(row[4]).toUpperCase() === 'YES' &&
      row[8] && activeStaffIds[String(row[8]).trim()])
    .filter(row => {
      const staffId = String(row[8]).trim();
      if (seen[staffId]) return false;
      seen[staffId] = true;
      return true;
    })
    .map(row => ({staffId:String(row[8]).trim(),name:String(row[0] || activeStaffIds[String(row[8]).trim()])}));
}

function getAdminAccess_(user) {
  if (user.canView !== 'YES') return {readPages:[],writePages:[],canManageStaff:false};
  const role = normalizeAdminRole_(user.role);
  const canManageUsers = ['superadmin','admin'].includes(role);
  const privileged = ['superadmin','admin','hr','adminhr','administrator','humanresources'].includes(role);
  const technical = role === 'technical';
  const accountant = role === 'accountant';
  const allPages = ['dashboard','attendance','staff','requisitions','inventory','assets','expenses','leave','projects','leads','reports','activity'];
  const defaultRead = privileged ? allPages : allPages.filter(page => page !== 'staff');
  const defaultWrite = privileged
    ? allPages.slice()
    : technical ? allPages.filter(page => page !== 'staff')
    : accountant ? ['requisitions','expenses'] : [];
  const readPages = parseAdminPageList_(user.accessPages,defaultRead)
    .filter(page => allPages.includes(page));
  if (canManageUsers && !readPages.includes('dashboard')) readPages.unshift('dashboard');
  const writePages = parseAdminPageList_(user.writePages,defaultWrite)
    .filter(page => allPages.includes(page));
  const allowedWrite = writePages.filter(page => readPages.includes(page));
  return {
    readPages:readPages.concat(canManageUsers ? ['users'] : []),
    writePages:allowedWrite,
    canManageStaff:privileged,
    canManageUsers:canManageUsers
  };
}

function normalizeAdminRole_(role) {
  return String(role || '').toLowerCase().replace(/[^a-z]/g,'');
}

function parseAdminPageList_(value,defaults) {
  if (!String(value || '').trim()) return defaults.slice();
  return String(value).split(',').map(page => page.trim().toLowerCase()).filter(Boolean);
}

function adminCanReadPage_(user,page) {
  return Boolean(user && user.valid && user.canView === 'YES' && user.access &&
    user.access.readPages.includes(String(page || '').toLowerCase()));
}

function adminCanWritePage_(user,page) {
  return Boolean(user && user.valid && user.canView === 'YES' && user.access &&
    user.access.writePages.includes(String(page || '').toLowerCase()));
}

function adminCanApprovePage_(user,page) {
  return adminCanWritePage_(user,page) &&
    Boolean(user.access.canManageStaff || user.canOverride === 'YES');
}

function adminCanDisbursePage_(user,page) {
  const role = normalizeAdminRole_(user && user.role);
  return adminCanWritePage_(user,page) &&
    Boolean(user && user.access &&
      (user.access.canManageStaff || ['accountant','accounts','finance'].includes(role)));
}

function adminCanManageUsers_(user) {
  return Boolean(user && user.valid && user.canView === 'YES' && user.access &&
    user.access.canManageUsers && user.access.readPages.includes('users'));
}

function requireAdminPageWrite_(user,page) {
  if (!adminCanWritePage_(user,page)) {
    throw new Error('You do not have permission to make changes on the '+page+' page.');
  }
}

function requireAdminPageRead_(user,page) {
  if (!adminCanReadPage_(user,page)) {
    throw new Error('You do not have access to the '+page+' page.');
  }
}

function normalizePin_(value) {
  const pin = String(value == null ? '' : value).trim();
  return /^\d{1,4}$/.test(pin) ? pin.padStart(4, '0') : pin;
}

function isValidPin_(pin) {
  return /^\d{4}$/.test(String(pin || ''));
}

function updateStaffPin_(row, pin) {
  const normalizedPin = String(pin == null ? '' : pin).trim();
  if (!isValidPin_(normalizedPin)) throw new Error('PIN must be exactly 4 digits.');
  const cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET).getRange(row, 5);
  cell.setNumberFormat('@');
  cell.setValue(normalizedPin);
}

function updateStaffDeviceToken_(row, token) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET).getRange(row, 6).setValue(sha256Hex_(token));
}

function deviceTokenMatches_(storedToken, submittedToken) {
  const stored = String(storedToken || '');
  const submitted = String(submittedToken || '');
  if (!stored || !submitted) return false;
  if (/^[a-f0-9]{64}$/i.test(stored)) return stored.toLowerCase() === sha256Hex_(submitted);
  return stored === submitted;
}

function deviceTokenFingerprint_(storedToken) {
  const token = String(storedToken || '');
  return /^[a-f0-9]{64}$/i.test(token) ? token.toLowerCase() : (token ? sha256Hex_(token) : '');
}

function migrateDeviceTokens_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName(STAFF_SHEET);
  if (staffSheet && staffSheet.getLastRow() > 1) {
    const range = staffSheet.getRange(2,6,staffSheet.getLastRow()-1,1);
    range.setValues(range.getValues().map(row => [deviceTokenFingerprint_(row[0])]));
  }
  const logSheet = ss.getSheetByName(LOG_SHEET);
  if (logSheet && logSheet.getLastRow() > 1) {
    const range = logSheet.getRange(2,21,logSheet.getLastRow()-1,1);
    range.setValues(range.getValues().map(row => [deviceTokenFingerprint_(row[0])]));
  }
}

function migratePinsToFourDigits() {
  migratePinColumn_(STAFF_SHEET, 5);
  migratePinColumn_(ADMIN_USERS_SHEET, 2);
}

function migratePinColumn_(sheetName, column) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  const range = sheet.getRange(2, column, sheet.getLastRow() - 1, 1);
  const values = range.getDisplayValues().map(function(row) {
    const original = String(row[0] || '').trim();
    const normalized = normalizePin_(original);
    return [isValidPin_(normalized) ? normalized : original];
  });
  range.setNumberFormat('@');
  range.setValues(values);
}
