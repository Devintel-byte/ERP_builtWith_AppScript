function getAdminUser_(pin) {
  const credential = String(pin == null ? '' : pin).trim().split('|');
  const suppliedPin = credential[0];
  const selectedStaffId = String(credential[1] || '').trim();
  if (!isValidPin_(suppliedPin)) return {valid:false};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET);
  if (!sh || sh.getLastRow() < 2) return {valid:false};
  const rows = sh.getRange(2,1,sh.getLastRow()-1,Math.max(12,sh.getLastColumn())).getValues();
  if (isPinLocked_(selectedStaffId)) return {valid:false, locked:true};
  const matched = findAdminCredentialMatch_(rows,suppliedPin,selectedStaffId,getStaffPinMapFor_(selectedStaffId));
  if (!matched) {
    recordPinFailure_(selectedStaffId);
    return {valid:false};
  }
  clearPinFailures_(selectedStaffId);
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

/**
 * Finds the Admin Users row a sign-in belongs to.
 *
 * A person has ONE PIN, and it lives in Staff List. When an access account is
 * linked to a staff member (Staff ID column), that staff member's PIN is the
 * credential and the Staff ID chosen at sign-in must be the linked one. The
 * Admin Users PIN column is only consulted for legacy rows with no Staff ID, so
 * the two stores cannot drift apart and a reset or recovery that changes the
 * staff PIN changes admin sign-in with it.
 *
 * `staffPins` is {staffId: pin} for active staff with a valid PIN. When it is
 * not supplied every row falls back to its own PIN column, which keeps older
 * callers and tests working.
 */
function findAdminCredentialMatch_(rows,suppliedPin,selectedStaffId,staffPins) {
  const useStaffPins = Boolean(staffPins) && typeof staffPins === 'object';
  const matches = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[2]).toUpperCase() !== 'YES') continue;
    const linkedId = String(row[8] || '').trim();
    let pinMatches;
    if (useStaffPins && linkedId) {
      if (selectedStaffId !== linkedId) continue;
      pinMatches = Object.prototype.hasOwnProperty.call(staffPins,linkedId) && staffPins[linkedId] === suppliedPin;
    } else {
      pinMatches = normalizePin_(row[1]) === suppliedPin;
    }
    if (pinMatches) matches.push({row:row,index:i});
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
  const allPages = ['dashboard','attendance','staff','requisitions','inventory','assets','expenses','leave','projects','leads','fieldduty','reports','activity'];
  const defaultRead = privileged ? allPages : allPages.filter(page => page !== 'staff');
  const defaultWrite = privileged
    ? allPages.slice()
    : technical ? allPages.filter(page => page !== 'staff')
    : accountant ? ['requisitions','expenses'] : [];
  const readPages = parseAdminPageList_(user.accessPages,defaultRead)
    .filter(page => allPages.includes(page));
  if (canManageUsers && !readPages.includes('dashboard')) readPages.unshift('dashboard');
  // Admin and Super Admin can approve field duty, so they always reach its page even
  // when an explicit Access Pages list (which replaces the role defaults) predates it.
  if (canManageUsers && !readPages.includes('fieldduty')) readPages.push('fieldduty');
  const writePages = parseAdminPageList_(user.writePages,defaultWrite)
    .filter(page => allPages.includes(page));
  const allowedWrite = writePages.filter(page => readPages.includes(page));
  if (canManageUsers && !allowedWrite.includes('fieldduty')) allowedWrite.push('fieldduty');
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

/**
 * {staffId: pin} for the one staff member signing in, if they are active and hold a
 * valid PIN. A linked admin row is only ever matched against the Staff ID chosen at
 * sign-in, so there is no reason to read the whole roster on every admin request.
 */
function getStaffPinMapFor_(staffId) {
  const map = Object.create(null);
  const id = String(staffId || '').trim();
  if (!id) return map;
  const staff = getStaffById_(id);
  const pin = staff ? normalizePin_(staff.pin) : '';
  if (isValidPin_(pin)) map[id] = pin;
  return map;
}

/**
 * The three sample accounts shipped with placeholder PINs. They are kept here as
 * digests only so the digits are not repeated in source; that is tidiness, not
 * secrecy, because four digits are trivially reversible.
 *
 * A placeholder still signs in until the person creates their own PIN, but
 * attendance treats it as "not set": the person is asked to choose a PIN, which
 * then replaces the placeholder for every page, and the placeholder can never be
 * chosen again.
 */
function startingPinDigests_() {
  return {
    'MS-001':'4b9a7f50c0bb198c6f5414c5a8459f5d216d34ab521ea94c060ea35cac66f900',
    'MS-002':'165940940a02a187e4463ff467090930038c5af8fc26107bf301e714f599a1da',
    'MS-003':'446e21f212ab200933c4c9a0802e1ff0c410bbd75fca10168746fc49883096db'
  };
}

/** True when `pin` is the placeholder that shipped for this staff ID. */
function isStartingPin_(staffId, pin) {
  const digests = startingPinDigests_();
  const id = String(staffId || '').trim();
  if (!Object.prototype.hasOwnProperty.call(digests,id)) return false;
  const normalized = normalizePin_(pin);
  return isValidPin_(normalized) && sha256Hex_(normalized) === digests[id];
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

/** Anything other than the explicit opt-in keeps the strict behaviour. */
function normalizeDeviceMismatchPolicy_(value) {
  return String(value == null ? '' : value).trim().toUpperCase() === 'REBIND_AT_OFFICE' ? 'REBIND_AT_OFFICE' : 'BLOCK';
}

/**
 * Decides what to do with the device token a browser sent.
 *
 * Outcomes: 'match', 'register' (nothing on file yet), 'rebind' (replace the
 * registration), or 'reject'. `issue` says why a request did not match:
 * 'missing' means this browser sent no token at all (storage cleared, private
 * tab, in-app browser, a different browser), 'mismatch' means it sent the wrong
 * one. 'missing' is what an iPhone that discarded its site data produces, so the
 * two get different guidance.
 *
 * A rebind needs the opt-in policy, a correct PIN (checked by the caller before
 * this runs) and proof the person is physically at an approved location. The
 * location condition is what keeps the device lock useful: someone clocking in
 * for a colleague from elsewhere is still refused.
 */
function resolveDeviceCheck_(storedToken, submittedToken, options) {
  const opts = options || {};
  const stored = String(storedToken || '');
  if (!stored) return {outcome: opts.autoRegister ? 'register' : 'reject', issue: 'unregistered'};
  if (deviceTokenMatches_(stored, submittedToken)) return {outcome: 'match', issue: ''};
  const issue = String(submittedToken || '') ? 'mismatch' : 'missing';
  const canRebind = normalizeDeviceMismatchPolicy_(opts.policy) === 'REBIND_AT_OFFICE' && opts.insideApprovedLocation === true;
  return {outcome: canRebind ? 'rebind' : 'reject', issue: issue};
}

/** Coarse browser label for support logs; the full user agent is already recorded separately. */
function classifyUserAgent_(userAgent) {
  const ua = String(userAgent || '');
  if (!/iPhone|iPad|iPod/.test(ua)) return /Android/.test(ua) ? 'Android' : (ua ? 'Other' : 'Unknown');
  if (/CriOS|FxiOS|EdgiOS/.test(ua)) return 'iOS (Chrome/Firefox/Edge)';
  if (/FBAN|FBAV|Instagram|Line\/|MicroMessenger|GSA\//.test(ua) || !/Safari\//.test(ua)) return 'iOS (in-app browser or home-screen app)';
  return 'iOS Safari';
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
  requireOwnerContext_();
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
