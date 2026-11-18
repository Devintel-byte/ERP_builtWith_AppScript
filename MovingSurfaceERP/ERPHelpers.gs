function erpDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Africa/Lagos', 'yyyy-MM-dd');
  }
  return String(v);
}

function erpRequireAdmin_(pin,page) {
  const user = getAdminUser_(pin);
  if (!user.valid) throw new Error('Invalid admin PIN.');
  if (page) requireAdminPageWrite_(user,page);
  return user;
}

function erpNextId_(sheetName, prefix, idCol) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const col = idCol || 1;
  const lastRow = sh.getLastRow();
  let max = 0;
  if (lastRow >= 2) {
    sh.getRange(2,col,lastRow-1,1).getValues().flat().forEach(v => {
      const m = String(v||'').match(/(\d+)$/);
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return prefix + '-' + String(max + 1).padStart(4,'0');
}

function erpGetStaffByAnyId_(staffId) {
  const staff = getStaffById_(staffId);
  if (!staff) throw new Error('Staff not found.');
  return staff;
}

/** Parses a money input, rejecting negatives and non-numbers. */
function erpNonNegativeAmount_(value, label) {
  if (value === undefined || value === null || String(value).trim() === '') return 0;
  const parsed = Number(String(value).replace(/,/g, '').trim());
  if (!isFinite(parsed)) throw new Error(label + ' must be a number.');
  if (parsed < 0) throw new Error(label + ' cannot be negative.');
  return Math.round(parsed * 100) / 100;
}

/** Parses a percentage input and clamps it to 0-100. */
function erpPercent_(value, label) {
  if (value === undefined || value === null || String(value).trim() === '') return 0;
  const parsed = Number(String(value).replace(/%/g, '').trim());
  if (!isFinite(parsed)) throw new Error(label + ' must be a number.');
  return Math.min(100, Math.max(0, Math.round(parsed)));
}

/**
 * Normalizes a yyyy-MM-dd input to a date-only string, rejecting anything the
 * reminder scheduler could not compare reliably.
 */
function erpDateInput_(value, label) {
  const supplied = String(value == null ? '' : value).trim();
  if (!supplied) return '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(supplied)) throw new Error(label + ' must be a date in yyyy-mm-dd format.');
  const parsed = new Date(supplied + 'T00:00:00');
  if (isNaN(parsed.getTime())) throw new Error(label + ' is not a valid date.');
  return supplied;
}

/** Today's date in the configured timezone, as yyyy-MM-dd. */
function erpToday_() {
  return Utilities.formatDate(new Date(), erpTimezone_(), 'yyyy-MM-dd');
}

function erpTimezone_() {
  return Session.getScriptTimeZone() || 'Africa/Lagos';
}

/** Whole days from the first date string to the second. Both are yyyy-MM-dd. */
function erpDaysBetween_(fromDateKey, toDateKey) {
  const from = new Date(String(fromDateKey) + 'T00:00:00');
  const to = new Date(String(toDateKey) + 'T00:00:00');
  if (isNaN(from.getTime()) || isNaN(to.getTime())) return null;
  return Math.round((to.getTime() - from.getTime()) / 86400000);
}

/** Adds days to a yyyy-MM-dd date key. */
function erpAddDays_(dateKey, days) {
  const base = new Date(String(dateKey) + 'T00:00:00');
  if (isNaN(base.getTime())) return '';
  base.setDate(base.getDate() + Number(days || 0));
  return Utilities.formatDate(base, erpTimezone_(), 'yyyy-MM-dd');
}
