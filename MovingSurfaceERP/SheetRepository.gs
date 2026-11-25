function getSettings_() {
  const raw = getKeyValueSheet_(SETTINGS_SHEET);
  const closeTime = normalizeTimeSetting_(raw.CLOSE_TIME,'17:00','CLOSE_TIME');
  const overtimeGraceMinutes = nonNegativeSetting_(raw.OVERTIME_GRACE_MINUTES,15,'OVERTIME_GRACE_MINUTES');
  if (toMinutes_(closeTime)+overtimeGraceMinutes >= 1440) {
    throw new Error('Settings CLOSE_TIME plus OVERTIME_GRACE_MINUTES must be earlier than midnight.');
  }
  return {
    TIMEZONE: raw.TIMEZONE || 'Africa/Lagos',
    START_TIME: raw.START_TIME || '09:00',
    GRACE_MINUTES: Number(raw.GRACE_MINUTES || 15),
    CLOSE_TIME: closeTime,
    OVERTIME_GRACE_MINUTES: overtimeGraceMinutes,
    OVERTIME_PAY_RATE: nonNegativeSetting_(raw.OVERTIME_PAY_RATE,0,'OVERTIME_PAY_RATE'),
    HALF_DAY_CUTOFF: raw.HALF_DAY_CUTOFF || '10:00',
    ABSENT_CUTOFF: raw.ABSENT_CUTOFF || '12:00',
    REQUIRE_GPS: raw.REQUIRE_GPS || 'YES',
    DEFAULT_RADIUS_METRES: Number(raw.DEFAULT_RADIUS_METRES || 300),
    REQUIRE_PIN: raw.REQUIRE_PIN || 'YES',
    ALLOW_SELF_PIN: raw.ALLOW_SELF_PIN || 'YES',
    REQUIRE_DEVICE_TOKEN: raw.REQUIRE_DEVICE_TOKEN || 'YES',
    AUTO_REGISTER_DEVICE: raw.AUTO_REGISTER_DEVICE || 'YES',
    DEVICE_MISMATCH_POLICY: normalizeDeviceMismatchPolicy_(raw.DEVICE_MISMATCH_POLICY),
    DEFAULT_LATE_PENALTY: Number(raw.DEFAULT_LATE_PENALTY || 1000),
    HALF_DAY_PERCENT: Number(raw.HALF_DAY_PERCENT || 50),
    WORKDAYS: raw.WORKDAYS || 'MON,TUE,WED,THU,FRI'
  };
}

function normalizeTimeSetting_(value,defaultValue,name) {
  if (value === undefined || value === null || value === '') return defaultValue;

  if (value instanceof Date && !isNaN(value.getTime())) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const timezone = spreadsheet.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Etc/UTC';
    return Utilities.formatDate(value,timezone,'HH:mm');
  }

  if (typeof value === 'number' && Number.isFinite(value) && value >= 0) {
    const fractionalDay = value % 1;
    const minutes = Math.round(fractionalDay*1440)%1440;
    return String(Math.floor(minutes/60)).padStart(2,'0')+':'+String(minutes%60).padStart(2,'0');
  }

  const text = String(value).trim();
  const match = text.match(/^(\d{1,2}):([0-5]\d)(?::[0-5]\d)?\s*(AM|PM)?$/i);
  if (!match) throw new Error('Settings '+name+' must be a valid time, such as 17:00 or 5:00 PM.');
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const meridiem = String(match[3] || '').toUpperCase();
  if (meridiem) {
    if (hour < 1 || hour > 12) throw new Error('Settings '+name+' must use a valid 12-hour time.');
    hour = hour%12+(meridiem === 'PM' ? 12 : 0);
  } else if (hour > 23) {
    throw new Error('Settings '+name+' must use a valid 24-hour time.');
  }
  return String(hour).padStart(2,'0')+':'+String(minute).padStart(2,'0');
}

function nonNegativeSetting_(value,defaultValue,name) {
  if (value === undefined || value === null || value === '') return defaultValue;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 0) throw new Error('Settings '+name+' must be a non-negative number.');
  return parsed;
}

function getReqSettings_() {
  const raw = getKeyValueSheet_(REQ_SETTINGS_SHEET);
  return {
    MD_EMAIL: raw.MD_EMAIL || 'joemugbeh@gmail.com',
    ACCOUNTS_EMAIL: raw.ACCOUNTS_EMAIL || 'accounts@movingsurface.com.ng',
    HR_EMAIL: raw.HR_EMAIL || '',
    COMPANY_NAME: raw.COMPANY_NAME || 'Moving Surface Interactive Media Limited',
    ALLOW_STAFF_EMAIL_COPY: raw.ALLOW_STAFF_EMAIL_COPY || 'YES',
    DEFAULT_CURRENCY: raw.DEFAULT_CURRENCY || 'NGN',
    HIGH_VALUE_THRESHOLD: Number(raw.HIGH_VALUE_THRESHOLD || 500000)
  };
}

function getKeyValueSheet_(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r => {
    if (r[0] !== '') out[String(r[0]).trim()] = r[1];
  });
  return out;
}

function getActiveStaff_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,29).getValues()
    .map((r,i) => staffObj_(r,i+2))
    .filter(s => String(s.status).toLowerCase() === 'active');
}

function getAllStaff_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,29).getValues()
    .map((row,index) => staffObj_(row,index+2))
    .filter(staff => staff.id);
}

function getStaffById_(staffId) {
  const staff = getStaffByIdAnyStatus_(staffId);
  return staff && String(staff.status).toLowerCase() === 'active' ? staff : null;
}

function getStaffByIdAnyStatus_(staffId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET);
  if (!sh || sh.getLastRow() < 2 || !String(staffId || '').trim()) return null;
  const match = sh.getRange(2,1,sh.getLastRow()-1,1)
    .createTextFinder(String(staffId))
    .matchCase(true)
    .matchEntireCell(true)
    .findNext();
  if (!match) return null;
  const row = match.getRow();
  return staffObj_(sh.getRange(row,1,1,29).getValues()[0], row);
}

function staffObj_(r,row) {
  return {
    row:row, id:String(r[0] || ''), name:String(r[1] || ''), department:String(r[2] || ''), location:String(r[3] || ''),
    pin:String(r[4] || ''), deviceToken:String(r[5] || ''), companyEmail:String(r[6] || ''), personalEmail:String(r[7] || ''),
    requireGps:String(r[8] || 'YES'), requireDevice:String(r[9] || 'YES'), monthlySalary:Number(r[10] || 0), latePenalty:Number(r[11] || 1000),
    leaveAllowance:r[12] || '', phone:String(r[13] || ''), notes:String(r[14] || ''),
    dateJoined:r[15] || '', employmentType:String(r[16] || ''), status:String(r[17] || ''), exitDate:r[18] || '',
    designation:String(r[19] || ''), manager:String(r[20] || ''), bankName:String(r[21] || ''),
    accountNumber:String(r[22] || ''), accountName:String(r[23] || ''),
    annualLeaveEntitlement:Number(r[24] || 0), sickLeaveEntitlement:Number(r[25] || 0),
    workDaysPerMonth:Number(r[26] || 22), qrCode:String(r[27] || ''), canSubmitRequisition:String(r[28] || 'YES')
  };
}


function generateNextStaffId_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET);
  const ids = sh.getLastRow() < 2 ? [] : sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat().map(String);
  let max = 0;
  ids.forEach(id => { const m = id.match(/^MS-(\d+)$/i); if (m) max = Math.max(max, Number(m[1])); });
  return 'MS-' + String(max + 1).padStart(3,'0');
}

function getAttendanceStateForDay_(dateKey,staffId,timezone) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return {checkedIn:false,checkedOut:false};
  return attendanceStateFromRows_(sh.getRange(2,1,sh.getLastRow()-1,5).getValues(),dateKey,staffId,timezone);
}

function attendanceStateFromRows_(rows,dateKey,staffId,timezone) {
  const state = {checkedIn:false,checkedOut:false};
  rows.forEach(function(row) {
    if (attendanceDateKey_(row[0],timezone) !== String(dateKey) ||
        String(row[1] || '').trim().toUpperCase() !== String(staffId || '').trim().toUpperCase()) return;
    const action = normalizeAttendanceAction_(row[4]);
    if (action === 'checkin') state.checkedIn = true;
    if (action === 'checkout' || action === 'checkoutovertime') state.checkedOut = true;
  });
  return state;
}

function attendanceDateKey_(value,timezone) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value,timezone || Session.getScriptTimeZone() || 'Africa/Lagos','yyyy-MM-dd');
  }
  const text = String(value == null ? '' : value).trim();
  const isoDate = text.match(/^(\d{4}-\d{2}-\d{2})(?:$|[T\s])/);
  return isoDate ? isoDate[1] : text;
}

function normalizeAttendanceAction_(action) {
  return String(action || '').trim().toLowerCase().replace(/[^a-z]/g,'');
}

function attendanceCheckoutType_(action,storedType) {
  const normalizedType=String(storedType || '').trim().toLowerCase();
  if (normalizedType==='regular' || normalizedType==='overtime') {
    return normalizedType==='overtime'?'Overtime':'Regular';
  }
  const normalizedAction=normalizeAttendanceAction_(action);
  if (normalizedAction==='checkout') return 'Regular';
  if (normalizedAction==='checkoutovertime') return 'Overtime';
  return '';
}

function alreadySubmitted_(dateKey,staffId,action,timezone) {
  const state = getAttendanceStateForDay_(dateKey,staffId,timezone);
  const normalizedAction = normalizeAttendanceAction_(action);
  return normalizedAction === 'checkin' ? state.checkedIn :
    (normalizedAction === 'checkout' || normalizedAction === 'checkoutovertime') && state.checkedOut;
}

function getAttendanceStatus_(now, action, settings) {
  // IMPORTANT: a staff who signs in should not be marked Absent.
  // Absent should only be used in reports for people who did not sign in at all.
  if (action !== 'Check In') return 'N/A';
  const t = toMinutes_(Utilities.formatDate(now, settings.TIMEZONE, 'HH:mm'));
  const start = toMinutes_(settings.START_TIME);
  const lateLimit = start + Number(settings.GRACE_MINUTES);
  const halfDay = toMinutes_(settings.HALF_DAY_CUTOFF);
  if (t <= lateLimit) return 'On Time';
  if (t < halfDay) return 'Late';
  return 'Half Day';
}

function getMatchedLocation_(lat,lng,settings) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOCATIONS_SHEET);
  let best = {name:'No approved location', distance:999999999, radius:settings.DEFAULT_RADIUS_METRES, allowed:false};
  if (!sh || sh.getLastRow() < 2 || isNaN(lat) || isNaN(lng)) return best;
  sh.getRange(2,1,sh.getLastRow()-1,5).getValues().forEach(r => {
    if (String(r[4]).toUpperCase() !== 'YES') return;
    const radius = Number(r[3] || settings.DEFAULT_RADIUS_METRES);
    const dist = getDistanceMetres_(lat,lng,Number(r[1]),Number(r[2]));
    if (dist < best.distance) best = {name:String(r[0] || 'Approved Location'), distance:dist, radius:radius, allowed:dist <= radius};
  });
  return best;
}
