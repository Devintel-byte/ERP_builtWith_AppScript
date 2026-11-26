/**
 * Field duty: staff working away from the office on installation or event jobs,
 * for days or weeks, sometimes outside the country.
 *
 * An APPROVED field duty does three things:
 *  - its working days are excused in payroll: counted as worked, never as absent,
 *    and paid in full (see computeExcusedDays_);
 *  - the person is shown as "on field duty" instead of "not signed in";
 *  - checking in during it is optional, needs no office location, and is never
 *    late, half-day or absent, because those rules are Lagos office hours.
 *
 * A PENDING, REJECTED or CANCELLED duty excuses nothing.
 *
 * Who approves: Admin and Super Admin always; the project's manager when the duty
 * names a project; nobody approves their own duty except a Super Admin, so a
 * workbook with a single Super Admin cannot deadlock. See fieldDutyApprovalDecision_.
 */

const FD_COL = {
  dutyId: 1, staffId: 2, staffName: 3, startDate: 4, endDate: 5, site: 6, projectId: 7, projectName: 8,
  reason: 9, status: 10, requestedBy: 11, decidedBy: 12, decidedAt: 13, decisionNote: 14,
  createdAt: 15, lastUpdated: 16
};
const FD_WIDTH = 16;
const FIELD_DUTY_MAX_DAYS = 366;

/*************** PURE RULES (no sheet access, so they can be tested directly) ***************/

function normalizePersonName_(name) {
  return String(name == null ? '' : name).toLowerCase().replace(/[^a-z0-9\s]/g, '').replace(/\s+/g, ' ').trim();
}

/** Working weekdays from a WORKDAYS setting such as "MON,TUE,WED", as getUTCDay() numbers. */
function parseWorkdaySetting_(text) {
  const names = {SUN: 0, MON: 1, TUE: 2, WED: 3, THU: 4, FRI: 5, SAT: 6};
  const days = [];
  String(text == null ? '' : text).split(',').forEach(function (part) {
    const key = part.trim().slice(0, 3).toUpperCase();
    if (Object.prototype.hasOwnProperty.call(names, key) && days.indexOf(names[key]) === -1) days.push(names[key]);
  });
  return days.length ? days : [1, 2, 3, 4, 5];
}

/** Every working day in a yyyy-MM month, as yyyy-MM-dd keys. Plain calendar arithmetic, no timezone involved. */
function fieldDutyMonthWorkdayKeys_(month, workdays) {
  const match = /^(\d{4})-(\d{2})$/.exec(String(month || ''));
  if (!match) return [];
  const year = Number(match[1]), monthIndex = Number(match[2]) - 1;
  const keys = [];
  for (let day = 1; day <= 31; day++) {
    const date = new Date(Date.UTC(year, monthIndex, day));
    if (date.getUTCMonth() !== monthIndex) break;
    if (workdays.indexOf(date.getUTCDay()) !== -1) {
      keys.push(match[1] + '-' + match[2] + '-' + ('0' + day).slice(-2));
    }
  }
  return keys;
}

/** A yyyy-MM-dd key from a sheet cell: a Date (read in `timezone`), a key, or ''. */
function fieldDutyDateKey_(value, timezone) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? '' : Utilities.formatDate(value, timezone || 'Africa/Lagos', 'yyyy-MM-dd');
  }
  const text = String(value == null ? '' : value).trim();
  return /^\d{4}-\d{2}-\d{2}/.test(text) ? text.slice(0, 10) : '';
}

function fieldDutyRangesOverlap_(aStart, aEnd, bStart, bEnd) {
  return aStart <= bEnd && bStart <= aEnd;
}

/**
 * Working days in a month that should not count against someone, per staff member.
 *
 * Each working day is classified once, in this order of precedence:
 *   1. the person checked in      -> present; nothing is excused
 *   2. an approved field duty     -> fieldDutyDays
 *   3. approved paid leave        -> leaveDays
 *   4. a public holiday           -> holidayDays
 *
 * Leave is counted in WORKING days, not the calendar days stored on the leave
 * request, so a leave that spans a weekend does not excuse the weekend twice.
 * Ranges are clipped to the month, and every staff id supplied gets an entry,
 * including people with no attendance rows at all (someone on duty all month).
 *
 * input: {month, workdays:[0-6], holidays:[keys], duties:[{staffId,start,end}],
 *         leaves:[{staffId,start,end}], checkIns:{staffId:{dateKey:true}}, staffIds:[...]}
 */
function computeExcusedDays_(input) {
  const days = fieldDutyMonthWorkdayKeys_(input.month, input.workdays || [1, 2, 3, 4, 5]);
  const holidays = {};
  (input.holidays || []).forEach(function (key) { holidays[key] = true; });
  const result = {};
  (input.staffIds || []).forEach(function (staffId) {
    const id = String(staffId);
    const counts = {fieldDutyDays: 0, leaveDays: 0, holidayDays: 0};
    const present = (input.checkIns && input.checkIns[id]) || {};
    const covers = function (ranges, key) {
      return (ranges || []).some(function (r) { return String(r.staffId) === id && r.start <= key && key <= r.end; });
    };
    days.forEach(function (key) {
      if (present[key]) return;
      if (covers(input.duties, key)) counts.fieldDutyDays++;
      else if (covers(input.leaves, key)) counts.leaveDays++;
      else if (holidays[key]) counts.holidayDays++;
    });
    result[id] = counts;
  });
  return result;
}

/**
 * May this approver decide this duty?
 *
 * approver: {staffId, isAdmin, isSuperAdmin}
 * duty:     {staffId, projectId}
 * context:  {managerStaffIds} the active staff whose name equals the project's manager
 *
 * Projects store the manager as free text, so the link is a normalised name match.
 * If it matches nobody, or more than one person, nobody gets approval through it:
 * the check fails safe rather than guessing.
 */
function fieldDutyApprovalDecision_(approver, duty, context) {
  const ownDuty = String(duty.staffId || '') !== '' && String(duty.staffId) === String(approver.staffId || '');
  if (ownDuty && !approver.isSuperAdmin) {
    return {allowed: false, reason: 'You cannot approve your own field duty. Another Admin or Super Admin must.'};
  }
  if (approver.isAdmin || approver.isSuperAdmin) return {allowed: true, reason: ''};
  const managers = (context && context.managerStaffIds) || [];
  if (duty.projectId && managers.length === 1 && managers[0] === String(approver.staffId || '')) {
    return {allowed: true, reason: ''};
  }
  if (duty.projectId && managers.length > 1) {
    return {allowed: false, reason: 'More than one staff member matches this project\'s manager name, so it cannot be approved by the project manager. An Admin or Super Admin must.'};
  }
  return {allowed: false, reason: 'Only an Admin, Super Admin or the project\'s manager can approve field duty.'};
}

/*************** SHEET ACCESS ***************/

function fieldDutySheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FIELD_DUTY_SHEET);
  if (!sheet) throw new Error('Field Duty sheet is missing. Run Setup / Repair ERP Upgrade first.');
  return sheet;
}

function fieldDutyRowToObject_(row, sheetRow, today) {
  const tz = erpTimezone_();
  const start = fieldDutyDateKey_(row[FD_COL.startDate - 1], tz);
  const end = fieldDutyDateKey_(row[FD_COL.endDate - 1], tz);
  const status = String(row[FD_COL.status - 1] || '');
  const approved = status === 'Approved';
  return {
    row: sheetRow,
    dutyId: String(row[FD_COL.dutyId - 1] || ''),
    staffId: String(row[FD_COL.staffId - 1] || ''),
    staffName: String(row[FD_COL.staffName - 1] || ''),
    startDate: start, endDate: end,
    site: String(row[FD_COL.site - 1] || ''),
    projectId: String(row[FD_COL.projectId - 1] || ''),
    projectName: String(row[FD_COL.projectName - 1] || ''),
    reason: String(row[FD_COL.reason - 1] || ''),
    status: status,
    requestedBy: String(row[FD_COL.requestedBy - 1] || ''),
    decidedBy: String(row[FD_COL.decidedBy - 1] || ''),
    decidedAt: erpDate_(row[FD_COL.decidedAt - 1]),
    decisionNote: String(row[FD_COL.decisionNote - 1] || ''),
    phase: !start || !end ? '' : today < start ? 'Upcoming' : today > end ? 'Ended' : 'Active',
    isApproved: approved
  };
}

function getFieldDutyRecords_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FIELD_DUTY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const today = erpToday_();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, FD_WIDTH).getValues()
    .map(function (row, index) { return fieldDutyRowToObject_(row, index + 2, today); })
    .filter(function (record) { return record.dutyId; });
}

function findFieldDuty_(dutyId) {
  const wanted = String(dutyId || '').trim();
  if (!wanted) return null;
  const matches = getFieldDutyRecords_().filter(function (record) { return record.dutyId === wanted; });
  return matches.length ? matches[0] : null;
}

/** The approved duty covering a staff member on a given day, or null. */
function getApprovedFieldDutyOn_(staffId, dateKey) {
  const id = String(staffId || '');
  const matches = getFieldDutyRecords_().filter(function (r) {
    return r.isApproved && r.staffId === id && r.startDate <= dateKey && dateKey <= r.endDate;
  });
  return matches.length ? matches[0] : null;
}

/** {staffId: duty} for everyone on approved duty on a given day. */
function getApprovedFieldDutyMapOn_(dateKey) {
  const map = Object.create(null);
  getFieldDutyRecords_().forEach(function (r) {
    if (r.isApproved && r.startDate <= dateKey && dateKey <= r.endDate) map[r.staffId] = r;
  });
  return map;
}

/** Holiday date keys for a month. A holiday is on unless its Active cell says NO. */
function getMonthHolidayKeys_(month, timezone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOLIDAYS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const keys = [];
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function (row) {
    const key = fieldDutyDateKey_(row[0], timezone);
    const off = ['NO', 'N', 'FALSE', 'INACTIVE'].indexOf(String(row[2] == null ? '' : row[2]).trim().toUpperCase()) !== -1;
    if (key && key.slice(0, 7) === month && !off) keys.push(key);
  });
  return keys;
}

/** Approved, paid leave ranges. Unpaid leave stays an absence deduction, so it is not excused. */
function getApprovedPaidLeaveRanges_(timezone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leave Requests');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const ranges = [];
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues().forEach(function (r) {
    if (String(r[10] || '') !== 'Approved') return;
    if (String(r[5] || '').trim().toLowerCase() === 'unpaid') return;
    const start = fieldDutyDateKey_(r[6], timezone), end = fieldDutyDateKey_(r[7], timezone);
    if (start && end) ranges.push({staffId: String(r[2] || ''), start: start, end: end});
  });
  return ranges;
}

/** Excused days for this month, for the given staff, reading each source sheet once. */
function getExcusedDayCountsForMonth_(settings, now, staffIds, checkIns) {
  const tz = settings.TIMEZONE;
  const month = Utilities.formatDate(now, tz, 'yyyy-MM');
  return computeExcusedDays_({
    month: month,
    workdays: parseWorkdaySetting_(settings.WORKDAYS),
    holidays: getMonthHolidayKeys_(month, tz),
    duties: getFieldDutyRecords_().filter(function (r) { return r.isApproved; })
      .map(function (r) { return {staffId: r.staffId, start: r.startDate, end: r.endDate}; }),
    leaves: getApprovedPaidLeaveRanges_(tz),
    checkIns: checkIns || {},
    staffIds: staffIds
  });
}

/*************** DASHBOARD / PORTAL ***************/

/** What the portal needs to render a duty: the record plus what THIS viewer may do with it. */
function fieldDutyRecordsForUser_(user) {
  const staffList = getActiveStaff_();
  const projects = getERPProjects_();
  const today = erpToday_();
  const approver = fieldDutyApproverFromUser_(user);
  return getFieldDutyRecords_().map(function (record) {
    const decision = fieldDutyApprovalDecision_(approver, record, fieldDutyApprovalContext_(record.projectId, staffList, projects));
    const pending = record.status === 'Pending Approval';
    const started = record.startDate && record.startDate <= today;
    const mine = record.requestedBy === user.name;
    return Object.assign({}, record, {
      canApprove: decision.allowed && pending,
      canChangeEnd: decision.allowed && (pending || record.isApproved) && record.phase !== 'Ended',
      canCancel: (pending && (decision.allowed || mine)) || (record.isApproved && !started && (decision.allowed || mine)),
      approvalNote: decision.allowed ? '' : decision.reason
    });
  }).sort(function (a, b) { return a.startDate < b.startDate ? 1 : a.startDate > b.startDate ? -1 : 0; });
}

function fieldDutyApproverFromUser_(user) {
  const role = normalizeAdminRole_(user && user.role);
  return {
    staffId: String((user && user.staffId) || ''),
    isAdmin: Boolean(user && user.access && user.access.canManageUsers),
    isSuperAdmin: role === 'superadmin'
  };
}

function fieldDutyApprovalContext_(projectId, staffList, projects) {
  if (!projectId) return {managerStaffIds: []};
  const project = (projects || []).filter(function (p) { return p.projectId === projectId; })[0];
  const managerName = project ? normalizePersonName_(project.manager) : '';
  if (!managerName) return {managerStaffIds: []};
  return {
    managerStaffIds: (staffList || [])
      .filter(function (s) { return normalizePersonName_(s.name) === managerName; })
      .map(function (s) { return String(s.id); })
  };
}

/** Project choices for the duty form, independent of whether the viewer may open the Projects page. */
function fieldDutyProjectOptions_(projects) {
  return (projects || []).filter(function (p) { return p.projectId; })
    .map(function (p) { return {projectId: p.projectId, label: p.projectName + ' - ' + p.projectId, manager: p.manager}; });
}

/*************** WRITE API ***************/

function fieldDutyAuthorize_(user, duty) {
  return fieldDutyApprovalDecision_(fieldDutyApproverFromUser_(user), duty,
    fieldDutyApprovalContext_(duty.projectId, getActiveStaff_(), getERPProjects_()));
}

function erpCreateFieldDuty(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'fieldduty');
    const staff = getStaffById_(p.staffId);
    if (!staff) return {success: false, message: 'Select an active staff member.'};
    const start = erpDateInput_(p.startDate, 'Start date');
    const end = erpDateInput_(p.endDate, 'End date');
    if (!start || !end) return {success: false, message: 'Start and end dates are required.'};
    const length = erpDaysBetween_(start, end);
    if (length === null || length < 0) return {success: false, message: 'The end date cannot be before the start date.'};
    if (length + 1 > FIELD_DUTY_MAX_DAYS) return {success: false, message: 'A single field duty can cover at most ' + FIELD_DUTY_MAX_DAYS + ' days. Split a longer assignment.'};
    const site = String(p.site || '').trim();
    if (!site) return {success: false, message: 'Enter the site or city where the work is.'};
    const reason = String(p.reason || '').trim();
    if (!reason) return {success: false, message: 'Enter the reason for the field duty, such as the job it is for.'};

    let project = null;
    const projectId = String(p.projectId || '').trim();
    if (projectId) {
      project = getERPProjects_().filter(function (item) { return item.projectId === projectId; })[0];
      if (!project) return {success: false, message: 'That project no longer exists.'};
    }

    const clash = getFieldDutyRecords_().filter(function (r) {
      return r.staffId === staff.id && ['Pending Approval', 'Approved'].indexOf(r.status) !== -1 &&
        fieldDutyRangesOverlap_(start, end, r.startDate, r.endDate);
    })[0];
    if (clash) {
      return {success: false, message: staff.name + ' already has field duty ' + clash.dutyId + ' (' + clash.startDate + ' to ' + clash.endDate + ') overlapping these dates.'};
    }

    const backdated = start < erpToday_();
    const duty = {staffId: staff.id, projectId: projectId};
    const decision = fieldDutyAuthorize_(user, duty);
    const status = decision.allowed ? 'Approved' : 'Pending Approval';
    const now = new Date();
    const sheet = fieldDutySheet_();
    const id = erpNextId_(FIELD_DUTY_SHEET, 'DUTY', FD_COL.dutyId);
    const row = new Array(FD_WIDTH).fill('');
    row[FD_COL.dutyId - 1] = id;
    row[FD_COL.staffId - 1] = staff.id;
    row[FD_COL.staffName - 1] = staff.name;
    row[FD_COL.startDate - 1] = start;
    row[FD_COL.endDate - 1] = end;
    row[FD_COL.site - 1] = site;
    row[FD_COL.projectId - 1] = projectId;
    row[FD_COL.projectName - 1] = project ? project.projectName : '';
    row[FD_COL.reason - 1] = reason.slice(0, 500);
    row[FD_COL.status - 1] = status;
    row[FD_COL.requestedBy - 1] = user.name;
    row[FD_COL.decidedBy - 1] = decision.allowed ? user.name : '';
    row[FD_COL.decidedAt - 1] = decision.allowed ? now : '';
    row[FD_COL.decisionNote - 1] = decision.allowed ? 'Approved on creation by ' + user.name + '.' : '';
    row[FD_COL.createdAt - 1] = now;
    row[FD_COL.lastUpdated - 1] = now;
    const target = sheet.getLastRow() + 1;
    // Dates stay text so the sheet never reinterprets them as locale dates.
    sheet.getRange(target, FD_COL.startDate, 1, 2).setNumberFormat('@');
    sheet.getRange(target, 1, 1, FD_WIDTH).setValues([row]);

    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'CREATE', 'Field Duty', id,
      staff.id + ' ' + start + ' to ' + end + ' at ' + site + '; ' + status + (backdated ? '; BACKDATED entry' : ''));
    return {
      success: true, dutyId: id, status: status,
      message: 'Field duty ' + id + ' recorded for ' + staff.name + (status === 'Approved'
        ? ' and approved. Those working days count as present, with full pay.'
        : '. It is awaiting approval and excuses nothing until approved.' + (decision.reason ? ' ' + decision.reason : '')) +
        (backdated ? ' It starts in the past, so the days already gone are excused too once approved.' : '')
    };
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

function erpDecideFieldDuty(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'fieldduty');
    const duty = findFieldDuty_(p.dutyId);
    if (!duty) return {success: false, message: 'Field duty not found.'};
    if (duty.status !== 'Pending Approval') return {success: false, message: 'This field duty is not awaiting approval.'};
    if (p.decision !== 'Approved' && p.decision !== 'Rejected') return {success: false, message: 'Choose Approved or Rejected.'};
    const decision = fieldDutyAuthorize_(user, duty);
    if (!decision.allowed) return {success: false, message: decision.reason};
    const note = String(p.note || '').trim().slice(0, 300);
    if (p.decision === 'Rejected' && !note) return {success: false, message: 'Give a reason when rejecting, so the requester knows why.'};
    const sheet = fieldDutySheet_();
    sheet.getRange(duty.row, FD_COL.status).setValue(p.decision);
    sheet.getRange(duty.row, FD_COL.decidedBy).setValue(user.name);
    sheet.getRange(duty.row, FD_COL.decidedAt).setValue(new Date());
    sheet.getRange(duty.row, FD_COL.decisionNote).setValue(note);
    sheet.getRange(duty.row, FD_COL.lastUpdated).setValue(new Date());
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Field Duty', duty.dutyId, 'Status changed to ' + p.decision);
    return {success: true, message: 'Field duty ' + duty.dutyId + ' ' + p.decision.toLowerCase() + '.'};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/** Change the end date of a pending or approved duty: return early, or extend. Approvers only. */
function erpChangeFieldDutyEnd(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'fieldduty');
    const duty = findFieldDuty_(p.dutyId);
    if (!duty) return {success: false, message: 'Field duty not found.'};
    if (['Pending Approval', 'Approved'].indexOf(duty.status) === -1) return {success: false, message: 'Only a pending or approved field duty can be changed.'};
    const decision = fieldDutyAuthorize_(user, duty);
    if (!decision.allowed) return {success: false, message: decision.reason};
    const end = erpDateInput_(p.endDate, 'New end date');
    if (!end) return {success: false, message: 'Enter the new end date.'};
    const length = erpDaysBetween_(duty.startDate, end);
    if (length === null || length < 0) return {success: false, message: 'The end date cannot be before the start date (' + duty.startDate + ').'};
    if (length + 1 > FIELD_DUTY_MAX_DAYS) return {success: false, message: 'A single field duty can cover at most ' + FIELD_DUTY_MAX_DAYS + ' days.'};
    if (end === duty.endDate) return {success: false, message: 'That is already the end date.'};
    const clash = getFieldDutyRecords_().filter(function (r) {
      return r.dutyId !== duty.dutyId && r.staffId === duty.staffId && ['Pending Approval', 'Approved'].indexOf(r.status) !== -1 &&
        fieldDutyRangesOverlap_(duty.startDate, end, r.startDate, r.endDate);
    })[0];
    if (clash) return {success: false, message: 'That would overlap ' + clash.dutyId + ' (' + clash.startDate + ' to ' + clash.endDate + ').'};
    const sheet = fieldDutySheet_();
    sheet.getRange(duty.row, FD_COL.endDate).setNumberFormat('@').setValue(end);
    sheet.getRange(duty.row, FD_COL.lastUpdated).setValue(new Date());
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Field Duty', duty.dutyId,
      'End date changed from ' + duty.endDate + ' to ' + end + (String(p.note || '').trim() ? '; ' + String(p.note).trim().slice(0, 200) : ''));
    return {success: true, message: 'Field duty ' + duty.dutyId + ' now ends on ' + end + '.'};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cancel a duty. A pending one, or an approved one that has not started, can be
 * cancelled by whoever requested it or by an approver. Once an approved duty has
 * started it cannot be cancelled, only shortened: cancelling would un-excuse days
 * already worked and possibly already paid.
 */
function erpCancelFieldDuty(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'fieldduty');
    const duty = findFieldDuty_(p.dutyId);
    if (!duty) return {success: false, message: 'Field duty not found.'};
    if (['Pending Approval', 'Approved'].indexOf(duty.status) === -1) return {success: false, message: 'This field duty is already ' + duty.status.toLowerCase() + '.'};
    const decision = fieldDutyAuthorize_(user, duty);
    if (!decision.allowed && duty.requestedBy !== user.name) {
      return {success: false, message: 'Only the person who requested this, or an approver, can cancel it.'};
    }
    if (duty.isApproved && duty.startDate <= erpToday_()) {
      return {success: false, message: 'This field duty has already started, so it cannot be cancelled. Change its end date instead; days already worked stay excused.'};
    }
    const sheet = fieldDutySheet_();
    sheet.getRange(duty.row, FD_COL.status).setValue('Cancelled');
    sheet.getRange(duty.row, FD_COL.decisionNote).setValue('Cancelled by ' + user.name + (String(p.note || '').trim() ? ': ' + String(p.note).trim().slice(0, 200) : '.'));
    sheet.getRange(duty.row, FD_COL.lastUpdated).setValue(new Date());
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Field Duty', duty.dutyId, 'Cancelled');
    return {success: true, message: 'Field duty ' + duty.dutyId + ' cancelled.'};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}
