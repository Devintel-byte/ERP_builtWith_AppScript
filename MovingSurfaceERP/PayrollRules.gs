function getPenaltySummary_(staff, attendanceStatus, action, now, settings) {
  const payroll = getStaffPayrollSummary_(staff, settings, now);
  let todayPenalty = 0;
  if (action === 'Check In') {
    if (attendanceStatus === 'Late') todayPenalty = staff.latePenalty || settings.DEFAULT_LATE_PENALTY;
    if (attendanceStatus === 'Half Day') todayPenalty = payroll.dailyRate * (settings.HALF_DAY_PERCENT/100);
    if (attendanceStatus === 'Absent') todayPenalty = payroll.dailyRate;
  }
  return {
    todayPenalty:Math.round(todayPenalty),
    monthLatePenalty:Math.round(payroll.latePenaltyTotal),
    monthHalfDayDeduction:Math.round(payroll.halfDayDeduction),
    monthAbsenceDeduction:Math.round(payroll.absenceDeduction),
    monthTotalDeduction:Math.round(payroll.totalDeduction)
  };
}

function getStaffPayrollSummary_(staff, settings, now) {
  settings = settings || getSettings_();
  now = now || new Date();
  const counts = getCurrentMonthAttendanceCounts_(settings, now, staff.id)[staff.id] || {presentDays:0,lateCount:0,halfDayCount:0};
  return calculateStaffPayrollSummary_(staff, counts, settings);
}

/**
 * Month counts per staff member: check-ins, lateness, overtime, and the working
 * days excused by approved field duty, approved paid leave and public holidays.
 *
 * Every staff member being counted gets an entry, including one with no attendance
 * rows at all (someone on field duty all month), because callers fall back to an
 * empty object for a missing entry and would silently lose their excused days.
 * Pass `onlyStaffId` to count a single person, which keeps a check-in cheap.
 */
function getCurrentMonthAttendanceCounts_(settings, now, onlyStaffId) {
  const counts = scanMonthAttendanceLog_(settings, now);
  const staffIds = onlyStaffId ? [String(onlyStaffId)] : getActiveStaff_().map(staff => String(staff.id));
  const checkIns = Object.create(null);
  Object.keys(counts).forEach(function(id) { checkIns[id] = counts[id].checkInDays || {}; });
  const excused = getExcusedDayCountsForMonth_(settings, now, staffIds, checkIns);
  const monthWorkdays = fieldDutyMonthWorkdayKeys_(Utilities.formatDate(now, settings.TIMEZONE, 'yyyy-MM'),
    parseWorkdaySetting_(settings.WORKDAYS)).length;
  staffIds.forEach(function(id) {
    if (!counts[id]) counts[id] = {presentDays:0,lateCount:0,halfDayCount:0,overtimeDays:0,overtimeDateKeys:{},checkInDays:{}};
    const e = excused[id] || {};
    counts[id].monthWorkdays = monthWorkdays;
    counts[id].fieldDutyDays = Number(e.fieldDutyDays || 0);
    counts[id].leaveDays = Number(e.leaveDays || 0);
    counts[id].holidayDays = Number(e.holidayDays || 0);
  });
  return counts;
}

function scanMonthAttendanceLog_(settings, now) {
  const month = Utilities.formatDate(now,settings.TIMEZONE,'yyyy-MM');
  const counts = Object.create(null);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return counts;

  const dateValues = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  const monthRows = [];
  dateValues.forEach((row,index) => {
    if (formatAttendanceDate_(row[0],settings).slice(0,7) === month) monthRows.push(index+2);
  });
  if (!monthRows.length) return counts;
  const firstRow = Math.min.apply(null,monthRows);
  const lastRow = Math.max.apply(null,monthRows);
  sh.getRange(firstRow,1,lastRow-firstRow+1,14).getValues().forEach(function(row) {
    const staffId = String(row[1] || '');
    const dateKey = formatAttendanceDate_(row[0],settings);
    if (!staffId || dateKey.slice(0,7) !== month) return;
    if (!counts[staffId]) counts[staffId] = {presentDays:0,lateCount:0,halfDayCount:0,overtimeDays:0,overtimeDateKeys:{},checkInDays:{}};
    if (String(row[4]) === 'Check In') {
      counts[staffId].presentDays++;
      counts[staffId].checkInDays[dateKey] = true;
      if (String(row[13]) === 'Late') counts[staffId].lateCount++;
      if (String(row[13]) === 'Half Day') counts[staffId].halfDayCount++;
    }
    if (String(row[4]) === 'Check Out Overtime' && !counts[staffId].overtimeDateKeys[dateKey]) {
      counts[staffId].overtimeDateKeys[dateKey] = true;
      counts[staffId].overtimeDays++;
    }
  });
  return counts;
}

function calculateStaffPayrollSummary_(staff, counts, settings) {
  const requiredDays = Number(staff.workDaysPerMonth || 22);
  const presentDays = Number(counts.presentDays || 0);
  const lateCount = Number(counts.lateCount || 0);
  const halfDayCount = Number(counts.halfDayCount || 0);
  const overtimeDays = Number(counts.overtimeDays || 0);
  // Days that must not count against someone: approved field duty, approved paid
  // leave and public holidays (see computeExcusedDays_). Leave used to be fixed at 0,
  // so approved leave was deducted as absence.
  const fieldDutyDays = Number(counts.fieldDutyDays || 0);
  const leaveDays = Number(counts.leaveDays || 0);
  const holidayDays = Number(counts.holidayDays || 0);
  // Nobody can be expected to work more days than the month has. A staff member's
  // required days are a fixed 22, but a month often has 20 or 21 working days, and
  // without this cap someone excused or present for EVERY working day would still be
  // short by the difference and lose a day's pay. The daily rate stays on the required
  // days, so salaries do not move; only the expectation is capped. When the month's
  // working days are unknown the required days are used, as before.
  const monthWorkdays = Number(counts.monthWorkdays || 0);
  const expectedDays = monthWorkdays > 0 ? Math.min(requiredDays, monthWorkdays) : requiredDays;
  const absentDays = Math.max(0, expectedDays - presentDays - fieldDutyDays - leaveDays - holidayDays);
  const dailyRate = requiredDays > 0 ? Number(staff.monthlySalary || 0) / requiredDays : 0;
  const latePenaltyTotal = lateCount * (staff.latePenalty || settings.DEFAULT_LATE_PENALTY);
  const halfDayDeduction = halfDayCount * dailyRate * (settings.HALF_DAY_PERCENT/100);
  const absenceDeduction = absentDays * dailyRate;
  const totalDeduction = latePenaltyTotal + halfDayDeduction + absenceDeduction;
  const overtimePay = overtimeDays * Number(settings.OVERTIME_PAY_RATE || 0);
  return {
    requiredDays, presentDays, leaveDays, fieldDutyDays, holidayDays, absentDays, lateCount, halfDayCount, overtimeDays, dailyRate,
    latePenaltyTotal, halfDayDeduction, absenceDeduction, totalDeduction,
    overtimePay,
    netPay:Number(staff.monthlySalary || 0)-totalDeduction+overtimePay
  };
}
