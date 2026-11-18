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
  const counts = getCurrentMonthAttendanceCounts_(settings, now)[staff.id] || {presentDays:0,lateCount:0,halfDayCount:0};
  return calculateStaffPayrollSummary_(staff, counts, settings);
}

function getCurrentMonthAttendanceCounts_(settings, now) {
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
    if (!counts[staffId]) counts[staffId] = {presentDays:0,lateCount:0,halfDayCount:0,overtimeDays:0,overtimeDateKeys:{}};
    if (String(row[4]) === 'Check In') {
      counts[staffId].presentDays++;
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
  const leaveDays = 0;
  const absentDays = Math.max(0, requiredDays - presentDays - leaveDays);
  const dailyRate = requiredDays > 0 ? Number(staff.monthlySalary || 0) / requiredDays : 0;
  const latePenaltyTotal = lateCount * (staff.latePenalty || settings.DEFAULT_LATE_PENALTY);
  const halfDayDeduction = halfDayCount * dailyRate * (settings.HALF_DAY_PERCENT/100);
  const absenceDeduction = absentDays * dailyRate;
  const totalDeduction = latePenaltyTotal + halfDayDeduction + absenceDeduction;
  const overtimePay = overtimeDays * Number(settings.OVERTIME_PAY_RATE || 0);
  return {
    requiredDays, presentDays, leaveDays, absentDays, lateCount, halfDayCount, overtimeDays, dailyRate,
    latePenaltyTotal, halfDayDeduction, absenceDeduction, totalDeduction,
    overtimePay,
    netPay:Number(staff.monthlySalary || 0)-totalDeduction+overtimePay
  };
}
