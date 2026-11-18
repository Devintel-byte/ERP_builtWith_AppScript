function syncAccountantViewRows() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTANT_VIEW_SHEET);
  const staff = getActiveStaff_();
  const settings = getSettings_();
  const counts = getCurrentMonthAttendanceCounts_(settings, new Date());
  const rows = staff.map(s => {
    const payroll = calculateStaffPayrollSummary_(s, counts[s.id] || {}, settings);
    return [
      Utilities.formatDate(new Date(), settings.TIMEZONE, 'yyyy-MM'),
      s.id, s.name, s.department, s.designation, s.employmentType, s.status,
      s.monthlySalary, payroll.requiredDays, payroll.presentDays, payroll.leaveDays, payroll.absentDays,
      payroll.lateCount, payroll.halfDayCount, payroll.dailyRate, payroll.latePenaltyTotal,
      payroll.halfDayDeduction, payroll.absenceDeduction, payroll.totalDeduction,
      payroll.overtimeDays, payroll.overtimePay, payroll.netPay
    ];
  });
  const oldLastRow = sh.getLastRow();
  if (rows.length) sh.getRange(2,1,rows.length,22).setValues(rows);
  if (oldLastRow > rows.length + 1) {
    sh.getRange(rows.length + 2,1,oldLastRow - rows.length - 1,22).clearContent();
  }
}

function getPayrollRows_(staffList) {
  const settings = getSettings_();
  const month = Utilities.formatDate(new Date(), settings.TIMEZONE, 'yyyy-MM');
  const counts = getCurrentMonthAttendanceCounts_(settings, new Date());
  return (staffList || getActiveStaff_()).map(function(staff) {
    const payroll = calculateStaffPayrollSummary_(staff, counts[staff.id] || {}, settings);
    return {
      month:month,staffId:staff.id,name:staff.name,salary:staff.monthlySalary,
      required:payroll.requiredDays,present:payroll.presentDays,absent:payroll.absentDays,
      late:payroll.lateCount,halfDay:payroll.halfDayCount,overtimeDays:payroll.overtimeDays,overtimePay:payroll.overtimePay,
      deduction:payroll.totalDeduction,net:payroll.netPay
    };
  });
}

function getLiveTodayAttendance_(staffList) {
  const settings = getSettings_();
  const now = new Date();
  const todayKey = Utilities.formatDate(now, settings.TIMEZONE, 'yyyy-MM-dd');
  const currentTime = Utilities.formatDate(now, settings.TIMEZONE, 'hh:mm:ss a');
  staffList = staffList || getActiveStaff_();
  const map = {};
  staffList.forEach(s => {
    map[s.id] = {staffId:s.id, name:s.name, department:s.department, workLocation:s.location, checkIn:'', checkOut:'', overtimeCheckout:false, rawCheckIn:0, rawCheckOut:0, status:'Not Signed In', matchedLocation:'', gpsStatus:''};
  });

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (sh && sh.getLastRow() >= 2) {
    const matches = sh.getRange(2,1,sh.getLastRow()-1,1)
      .createTextFinder(todayKey)
      .matchEntireCell(true)
      .findAll();
    let rows = [];
    if (matches.length) {
      const matchRows = matches.map(cell => cell.getRow());
      const firstRow = Math.min.apply(null, matchRows);
      const lastRow = Math.max.apply(null, matchRows);
      rows = sh.getRange(firstRow,1,lastRow-firstRow+1,14).getValues();
    }
    rows.forEach(r => {
      const staffId = String(r[1] || '');
      if (!map[staffId]) return;
      const ts = r[5] ? new Date(r[5]) : null;
      if (!ts || isNaN(ts.getTime())) return;
      if (Utilities.formatDate(ts, settings.TIMEZONE, 'yyyy-MM-dd') !== todayKey) return;
      const action = String(r[4] || '');
      const raw = ts.getTime();
      const timeText = Utilities.formatDate(ts, settings.TIMEZONE, 'hh:mm:ss a');
      if (action === 'Check In' && (!map[staffId].rawCheckIn || raw < map[staffId].rawCheckIn)) {
        map[staffId].checkIn = timeText;
        map[staffId].rawCheckIn = raw;
        map[staffId].status = String(r[13] || 'Present');
        map[staffId].matchedLocation = String(r[10] || '');
        map[staffId].gpsStatus = String(r[11] || '');
      }
      if ((action === 'Check Out' || action === 'Check Out Overtime') && (!map[staffId].rawCheckOut || raw > map[staffId].rawCheckOut)) {
        map[staffId].checkOut = timeText;
        map[staffId].rawCheckOut = raw;
        map[staffId].overtimeCheckout = action === 'Check Out Overtime';
      }
    });
  }

  const all = Object.values(map);
  const signedIn = all.filter(r => r.checkIn && !r.checkOut).sort((a,b)=>a.rawCheckIn-b.rawCheckIn);
  const checkedOut = all.filter(r => r.checkIn && r.checkOut).sort((a,b)=>a.rawCheckOut-b.rawCheckOut);
  const notSignedIn = all.filter(r => !r.checkIn).sort((a,b)=>String(a.name).localeCompare(String(b.name)));

  return {
    date:todayKey, currentTime:currentTime, totalStaff:all.length,
    signedInCount:signedIn.length, checkedOutCount:checkedOut.length, overtimeCheckoutCount:all.filter(r=>r.overtimeCheckout).length, notSignedInCount:notSignedIn.length,
    lateCount:all.filter(r=>r.status==='Late').length, halfDayCount:all.filter(r=>r.status==='Half Day').length,
    signedIn:signedIn, checkedOut:checkedOut, notSignedIn:notSignedIn
  };
}

/*************** HELPERS ***************/
