function fixOldAttendanceRowsMarkedAbsent() {
  requireOwnerContext_();
  // Optional one-time cleanup: changes old successful Check In rows with status Absent to Half Day.
  // This does not affect no-show absence calculations; it only fixes rows where someone actually signed in.
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  const last = sh.getLastRow();
  const actions = sh.getRange(2,5,last-1,1).getValues();
  const statuses = sh.getRange(2,14,last-1,1).getValues();
  let count = 0;
  for (let i=0;i<actions.length;i++) {
    if (String(actions[i][0]) === 'Check In' && String(statuses[i][0]) === 'Absent') {
      statuses[i][0] = 'Half Day';
      count++;
    }
  }
  sh.getRange(2,14,last-1,1).setValues(statuses);
  Logger.log('Fixed old attendance rows: ' + count);
}
