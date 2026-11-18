function setupMovingSurfaceSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEET_HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = SHEET_HEADERS[name];
    const current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    let needsHeader = false;
    for (let i = 0; i < headers.length; i++) {
      if (String(current[i] || '') !== headers[i]) {
        needsHeader = true;
        break;
      }
    }
    if (needsHeader) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  });

  seedDefaults_();
  backfillAttendanceCheckoutFields_();
  migratePinsToFourDigits();
  migrateDeviceTokens_();
  invalidatePublicStaffOptionsCache_();
  installMailQueueWorker_();
  syncAccountantViewRows();

  Logger.log('Moving Surface HR system setup complete.');
  Logger.log('Use ?staff=1 for staff portal, ?admin=1 for Admin/Accounts and ?req=1 for staff requisition.');
}

function backfillAttendanceCheckoutFields_() {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet || sheet.getLastRow()<2) return;
  const settings=getSettings_();
  const range=sheet.getRange(2,1,sheet.getLastRow()-1,25);
  const rows=range.getValues();
  let changed=false;
  rows.forEach(function(row) {
    const type=attendanceCheckoutType_(row[4],row[24]);
    if (!type) return;
    if (!row[23]) {
      row[23]=formatAttendanceTime_(row[5],row[6],settings);
      changed=true;
    }
    if (String(row[24] || '')!==type) {
      row[24]=type;
      changed=true;
    }
  });
  if (changed) range.setValues(rows);
}

function seedDefaults_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const settingDefaults = [
    ['TIMEZONE','Africa/Lagos','Use Africa/Lagos for Nigeria attendance time'],
    ['START_TIME','09:00','Official resumption time'],
    ['GRACE_MINUTES','15','Late starts after this grace period'],
    ['CLOSE_TIME','17:00','Official close of work; use HH:mm in the configured timezone'],
    ['OVERTIME_GRACE_MINUTES','15','Overtime checkout is valid only after closing time plus this grace period'],
    ['OVERTIME_PAY_RATE','0','Optional overtime pay per overtime day; set 0 to disable overtime pay'],
    ['HALF_DAY_CUTOFF','10:00','Check-in from this time is half day'],
    ['ABSENT_CUTOFF','12:00','Check-in from this time is marked absent'],
    ['REQUIRE_GPS','YES','YES rejects attendance outside approved locations'],
    ['DEFAULT_RADIUS_METRES','300','Default allowed office radius'],
    ['REQUIRE_PIN','YES','Staff must enter PIN'],
    ['ALLOW_SELF_PIN','YES','Staff can create first PIN on first use'],
    ['REQUIRE_DEVICE_TOKEN','YES','Locks staff application access to the registered phone/browser'],
    ['AUTO_REGISTER_DEVICE','YES','Registers first device automatically'],
    ['DEFAULT_LATE_PENALTY','1000','Default late penalty when staff row is blank'],
    ['HALF_DAY_PERCENT','50','Half day deduction percent of daily rate'],
    ['WORKDAYS','MON,TUE,WED,THU,FRI','Working days used for payroll']
  ];
  const existingSettings = getKeyValueSheet_(SETTINGS_SHEET);
  const missingSettings = settingDefaults.filter(row => !Object.prototype.hasOwnProperty.call(existingSettings, row[0]));
  if (missingSettings.length) {
    settings.getRange(settings.getLastRow()+1,1,missingSettings.length,3).setValues(missingSettings);
  }

  const reqSettings = ss.getSheetByName(REQ_SETTINGS_SHEET);
  if (reqSettings.getLastRow() < 2) {
    reqSettings.getRange(2,1,7,3).setValues([
      ['MD_EMAIL','joemugbeh@gmail.com','First approval email goes here'],
      ['ACCOUNTS_EMAIL','accounts@movingsurface.com.ng','Accounts notification after approval'],
      ['HR_EMAIL','','Optional copy email'],
      ['COMPANY_NAME','Moving Surface Interactive Media Limited','Used in email headings'],
      ['ALLOW_STAFF_EMAIL_COPY','YES','Send staff notifications where email exists'],
      ['DEFAULT_CURRENCY','NGN','Used for estimates'],
      ['HIGH_VALUE_THRESHOLD','500000','Flags requests above this amount']
    ]);
  }

  const admin = ss.getSheetByName(ADMIN_USERS_SHEET);
  if (admin.getLastRow() < 2) {
    admin.getRange(2,1,3,12).setValues([
      ['Kome Vreno','2012','YES','Super Admin','YES','YES','YES','YES','','','',''],
      ['Accountant','1012','YES','Accountant','YES','NO','NO','NO','','','',''],
      ['Admin Assistant','0003','YES','Admin Assistant','YES','NO','NO','NO','','','','']
    ]);
  }

  const staff = ss.getSheetByName(STAFF_SHEET);
  if (staff.getLastRow() < 2) {
    staff.getRange(2,1,3,29).setValues([
      ['MS-001','Kome Vreno','Management','Abuja','2012','','joemugbeh@gmail.com','','YES','YES',0,1000,'','','Owner/admin sample',new Date(2026,0,1),'Permanent','Active','','MD/CEO','','','','',20,20,22,'','YES'],
      ['MS-002','Accountant','Accounts','Abuja','1012','','accounts@movingsurface.com.ng','','YES','YES',0,1000,'','','Accounts sample',new Date(2026,0,1),'Permanent','Active','','Accountant','','','','',20,20,22,'','YES'],
      ['MS-003','Admin Assistant','Admin','Abuja','0003','','','','YES','YES',0,1000,'','','Admin sample',new Date(2026,0,1),'Permanent','Active','','Admin Assistant','','','','',20,20,22,'','YES']
    ]);
  }

  const loc = ss.getSheetByName(LOCATIONS_SHEET);
  if (loc.getLastRow() < 2) {
    loc.getRange(2,1,2,5).setValues([
      ['Abuja Office',9.0765,7.3986,500,'YES'],
      ['Lagos Office',6.5244,3.3792,500,'YES']
    ]);
  }
}
