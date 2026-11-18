function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Moving Surface ERP')
    .addItem('Setup / Repair Core System', 'setupMovingSurfaceSystem')
    .addItem('Setup / Repair ERP Upgrade', 'setupMovingSurfaceERPUpgrade')
    .addItem('Migrate Projects / Leads Schema', 'migrateProjectsAndLeadsSchema')
    .addItem('Sync Accountant View Rows', 'syncAccountantViewRows')
    .addItem('Reconcile Project Spend Rollups', 'reconcileAllProjectRollups')
    .addItem('Process Pending Emails Now', 'processPendingMailQueue')
    .addItem('Run Lead Follow-up Reminders Now', 'runLeadReminderSweepNow')
    .addItem('Install Reminder / Reconcile Triggers', 'installLeadReminderWorker')
    .addSeparator()
    .addItem('Test Admin Login', 'testAdminLogin')
    .addItem('Test ERP Portal Data', 'testERPPortalData')
    .addItem('Test Routing and PIN Rules', 'testRoutingAndPinRules')
    .addItem('Test Payroll Summary Rules', 'testPayrollSummaryRules')
    .addItem('Test Overtime Rules', 'testOvertimeRules')
    .addItem('Test Device Token Fingerprints', 'testDeviceTokenFingerprintRules')
    .addItem('Test Projects / Leads Pipeline', 'testProjectsAndLeadsPipeline')
    .addItem('Test Inventory Rules', 'testInventoryModule')
    .addItem('Show Web App Link Help', 'showWebAppHelp')
    .addToUi();
}

function showWebAppHelp() {
  const deployed = ScriptApp.getService().getUrl();
  const url = deployed || 'Deploy as Web App first, then copy the /exec URL.';
  SpreadsheetApp.getUi().alert(
    'Web App Links\n\nAttendance:\n' + url +
    '\n\nStaff Self-Service Portal:\n' + (deployed ? getWebAppRouteUrl_('staff') : url) +
    '\n\nStaff Requisition only:\n' + (deployed ? getWebAppRouteUrl_('req') : url) +
    '\n\nStaff account recovery:\n' + (deployed ? getWebAppRouteUrl_('recovery') : url) +
    '\n\nAdmin / Accounts / ERP:\n' + (deployed ? getWebAppRouteUrl_('admin') : url)
  );
}

function testAdminLogin() {
  ['2012', '2000', '1234', 'wrong'].forEach(pin => {
    Logger.log('PIN ' + pin + ' => ' + JSON.stringify(getAdminUser_(pin)));
  });
}

function testRoutes() {
  Logger.log('Attendance page title: ' + renderPage_('AttendanceView', 'Moving Surface Attendance').getTitle());
  Logger.log('Admin page title: ' + renderPage_('AdminPortalView', 'Moving Surface ERP').getTitle());
  Logger.log('Requisition page title: ' + renderPage_('RequisitionView', 'Moving Surface Requisition').getTitle());
}

function testAdminPortalData() {
  const result = adminGetPortalData({adminPin: '2012'});
  Logger.log(JSON.stringify(result).substring(0, 5000));
}
