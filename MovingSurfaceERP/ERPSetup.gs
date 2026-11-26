function setupMovingSurfaceERPUpgrade() {
  setupMovingSurfaceSystem();
  // The schema migration must run before headers are enforced: relabelling the
  // legacy Projects columns in place would strand the data sitting under them.
  migrateProjectsAndLeadsSchema();
  setupERPModuleSheets_();
  installLeadReminderWorker_();
  installProjectReconcileWorker_();
  reconcileAllProjectRollups();
  Logger.log('Moving Surface ERP upgrade setup complete.');
  Logger.log('Attendance: Web App URL');
  Logger.log('Staff portal: Web App URL?staff=1');
  Logger.log('Account recovery: Web App URL?recover=1');
  Logger.log('Staff requisition: Web App URL?req=1');
  Logger.log('Admin / ERP: Web App URL?admin=1');
}

function setupERPModuleSheets_() {
  ensureERPHeaders_('ERP Settings', ['Setting','Value','Notes'], [
    ['COMPANY_NAME','Moving Surface Interactive Media Limited','Used in dashboard and emails'],
    ['LOW_STOCK_ALERT','YES','Show inventory low stock warning'],
    ['DEFAULT_APPROVAL_LIMIT','500000','High-value request threshold'],
    ['ERP_VERSION','4.0','Projects / Leads pipeline upgrade'],
    ['TIMEZONE','Africa/Lagos','Used for date/time display'],
    ['LEAD_REMINDERS_ENABLED','YES','Set to NO to suspend all lead follow-up reminders'],
    ['LEAD_WEEKLY_DIGEST_DAY','Monday','Weekday the per-owner follow-up digest is sent'],
    ['LEAD_REMINDER_HORIZON_DAYS','28','How far ahead the weekly digest looks'],
    ['LEAD_FINAL_REMINDER_DAYS','3','Days before a follow-up date that the final reminder fires']
  ]);

  ensureERPHeaders_('ERP Dashboard', ['Metric','Value','Notes'], [
    ['Active Staff','=COUNTIF(\'Staff List\'!R:R,"Active")','From Staff List'],
    ['Open Requisitions','=COUNTIF(\'Item Requisitions\'!Q:Q,"Pending MD Approval")+COUNTIF(\'Item Requisitions\'!Q:Q,"Approved - Awaiting Accounts")','Pending + awaiting accounts'],
    ['Open Expense Claims','=COUNTIF(\'Expense Claims\'!J:J,"Pending Approval")+COUNTIF(\'Expense Claims\'!J:J,"Approved - Awaiting Accounts")','Requires action'],
    ['Assets Assigned','=COUNTIF(Assets!I:I,"Assigned")','Asset register'],
    ['Inventory Low Stock','=COUNTIF(\'Inventory Items\'!K:K,"Low Stock")','Inventory levels'],
    ['Active Projects','=COUNTIF(Projects!I:I,"Active")','Projects sheet (Status is column I)'],
    ['Open Leads','=COUNTA(Leads!A2:A1000)-COUNTIF(Leads!I:I,"Lost")-COUNTIF(Leads!I:I,"Dormant")-COUNTIF(Leads!I:I,"Delivered")','Leads still in play'],
    ['Pipeline Value','=SUMIF(Leads!I:I,"<>Lost",Leads!K:K)','Estimated value of leads not lost'],
    ['Weighted Pipeline','=SUM(Leads!M:M)','Estimated value x probability'],
    ['Contracted Revenue','=SUM(Projects!J:J)','Total project contract value'],
    ['Project Actual Spend','=SUM(Projects!M:M)','Rolled up from Project Cost Lines'],
    ['Outstanding Receivable','=SUM(Projects!S:S)','Invoiced less received']
  ]);
  ensureERPDashboardMetric_(
    'Overtime Days This Month',
    '=COUNTIFS(\'Attendance Log\'!E:E,"Check Out Overtime",\'Attendance Log\'!A:A,">="&TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1),"yyyy-mm-dd"),\'Attendance Log\'!A:A,"<"&TEXT(EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1),"yyyy-mm-dd"))',
    'Recorded overtime checkouts this month; one checkout per staff member per day'
  );

  ensureERPHeaders_('Assets', ['Asset ID','Asset Name','Category','Brand/Model','Serial Number','Purchase Date','Purchase Cost','Location','Status','Assigned To Staff ID','Assigned To Name','Condition','Return Due Date','Notes','Last Updated']);
  ensureERPHeaders_('Asset Movements', ['Timestamp','Movement ID','Asset ID','Asset Name','Action','From','To / Staff ID','Staff Name','Project / Event','Condition Out','Condition In','Handled By','Notes']);
  ensureERPHeaders_('Inventory Items', ['Item ID','Item Name','Category','Unit','Opening Stock','Stock In','Stock Out','Current Stock','Reorder Level','Location','Stock Status','Preferred Vendor','Notes','Last Updated']);
  ensureERPHeaders_('Inventory Transactions', ['Timestamp','Transaction ID','Item ID','Item Name','Transaction Type','Quantity','Project / Event','Requested By','Approved By','Reference','Notes']);
  ensureERPHeaders_('Expense Claims', ['Timestamp','Claim ID','Staff ID','Staff Name','Department','Expense Date','Category','Project / Event','Amount','Status','Approver','Approval Date','Accounts Status','Paid Date','Receipt Link','Purpose','Payment Reference','Last Updated']);
  ensureERPHeaders_('Leave Requests', ['Timestamp','Leave ID','Staff ID','Staff Name','Department','Leave Type','Start Date','End Date','Days','Reason','Status','Approver','Approval Date','Notes','Last Updated']);
  // Seed lists only fill an empty sheet, so existing workbooks need these
  // upserted by key rather than appended with the seed rows above.
  ensureERPSetting_('ERP_VERSION','4.0','Projects / Leads pipeline upgrade');
  ensureERPSetting_('LEAD_REMINDERS_ENABLED','YES','Set to NO to suspend all lead follow-up reminders');
  ensureERPSetting_('LEAD_WEEKLY_DIGEST_DAY','Monday','Weekday the per-owner follow-up digest is sent');
  ensureERPSetting_('LEAD_REMINDER_HORIZON_DAYS','28','How far ahead the weekly digest looks');
  ensureERPSetting_('LEAD_FINAL_REMINDER_DAYS','3','Days before a follow-up date that the final reminder fires');

  ensureERPDashboardMetric_('Open Leads',
    '=COUNTA(Leads!A2:A1000)-COUNTIF(Leads!I:I,"Lost")-COUNTIF(Leads!I:I,"Dormant")-COUNTIF(Leads!I:I,"Delivered")',
    'Leads still in play');
  ensureERPDashboardMetric_('Pipeline Value','=SUMIF(Leads!I:I,"<>Lost",Leads!K:K)','Estimated value of leads not lost');
  ensureERPDashboardMetric_('Weighted Pipeline','=SUM(Leads!M:M)','Estimated value x probability');
  ensureERPDashboardMetric_('Contracted Revenue','=SUM(Projects!J:J)','Total project contract value');
  ensureERPDashboardMetric_('Project Actual Spend','=SUM(Projects!M:M)','Rolled up from Project Cost Lines');
  ensureERPDashboardMetric_('Outstanding Receivable','=SUM(Projects!S:S)','Invoiced less received');
  ensureERPDashboardMetric_('Overdue Follow-ups',
    '=COUNTIFS(Leads!P:P,"<"&TODAY(),Leads!P:P,"<>",Leads!Q:Q,"<>Off")',
    'Scheduled follow-ups now past their date');
  ensureERPDashboardMetric_('Active Projects','=COUNTIF(Projects!I:I,"Active")','Projects sheet (Status is column I)');

  ensureERPHeaders_(PROJECTS_SHEET, projectsHeaders_());
  ensureERPHeaders_(PROJECT_COSTS_SHEET, projectCostHeaders_());
  ensureERPHeaders_(CLIENTS_SHEET, clientsHeaders_());
  ensureERPHeaders_(LEADS_SHEET, leadsHeaders_());
  ensureERPHeaders_(FIELD_DUTY_SHEET, fieldDutyHeaders_());
  ensureERPHeaders_(LEAD_ACTIVITY_SHEET, leadActivityHeaders_());
  ensureERPHeaders_(LEAD_REMINDER_LOG_SHEET, leadReminderLogHeaders_());
  ensureERPHeaders_('Vendors', ['Vendor ID','Vendor Name','Category','Contact Person','Email','Phone','Location','Bank Details','Status','Notes','Last Updated']);
  ensureERPHeaders_('Purchase Orders', ['PO ID','Requisition ID','Vendor ID','Vendor Name','Project / Event','Total Amount','Status','Prepared By','Approved By','Issue Date','Expected Delivery','Delivery Status','Notes','Last Updated']);
  ensureERPHeaders_('ERP Notifications', ['Timestamp','Notification ID','Module','Target Role','Target Email','Subject','Status','Related ID','Notes']);

  seedERPExamples_();
}

function ensureERPDashboardMetric_(label,formula,notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ERP Dashboard');
  const match = sheet.getRange(2,1,Math.max(1,sheet.getLastRow()-1),1)
    .createTextFinder(String(label))
    .matchCase(true)
    .matchEntireCell(true)
    .findNext();
  const row = match ? match.getRow() : sheet.getLastRow()+1;
  sheet.getRange(row,1,1,3).setValues([[label,formula,notes]]);
}

/** Upserts one ERP Settings row by key, leaving an existing value in place. */
function ensureERPSetting_(key,defaultValue,notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERP_SETTINGS_SHEET);
  if (!sheet) return;
  const match = sheet.getRange(1,1,Math.max(1,sheet.getLastRow()),1)
    .createTextFinder(String(key)).matchCase(true).matchEntireCell(true).findNext();
  if (match) {
    // Keep whatever the business already configured; only fill a blank value.
    if (String(sheet.getRange(match.getRow(),2).getValue() || '').trim() === '') {
      sheet.getRange(match.getRow(),2).setValue(defaultValue);
    }
    sheet.getRange(match.getRow(),3).setValue(notes);
    return;
  }
  sheet.appendRow([key,defaultValue,notes]);
}

function ensureERPHeaders_(sheetName, headers, seedRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setBackground('#0F766E').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sh.autoResizeColumns(1, headers.length);
  if (seedRows && sh.getLastRow() < 2) {
    sh.getRange(2, 1, seedRows.length, headers.length).setValues(seedRows);
  }
}

function seedERPExamples_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const assets = ss.getSheetByName('Assets');
  if (assets.getLastRow() < 2) {
    assets.appendRow(['MS-AST-0001','Datapath FX4','Video Processor','Datapath FX4','','','', 'Abuja','Available','','','Good','','Sample asset - edit or delete', new Date()]);
    assets.appendRow(['MS-AST-0002','Projector','Projector','Panasonic / Barco','','','', 'Abuja','Available','','','Good','','Sample asset - edit or delete', new Date()]);
  }

  const inv = ss.getSheetByName('Inventory Items');
  if (inv.getLastRow() < 2) {
    // Computed values, not row-pinned formulas, and categories from the list.
    inv.appendRow(['INV-0001','HDMI Cable 10m','Cable','pcs',10,0,0,10,3,'Abuja','OK','','Sample item - edit or delete',new Date()]);
    inv.appendRow(['INV-0002','SDI Cable 20m','Cable','pcs',6,0,0,6,2,'Abuja','OK','','Sample item - edit or delete',new Date()]);
  }

  const clients = ss.getSheetByName(CLIENTS_SHEET);
  if (clients.getLastRow() < 2) {
    clients.appendRow(['CL-0001','dmg events','Leigh / Chloe','','','','Events','','','Active','','','','','NOG content/projects',new Date()]);
    clients.appendRow(['CL-0002','NASRDA','','','','','Government','Abuja','','Active','','','','','Experience Center',new Date()]);
  }

  const projects = ss.getSheetByName(PROJECTS_SHEET);
  if (projects.getLastRow() < 2) {
    // Contract Value and Expenditure Budget start at zero on purpose: the
    // derived columns are recomputed from Project Cost Lines, never typed.
    projects.appendRow(['PRJ-0001','NOG 2026','CL-0001','dmg events','','','','','Active',0,0,0,0,0,0,0,0,0,0,'Abuja','Sample project - set Contract Value and Expenditure Budget',new Date(),new Date()]);
    projects.appendRow(['PRJ-0002','NASRDA Experience Center','CL-0002','NASRDA','','','','','Active',0,0,0,0,0,0,0,0,0,0,'Abuja','Sample project - set Contract Value and Expenditure Budget',new Date(),new Date()]);
  }

  const leads = ss.getSheetByName(LEADS_SHEET);
  if (leads.getLastRow() < 2) {
    leads.appendRow(['LEAD-0001','CL-0001','dmg events','NOG 2027 content package','Leigh / Chloe','','','Existing Client','Qualified',25,0,0,0,'','','','Weekly','','','','','Sample lead - set the opportunity value and follow-up date',new Date(),new Date()]);
  }

  applyERPValidations_();
}

function applyERPValidations_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setERPValidation_(ss.getSheetByName('Assets'), 'I2:I1000', ['Available','Assigned','Maintenance','Damaged','Retired']);
  setERPValidation_(ss.getSheetByName(INVENTORY_TX_SHEET), 'E2:E1000', inventoryTransactionTypes_());
  // Category and Unit are picked from a list in the portal, so the sheet enforces
  // the same vocabulary for anyone editing rows directly.
  setERPValidation_(ss.getSheetByName(INVENTORY_ITEMS_SHEET), 'C2:C1000', inventoryCategories_());
  setERPValidation_(ss.getSheetByName(INVENTORY_ITEMS_SHEET), 'D2:D1000', inventoryUnits_());
  setERPValidation_(ss.getSheetByName('Expense Claims'), 'J2:J1000', ['Pending Approval','Approved - Awaiting Accounts','Rejected','Paid / Closed']);
  setERPValidation_(ss.getSheetByName('Expense Claims'), 'M2:M1000', ['Not Submitted','Awaiting Accounts','Paid','Rejected']);
  setERPValidation_(ss.getSheetByName('Leave Requests'), 'K2:K1000', ['Pending Approval','Approved','Rejected','Cancelled']);
  if (ss.getSheetByName(FIELD_DUTY_SHEET)) {
    setERPValidation_(ss.getSheetByName(FIELD_DUTY_SHEET), 'J2:J1000', fieldDutyStatuses_());
    // Kept as text so the sheet never reinterprets a duty date as a locale date.
    ss.getSheetByName(FIELD_DUTY_SHEET).getRange('D2:E1000').setNumberFormat('@');
  }
  // Projects.Status moved from column G to column I in the new layout.
  setERPValidation_(ss.getSheetByName(PROJECTS_SHEET), 'I2:I1000', projectStatuses_());
  setERPValidation_(ss.getSheetByName(LEADS_SHEET), 'I2:I1000', leadStages_());
  setERPValidation_(ss.getSheetByName(LEADS_SHEET), 'H2:H1000', leadSources_());
  setERPValidation_(ss.getSheetByName(LEADS_SHEET), 'Q2:Q1000', leadReminderCadences_());
  setERPValidation_(ss.getSheetByName(CLIENTS_SHEET), 'J2:J1000', clientRelationshipStatuses_());
  setERPValidation_(ss.getSheetByName(LEAD_ACTIVITY_SHEET), 'E2:E1000', leadActivityTypes_());
  setERPValidation_(ss.getSheetByName(PROJECT_COSTS_SHEET), 'D2:D5000', projectCostCategories_());
  setERPValidation_(ss.getSheetByName(PROJECT_COSTS_SHEET), 'M2:M5000', projectCostSourceModules_());
  setERPValidation_(ss.getSheetByName(PROJECT_COSTS_SHEET), 'O2:O5000', projectCostPaymentStatuses_());
  // Expense categories share the project cost vocabulary so a claim posted to a
  // project lands in a category the project report already understands.
  setERPValidation_(ss.getSheetByName(EXPENSE_CLAIMS_SHEET), 'G2:G1000', projectCostCategories_());
}

function setERPValidation_(sh, rangeA1, values) {
  if (!sh) return;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
  sh.getRange(rangeA1).setDataValidation(rule);
}
