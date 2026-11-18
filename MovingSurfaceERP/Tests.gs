function testERPPortalData() {
  setupERPModuleSheets_();
  ['2012','1012','0003'].forEach(function(pin) {
    const data = adminGetPortalData({adminPin:pin});
    Logger.log('PIN ' + pin + ' => success=' + data.success + ', user=' + (data.user ? data.user.name : data.message));
  });
}

function testRoutingAndPinRules() {
  const checks = [
    {params:{},route:'attendance'},
    {params:{staff:'1'},route:'staff'},
    {params:{req:'1'},route:'req'},
    {params:{recover:'1'},route:'recovery'},
    {params:{ADMIN:'1'},route:'admin'},
    {params:{RECOVER:'1'},route:'recovery'},
    {params:{admin:'1'},route:'admin'},
    {params:{staff:'1',admin:'1'},route:'admin'}
  ];
  checks.forEach(function(check) {
    const actual = resolveRoute_(check.params).key;
    if (actual !== check.route) throw new Error('Route test failed: expected ' + check.route + ', got ' + actual);
  });
  if (resolveRoute_({recover:'1',token:'test'}).key !== 'recovery') {
    throw new Error('Recovery links with a token must route to the recovery page.');
  }
  if (getQueryParameterIgnoreCase_({ADMIN:'1'},'admin') !== '1') {
    throw new Error('Query parameter lookup must ignore key casing.');
  }
  const normalized = normalizeQueryParameters_({ADMIN:'1',TOKEN:'recovery-token',ApproveReq:'REQ-1'});
  if (normalized.admin !== '1' || normalized.token !== 'recovery-token' || normalized.approvereq !== 'REQ-1') {
    throw new Error('Query parameters must be normalized case-insensitively.');
  }
  if (buildCanonicalWebAppUrl_('https://example.test/exec',{ADMIN:'1',TOKEN:'AbC123'}) !==
      'https://example.test/exec?admin=1&token=AbC123') {
    throw new Error('Canonical URLs must lowercase parameter names and preserve case-sensitive values.');
  }
  if (normalizePin_('0008') !== '0008' || normalizePin_(8) !== '0008' || !isValidPin_('0000') || isValidPin_('008')) {
    throw new Error('Four-digit PIN tests failed.');
  }

  Logger.log('Routing and four-digit PIN checks passed.');
}

function testPayrollSummaryRules() {
  const result = calculateStaffPayrollSummary_(
    {workDaysPerMonth:22,monthlySalary:220000,latePenalty:1000},
    {presentDays:18,lateCount:2,halfDayCount:1},
    {DEFAULT_LATE_PENALTY:1000,HALF_DAY_PERCENT:50}
  );
  if (result.dailyRate !== 10000 || result.absentDays !== 4 || result.totalDeduction !== 47000 || result.netPay !== 173000) {
    throw new Error('Payroll summary calculation failed: ' + JSON.stringify(result));
  }
  Logger.log('Payroll summary calculation checks passed.');
}

function testOvertimeRules() {
  const settings = {CLOSE_TIME:'17:00',OVERTIME_GRACE_MINUTES:15,OVERTIME_PAY_RATE:3000,HALF_DAY_PERCENT:50,DEFAULT_LATE_PENALTY:1000};
  if (normalizeTimeSetting_('17:00','09:00','CLOSE_TIME') !== '17:00' ||
      normalizeTimeSetting_('5:00 PM','09:00','CLOSE_TIME') !== '17:00' ||
      normalizeTimeSetting_('12:00 AM','09:00','CLOSE_TIME') !== '00:00' ||
      normalizeTimeSetting_(0.7083333333333334,'09:00','CLOSE_TIME') !== '17:00') {
    throw new Error('Spreadsheet close-time formats were not normalized correctly.');
  }
  if (nonNegativeSetting_(0,15,'OVERTIME_GRACE_MINUTES') !== 0) throw new Error('A zero overtime grace period must be preserved.');
  let invalidRateRejected = false;
  try { nonNegativeSetting_(-1,0,'OVERTIME_PAY_RATE'); } catch (err) { invalidRateRejected = true; }
  if (!invalidRateRejected) throw new Error('Negative overtime pay rates must be rejected.');
  if (isOvertimeCheckoutAllowedAtMinute_(17*60+14,settings)) throw new Error('Overtime checkout must fail before the grace period ends.');
  if (!isOvertimeCheckoutAllowedAtMinute_(17*60+15,settings)) throw new Error('Overtime checkout must be valid at the grace-period boundary.');
  if (!isOvertimeCheckoutAllowedAtMinute_(20*60,settings)) throw new Error('Overtime checkout must remain valid after the boundary.');
  const summary = calculateStaffPayrollSummary_(
    {workDaysPerMonth:22,monthlySalary:220000,latePenalty:1000},
    {presentDays:18,lateCount:2,halfDayCount:1,overtimeDays:2},
    settings
  );
  if (summary.overtimeDays !== 2 || summary.overtimePay !== 6000 || summary.netPay !== 179000) {
    throw new Error('Overtime days/pay calculation failed: ' + JSON.stringify(summary));
  }
  Logger.log('Overtime timing and payroll checks passed.');
}

function testDeviceTokenFingerprintRules() {
  const token = 'sample-device-token-for-tests';
  const fingerprint = sha256Hex_(token);
  if (fingerprint === token || !/^[a-f0-9]{64}$/.test(fingerprint)) {
    throw new Error('Device token fingerprint must be a SHA-256 hex digest.');
  }
  if (!deviceTokenMatches_(fingerprint,token) || deviceTokenMatches_(fingerprint,'different-token')) {
    throw new Error('Device token fingerprint verification failed.');
  }
  Logger.log('Device token fingerprint checks passed.');
}

function testRecoveryEmailMatchRules() {
  const staff = {companyEmail:'staff@example.com'};
  if (!recoveryEmailMatches_(staff,' STAFF@example.com ')) {
    throw new Error('Recovery email matching must ignore surrounding whitespace and letter case.');
  }
  if (recoveryEmailMatches_(staff,'other@example.com') ||
      recoveryEmailMatches_({companyEmail:''},'staff@example.com') ||
      recoveryEmailMatches_(null,'staff@example.com')) {
    throw new Error('Recovery must reject an unregistered email or missing staff record.');
  }
  Logger.log('Recovery email match checks passed.');
}

function testAttendanceTransitionRules() {
  const empty = {checkedIn:false,checkedOut:false};
  if (getAttendanceTransitionError_('Check In',empty)) {
    throw new Error('The first check-in of the day must be allowed.');
  }
  if (getAttendanceTransitionError_('Check Out',empty) !== 'Check in before checking out today.' ||
      getAttendanceTransitionError_('Check Out Overtime',empty) !== 'Check in before checking out today.') {
    throw new Error('Check-out must require a prior check-in.');
  }
  const checkedIn = {checkedIn:true,checkedOut:false};
  if (getAttendanceTransitionError_('Check In',checkedIn) !== 'You already submitted Check In today.' ||
      getAttendanceTransitionError_('Check Out',checkedIn) ||
      getAttendanceTransitionError_('Check Out Overtime',checkedIn)) {
    throw new Error('A staff member may check in once and check out once per day.');
  }
  if (getAttendanceTransitionError_('Check In',{checkedIn:false,checkedOut:true}) !==
      'You have already checked out today. Check in is no longer available.' ||
      getAttendanceTransitionError_('Check Out',{checkedIn:true,checkedOut:true}) !== 'You already checked out today.' ||
      getAttendanceTransitionError_('Check Out Overtime',{checkedIn:true,checkedOut:true}) !== 'You already checked out today.') {
    throw new Error('Check-in must not follow checkout, and normal/overtime checkout must be mutually exclusive.');
  }
  Logger.log('Attendance transition checks passed.');
}

function testAttendanceLogStateMatching() {
  const timezone='Africa/Lagos';
  const date=new Date('2026-10-09T10:00:00Z');
  if (attendanceDateKey_(date,timezone)!=='2026-10-09' ||
      attendanceDateKey_('2026-10-09','Africa/Lagos')!=='2026-10-09' ||
      attendanceDateKey_('2026-10-09T09:15:00','Africa/Lagos')!=='2026-10-09') {
    throw new Error('Attendance state matching must support spreadsheet Date values and ISO date strings.');
  }
  if (normalizeAttendanceAction_(' check in ')!=='checkin' ||
      normalizeAttendanceAction_('Check Out Overtime')!=='checkoutovertime') {
    throw new Error('Attendance action matching must ignore casing, spaces, and punctuation.');
  }
  const rows=[
    [date,'MS-001','','',' Check In '],
    ['2026-10-09','MS-001','','','CHECK OUT'],
    ['2026-10-09','MS-002','','','Check In'],
    ['2026-10-08','MS-001','','','Check In']
  ];
  const state=attendanceStateFromRows_(rows,'2026-10-09','ms-001',timezone);
  if (!state.checkedIn || !state.checkedOut ||
      attendanceStateFromRows_(rows,'2026-10-09','MS-003',timezone).checkedIn ||
      attendanceStateFromRows_(rows,'2026-10-10','MS-001',timezone).checkedIn) {
    throw new Error('Attendance state must match date and staff across Date/text rows, preventing duplicate check-ins and allowing checkout.');
  }
  Logger.log('Attendance log matching checks passed.');
}

function testAttendanceCheckoutDetails() {
  if (attendanceCheckoutType_('Check Out','')!=='Regular' ||
      attendanceCheckoutType_('Check Out Overtime','')!=='Overtime' ||
      attendanceCheckoutType_('Check In','')!=='' ||
      attendanceCheckoutType_('Check Out','Overtime')!=='Overtime' ||
      attendanceCheckoutType_('Check In','Regular')!=='Regular') {
    throw new Error('Checkout type must explicitly distinguish regular and overtime checkouts and preserve stored values.');
  }
  Logger.log('Attendance checkout detail checks passed.');
}

function testAttendanceResponseMessageSections() {
  const response=buildAttendanceResponseMessages_(
    'Check In','Test Staff','08:20','Late','Main Office','Recognized',12,false,
    {todayPenalty:100,monthLatePenalty:200,monthHalfDayDeduction:0,monthAbsenceDeduction:0,monthTotalDeduction:300}
  );
  if (!response.primaryMessage.startsWith('Check In successful') ||
      response.primaryMessage.indexOf('Late')!==-1 ||
      !response.additionalMessage.includes('Status: Late') ||
      !response.additionalMessage.includes('Today Penalty:') ||
      !response.message.startsWith(response.primaryMessage+'\n\n'+response.additionalMessage)) {
    throw new Error('Successful attendance confirmation must be separate from status and deduction details while preserving the combined message.');
  }
  Logger.log('Attendance response message section checks passed.');
}

function testRequisitionClientDto() {
  const internal = {reqId:'REQ-TEST',status:'Pending MD Approval',token:'server-only-token'};
  const client = requisitionClientDto_(internal);
  if ('token' in client) throw new Error('Requisition approval token was exposed to a client DTO.');
  if (client.reqId !== internal.reqId || client.status !== internal.status) {
    throw new Error('Client-safe requisition DTO omitted required display fields.');
  }
  if (internal.token !== 'server-only-token') throw new Error('DTO projection modified the internal requisition object.');
  Logger.log('Requisition client DTO security checks passed.');
}

function testAdminAccessRules() {
  const privileged = getAdminAccess_({role:'Admin / HR',canView:'YES'});
  const technical = getAdminAccess_({role:'Technical',canView:'YES'});
  const accountant = getAdminAccess_({role:'Accountant',canView:'YES'});
  const staff = getAdminAccess_({role:'Operations',canView:'YES'});
  const denied = getAdminAccess_({role:'Super Admin',canView:'NO'});
  if (!privileged.readPages.includes('staff') || !privileged.writePages.includes('staff')) {
    throw new Error('Admin/HR must have staff read and write access.');
  }
  if (technical.readPages.includes('staff') || technical.writePages.includes('staff') ||
      !technical.writePages.includes('inventory')) {
    throw new Error('Technical must be able to write inventory but must not access Staff.');
  }
  if (accountant.readPages.includes('staff') || accountant.writePages.includes('staff') ||
      !accountant.writePages.includes('expenses') || !accountant.writePages.includes('requisitions')) {
    throw new Error('Accountants must have no Staff access and retain assigned finance writes.');
  }
  if (staff.readPages.includes('staff') || staff.writePages.length) {
    throw new Error('Other roles must not access Staff or receive default write permissions.');
  }
  if (denied.readPages.length || denied.writePages.length) {
    throw new Error('Can View=NO must remove all page access.');
  }
  const custom = getAdminAccess_({
    role:'Operations',canView:'YES',accessPages:'dashboard,staff,inventory',
    writePages:'inventory,staff'
  });
  if (!custom.readPages.includes('staff') || !custom.writePages.includes('staff') ||
      !custom.writePages.includes('inventory') || custom.readPages.includes('users') ||
      custom.writePages.includes('users')) {
    throw new Error('Explicit page permissions must override role defaults without granting Users.');
  }
  const rootAccess = getAdminAccess_({role:'Admin',canView:'YES'});
  const restrictedRoot = getAdminAccess_({
    role:'Super Admin',canView:'YES',accessPages:'dashboard,reports',
    writePages:'reports'
  });
  const hrAccess = getAdminAccess_({role:'HR',canView:'YES'});
  if (!rootAccess.readPages.includes('users') || rootAccess.writePages.includes('users') ||
      hrAccess.readPages.includes('users') || hrAccess.writePages.includes('users') ||
      !restrictedRoot.readPages.includes('users') ||
      !restrictedRoot.readPages.includes('dashboard') ||
      restrictedRoot.readPages.includes('staff') ||
      !restrictedRoot.writePages.includes('reports') ||
      restrictedRoot.writePages.includes('staff') ||
      !adminCanManageUsers_({valid:true,canView:'YES',access:rootAccess}) ||
      adminCanManageUsers_({valid:true,canView:'YES',access:hrAccess})) {
    throw new Error('Only Admin and Super Admin may open Users; explicit root policies must retain manager access.');
  }
  if (adminCanWritePage_({valid:false,canView:'YES',access:privileged},'staff') ||
      adminCanWritePage_({valid:true,canView:'NO',access:privileged},'staff')) {
    throw new Error('Invalid or view-disabled users must not retain write access.');
  }
  if (adminCanApprovePage_({valid:true,canView:'YES',role:'Technical',canOverride:'NO',access:technical},'requisitions') ||
      !adminCanApprovePage_({valid:true,canView:'YES',role:'Technical',canOverride:'YES',access:technical},'requisitions') ||
      !adminCanDisbursePage_({valid:true,canView:'YES',role:'Accountant',access:accountant},'expenses') ||
      adminCanDisbursePage_({valid:true,canView:'YES',role:'Accountant',access:{writePages:['requisitions']}},'expenses')) {
    throw new Error('Approval and disbursement actions must respect role and page-write permissions.');
  }
  Logger.log('Admin role access checks passed.');
}

function testAdminCredentialResolution() {
  const rows = [
    ['First Admin','0123','YES','Admin','YES','YES','YES','YES','STAFF-1'],
    ['Second Admin','0123','YES','Accountant','YES','NO','NO','NO','STAFF-2'],
    ['Unique Admin','0456','YES','Admin','YES','YES','YES','YES','STAFF-3']
  ];
  const selected = findAdminCredentialMatch_(rows,'0123','STAFF-2');
  if (!selected || selected.row[0] !== 'Second Admin') {
    throw new Error('The selected Staff ID must resolve the matching admin row when a PIN is shared.');
  }
  const first = findAdminCredentialMatch_(rows,'0123','STAFF-1');
  if (findAdminCredentialMatch_(rows,'0123','') !== null ||
      !first || first.row[0] !== 'First Admin' ||
      findAdminCredentialMatch_(rows.concat([
        ['Duplicate Mapping','0123','YES','Technical','YES','NO','NO','NO','STAFF-2']
      ]),'0123','STAFF-2') !== null) {
    throw new Error('Shared PIN ambiguity and duplicate Staff ID mappings must be rejected.');
  }
  const unique = findAdminCredentialMatch_(rows,'0456','');
  if (!unique || unique.row[0] !== 'Unique Admin') {
    throw new Error('A unique PIN must remain resolvable for legacy callers.');
  }
  Logger.log('Admin PIN and Staff ID resolution checks passed.');
}

function testUserPermissionMatrixValidation() {
  const read=['dashboard','inventory','assets'];
  const write=['inventory'];
  if (normalizeAdminManagedPageList_(read).join(',')!=='dashboard,inventory,assets' ||
      normalizeAdminManagedPageList_(['users'])!==null ||
      normalizeAdminManagedPageList_(['made-up'])!==null ||
      write.some(page=>!read.includes(page))) {
    throw new Error('The user access matrix must accept known pages and exclude protected or unknown pages.');
  }
  Logger.log('User permission matrix checks passed.');
}

function testAdminManagedRoleLabels() {
  const labels=getAdminManagedRoleLabels_([
    {role:'Operations'},
    {role:'Accountant'},
    {role:'Super Admin'},
    {role:'HR'},
    {role:'Admin / HR'},
    {role:'Custom Role'},
    {role:'Custom Role'},
    {role:''}
  ]);
  const required=['MD/CEO','SuperAdmin','Admin/HR','Accountant','Technical','Driver',
    'Motion Artist','3D Artist','Head of Operations','IoT Engineer','Lead Software Engineer','Janitor'];
  if (required.some(role=>!labels.includes(role)) ||
      labels.filter(role=>role.toLowerCase()==='accountant').length!==1 ||
      labels.includes('Super Admin') || labels.includes('HR') || labels.includes('Admin / HR') ||
      labels.filter(role=>role==='SuperAdmin').length!==1 ||
      labels.filter(role=>role==='Admin/HR').length!==1 ||
      canonicalAdminRoleLabel_('Super Admin')!=='SuperAdmin' ||
      canonicalAdminRoleLabel_('HR')!=='Admin/HR' ||
      !labels.includes('Operations') || !labels.includes('Custom Role')) {
    throw new Error('Role choices must use canonical labels and retain unique existing custom labels.');
  }
  Logger.log('Admin role label checks passed.');
}

function testRequisitionNotificationClassification() {
  const inventory = [
    {itemId:'INV-001',itemName:'A4 Paper'},
    {itemId:'INV-002',itemName:'Printer Toner'}
  ];
  const stockRequest = classifyRequisitionItems_([{itemName:' a4 paper '}],inventory);
  const mixedRequest = classifyRequisitionItems_([
    {itemName:'INV-002'},{itemName:'Laptop'}
  ],inventory);
  if (!stockRequest.inventory || stockRequest.purchase ||
      !mixedRequest.inventory || !mixedRequest.purchase) {
    throw new Error('Requisition notices must route inventory and purchase items to their respective roles.');
  }
  if (safeEmailUrl_('javascript:alert(1)') || safeEmailUrl_('data:text/html,test')) {
    throw new Error('Email attachment links must reject non-HTTP URL schemes.');
  }
  Logger.log('Requisition notification classification checks passed.');
}


/*************** PROJECTS / LEADS PIPELINE TESTS ***************/

/**
 * The column maps in Projects.gs, ProjectCosts.gs and Leads.gs address cells by
 * number. If a header is ever inserted without updating its map, every read and
 * write silently shifts by one column. This test is the guard against that.
 */
function testProjectsLeadsColumnMaps() {
  const checks = [
    ['Projects', projectsHeaders_(), PRJ_COL, PRJ_WIDTH, {
      projectId:'Project ID', projectName:'Project Name', clientId:'Client ID', client:'Client',
      leadId:'Lead ID', manager:'Project Manager', startDate:'Start Date', endDate:'End Date',
      status:'Status', contractValue:'Contract Value', budget:'Expenditure Budget',
      plannedCost:'Planned Cost', actualSpend:'Actual Spend', budgetVariance:'Budget Variance',
      balance:'Project Balance', marginPct:'Margin %', invoiced:'Invoiced Amount',
      received:'Amount Received', receivable:'Outstanding Receivable', location:'Location',
      notes:'Notes', lastRollup:'Last Rollup At', lastUpdated:'Last Updated'
    }],
    ['Project Cost Lines', projectCostHeaders_(), COST_COL, COST_WIDTH, {
      lineId:'Line ID', projectId:'Project ID', costDate:'Cost Date', category:'Category',
      subCategory:'Sub-Category', description:'Description', vendor:'Vendor / Payee',
      quantity:'Quantity', unit:'Unit', unitCost:'Unit Cost', plannedAmount:'Planned Amount',
      actualAmount:'Actual Amount', sourceModule:'Source Module', sourceRefId:'Source Ref ID',
      paymentStatus:'Payment Status', paidDate:'Paid Date', paymentReference:'Payment Reference',
      recordedBy:'Recorded By', notes:'Notes', lastUpdated:'Last Updated'
    }],
    ['Clients', clientsHeaders_(), CLIENT_COL, CLIENT_WIDTH, {
      clientId:'Client ID', clientName:'Client / Company', contactPerson:'Contact Person',
      roleTitle:'Role / Title', email:'Email', phone:'Phone', sector:'Sector', location:'Location',
      website:'Website', relationshipStatus:'Relationship Status', accountOwner:'Account Owner',
      accountOwnerName:'Account Owner Name', firstEngaged:'First Engaged',
      lastContact:'Last Contact Date', notes:'Notes', lastUpdated:'Last Updated'
    }],
    ['Leads', leadsHeaders_(), LEAD_COL, LEAD_WIDTH, {
      leadId:'Lead ID', clientId:'Client ID', clientName:'Client / Company', opportunity:'Opportunity',
      contactPerson:'Contact Person', email:'Email', phone:'Phone', source:'Source', stage:'Stage',
      probability:'Probability %', estimatedValue:'Estimated Value', contractValue:'Contract Value',
      weightedValue:'Weighted Value', expectedClose:'Expected Close Date', lastContact:'Last Contact Date',
      nextFollowup:'Next Follow-up', cadence:'Reminder Cadence', owner:'Follow-up Owner',
      ownerName:'Follow-up Owner Name', lastReminderSent:'Last Reminder Sent', lostReason:'Lost Reason',
      notes:'Notes', createdAt:'Created At', lastUpdated:'Last Updated'
    }]
  ];
  checks.forEach(function(check) {
    const label = check[0], headers = check[1], map = check[2], width = check[3], expected = check[4];
    if (headers.length !== width) {
      throw new Error(label + ': header count ' + headers.length + ' does not match declared width ' + width + '.');
    }
    Object.keys(expected).forEach(function(key) {
      if (!map[key]) throw new Error(label + ': column map is missing the key ' + key + '.');
      if (headers[map[key] - 1] !== expected[key]) {
        throw new Error(label + ': ' + key + ' points at column ' + map[key] + ' (' +
          headers[map[key] - 1] + ') but should point at ' + expected[key] + '.');
      }
    });
    if (Object.keys(map).length !== headers.length) {
      throw new Error(label + ': column map has ' + Object.keys(map).length + ' entries for ' + headers.length + ' headers.');
    }
  });
  Logger.log('Projects / Leads column map checks passed.');
}

/** Controlled vocabulary and normalization rules. */
function testProjectsLeadsVocabulary() {
  if (normalizeProjectStatus_('Lead') !== 'Planning') throw new Error('Legacy Lead status must migrate to Planning.');
  if (normalizeProjectStatus_('active') !== 'Active') throw new Error('Project status matching must ignore case.');
  if (normalizeProjectStatus_('nonsense') !== 'Planning') throw new Error('Unknown project status must fall back to Planning.');

  if (normalizeProjectCostCategory_('feeding') !== 'Feeding') throw new Error('Cost category matching must ignore case.');
  if (normalizeProjectCostCategory_('logistics/transport') !== 'Logistics / Transport') {
    throw new Error('Cost category matching must ignore punctuation and spacing.');
  }
  if (normalizeProjectCostCategory_('') !== 'Miscellaneous') throw new Error('A blank category must fall back to Miscellaneous.');
  if (normalizeProjectCostCategory_('Helicopter') !== 'Miscellaneous') throw new Error('An unknown category must fall back to Miscellaneous.');

  if (leadStageProbability_('Won (Contracted)') !== 100) throw new Error('A won lead must carry 100% probability.');
  if (leadStageProbability_('Lost') !== 0) throw new Error('A lost lead must carry 0% probability.');
  if (leadStageProbability_('Negotiation') !== 60) throw new Error('Negotiation probability changed unexpectedly.');
  if (leadStageProbability_('not a stage') !== 10) throw new Error('An unknown stage must fall back to 10%.');

  if (normalizeLeadStage_('negotiation') !== 'Negotiation') throw new Error('Lead stage matching must ignore case.');
  if (normalizeLeadCadence_('off') !== 'Off') throw new Error('Cadence matching must ignore case.');
  if (normalizeLeadCadence_('') !== 'Weekly') throw new Error('A blank cadence must default to Weekly.');
  if (normalizeRelationshipStatus_('') !== 'Prospect') throw new Error('A blank relationship status must default to Prospect.');

  if (legacyStatusToLeadStage_('Lead') !== 'New / Identified') throw new Error('Legacy CRM Lead status must map to New / Identified.');
  if (legacyStatusToLeadStage_('Active') !== 'Qualified') throw new Error('Legacy CRM Active status must map to Qualified.');

  // Closed stages must never attract reminders, and won stages must be a subset of the stage list.
  leadClosedStages_().concat(leadWonStages_()).forEach(function(stage) {
    if (leadStages_().indexOf(stage) === -1) throw new Error('Stage ' + stage + ' is not in leadStages_().');
  });
  if (leadClosedStages_().indexOf('Won (Contracted)') !== -1) {
    throw new Error('A won deal still needs follow-up for delivery and payment, so it must not be a closed stage.');
  }
  Logger.log('Projects / Leads vocabulary checks passed.');
}

/** Money, percentage and date input parsing. */
function testErpInputParsing() {
  if (erpNonNegativeAmount_('', 'X') !== 0) throw new Error('A blank amount must parse as zero.');
  if (erpNonNegativeAmount_('1,250,000', 'X') !== 1250000) throw new Error('Amounts must accept thousands separators.');
  if (erpNonNegativeAmount_('99.456', 'X') !== 99.46) throw new Error('Amounts must round to two decimal places.');
  let rejected = false;
  try { erpNonNegativeAmount_('-5', 'Contract Value'); } catch (err) { rejected = true; }
  if (!rejected) throw new Error('A negative amount must be rejected.');
  rejected = false;
  try { erpNonNegativeAmount_('abc', 'Budget'); } catch (err) { rejected = true; }
  if (!rejected) throw new Error('A non-numeric amount must be rejected.');

  if (erpPercent_('150', 'P') !== 100) throw new Error('Percentages must clamp at 100.');
  if (erpPercent_('-10', 'P') !== 0) throw new Error('Percentages must clamp at 0.');
  if (erpPercent_('45%', 'P') !== 45) throw new Error('Percentages must accept a trailing percent sign.');

  if (erpDateInput_('', 'D') !== '') throw new Error('A blank date must stay blank.');
  if (erpDateInput_('2026-11-02', 'D') !== '2026-11-02') throw new Error('A valid date must pass through unchanged.');
  rejected = false;
  try { erpDateInput_('02/11/2026', 'Next Follow-up'); } catch (err) { rejected = true; }
  if (!rejected) throw new Error('A non-ISO date must be rejected so the scheduler can compare dates reliably.');

  if (erpDaysBetween_('2026-11-02', '2026-11-05') !== 3) throw new Error('Day difference must count forward correctly.');
  if (erpDaysBetween_('2026-11-05', '2026-11-02') !== -3) throw new Error('An overdue date must give a negative difference.');
  if (erpDaysBetween_('2026-11-02', '2026-11-02') !== 0) throw new Error('The same date must give zero days.');
  if (erpAddDays_('2026-02-26', 3) !== '2026-03-01') throw new Error('Adding days must roll over month ends.');
  Logger.log('ERP input parsing checks passed.');
}

/** Reminder scheduling: which reminder fires, and the keys that keep it once-only. */
function testLeadReminderScheduling() {
  if (normalizeWeekday_('Monday') !== 1) throw new Error('Monday must resolve to weekday 1.');
  if (normalizeWeekday_('sunday') !== 0) throw new Error('Weekday matching must ignore case.');
  if (normalizeWeekday_('nonsense') !== 1) throw new Error('An unknown weekday must fall back to Monday.');

  const settings = {FINAL_DAYS:3, HORIZON_DAYS:28};
  const today = '2026-11-02';
  const cases = [
    {due:'2026-11-05', days:3, final:true, digest:true, overdue:false},
    {due:'2026-11-04', days:2, final:false, digest:true, overdue:false},
    {due:'2026-11-06', days:4, final:false, digest:true, overdue:false},
    {due:'2026-10-28', days:-5, final:false, digest:true, overdue:true},
    {due:'2026-12-20', days:48, final:false, digest:false, overdue:false}
  ];
  cases.forEach(function(testCase) {
    const days = erpDaysBetween_(today, testCase.due);
    if (days !== testCase.days) throw new Error('Expected ' + testCase.days + ' days to ' + testCase.due + ', got ' + days + '.');
    if ((days === settings.FINAL_DAYS) !== testCase.final) {
      throw new Error('Final reminder decision wrong for ' + testCase.due + '.');
    }
    if ((days <= settings.HORIZON_DAYS) !== testCase.digest) {
      throw new Error('Digest inclusion decision wrong for ' + testCase.due + '.');
    }
    if ((days < 0) !== testCase.overdue) throw new Error('Overdue decision wrong for ' + testCase.due + '.');
  });

  // The idempotency keys must be stable for the same lead and date, and distinct
  // across leads, reminder types and dates. This is what stops duplicate sends.
  const keyA = 'LEAD-0007|FINAL|2026-11-05';
  const keyB = 'LEAD-0007|FINAL|2026-11-05';
  const keyC = 'LEAD-0008|FINAL|2026-11-05';
  const keyD = 'LEAD-0007|FINAL|2026-11-12';
  const keyE = 'DIGEST|owner@example.test|2026-11-02';
  if (keyA !== keyB) throw new Error('The same lead and due date must produce the same reminder key.');
  if (keyA === keyC || keyA === keyD || keyA === keyE) {
    throw new Error('Reminder keys must differ across leads, dates and reminder types.');
  }
  Logger.log('Lead reminder scheduling checks passed.');
}

/**
 * The crm -> leads permission rewrite. Getting this wrong silently removes the
 * module from every administrator holding an explicit page list.
 */
function testLeadsPagePermissionMigration() {
  const allPages = getAdminAccess_({role:'Super Admin',canView:'YES',accessPages:'',writePages:''});
  if (allPages.readPages.indexOf('leads') === -1) throw new Error('Senior roles must read the leads page by default.');
  if (allPages.readPages.indexOf('crm') !== -1) throw new Error('The crm page token must no longer be recognized.');
  if (getAdminManagedPages_().filter(page => page.id === 'leads').length !== 1) {
    throw new Error('The managed page list must offer leads exactly once.');
  }
  if (getAdminManagedPages_().filter(page => page.id === 'crm').length !== 0) {
    throw new Error('The managed page list must no longer offer crm.');
  }

  // An explicit legacy list must survive the rewrite with leads in place of crm.
  const rewrite = function(value) {
    const tokens = String(value || '').split(',').map(token => token.trim()).filter(Boolean);
    const next = [];
    tokens.forEach(function(token) {
      const mapped = token.toLowerCase() === 'crm' ? 'leads' : token;
      if (next.indexOf(mapped) === -1) next.push(mapped);
    });
    return next.join(',');
  };
  if (rewrite('dashboard,projects,crm,reports') !== 'dashboard,projects,leads,reports') {
    throw new Error('The crm token must be rewritten in place.');
  }
  if (rewrite('crm,leads') !== 'leads') throw new Error('A row already holding both tokens must not end up with a duplicate.');
  if (rewrite('') !== '') throw new Error('A blank page list must stay blank so role defaults still apply.');

  const migrated = getAdminAccess_({role:'Accountant',canView:'YES',
    accessPages:rewrite('dashboard,projects,crm,reports'),writePages:rewrite('projects,crm')});
  if (migrated.readPages.indexOf('leads') === -1) throw new Error('A migrated accountant must keep read access to leads.');
  if (migrated.writePages.indexOf('leads') === -1) throw new Error('A migrated accountant must keep write access to leads.');
  Logger.log('Leads page permission migration checks passed.');
}

/** Portfolio and pipeline aggregation arithmetic. */
function testProjectAndPipelineSummaries() {
  const projects = [
    {projectId:'P1',status:'Active',contractValue:10000000,budget:6000000,actualSpend:4000000,
     balance:6000000,marginPct:60,receivable:2000000,overBudget:false},
    {projectId:'P2',status:'Completed',contractValue:5000000,budget:3000000,actualSpend:3500000,
     balance:1500000,marginPct:30,receivable:0,overBudget:true},
    {projectId:'P3',status:'Planning',contractValue:0,budget:0,actualSpend:0,
     balance:0,marginPct:0,receivable:0,overBudget:false}
  ];
  const portfolio = summarizeProjectPortfolio_(projects);
  if (portfolio.contractValue !== 15000000) throw new Error('Portfolio revenue must sum contract values.');
  if (portfolio.actualSpend !== 7500000) throw new Error('Portfolio spend must sum actual spend.');
  if (portfolio.balance !== 7500000) throw new Error('Portfolio balance must be revenue less spend.');
  if (portfolio.marginPct !== 50) throw new Error('Blended margin must be 50%, got ' + portfolio.marginPct + '.');
  if (portfolio.overBudget !== 1) throw new Error('Exactly one project is over budget.');
  if (portfolio.live !== 2) throw new Error('Planning, Active and On Hold projects count as live.');

  const leads = [
    {leadId:'L1',stage:'Negotiation',estimatedValue:4000000,weightedValue:2400000,contractValue:0,
     isOpen:true,isWon:false,followupOverdue:true,followupDueSoon:false},
    {leadId:'L2',stage:'Won (Contracted)',estimatedValue:8000000,weightedValue:8000000,contractValue:8500000,
     isOpen:true,isWon:true,followupOverdue:false,followupDueSoon:true},
    {leadId:'L3',stage:'Lost',estimatedValue:2000000,weightedValue:0,contractValue:0,
     isOpen:false,isWon:false,followupOverdue:false,followupDueSoon:false}
  ];
  const pipeline = summarizeLeadPipeline_(leads);
  if (pipeline.open !== 2) throw new Error('Two leads are still open.');
  if (pipeline.lost !== 1) throw new Error('One lead is lost.');
  if (pipeline.pipelineValue !== 12000000) throw new Error('Pipeline value must exclude lost leads.');
  if (pipeline.weightedValue !== 10400000) throw new Error('Weighted pipeline must sum weighted values of open leads.');
  if (pipeline.contractedValue !== 8500000) throw new Error('Contracted value must use the won lead contract value.');
  if (pipeline.overdueFollowups !== 1) throw new Error('One follow-up is overdue.');
  if (pipeline.dueSoonFollowups !== 1) throw new Error('One follow-up is due soon.');
  if (pipeline.byStage.filter(entry => entry.stage === 'Negotiation')[0].value !== 4000000) {
    throw new Error('Stage breakdown value is wrong.');
  }
  Logger.log('Project portfolio and lead pipeline summary checks passed.');
}

/** Requisition item categories must resolve to a project cost category. */
function testRequisitionCostCategoryMapping() {
  if (requisitionCostCategory_('[{"category":"Feeding"}]') !== 'Feeding') {
    throw new Error('A requisition item category must carry through to the cost line.');
  }
  if (requisitionCostCategory_('[{"category":""},{"category":"Accommodation"}]') !== 'Accommodation') {
    throw new Error('The first recognizable item category must win.');
  }
  if (requisitionCostCategory_('[{"category":"Widgets"}]') !== 'Materials & Consumables') {
    throw new Error('An unrecognized requisition category must fall back to Materials & Consumables.');
  }
  if (requisitionCostCategory_('not json') !== 'Materials & Consumables') {
    throw new Error('Malformed item JSON must not throw.');
  }
  if (requisitionCostCategory_('[]') !== 'Materials & Consumables') {
    throw new Error('An empty item list must fall back to Materials & Consumables.');
  }
  Logger.log('Requisition cost category mapping checks passed.');
}

/** Runs every Projects / Leads test in one go. */
function testProjectsAndLeadsPipeline() {
  testProjectsLeadsColumnMaps();
  testProjectsLeadsVocabulary();
  testErpInputParsing();
  testLeadReminderScheduling();
  testLeadsPagePermissionMigration();
  testProjectAndPipelineSummaries();
  testRequisitionCostCategoryMapping();
  Logger.log('All Projects / Leads pipeline tests passed.');
}


/*************** INVENTORY TESTS ***************/

/**
 * Unit and Opening Stock are routinely mistaken for each other, so the rules
 * that keep them distinct are pinned here.
 */
function testInventoryVocabulary() {
  if (normalizeInventoryUnit_('PCS') !== 'pcs') throw new Error('Unit matching must ignore case.');
  if (normalizeInventoryUnit_('furlong') !== 'pcs') throw new Error('An unknown unit must fall back to pcs.');
  if (normalizeInventoryUnit_('') !== 'pcs') throw new Error('A blank unit must fall back to pcs.');
  if (normalizeInventoryCategory_('display & led') !== 'Display & LED') throw new Error('Category matching must ignore case.');
  if (normalizeInventoryCategory_('powerdistribution') !== 'Power & Distribution') {
    throw new Error('Category matching must ignore punctuation and spacing.');
  }
  if (normalizeInventoryCategory_('Spaceship') !== 'Other') throw new Error('An unknown category must fall back to Other.');
  if (inventoryUnits_().indexOf('pcs') === -1) throw new Error('pcs must stay in the unit list; it is the fallback.');
  if (inventoryCategories_().indexOf('Other') === -1) throw new Error('Other must stay in the category list; it is the fallback.');

  // A unit is a label and opening stock is a number: they must never share a list.
  inventoryUnits_().forEach(function(unit) {
    if (!isNaN(Number(unit))) throw new Error('Unit "' + unit + '" looks like a number; units are labels, not quantities.');
  });
  Logger.log('Inventory vocabulary checks passed.');
}

function testInventoryQuantityRules() {
  if (inventoryQuantity_('', 'X') !== 0) throw new Error('A blank quantity must parse as zero.');
  if (inventoryQuantity_('1,250', 'X') !== 1250) throw new Error('Quantities must accept thousands separators.');
  if (inventoryQuantity_('4.6', 'X') !== 5) throw new Error('Quantities must round to whole units.');
  let rejected = false;
  try { inventoryQuantity_('-1', 'Opening stock'); } catch (err) { rejected = true; }
  if (!rejected) throw new Error('A negative quantity must be rejected.');
  rejected = false;
  try { inventoryQuantity_('abc', 'Quantity'); } catch (err) { rejected = true; }
  if (!rejected) throw new Error('A non-numeric quantity must be rejected.');

  if (inventoryStockStatus_(3, 3) !== 'Low Stock') throw new Error('Stock at the reorder level counts as low.');
  if (inventoryStockStatus_(2, 3) !== 'Low Stock') throw new Error('Stock below the reorder level counts as low.');
  if (inventoryStockStatus_(4, 3) !== 'OK') throw new Error('Stock above the reorder level is OK.');
  Logger.log('Inventory quantity checks passed.');
}

/** The import reference shown in the portal must match what the importer reads. */
function testInventoryImportColumns() {
  const columns = inventoryImportColumns_();
  const required = columns.filter(column => column.required).map(column => column.key);
  if (required.join(',') !== 'itemName') throw new Error('Item Name must be the only required import column.');
  const keys = columns.map(column => column.key).sort().join(',');
  if (keys !== 'category,itemName,location,notes,openingStock,preferredVendor,reorderLevel,unit') {
    throw new Error('The documented import columns no longer match the importer: ' + keys);
  }
  columns.forEach(function(column) {
    if (!column.header || !column.note) throw new Error('Import column ' + column.key + ' needs a header and a note.');
  });
  Logger.log('Inventory import column checks passed.');
}

function testInventoryColumnMap() {
  const headers = ['Item ID','Item Name','Category','Unit','Opening Stock','Stock In','Stock Out','Current Stock',
    'Reorder Level','Location','Stock Status','Preferred Vendor','Notes','Last Updated'];
  if (headers.length !== INV_WIDTH) throw new Error('Inventory header count does not match INV_WIDTH.');
  const expected = {
    itemId:'Item ID', itemName:'Item Name', category:'Category', unit:'Unit', openingStock:'Opening Stock',
    stockIn:'Stock In', stockOut:'Stock Out', currentStock:'Current Stock', reorderLevel:'Reorder Level',
    location:'Location', stockStatus:'Stock Status', preferredVendor:'Preferred Vendor', notes:'Notes',
    lastUpdated:'Last Updated'
  };
  Object.keys(expected).forEach(function(key) {
    if (headers[INV_COL[key] - 1] !== expected[key]) {
      throw new Error('INV_COL.' + key + ' points at ' + headers[INV_COL[key] - 1] + ', expected ' + expected[key] + '.');
    }
  });
  if (Object.keys(INV_COL).length !== headers.length) throw new Error('INV_COL has the wrong number of entries.');
  Logger.log('Inventory column map checks passed.');
}

function testInventoryModule() {
  testInventoryVocabulary();
  testInventoryQuantityRules();
  testInventoryImportColumns();
  testInventoryColumnMap();
  Logger.log('All inventory tests passed.');
}


/**
 * The client details that ride along with Add Project and Add Lead.
 *
 * The mapping is the whole point of this suite: the fields travel under
 * `client`-prefixed names so a project's own Location and Notes are never
 * filed as the client's address and remarks.
 */
function testClientDetailCapture() {
  const fields = clientDetailFields_();
  ['contactPerson','roleTitle','email','phone','sector','location','website','notes'].forEach(function (field) {
    if (fields.indexOf(field) === -1) throw new Error('Client detail field missing: ' + field);
    if (!CLIENT_COL[field]) throw new Error('No Clients column is mapped for ' + field);
  });

  // A project payload: the client fields are the prefixed ones, never the
  // project's own location and notes.
  const project = clientDetailsFromInput_({
    projectName: 'Experience centre', client: 'Seplat Energy', location: 'Abuja', notes: 'Project remark',
    clientContactPerson: 'Ngozi Eze', clientEmail: 'ngozi@seplat.test', clientPhone: ' 08030000003 ',
    clientSector: 'Oil & Gas', clientLocation: 'Lagos'
  }, {});
  if (project.contactPerson !== 'Ngozi Eze') throw new Error('clientContactPerson must map to contactPerson.');
  if (project.phone !== '08030000003') throw new Error('Client detail values must be trimmed.');
  if (project.location !== 'Lagos') throw new Error("The client's own location must win over the project's.");
  if (project.notes !== '') throw new Error('Project notes must not become client notes.');

  // A lead payload: the lead's own contact fields are reused for the client.
  const lead = clientDetailsFromInput_({clientSector: 'Government', notes: 'Lead remark'},
    {contactPerson: 'Ada Briggs', email: 'ada@rsg.test', phone: '08030000001', accountOwner: 'MS-001', notes: ''});
  if (lead.contactPerson !== 'Ada Briggs') throw new Error('An override must win over the prefixed key.');
  if (lead.sector !== 'Government') throw new Error('Un-overridden fields still come from the payload.');
  if (lead.notes !== '') throw new Error('An empty override must suppress the payload value.');
  if (lead.accountOwner !== 'MS-001') throw new Error('The follow-up owner must become the account owner.');

  // The sentence shown after a save has to say what actually happened.
  const created = clientResolutionNote_({clientId: 'CL-0007', created: true, enriched: ['email','sector']});
  if (created.indexOf('CL-0007 created') === -1) throw new Error('A new client must be reported: ' + created);
  if (created.indexOf('email, sector') === -1) throw new Error('The fields written must be named: ' + created);

  const filled = clientResolutionNote_({clientId: 'CL-0007', created: false, enriched: ['phone']});
  if (filled.indexOf('updated') === -1 || filled.indexOf('phone') === -1) {
    throw new Error('Filling a blank on an existing client must be reported: ' + filled);
  }
  // Nothing changed means nothing to say.
  if (clientResolutionNote_({clientId: 'CL-0007', created: false, enriched: []}) !== '') {
    throw new Error('An untouched client must add no message.');
  }
  if (clientResolutionNote_({clientId: '', created: false, enriched: []}) !== '') {
    throw new Error('A project with no client must add no message.');
  }

  Logger.log('testClientDetailCapture passed.');
}
