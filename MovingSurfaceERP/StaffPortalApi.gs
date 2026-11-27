function staffPortalGetData(payload) {
  try {
    const staff = validateStaffPortalLogin_(payload.staffId, payload.pin, payload.deviceToken);
    if (payload.logSignIn === true) {
      recordActivity_('Staff',staff.name,staff.id,'SIGN IN','Staff Portal','', 'Successful staff portal sign-in');
    }
    const data = getStaffPortalData_(staff);
    return Object.assign({success:true, staff:{id:staff.id,name:staff.name,department:staff.department,designation:staff.designation,location:staff.location}}, data);
  } catch (err) {
    return {success:false, message:'Staff portal error: ' + (err.message || err)};
  }
}

function getPublicStaffOptions() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'PUBLIC_ACTIVE_STAFF_OPTIONS';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const options = getActiveStaff_().map(function(staff) {
    return {
      id:staff.id,
      name:staff.name,
      pinSet:isValidPin_(normalizePin_(staff.pin)) && !isStartingPin_(staff.id, staff.pin),
      deviceSet:Boolean(staff.deviceToken)
    };
  });
  cache.put(cacheKey, JSON.stringify(options), 60);
  return options;
}

function invalidatePublicStaffOptionsCache_() {
  CacheService.getScriptCache().remove('PUBLIC_ACTIVE_STAFF_OPTIONS');
}

/**
 * Signs a staff member in to the portal, requisitions, leave and expenses.
 *
 * The Staff ID and PIN are the whole credential, from any device and any
 * location. Only marking attendance checks the registered device and GPS.
 * Clients used to send a device token here as well; it is ignored.
 */
function validateStaffPortalLogin_(staffId, pin) {
  const staff = getStaffById_(staffId);
  if (!staff) throw new Error('Staff not found or inactive.');
  if (!staff.pin) throw new Error('Your PIN is not set. Please use attendance first to create your PIN, or contact Admin.');
  const suppliedPin = String(pin == null ? '' : pin).trim();
  if (isPinLocked_(staff.id)) throw new Error(PIN_LOCKED_MESSAGE);
  if (!isValidPin_(suppliedPin) || normalizePin_(staff.pin) !== suppliedPin) {
    if (isValidPin_(suppliedPin)) recordPinFailure_(staff.id);
    throw new Error('Invalid Staff ID or PIN. Enter all four digits, including any leading zero.');
  }
  clearPinFailures_(staff.id);
  return staff;
}

function getStaffPortalData_(staff) {
  const reqs = getAllRequisitions_().filter(r => String(r.staffId) === String(staff.id));
  const leave = getERPLeaveRequests_().filter(r => String(r.staffId) === String(staff.id));
  const expenses = getERPExpenseClaims_().filter(r => String(r.staffId) === String(staff.id));
  const assets = getERPAssets_().filter(r => String(r.assignedToStaffId) === String(staff.id));
  const attendance = getRecentAttendanceForStaff_(staff.id, 20);
  const summary = {
    pendingReq: reqs.filter(r => String(r.status).indexOf('Pending') >= 0 || String(r.status).indexOf('Awaiting') >= 0).length,
    openLeave: leave.filter(r => ['Pending Approval','Approved'].includes(String(r.status))).length,
    openExpense: expenses.filter(r => String(r.status) !== 'Paid / Closed' && String(r.status) !== 'Rejected').length,
    assignedAssets: assets.length
  };
  return {requisitions:reqs.map(requisitionClientDto_), leaveRequests:leave, expenses:expenses, assets:assets,
    attendance:attendance, summary:summary, refData:staffReferenceData_()};
}

function getRecentAttendanceForStaff_(staffId, limit) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const settings = getSettings_();
  const matches = sh.getRange(2,2,sh.getLastRow()-1,1)
    .createTextFinder(String(staffId))
    .matchCase(true)
    .matchEntireCell(true)
    .findAll();
  if (!matches.length) return [];
  const matchRows = matches.map(cell => cell.getRow());
  const firstRow = Math.min.apply(null, matchRows.slice(-Number(limit || 20)));
  const lastRow = Math.max.apply(null, matchRows);
  const values = sh.getRange(firstRow,1,lastRow-firstRow+1,25).getValues();
  return values.filter(r => String(r[1]) === String(staffId)).slice(-Number(limit || 20)).reverse().map(r => ({
    dateKey: formatAttendanceDate_(r[0], settings),
    action: String(r[4] || ''),
    time: formatAttendanceTime_(r[5], r[6], settings),
    checkoutTime:String(r[23] || (attendanceCheckoutType_(r[4],'') ? formatAttendanceTime_(r[5],r[6],settings) : '')),
    checkoutType:attendanceCheckoutType_(r[4],r[24]),
    status: normalizeAttendanceStatusForDisplay_(r[13]),
    location: String(r[10] || '')
  }));
}

function formatAttendanceDate_(dateValue, settings) {
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
    return Utilities.formatDate(dateValue, settings.TIMEZONE, 'yyyy-MM-dd');
  }
  return String(dateValue || '');
}

function formatAttendanceTime_(timestampValue, timeValue, settings) {
  // Prefer the real timestamp column. This prevents time-only spreadsheet cells
  // from displaying as Dec 31 1899 in the browser.
  if (timestampValue instanceof Date && !isNaN(timestampValue.getTime())) {
    return Utilities.formatDate(timestampValue, settings.TIMEZONE, 'hh:mm:ss a');
  }
  if (timeValue instanceof Date && !isNaN(timeValue.getTime())) {
    return Utilities.formatDate(timeValue, settings.TIMEZONE, 'hh:mm:ss a');
  }
  const s = String(timeValue || '').trim();
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(s)) return s;
  return s;
}

function normalizeAttendanceStatusForDisplay_(status) {
  status = String(status || '');
  // Old rows may have incorrectly stored Absent for a successful Check In.
  // New rows will no longer do this, but we keep this defensive display cleanup.
  return status;
}

function staffSubmitLeaveRequest(p) {
  try {
    const staff = validateStaffPortalLogin_(p.staffId, p.pin, p.deviceToken);
    if (!p.startDate || !p.endDate) return {success:false, message:'Start and end date are required.'};
    const start = new Date(p.startDate), end = new Date(p.endDate);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) return {success:false, message:'Invalid leave dates.'};
    if (end < start) return {success:false, message:'End date cannot be before start date.'};
    const days = Math.max(1, Math.round((end-start)/86400000)+1);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leave Requests');
    const id = erpNextId_('Leave Requests','LEV',2);
    sh.appendRow([new Date(),id,staff.id,staff.name,staff.department,String(p.leaveType||''),p.startDate,p.endDate,days,String(p.reason||''),'Pending Approval','','','Submitted by staff portal',new Date()]);
    recordActivity_('Staff',staff.name,staff.id,'CREATE','Leave Request',id,'Leave request submitted');
    return {success:true, message:'Leave request submitted successfully.\n\nLeave ID: ' + id + '\nStatus: Pending Approval'};
  } catch (err) { return {success:false, message:err.message}; }
}

function staffSubmitExpenseClaim(p) {
  try {
    const staff = validateStaffPortalLogin_(p.staffId, p.pin, p.deviceToken);
    const amount = Number(p.amount || 0);
    if (amount <= 0) return {success:false, message:'Amount is required.'};
    if (!p.purpose) return {success:false, message:'Purpose is required.'};
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expense Claims');
    const id = erpNextId_('Expense Claims','EXP',2);
    sh.appendRow([new Date(),id,staff.id,staff.name,staff.department,p.expenseDate||'',normalizeProjectCostCategory_(p.category),String(p.project||''),amount,'Pending Approval','', '', 'Not Submitted','',String(p.receipt||''),String(p.purpose||''),'',new Date(),resolveSubmittedProjectId_(p.projectId, p.project)]);
    recordActivity_('Staff',staff.name,staff.id,'CREATE','Expense Claim',id,'Expense claim submitted');
    return {success:true, message:'Expense claim submitted successfully.\n\nClaim ID: ' + id + '\nStatus: Pending Approval'};
  } catch (err) { return {success:false, message:err.message}; }
}

function testStaffPortalLogin() {
  Logger.log('Use a registered staff device token to test the staff portal.');
}

/**
 * The vocabularies the staff-facing forms offer.
 *
 * Expense and requisition categories come from the project cost list, so a
 * staff claim lands in a category the project report already understands
 * instead of collapsing to Miscellaneous on the way in.
 */
function staffReferenceData_() {
  return {
    costCategories: projectCostCategories_(),
    leaveTypes: ['Annual','Sick','Emergency','Compassionate','Unpaid'],
    priorities: ['Normal','Urgent','Low']
  };
}

/** Reference data for the standalone requisition page, which has no login yet. */
function getRequisitionReferenceData() {
  return staffReferenceData_();
}
