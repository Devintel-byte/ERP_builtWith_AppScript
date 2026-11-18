function adminCreateStaff(payload) {
  const user = getAdminUser_(payload.adminPin);
  if (!adminCanWritePage_(user,'staff') || user.canOnboard !== 'YES') return {success:false,message:'You do not have staff management permission.'};
  const fullName = String(payload.fullName || '').trim();
  const department = String(payload.department || '').trim();
  const location = String(payload.workLocation || '').trim();
  if (!fullName || !department || !location) return {success:false,message:'Name, department, and work location are required.'};
  const companyEmail = String(payload.companyEmail || '').trim();
  const personalEmail = String(payload.personalEmail || '').trim();
  if ((companyEmail && !isValidEmail_(companyEmail)) || (personalEmail && !isValidEmail_(personalEmail))) {
    return {success:false,message:'Enter a valid email address.'};
  }
  const salary = Number(payload.monthlySalary || 0);
  const workDays = Number(payload.workDaysPerMonth || 22);
  if (!Number.isFinite(salary) || salary < 0 || !Number.isFinite(workDays) || workDays <= 0) {
    return {success:false,message:'Salary and monthly work days must be valid non-negative/positive numbers.'};
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const staffId = String(payload.staffId || '').trim() || generateNextStaffId_();
    if (getStaffByIdAnyStatus_(staffId)) return {success:false,message:'Staff ID already exists: '+staffId};
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET);
    const row = [
      staffId,fullName,department,location,'','',companyEmail,personalEmail,
      String(payload.requireGps || 'YES'),String(payload.requireDevice || 'YES'),
      salary,Number(payload.latePenalty || 1000),Number(payload.leaveAllowance || 0),
      String(payload.phone || '').trim(),String(payload.notes || '').trim(),
      payload.dateJoined ? new Date(String(payload.dateJoined)+'T12:00:00') : new Date(),
      String(payload.employmentType || 'Permanent'),'Active','',
      String(payload.designation || '').trim(),String(payload.manager || '').trim(),
      String(payload.bankName || '').trim(),String(payload.accountNumber || '').trim(),
      String(payload.accountName || '').trim(),Number(payload.annualLeaveEntitlement || 20),
      Number(payload.sickLeaveEntitlement || 20),workDays,'',
      String(payload.canSubmitRequisition || 'YES')
    ];
    const targetRow = sheet.getLastRow()+1;
    sheet.getRange(targetRow,1,1,row.length).setValues([row]);
    sheet.getRange(targetRow,5).setNumberFormat('@');
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_MOVEMENT_LOG_SHEET)
      .appendRow([new Date(),staffId,fullName,'Joined','', 'Active',user.name,'Created from Admin Portal']);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'CREATE','Staff',staffId,'New staff profile created');
    invalidatePublicStaffOptionsCache_();
    syncAccountantViewRows();
    return {success:true,message:'Staff profile created for '+fullName+' ('+staffId+').'};
  } finally {
    lock.releaseLock();
  }
}

function adminUpdateStaff(payload) {
  const user = getAdminUser_(payload.adminPin);
  if (!adminCanWritePage_(user,'staff') || user.canOnboard !== 'YES') return {success:false,message:'You do not have staff management permission.'};
  const staff = getStaffByIdAnyStatus_(payload.staffId);
  if (!staff) return {success:false,message:'Staff record not found.'};

  const fullName = String(payload.fullName || '').trim();
  const department = String(payload.department || '').trim();
  const location = String(payload.workLocation || '').trim();
  if (!fullName || !department || !location) return {success:false,message:'Name, department, and work location are required.'};
  const companyEmail = String(payload.companyEmail || '').trim();
  const personalEmail = String(payload.personalEmail || '').trim();
  if ((companyEmail && !isValidEmail_(companyEmail)) || (personalEmail && !isValidEmail_(personalEmail))) {
    return {success:false,message:'Enter a valid email address.'};
  }
  const salary = Number(payload.monthlySalary || 0);
  const workDays = Number(payload.workDaysPerMonth || 22);
  if (!Number.isFinite(salary) || salary < 0 || !Number.isFinite(workDays) || workDays <= 0) {
    return {success:false,message:'Salary and monthly work days must be valid non-negative/positive numbers.'};
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET);
    const row = sheet.getRange(staff.row,1,1,29).getValues()[0];
    const oldName = String(row[1] || '');
    const fields = {
      1:fullName,2:department,3:location,6:companyEmail,7:personalEmail,
      8:String(payload.requireGps || 'YES'),9:String(payload.requireDevice || 'YES'),
      10:salary,11:Number(payload.latePenalty || 1000),12:Number(payload.leaveAllowance || 0),
      13:String(payload.phone || '').trim(),14:String(payload.notes || '').trim(),
      15:payload.dateJoined ? new Date(String(payload.dateJoined)+'T12:00:00') : row[15],
      16:String(payload.employmentType || 'Permanent'),19:String(payload.designation || '').trim(),
      20:String(payload.manager || '').trim(),21:String(payload.bankName || '').trim(),
      22:String(payload.accountNumber || '').trim(),23:String(payload.accountName || '').trim(),
      24:Number(payload.annualLeaveEntitlement || 20),25:Number(payload.sickLeaveEntitlement || 20),
      26:workDays,28:String(payload.canSubmitRequisition || 'YES')
    };
    Object.keys(fields).forEach(column => { row[Number(column)] = fields[column]; });
    sheet.getRange(staff.row,1,1,29).setValues([row]);
    sheet.getRange(staff.row,5).setNumberFormat('@');
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_MOVEMENT_LOG_SHEET)
      .appendRow([new Date(),staff.id,fullName,'Profile Updated',String(row[17] || ''),String(row[17] || ''),user.name,'Updated from Admin Portal; previous name: '+oldName]);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'UPDATE','Staff',staff.id,'Staff profile fields updated');
    invalidatePublicStaffOptionsCache_();
    syncAccountantViewRows();
    return {success:true,message:'Staff profile updated for '+fullName+'.'};
  } finally {
    lock.releaseLock();
  }
}

function adminSetStaffStatus(payload) {
  const user = getAdminUser_(payload.adminPin);
  if (!adminCanWritePage_(user,'staff') || user.canOnboard !== 'YES') return {success:false,message:'You do not have staff management permission.'};
  const targetStatus = String(payload.status || '');
  if (targetStatus !== 'Active' && targetStatus !== 'Inactive') return {success:false,message:'Status must be Active or Inactive.'};
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const staff = getStaffByIdAnyStatus_(payload.staffId);
    if (!staff) return {success:false,message:'Staff record not found.'};
    if (String(staff.status) === targetStatus) return {success:true,message:staff.name+' is already '+targetStatus.toLowerCase()+'.'};
    const now = new Date();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET);
    sheet.getRange(staff.row,18).setValue(targetStatus);
    sheet.getRange(staff.row,19).setValue(targetStatus === 'Inactive' ? now : '');
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_MOVEMENT_LOG_SHEET)
      .appendRow([now,staff.id,staff.name,targetStatus === 'Inactive' ? 'Disabled' : 'Reactivated',staff.status,targetStatus,user.name,String(payload.reason || '')]);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,targetStatus === 'Inactive' ? 'DISABLE' : 'REACTIVATE','Staff',staff.id,String(payload.reason || ''));
    invalidatePublicStaffOptionsCache_();
    syncAccountantViewRows();
    return {success:true,message:staff.name+' has been '+(targetStatus === 'Inactive' ? 'disabled' : 'reactivated')+'.'};
  } finally {
    lock.releaseLock();
  }
}

function adminDeleteStaff(payload) {
  const user = getAdminUser_(payload.adminPin);
  if (!adminCanWritePage_(user,'staff') || user.canOnboard !== 'YES') return {success:false,message:'You do not have staff management permission.'};
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const staff = getStaffByIdAnyStatus_(payload.staffId);
    if (!staff) return {success:false,message:'Staff record not found.'};
    const references = findStaffReferences_(staff.id);
    if (references.length) {
      return {success:false,message:'Permanent deletion is blocked because this staff record has linked history ('+references.join(', ')+'). Disable the staff member instead to preserve audit and payroll records.'};
    }
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'DELETE','Staff',staff.id,'Staff record permanently removed after confirming no linked records.');
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET).deleteRow(staff.row);
    invalidatePublicStaffOptionsCache_();
    syncAccountantViewRows();
    return {success:true,message:staff.name+' was permanently deleted.'};
  } finally {
    lock.releaseLock();
  }
}

function findStaffReferences_(staffId) {
  const checks = [
    [LOG_SHEET,2],
    [REQ_SHEET,3],
    ['Expense Claims',3],
    ['Leave Requests',3],
    ['Assets',10]
  ];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return checks.filter(function(check) {
    const sheet = ss.getSheetByName(check[0]);
    return sheet && sheet.getLastRow() > 1 &&
      sheet.getRange(2,check[1],sheet.getLastRow()-1,1)
        .createTextFinder(String(staffId))
        .matchCase(true)
        .matchEntireCell(true)
        .findNext();
  }).map(check => check[0]);
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || ''));
}
