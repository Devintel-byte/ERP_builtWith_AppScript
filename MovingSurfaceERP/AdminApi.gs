function adminDecisionRequisition(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = getAdminUser_(payload.adminPin);
    if (!user.valid || user.canView !== 'YES') return {success:false, message:'Invalid PIN.'};
    if (!adminCanWritePage_(user,'requisitions')) return {success:false,message:'You do not have permission to update requisitions.'};
    if (!adminCanApprovePage_(user,'requisitions')) return {success:false, message:'You do not have permission to approve/reject requisitions.'};

    const row = getRequisitionById_(payload.reqId);
    if (!row) return {success:false, message:'Request not found.'};
    if (row.status !== 'Pending MD Approval') return {success:false, message:'Request is not pending MD approval. Current status: '+row.status};

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQ_SHEET);
    const now = new Date();
    const decision = String(payload.decision || '');
    if (decision !== 'Approved' && decision !== 'Rejected') return {success:false, message:'Choose Approved or Rejected.'};
    const newStatus = decision === 'Approved' ? 'Approved - Awaiting Accounts' : 'Rejected by MD';

    sh.getRange(row.row, 11).setValue(newStatus);
    sh.getRange(row.row, 12).setValue(now);
    sh.getRange(row.row, 13).setValue(user.name);
    sh.getRange(row.row, 14).setValue(String(payload.comment || ''));
    sh.getRange(row.row, 22).setValue(now);

    logReqAudit_(row.reqId, decision, user.name, row.status, newStatus, String(payload.comment || ''));
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'UPDATE','Requisition',row.reqId,row.status+' -> '+newStatus);

    if (decision === 'Approved') {
      const accountsQueued = sendApprovedRequisitionToAccounts_(row.reqId);
      const staffQueued = notifyStaffRequisitionStatus_(row.reqId, newStatus);
      if (accountsQueued) sh.getRange(row.row, 15).setValue(now);
      if (staffQueued) sh.getRange(row.row, 20).setValue(now);
      const costPosted = postRequisitionProjectCost_(row, 'Committed', user.name, '');
      return {success:true, message:'Request '+row.reqId+' approved.'+(accountsQueued ? ' Accounts notification queued.' : ' Accounts notification could not be queued; check Mail Queue.')+(staffQueued ? ' Staff notification queued.' : '')+(costPosted ? ' '+costPosted : '')};
    }

    const staffQueued = notifyStaffRequisitionStatus_(row.reqId, newStatus);
    if (staffQueued) sh.getRange(row.row, 20).setValue(now);
    return {success:true, message:'Request '+row.reqId+' rejected.'+(staffQueued ? ' Staff notification queued.' : '')};
  } finally {
    lock.releaseLock();
  }
}

/**
 * Posts (or updates) the project cost line representing a requisition.
 *
 * Called on approval with status Committed and again on payment with status
 * Paid. The upsert key is the Request ID, so one line walks through its
 * lifecycle instead of being appended twice.
 */
function postRequisitionProjectCost_(row, paymentStatus, actorName, paymentReference) {
  try {
    if (!row.projectId) return '';
    return postProjectCostFromSource_({
      projectId:row.projectId,
      sourceModule:'Requisition',
      sourceRefId:row.reqId,
      costDate:erpToday_(),
      category:requisitionCostCategory_(row.itemsJson),
      description:'Requisition '+row.reqId+' - '+(row.itemsSummary ? String(row.itemsSummary).split('\n')[0] : String(row.purpose || '')),
      vendor:'',
      plannedAmount:Number(row.total || 0),
      actualAmount:paymentStatus === 'Paid' ? Number(row.total || 0) : 0,
      paymentStatus:paymentStatus,
      paymentReference:paymentReference,
      recordedBy:actorName
    });
  } catch (err) {
    Logger.log('Requisition cost line not posted for '+row.reqId+': '+(err.stack || err.message));
    return '';
  }
}

function adminMarkRequisitionPaid(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = getAdminUser_(payload.adminPin);
    if (!user.valid || user.canView !== 'YES') return {success:false, message:'Invalid PIN.'};
    if (!adminCanWritePage_(user,'requisitions')) return {success:false,message:'You do not have permission to update requisitions.'};
    if (!adminCanDisbursePage_(user,'requisitions')) return {success:false, message:'You do not have accounts/disbursement permission.'};

    const row = getRequisitionById_(payload.reqId);
    if (!row) return {success:false, message:'Request not found.'};
    if (row.status !== 'Approved - Awaiting Accounts') return {success:false, message:'Request is not awaiting accounts. Current status: '+row.status};

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQ_SHEET);
    const now = new Date();
    sh.getRange(row.row, 11).setValue('Closed - Paid / Disbursed');
    sh.getRange(row.row, 16).setValue('Paid / Disbursed');
    sh.getRange(row.row, 17).setValue(now);
    sh.getRange(row.row, 18).setValue(user.name);
    sh.getRange(row.row, 19).setValue(String(payload.reference || ''));
    sh.getRange(row.row, 22).setValue(now);

    logReqAudit_(row.reqId, 'Marked Paid / Disbursed', user.name, row.status, 'Closed - Paid / Disbursed', String(payload.reference || ''));
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'UPDATE','Requisition',row.reqId,'Marked paid/disbursed');
    const staffQueued = notifyStaffRequisitionStatus_(row.reqId, 'Paid / Disbursed');
    if (staffQueued) sh.getRange(row.row, 20).setValue(now);
    const costPosted = postRequisitionProjectCost_(row, 'Paid', user.name, String(payload.reference || ''));

    return {success:true, message:'Request '+row.reqId+' marked as paid/disbursed.'+(staffQueued ? ' Staff notification queued.' : '')+(costPosted ? ' '+costPosted : '')};
  } finally {
    lock.releaseLock();
  }
}

function adminAddNewStaff(payload) {
  const user = getAdminUser_(payload.adminPin);
  if (!user.valid) return {success:false, message:'Invalid PIN.'};
  if (!adminCanWritePage_(user,'staff')) return {success:false,message:'You do not have permission to manage staff accounts.'};
  if (user.canOnboard !== 'YES') return {success:false, message:'You do not have onboarding permission.'};

  const fullName = String(payload.fullName || '').trim();
  const department = String(payload.department || '').trim();
  if (!fullName || !department) return {success:false, message:'Full name and department are required.'};

  let staffId = String(payload.staffId || '').trim();
  if (!staffId) staffId = generateNextStaffId_();
  if (getStaffByIdAnyStatus_(staffId)) return {success:false, message:'Staff ID already exists: '+staffId};

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET);
  sh.appendRow([
    staffId, fullName, department, String(payload.workLocation || 'Abuja'), '', '', String(payload.companyEmail || ''), '',
    'YES','YES',Number(payload.monthlySalary || 0),1000,'','','Added from Admin Portal',new Date(),
    'Permanent','Active','',String(payload.designation || ''),String(payload.manager || ''),
    '','','',20,20,22,'','YES'
  ]);

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_MOVEMENT_LOG_SHEET).appendRow([new Date(), staffId, fullName, 'Joined', '', 'Active', user.name, 'Added from Admin Portal']);
  recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'CREATE','Staff',staffId,'New staff profile created');
  invalidatePublicStaffOptionsCache_();
  syncAccountantViewRows();

  return {success:true, message:'New staff added successfully.\n\nStaff ID: '+staffId+'\nName: '+fullName+'\n\nStaff can now use the attendance link to create PIN and register device.'};
}

function adminResetStaff(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = getAdminUser_(payload.adminPin);
    if (!user.valid) return {success:false, message:'Invalid PIN.'};
    if (!adminCanWritePage_(user,'staff')) return {success:false,message:'You do not have permission to manage staff accounts.'};
    if (user.canReset !== 'YES') return {success:false, message:'You do not have reset permission.'};

    const staff = getStaffById_(payload.staffId);
    if (!staff) return {success:false, message:'Staff not found.'};
    if (!['Forgot PIN','New Device','Forgot PIN + New Device'].includes(String(payload.resetType || ''))) {
      return {success:false, message:'Select a valid reset type.'};
    }

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_SHEET);
    let pinCleared = 'NO', deviceCleared = 'NO';
    if (payload.resetType === 'Forgot PIN' || payload.resetType === 'Forgot PIN + New Device') {
      sh.getRange(staff.row, 5).setValue('');
      pinCleared = 'YES';
    }
    if (payload.resetType === 'New Device' || payload.resetType === 'Forgot PIN + New Device') {
      sh.getRange(staff.row, 6).setValue('');
      deviceCleared = 'YES';
    }
    revokeStaffRecoveryRequests_(staff.id);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_RESET_LOG_SHEET).appendRow([new Date(), staff.id, staff.name, payload.resetType, pinCleared, deviceCleared, user.name, payload.reason || '', 'Reset completed from Admin Portal']);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'UPDATE','Staff',staff.id,'Credential/device reset: '+String(payload.resetType || ''));
    invalidatePublicStaffOptionsCache_();
    return {success:true, message:'Reset completed for '+staff.name+'\nPIN cleared: '+pinCleared+'\nDevice cleared: '+deviceCleared};
  } finally {
    lock.releaseLock();
  }
}

/*************** EMAIL APPROVAL ***************/

function adminGetPortalData(payload) {
  try {
    const user = getAdminUser_(payload.adminPin);
    if (!user.valid) return {success:false, message:'Invalid PIN or inactive admin user.'};
    if (user.canView !== 'YES') return {success:false, message:'This user does not have dashboard view permission.'};
    if (!adminCanReadPage_(user,'dashboard')) return {success:false,message:'This user does not have dashboard access.'};

    const allStaff = getAllStaff_();
    const staffRaw = allStaff.filter(s => String(s.status).toLowerCase() === 'active');
    const staffOptions = staffRaw.map(s => `<option value="${escapeHtml_(s.id)}">${escapeHtml_(s.name)} - ${escapeHtml_(s.id)}</option>`).join('');
    const staff = adminCanReadPage_(user,'staff') ? allStaff.map(s => ({
      id:s.id,name:s.name,department:s.department,location:s.location,
      pinSet:isValidPin_(normalizePin_(s.pin)) && !isStartingPin_(s.id,s.pin)?'YES':'NO',deviceSet:s.deviceToken?'YES':'NO',status:s.status||'Active',
      companyEmail:s.companyEmail,personalEmail:s.personalEmail,requireGps:s.requireGps,requireDevice:s.requireDevice,
      monthlySalary:s.monthlySalary,latePenalty:s.latePenalty,leaveAllowance:s.leaveAllowance,phone:s.phone,notes:s.notes,
      dateJoined:erpDate_(s.dateJoined),employmentType:s.employmentType,exitDate:erpDate_(s.exitDate),
      designation:s.designation,manager:s.manager,bankName:s.bankName,accountNumber:s.accountNumber,
      accountName:s.accountName,annualLeaveEntitlement:s.annualLeaveEntitlement,
      sickLeaveEntitlement:s.sickLeaveEntitlement,workDaysPerMonth:s.workDaysPerMonth,
      canSubmitRequisition:s.canSubmitRequisition
    })) : [];
    if (payload.logSignIn === true) {
      recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'SIGN IN','Admin Portal','', 'Successful admin portal sign-in');
    }

    const canApproveReq = adminCanApprovePage_(user,'requisitions');
    const canApproveExpense = adminCanApprovePage_(user,'expenses');
    const canApproveLeave = adminCanApprovePage_(user,'leave');
    const canDisburseReq = adminCanDisbursePage_(user,'requisitions');
    const canDisburseExpense = adminCanDisbursePage_(user,'expenses');

    const requisitions = safeRows_(getAllRequisitions_()).map(requisitionClientDto_);
    const liveAttendance = getLiveTodayAttendance_(staffRaw);
    const payrollRows = safeRows_(getPayrollRows_(staffRaw));

    const inventory = getERPInventoryItems_();
    const assets = getERPAssets_();
    const expenses = getERPExpenseClaims_();
    const leaveRequests = getERPLeaveRequests_();
    const projects = getERPProjects_();
    const clients = getERPClients_();
    const leads = getERPLeads_();

    const inventoryOptions = inventory.map(i => `<option value="${escapeHtml_(i.itemId)}">${escapeHtml_(i.itemName)} - ${escapeHtml_(i.itemId)}</option>`).join('');

    const canRead = page => adminCanReadPage_(user,page);
    const fieldDuty = canRead('fieldduty') ? fieldDutyRecordsForUser_(user) : [];
    const summary = {
      activeStaff:canRead('staff')?staffRaw.length:null,
      signedIn:canRead('attendance')?liveAttendance.signedInCount:null,
      overtimeDays:canRead('attendance')?payrollRows.reduce((sum,row) => sum+Number(row.overtimeDays || 0),0):null,
      pendingReq:canRead('requisitions')?requisitions.filter(r => r.status === 'Pending MD Approval').length:null,
      awaitingAccounts:canRead('requisitions')?requisitions.filter(r => r.status === 'Approved - Awaiting Accounts').length:null,
      openExpenses:canRead('expenses')?expenses.filter(r => r.status === 'Pending Approval' || r.status === 'Approved - Awaiting Accounts').length:null,
      assetsAssigned:canRead('assets')?assets.filter(r => r.status === 'Assigned').length:null,
      lowStock:canRead('inventory')?inventory.filter(r => r.stockStatus === 'Low Stock').length:null,
      activeProjects:canRead('projects')?projects.filter(r => r.status === 'Active').length:null,
      openLeads:canRead('leads')?leads.filter(r => r.isOpen).length:null,
      overdueFollowups:canRead('leads')?leads.filter(r => r.followupOverdue).length:null,
      dueSoonFollowups:canRead('leads')?leads.filter(r => r.followupDueSoon).length:null,
      pipelineValue:canRead('leads')?leads.filter(r => r.isOpen).reduce((sum,r) => sum+r.estimatedValue,0):null,
      weightedPipeline:canRead('leads')?Math.round(leads.filter(r => r.isOpen).reduce((sum,r) => sum+r.weightedValue,0)):null,
      projectBalance:canRead('projects')?projects.reduce((sum,r) => sum+r.balance,0):null,
      projectReceivable:canRead('projects')?projects.reduce((sum,r) => sum+r.receivable,0):null,
      projectsOverBudget:canRead('projects')?projects.filter(r => r.overBudget).length:null,
      fieldDutyPending:canRead('fieldduty')?fieldDuty.filter(r => r.status === 'Pending Approval').length:null,
      onFieldDuty:canRead('fieldduty')?fieldDuty.filter(r => r.isApproved && r.phase === 'Active').length:null
    };

    const clientUser = {name:user.name,role:user.role,canView:user.canView,canOverride:user.canOverride,
      canReset:user.canReset,canOnboard:user.canOnboard,access:user.access};
    return {
      success:true,
      user:clientUser,
      canApproveReq:canApproveReq,
      canApproveExpense:canApproveExpense,
      canApproveLeave:canApproveLeave,
      canDisburseReq:canDisburseReq,
      canDisburseExpense:canDisburseExpense,
      staffOptions:canRead('staff') || ['assets','expenses','leave','fieldduty'].some(page => adminCanWritePage_(user,page)) ? staffOptions : '',
      inventoryOptions:canRead('inventory') ? inventoryOptions : '',
      staff:staff,
      webAppUrl:getWebAppRouteUrl_('attendance') || '',
      sheetLinks:getSpreadsheetSheetLinks_(user),
      liveAttendance:canRead('attendance') ? liveAttendance : {totalStaff:0,signedInCount:0,checkedOutCount:0,overtimeCheckoutCount:0,notSignedInCount:0,date:'',currentTime:'',signedIn:[],checkedOut:[],notSignedIn:[]},
      payrollRows:canRead('reports') ? payrollRows : [],
      requisitions:canRead('requisitions') ? requisitions : [],
      inventory:canRead('inventory') ? inventory : [],
      inventoryRefData:canRead('inventory') ? {
        categories:inventoryCategories_(),units:inventoryUnits_(),
        transactionTypes:inventoryTransactionTypes_(),importColumns:inventoryImportColumns_()
      } : null,
      assets:canRead('assets') ? assets : [],
      expenses:canRead('expenses') ? expenses : [],
      leaveRequests:canRead('leave') ? leaveRequests : [],
      projects:canRead('projects') ? projects : [],
      projectSummary:canRead('projects') ? summarizeProjectPortfolio_(projects) : null,
      projectOptions:canRead('projects') ? getProjectOptionList_() : [],
      projectStatuses:canRead('projects') ? projectStatuses_() : [],
      costCategories:canRead('projects') ? projectCostCategories_() : [],
      clients:canRead('leads') ? clients : [],
      leads:canRead('leads') ? leads : [],
      leadSummary:canRead('leads') ? summarizeLeadPipeline_(leads) : null,
      leadRefData:canRead('leads') ? {
        stages:leadStages_(),sources:leadSources_(),cadences:leadReminderCadences_(),
        activityTypes:leadActivityTypes_(),relationshipStatuses:clientRelationshipStatuses_()
      } : null,
      fieldDuty:fieldDuty,
      fieldDutyProjects:canRead('fieldduty') ? fieldDutyProjectOptions_(projects) : [],
      summary:summary
    };
  } catch (err) {
    return {success:false, message:'Admin portal error: ' + (err.stack || err.message)};
  }
}

function adminGetUserManagementData(payload) {
  const user = getAdminUser_(payload && payload.adminPin);
  if (!adminCanManageUsers_(user)) {
    return {success:false,message:'Only Super Admin and Admin can manage portal users.'};
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET);
  if (!sheet) return {success:false,message:'Admin Users sheet is missing. Run setup / repair first.'};
  const staff = getAllStaff_().map(person => ({
    staffId:String(person.id || ''),name:String(person.name || ''),
    department:String(person.department || ''),status:String(person.status || 'Active'),
    email:String(person.companyEmail || '')
  }));
  const accounts = sheet.getLastRow() < 2 ? [] :
    sheet.getRange(2,1,sheet.getLastRow()-1,Math.max(12,sheet.getLastColumn())).getValues()
      .map((row,index)=>({values:row,rowNumber:index+2}))
      .filter(account=>account.values.some(value=>String(value || '').trim()))
      .map(entry => {
        const row=entry.values;
        const account = {
          row:entry.rowNumber,name:String(row[0] || ''),staffId:String(row[8] || '').trim(),
          role:String(row[3] || ''),roleLabel:canonicalAdminRoleLabel_(row[3]),active:String(row[2] || '').toUpperCase() === 'YES',
          canView:String(row[4] || '').toUpperCase() === 'YES',
          canOverride:String(row[5] || '').toUpperCase() === 'YES',
          canReset:String(row[6] || '').toUpperCase() === 'YES',
          canOnboard:String(row[7] || '').toUpperCase() === 'YES',
          email:String(row[9] || ''),accessPages:String(row[10] || ''),
          writePages:String(row[11] || '')
        };
        const access = getAdminAccess_({
          role:account.role,canView:'YES',
          accessPages:account.accessPages,writePages:account.writePages
        });
        account.readPages=access.readPages.filter(page => page !== 'users');
        account.writePages=access.writePages.filter(page => page !== 'users');
        return account;
      });
  return {
    success:true,staff:staff,accounts:accounts,pages:getAdminManagedPages_(),
    roles:getAdminManagedRoleLabels_(accounts)
  };
}

function adminSaveUserPrivileges(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = getAdminUser_(payload && payload.adminPin);
    if (!adminCanManageUsers_(user)) {
      return {success:false,message:'Only Super Admin and Admin can manage portal users.'};
    }
    const staffId = String(payload.staffId || '').trim();
    const staff = getStaffByIdAnyStatus_(staffId);
    if (!staff) return {success:false,message:'Select a staff member who exists in Staff List.'};
    const role = String(payload.role || '').trim();
    if (!role || role.length > 60) return {success:false,message:'Enter a role name of 1 to 60 characters.'};
    const email = String(payload.email || '').trim();
    if (email && !isValidEmail_(email)) return {success:false,message:'Enter a valid email address or leave Email blank.'};
    const canView = payload.canView === true || String(payload.canView).toUpperCase() === 'YES';
    const active = payload.active === true || String(payload.active).toUpperCase() === 'YES';
    const readPages = normalizeAdminManagedPageList_(payload.readPages);
    const writePages = normalizeAdminManagedPageList_(payload.writePages);
    if (readPages === null || writePages === null) {
      return {success:false,message:'One or more selected pages are not valid.'};
    }
    if (writePages.some(page => !readPages.includes(page))) {
      return {success:false,message:'Every page with write access must also have read access.'};
    }
    if (canView && !readPages.includes('dashboard')) {
      return {success:false,message:'Dashboard read access is required to sign in to the portal.'};
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET);
    if (!sheet) throw new Error('Admin Users sheet is missing. Run setup / repair first.');

    const requestedRow = Number(payload.accountRow || 0);
    let rowNumber = requestedRow;
    let previous = null;
    if (rowNumber) {
      if (rowNumber < 2 || rowNumber > sheet.getLastRow()) {
        return {success:false,message:'The selected access account no longer exists. Refresh and try again.'};
      }
      previous = sheet.getRange(rowNumber,1,1,Math.max(12,sheet.getLastColumn())).getValues()[0];
      if (String(previous[8] || '').trim() !== String(payload.expectedStaffId || '').trim()) {
        return {success:false,message:'The selected account changed. Refresh the Users page and try again.'};
      }
      if (rowNumber === user.row &&
          (!active || !canView || !['admin','superadmin'].includes(normalizeAdminRole_(role)) ||
           (user.staffId && user.staffId !== staffId))) {
        return {success:false,message:'You cannot remove your own active Admin or Super Admin access.'};
      }
    } else {
      if (sheet.getLastRow() >= 2) {
        const rows=sheet.getRange(2,1,sheet.getLastRow()-1,Math.max(12,sheet.getLastColumn())).getValues();
        const matches=rows.filter(row=>String(row[8] || '').trim()===staffId);
        if (matches.length) return {success:false,message:'This staff member already has an access account. Use Modify access.'};
      }
      rowNumber=sheet.getLastRow()+1;
      previous=['','', '', '', '', '', '', '', '', '', '', ''];
    }
    if (sheet.getLastRow() >= 2) {
      const mappedRows=sheet.getRange(2,9,sheet.getLastRow()-1,1).getValues();
      if (mappedRows.some((row,index)=>index+2!==rowNumber && String(row[0] || '').trim()===staffId)) {
        return {success:false,message:'Another access account is already linked to this staff member.'};
      }
    }

    const before=adminPrivilegeSummary_(previous);
    const nextValues=previous.slice(0,12);
    nextValues[0]=String(staff.name || '').trim();
    nextValues[2]=active?'YES':'NO';
    nextValues[3]=role;
    nextValues[4]=canView?'YES':'NO';
    nextValues[5]=payload.canOverride===true || String(payload.canOverride).toUpperCase()==='YES'?'YES':'NO';
    nextValues[6]=payload.canReset===true || String(payload.canReset).toUpperCase()==='YES'?'YES':'NO';
    nextValues[7]=payload.canOnboard===true || String(payload.canOnboard).toUpperCase()==='YES'?'YES':'NO';
    nextValues[8]=staffId;
    nextValues[9]=email || String(staff.companyEmail || '');
    nextValues[10]=readPages.join(',');
    nextValues[11]=writePages.join(',');
    sheet.getRange(rowNumber,1,1,12).setValues([nextValues]);
    const after=adminPrivilegeSummary_(nextValues);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,requestedRow?'UPDATE':'CREATE',
      'Admin User',staffId,'Portal access policy '+JSON.stringify({before:before,after:after}));
    invalidatePublicStaffOptionsCache_();
    const hasOwnPin = isValidPin_(normalizePin_(staff.pin)) && !isStartingPin_(staff.id, staff.pin);
    return {success:true,message:(requestedRow?'Access updated for ':'Access created for ')+staff.name+'. They sign in with their own PIN, the same one used everywhere.'+
      (hasOwnPin ? '' : ' They have not created it yet: ask them to open the attendance page and create their PIN first.')};
  } catch (err) {
    return {success:false,message:'User access could not be saved: '+(err.message || err)};
  } finally {
    lock.releaseLock();
  }
}

function getAdminManagedPages_() {
  return [
    {id:'dashboard',label:'Dashboard'},
    {id:'attendance',label:'Attendance'},
    {id:'staff',label:'Staff'},
    {id:'requisitions',label:'Requisitions'},
    {id:'inventory',label:'Inventory'},
    {id:'assets',label:'Assets'},
    {id:'expenses',label:'Expense claims'},
    {id:'leave',label:'Leave'},
    {id:'projects',label:'Projects'},
    {id:'leads',label:'Leads'},
    {id:'fieldduty',label:'Field duty'},
    {id:'reports',label:'Reports'},
    {id:'activity',label:'Activity log'}
  ];
}

function getAdminManagedRoleLabels_(accounts) {
  const labels = [
    'MD/CEO','SuperAdmin','Admin/HR','Accountant','Technical','Driver',
    'Motion Artist','3D Artist','Head of Operations','IoT Engineer',
    'Lead Software Engineer','Janitor'
  ];
  const seen = Object.create(null);
  labels.forEach(label => { seen[label.toLowerCase()] = true; });
  (accounts || []).forEach(account => {
    const label = canonicalAdminRoleLabel_(account.role);
    const key = label.toLowerCase();
    if (label && !seen[key]) {
      seen[key] = true;
      labels.push(label);
    }
  });
  return labels;
}

function canonicalAdminRoleLabel_(role) {
  const label = String(role || '').trim();
  const normalized = normalizeAdminRole_(label);
  if (normalized === 'superadmin') return 'SuperAdmin';
  if (normalized === 'adminhr' || normalized === 'hr') return 'Admin/HR';
  return label;
}

function normalizeAdminManagedPageList_(value) {
  if (!Array.isArray(value)) return null;
  const allowed=getAdminManagedPages_().map(page=>page.id);
  const normalized=[];
  for (let i=0;i<value.length;i++) {
    const page=String(value[i] || '').trim().toLowerCase();
    if (!allowed.includes(page)) return null;
    if (!normalized.includes(page)) normalized.push(page);
  }
  return allowed.filter(page=>normalized.includes(page));
}

function adminPrivilegeSummary_(row) {
  return {
    active:String(row[2] || '').toUpperCase()==='YES',
    role:String(row[3] || ''),
    canView:String(row[4] || '').toUpperCase()==='YES',
    canOverride:String(row[5] || '').toUpperCase()==='YES',
    canReset:String(row[6] || '').toUpperCase()==='YES',
    canOnboard:String(row[7] || '').toUpperCase()==='YES',
    staffId:String(row[8] || ''),
    email:String(row[9] || ''),
    readPages:String(row[10] || ''),
    writePages:String(row[11] || '')
  };
}

function safeRows_(rows) {
  return JSON.parse(JSON.stringify(rows || []));
}
