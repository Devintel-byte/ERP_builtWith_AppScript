function adminGetAttendancePage(payload) {
  const user = getAdminUser_(payload.adminPin);
  if (!adminCanReadPage_(user,'attendance')) return {success:false,message:'You do not have permission to view attendance.'};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {success:true,rows:[],page:1,pageSize:25,total:0,totalPages:0};

  const pageSize = 25;
  const page = Math.max(1,Math.floor(Number(payload.page) || 1));
  const query = String(payload.query || '').trim().toLowerCase();
  const action = String(payload.action || '');
  const sort = payload.sort === 'oldest' ? 1 : -1;
  const from = String(payload.from || '');
  const to = String(payload.to || '');
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,25).getValues();
  const settings = getSettings_();
  const staffById = Object.create(null);
  getAllStaff_().forEach(staff => { staffById[staff.id] = staff; });
  const filtered = values.map(function(row,index) {
    const dateKey = formatAttendanceDate_(row[0],settings);
    const stamp = row[5] instanceof Date ? row[5] : new Date(row[5]);
    const staff = staffById[String(row[1] || '')];
    return {
      dateKey:dateKey,
      timestamp:!isNaN(stamp.getTime()) ? Utilities.formatDate(stamp,Session.getScriptTimeZone() || 'Africa/Lagos','yyyy-MM-dd HH:mm:ss') : String(row[5] || ''),
      staffId:String(row[1] || ''),staffName:String(row[2] || ''),department:staff ? staff.department : '',
      location:String(row[3] || ''),action:String(row[4] || ''),time:String(row[6] || ''),
      checkoutTime:String(row[23] || (attendanceCheckoutType_(row[4],'') ? row[6] || '' : '')),
      checkoutType:attendanceCheckoutType_(row[4],row[24]),
      status:String(row[13] || ''),matchedLocation:String(row[10] || ''),gpsStatus:String(row[11] || ''),
      distance:Number(row[9] || 0),sheetRow:index+2,sortTime:!isNaN(stamp.getTime()) ? stamp.getTime() : 0
    };
  }).filter(function(row) {
    const text = [row.dateKey,row.timestamp,row.staffId,row.staffName,row.department,row.location,row.action,row.status,row.matchedLocation,row.gpsStatus].join(' ').toLowerCase();
    return (!from || row.dateKey >= from) && (!to || row.dateKey <= to) &&
      (!action || row.action === action) && (!query || text.includes(query));
  });
  filtered.sort((a,b) => sort*(a.sortTime-b.sortTime || a.sheetRow-b.sheetRow));
  const total = filtered.length;
  const totalPages = total ? Math.ceil(total/pageSize) : 0;
  return {success:true,rows:filtered.slice((page-1)*pageSize,page*pageSize),page:page,pageSize:pageSize,total:total,totalPages:totalPages};
}

function adminGetActivityPage(payload) {
  const user = getAdminUser_(payload.adminPin);
  if (!adminCanReadPage_(user,'activity')) return {success:false,message:'You do not have permission to view the activity log.'};
  return Object.assign({success:true},getActivityLogPage_(payload));
}

function adminGetAnalytics(payload) {
  const user = getAdminUser_(payload.adminPin);
  if (!adminCanReadPage_(user,'reports')) return {success:false,message:'You do not have permission to view reports.'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const from = String(payload.from || '');
  const to = String(payload.to || '');
  const staff = getAllStaff_();
  const requisitions = getAllRequisitions_();
  const expenses = getERPExpenseClaims_();
  const inventory = getERPInventoryItems_();
  const settings = getSettings_();
  const attendanceSheet = ss.getSheetByName(LOG_SHEET);
  const attendanceRows = attendanceSheet && attendanceSheet.getLastRow() > 1
  ? attendanceSheet.getRange(2,1,attendanceSheet.getLastRow()-1,25).getValues()
    : [];
  const selectedAttendance = attendanceRows.filter(function(row) {
    const key = formatAttendanceDate_(row[0],settings);
    return (!from || key >= from) && (!to || key <= to);
  });

  const byDate = Object.create(null);
  const byStatus = Object.create(null);
  const overtimeDatesByStaff = Object.create(null);
  let regularCheckouts=0,overtimeCheckouts=0;
  selectedAttendance.forEach(function(row) {
    const checkoutType=attendanceCheckoutType_(row[4],row[24]);
    if (checkoutType==='Overtime') {
      const key = String(row[1] || '')+'|'+formatAttendanceDate_(row[0],settings);
      overtimeDatesByStaff[key] = true;
    }
    if (checkoutType==='Regular') regularCheckouts++;
    if (checkoutType==='Overtime') overtimeCheckouts++;
    if (String(row[4]) !== 'Check In') return;
    const date = formatAttendanceDate_(row[0],settings);
    byDate[date] = (byDate[date] || 0)+1;
    const status = String(row[13] || 'Unknown');
    byStatus[status] = (byStatus[status] || 0)+1;
  });
  const dailyAttendance = Object.keys(byDate).sort().map(date => ({date:date,count:byDate[date]}));
  const activeStaff = staff.filter(row => String(row.status).toLowerCase() === 'active').length;
  const totalPayroll = staff.filter(row => String(row.status).toLowerCase() === 'active')
    .reduce((sum,row) => sum+Number(row.monthlySalary || 0),0);
  const reqByStatus = Object.create(null);
  let reqValue = 0;
  requisitions.forEach(function(row) {
    reqByStatus[row.status] = (reqByStatus[row.status] || 0)+1;
    reqValue += Number(row.total || 0);
  });
  const expenseByStatus = Object.create(null);
  let expenseAmount = 0;
  expenses.forEach(function(row) {
    expenseByStatus[row.status] = (expenseByStatus[row.status] || 0)+1;
    expenseAmount += Number(row.amount || 0);
  });
  const outOfStock = inventory.filter(row => row.stockStatus === 'Low Stock').length;
  return {
    success:true,
    generatedAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone() || 'Africa/Lagos','yyyy-MM-dd HH:mm:ss'),
    period:{from:from,to:to},
    metrics:{
      activeStaff:activeStaff,inactiveStaff:staff.length-activeStaff,
      attendanceCheckIns:selectedAttendance.filter(row => String(row[4]) === 'Check In').length,
      attendanceCheckOuts:regularCheckouts+overtimeCheckouts,
      regularCheckouts:regularCheckouts,overtimeCheckouts:overtimeCheckouts,
      overtimeDays:Object.keys(overtimeDatesByStaff).length,
      presentDays:dailyAttendance.length,late:byStatus.Late || 0,halfDay:byStatus['Half Day'] || 0,
      requisitions:requisitions.length,pendingRequisitions:reqByStatus['Pending MD Approval'] || 0,
      requisitionValue:reqValue,expenseClaims:expenses.length,expenseAmount:expenseAmount,
      inventoryLowStock:outOfStock,activeProjects:getERPProjects_().filter(row => row.status === 'Active').length,
      monthlySalaryCommitment:totalPayroll
    },
    dailyAttendance:dailyAttendance,
    attendanceByStatus:Object.keys(byStatus).map(status => ({label:status,value:byStatus[status]})),
    requisitionsByStatus:Object.keys(reqByStatus).map(status => ({label:status,value:reqByStatus[status]})),
    expensesByStatus:Object.keys(expenseByStatus).map(status => ({label:status,value:expenseByStatus[status]}))
  };
}

function getSpreadsheetSheetLinks_(user) {
  if (!user || !user.valid) return {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const links = {};
  const pageSheets = {
    attendance:[LOG_SHEET],
    reports:[ACCOUNTANT_VIEW_SHEET],
    requisitions:[REQ_SHEET],
    expenses:[EXPENSE_CLAIMS_SHEET],
    inventory:[INVENTORY_ITEMS_SHEET],
    projects:[PROJECTS_SHEET,PROJECT_COSTS_SHEET],
    leads:[LEADS_SHEET,CLIENTS_SHEET,LEAD_ACTIVITY_SHEET],
    staff:[STAFF_SHEET]
  };
  Object.keys(pageSheets).forEach(function(page) {
    if (!adminCanReadPage_(user,page)) return;
    pageSheets[page].forEach(function(name) {
      const sheet = ss.getSheetByName(name);
      if (sheet) links[name] = ss.getUrl()+'#gid='+sheet.getSheetId();
    });
  });
  return links;
}

function adminGetStaffReport(payload) {
  const user = getAdminUser_(payload.adminPin);
  if (!adminCanReadPage_(user,'staff') || !adminCanReadPage_(user,'reports')) {
    return {success:false,message:'You do not have permission to view individual staff records.'};
  }
  const staffId = String(payload.staffId || '').trim();
  const staff = getStaffByIdAnyStatus_(staffId);
  if (!staff) return {success:false,message:'Staff record not found.'};

  const settings = getSettings_();
  const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  const attendance = attendanceSheet && attendanceSheet.getLastRow() > 1
    ? attendanceSheet.getRange(2,1,attendanceSheet.getLastRow()-1,25).getValues()
      .filter(row => String(row[1]) === staff.id)
      .map(row => ({
        date:formatAttendanceDate_(row[0],settings),
        action:String(row[4] || ''),
        timestamp:row[5] instanceof Date ? Utilities.formatDate(row[5],settings.TIMEZONE,'yyyy-MM-dd HH:mm:ss') : String(row[5] || ''),
        checkoutTime:String(row[23] || (attendanceCheckoutType_(row[4],'') ? formatAttendanceTime_(row[5],row[6],settings) : '')),
        checkoutType:attendanceCheckoutType_(row[4],row[24]),
        location:String(row[3] || ''),
        status:String(row[13] || ''),
        penalty:Number(row[14] || 0),
        remarks:String(row[22] || '')
      }))
      .sort((a,b) => String(b.timestamp).localeCompare(String(a.timestamp)))
    : [];
  const requisitions = getAllRequisitions_().filter(row => row.staffId === staff.id).map(requisitionClientDto_);
  const leaveRequests = getERPLeaveRequests_().filter(row => row.staffId === staff.id);
  const expenses = getERPExpenseClaims_().filter(row => row.staffId === staff.id);
  const assets = getERPAssets_().filter(row => row.assignedToStaffId === staff.id);

  return {
    success:true,
    staff:{
      id:staff.id,name:staff.name,department:staff.department,location:staff.location,
      companyEmail:staff.companyEmail,phone:staff.phone,
      designation:staff.designation,manager:staff.manager,status:staff.status,
      employmentType:staff.employmentType,dateJoined:erpDate_(staff.dateJoined),
      exitDate:erpDate_(staff.exitDate),monthlySalary:staff.monthlySalary
    },
    attendance:attendance,
    requisitions:requisitions,
    leaveRequests:leaveRequests,
    expenses:expenses,
    assets:assets
  };
}
