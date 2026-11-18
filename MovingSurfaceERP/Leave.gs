function getERPLeaveRequests_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leave Requests');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,15).getValues().filter(r => r[1]).map(r => ({
    timestamp:erpDate_(r[0]), leaveId:String(r[1]||''), staffId:String(r[2]||''), staffName:String(r[3]||''), department:String(r[4]||''),
    leaveType:String(r[5]||''), startDate:erpDate_(r[6]), endDate:erpDate_(r[7]), days:Number(r[8]||0), reason:String(r[9]||''),
    status:String(r[10]||''), approver:String(r[11]||''), approvalDate:erpDate_(r[12]), notes:String(r[13]||'')
  }));
}


function erpSubmitLeaveRequest(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin,'leave');
    const staff = erpGetStaffByAnyId_(p.staffId);
    if (!p.startDate || !p.endDate) return {success:false, message:'Start and end date are required.'};
    const start = new Date(p.startDate), end = new Date(p.endDate);
    const days = Math.max(1, Math.round((end-start)/86400000)+1);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leave Requests');
    const id = erpNextId_('Leave Requests','LEV',2);
    sh.appendRow([new Date(),id,staff.id,staff.name,staff.department,String(p.leaveType||''),p.startDate,p.endDate,days,String(p.reason||''),'Pending Approval','','','',new Date()]);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'CREATE','Leave Request',id,'Leave request created for staff '+staff.id);
    return {success:true, message:'Leave request submitted: ' + id};
  } catch (err) { return {success:false, message:err.message}; }
}

function erpDecisionLeaveRequest(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin,'leave');
    if (!adminCanApprovePage_(user,'leave')) throw new Error('You do not have permission to approve or reject leave requests.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leave Requests');
    const data = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),15).getValues();
    let row=0;
    data.forEach((r,i)=>{ if(String(r[1])===String(p.leaveId)) row=i+2; });
    if (!row) return {success:false, message:'Leave request not found.'};
    if (String(data[row-2][10]) !== 'Pending Approval') return {success:false,message:'This leave request is not pending approval.'};
    if (p.decision !== 'Approved' && p.decision !== 'Rejected') return {success:false,message:'Choose Approved or Rejected.'};
    const newStatus = p.decision === 'Approved' ? 'Approved' : 'Rejected';
    sh.getRange(row,11).setValue(newStatus);
    sh.getRange(row,12).setValue(user.name);
    sh.getRange(row,13).setValue(new Date());
    sh.getRange(row,14).setValue(String(p.comment||''));
    sh.getRange(row,15).setValue(new Date());
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'UPDATE','Leave Request',String(p.leaveId),'Status changed to '+newStatus);
    return {success:true, message:'Leave request updated: ' + newStatus};
  } catch (err) { return {success:false, message:err.message}; }
}
