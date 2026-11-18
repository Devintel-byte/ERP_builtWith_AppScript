function getERPExpenseClaims_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expense Claims');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,19).getValues().filter(r => r[1]).map(r => ({
    timestamp:erpDate_(r[0]), claimId:String(r[1]||''), staffId:String(r[2]||''), staffName:String(r[3]||''), department:String(r[4]||''),
    expenseDate:erpDate_(r[5]), category:String(r[6]||''), project:String(r[7]||''), amount:Number(r[8]||0),
    status:String(r[9]||''), approver:String(r[10]||''), approvalDate:erpDate_(r[11]), accountsStatus:String(r[12]||''),
    paidDate:erpDate_(r[13]), receiptLink:String(r[14]||''), purpose:String(r[15]||''), paymentReference:String(r[16]||''),
    projectId:String(r[18]||'')
  }));
}


function erpSubmitExpenseClaim(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin,'expenses');
    const staff = erpGetStaffByAnyId_(p.staffId);
    const amount = Number(p.amount||0);
    if (amount <= 0) return {success:false, message:'Amount is required.'};
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expense Claims');
    const id = erpNextId_('Expense Claims','EXP',2);
    const projectId = resolveSubmittedProjectId_(p.projectId, p.project);
    sh.appendRow([new Date(),id,staff.id,staff.name,staff.department,p.expenseDate||'',normalizeProjectCostCategory_(p.category),String(p.project||''),amount,'Pending Approval','', '', 'Not Submitted','',String(p.receipt||''),String(p.purpose||''),'',new Date(),projectId]);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'CREATE','Expense Claim',id,'Expense claim created for staff '+staff.id);
    return {success:true, message:'Expense claim submitted: ' + id};
  } catch (err) { return {success:false, message:err.message}; }
}

function erpDecisionExpenseClaim(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin,'expenses');
    if (!adminCanApprovePage_(user,'expenses')) throw new Error('You do not have permission to approve or reject expense claims.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expense Claims');
    const data = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),19).getValues();
    let row=0;
    data.forEach((r,i)=>{ if(String(r[1])===String(p.claimId)) row=i+2; });
    if (!row) return {success:false, message:'Claim not found.'};
    if (String(data[row-2][9]) !== 'Pending Approval') return {success:false,message:'This expense claim is not pending approval.'};
    if (p.decision !== 'Approved' && p.decision !== 'Rejected') return {success:false,message:'Choose Approved or Rejected.'};
    const newStatus = p.decision === 'Approved' ? 'Approved - Awaiting Accounts' : 'Rejected';
    sh.getRange(row,10).setValue(newStatus);
    sh.getRange(row,11).setValue(user.name);
    sh.getRange(row,12).setValue(new Date());
    sh.getRange(row,13).setValue(newStatus === 'Approved - Awaiting Accounts' ? 'Awaiting Accounts' : 'Rejected');
    sh.getRange(row,18).setValue(new Date());
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'UPDATE','Expense Claim',String(p.claimId),'Status changed to '+newStatus);
    const costPosted = newStatus === 'Approved - Awaiting Accounts'
      ? postExpenseProjectCost_(data[row-2], 'Committed', user.name, '')
      : '';
    return {success:true, message:'Expense ' + p.claimId + ' updated: ' + newStatus + (costPosted ? ' ' + costPosted : '')};
  } catch (err) { return {success:false, message:err.message}; }
}

function erpMarkExpensePaid(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin,'expenses');
    if (!adminCanDisbursePage_(user,'expenses')) throw new Error('You do not have accounts/disbursement permission.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expense Claims');
    const data = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),19).getValues();
    let row=0;
    data.forEach((r,i)=>{ if(String(r[1])===String(p.claimId)) row=i+2; });
    if (!row) return {success:false, message:'Claim not found.'};
    if (String(data[row-2][9]) !== 'Approved - Awaiting Accounts') return {success:false,message:'This expense claim is not awaiting accounts.'};
    sh.getRange(row,10).setValue('Paid / Closed');
    sh.getRange(row,13).setValue('Paid');
    sh.getRange(row,14).setValue(new Date());
    sh.getRange(row,17).setValue(String(p.reference||''));
    sh.getRange(row,18).setValue(new Date());
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'UPDATE','Expense Claim',String(p.claimId),'Marked paid');
    const costPosted = postExpenseProjectCost_(data[row-2], 'Paid', user.name, String(p.reference||''));
    return {success:true, message:'Expense claim marked paid by ' + user.name + (costPosted ? ' ' + costPosted : '')};
  } catch (err) { return {success:false, message:err.message}; }
}

/**
 * Posts (or updates) the project cost line representing an expense claim.
 *
 * Keyed on the Claim ID, so approval (Committed) and payment (Paid) update one
 * line rather than appending two. Expense categories are drawn from the same
 * list as project cost categories, so a claim lands in a category the project
 * report already understands.
 */
function postExpenseProjectCost_(row, paymentStatus, actorName, paymentReference) {
  try {
    const projectId = String(row[18] || '').trim();
    if (!projectId) return '';
    const amount = Number(row[8] || 0);
    return postProjectCostFromSource_({
      projectId:projectId,
      sourceModule:'Expense Claim',
      sourceRefId:String(row[1] || ''),
      costDate:erpDate_(row[5]) || erpToday_(),
      category:normalizeProjectCostCategory_(row[6]),
      description:'Expense claim ' + String(row[1] || '') + ' - ' + String(row[15] || row[6] || ''),
      vendor:String(row[3] || ''),
      plannedAmount:amount,
      actualAmount:paymentStatus === 'Paid' ? amount : 0,
      paymentStatus:paymentStatus,
      paymentReference:paymentReference,
      recordedBy:actorName
    });
  } catch (err) {
    Logger.log('Expense cost line not posted for ' + String(row[1] || '') + ': ' + (err.stack || err.message));
    return '';
  }
}
