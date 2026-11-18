function getERPAssets_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assets');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,15).getValues().filter(r => r[0] || r[1]).map(r => ({
    assetId:String(r[0]||''), assetName:String(r[1]||''), category:String(r[2]||''), model:String(r[3]||''),
    serial:String(r[4]||''), purchaseDate:erpDate_(r[5]), purchaseCost:Number(r[6]||0), location:String(r[7]||''),
    status:String(r[8]||''), assignedToStaffId:String(r[9]||''), assignedToName:String(r[10]||''), condition:String(r[11]||''),
    returnDueDate:erpDate_(r[12]), notes:String(r[13]||'')
  }));
}


function erpAddAsset(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin,'assets');
    const name = String(p.assetName||'').trim();
    if (!name) return {success:false, message:'Asset name is required.'};
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assets');
    const id = erpNextId_('Assets','MS-AST',1);
    sh.appendRow([id,name,String(p.category||''),String(p.model||''),String(p.serial||''),'','',String(p.location||'Abuja'),'Available','','',String(p.condition||'Good'),'','',new Date()]);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'CREATE','Asset',id,'Asset added');
    return {success:true, message:'Asset added: ' + id};
  } catch (err) { return {success:false, message:err.message}; }
}

function erpAssignAsset(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin,'assets');
    const staff = erpGetStaffByAnyId_(p.staffId);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Assets');
    const data = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),15).getValues();
    let row=0, assetName='', oldLoc='';
    data.forEach((r,i)=>{ if(String(r[0])===String(p.assetId)){ row=i+2; assetName=String(r[1]||''); oldLoc=String(r[7]||''); }});
    if (!row) return {success:false, message:'Asset not found.'};
    sh.getRange(row,9).setValue('Assigned');
    sh.getRange(row,10).setValue(staff.id);
    sh.getRange(row,11).setValue(staff.name);
    sh.getRange(row,15).setValue(new Date());
    ss.getSheetByName('Asset Movements').appendRow([new Date(),erpNextId_('Asset Movements','MOV',2),p.assetId,assetName,'Assigned',oldLoc,staff.id,staff.name,String(p.project||''),'Good','',user.name,'',resolveSubmittedProjectId_(p.projectId, p.project)]);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'UPDATE','Asset',String(p.assetId),'Assigned to staff '+staff.id);
    return {success:true, message:'Asset assigned to ' + staff.name};
  } catch (err) { return {success:false, message:err.message}; }
}

function erpReturnAsset(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin,'assets');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Assets');
    const data = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),15).getValues();
    let row=0, assetName='', staffId='', staffName='';
    data.forEach((r,i)=>{ if(String(r[0])===String(p.assetId)){ row=i+2; assetName=String(r[1]||''); staffId=String(r[9]||''); staffName=String(r[10]||''); }});
    if (!row) return {success:false, message:'Asset not found.'};
    sh.getRange(row,9).setValue('Available');
    sh.getRange(row,10).setValue('');
    sh.getRange(row,11).setValue('');
    sh.getRange(row,12).setValue(String(p.condition||'Good'));
    sh.getRange(row,15).setValue(new Date());
    ss.getSheetByName('Asset Movements').appendRow([new Date(),erpNextId_('Asset Movements','MOV',2),p.assetId,assetName,'Returned',staffId,'Store',staffName,String(p.project||''),'',String(p.condition||'Good'),user.name,'',resolveSubmittedProjectId_(p.projectId, p.project)]);
    recordActivity_('Admin',user.name,'ADMIN-ROW-'+user.row,'UPDATE','Asset',String(p.assetId),'Asset returned; condition '+String(p.condition||'Good'));
    return {success:true, message:'Asset returned.'};
  } catch (err) { return {success:false, message:err.message}; }
}
