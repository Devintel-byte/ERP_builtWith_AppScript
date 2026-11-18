function submitItemRequisition(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let staff;
    try {
      staff = validateStaffPortalLogin_(payload.staffId, payload.pin, payload.deviceToken);
    } catch (err) {
      return {success:false, message:err.message};
    }
    if (String(staff.canSubmitRequisition || 'YES').toUpperCase() !== 'YES') return {success:false, message:'You are not enabled to submit requisitions.'};

    const purpose = String(payload.purpose || '').trim();
    const items = Array.isArray(payload.items) ? payload.items : [];
    const cleanItems = items.map(it => ({
      itemName: String(it.itemName || '').trim(),
      quantity: Number(it.qty || it.quantity || 0),
      unitCost: Number(it.unitCost || 0),
      total: Number(it.total || 0) || (Number(it.qty || it.quantity || 0) * Number(it.unitCost || 0)),
      category: String(it.category || ''),
      vendor: String(it.vendor || '')
    })).filter(it => it.itemName && it.quantity > 0 && it.unitCost >= 0);

    if (!purpose) return {success:false, message:'Purpose / Justification is required.'};
    if (cleanItems.length === 0) return {success:false, message:'Please add at least one valid item.'};

    const total = cleanItems.reduce((sum, it) => sum + Number(it.total || 0), 0);
    const reqId = generateRequisitionId_();
    const token = Utilities.getUuid();
    const now = new Date();
    const itemsSummary = cleanItems.map((it, i) => (i+1)+'. '+it.itemName+' x '+it.quantity+' @ NGN '+formatMoney_(it.unitCost)+' = NGN '+formatMoney_(it.total)).join('\n');

    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQ_SHEET).appendRow([
      now, reqId, staff.id, staff.name, staff.department, String(payload.project || ''), String(payload.priority || 'Normal'),
      payload.neededBefore || '', purpose, String(payload.attachment || ''), 'Pending MD Approval',
      '', '', '', '', 'Not Disbursed', '', '', '', '', token, now, total, JSON.stringify(cleanItems), itemsSummary,
      resolveSubmittedProjectId_(payload.projectId, payload.project)
    ]);

    logReqAudit_(reqId, 'Submitted', staff.name, '', 'Pending MD Approval', 'Submitted by staff');
    recordActivity_('Staff',staff.name,staff.id,'CREATE','Requisition',reqId,'Requisition submitted for approval');
    const notificationQueued = sendRequisitionToMd_(reqId);

    return {success:true, message:'Requisition submitted successfully.\n\nRequest ID: '+reqId+'\nItems: '+cleanItems.length+'\nTotal: NGN '+formatMoney_(total)+'\nStatus: Pending MD Approval'+(notificationQueued ? '\nApproval email queued for delivery.' : '\nThe request was saved, but its approval email could not be queued. Please contact Admin.')};
  } catch (err) {
    return {success:false, message:'Error: '+(err.stack || err.message)};
  } finally {
    lock.releaseLock();
  }
}

/*************** ADMIN PORTAL ***************/

/**
 * Resolves the Project ID a staff requisition should be keyed to.
 *
 * A valid Project ID is taken as given. Otherwise the free-text Project / Event
 * the staff member typed is matched against project names, so existing habits
 * still produce a usable key instead of an orphaned cost.
 */
function resolveSubmittedProjectId_(projectId, projectText) {
  const supplied = String(projectId || '').trim();
  if (supplied && findProjectRow_(supplied)) return supplied;
  const text = String(projectText || '').trim();
  if (!text) return '';
  const match = getERPProjects_().filter(project =>
    project.projectName.toLowerCase() === text.toLowerCase() || project.projectId.toLowerCase() === text.toLowerCase());
  return match.length ? match[0].projectId : '';
}
