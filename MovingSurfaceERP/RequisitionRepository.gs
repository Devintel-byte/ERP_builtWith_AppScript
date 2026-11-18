function generateRequisitionId_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQ_SHEET);
  const year = Utilities.formatDate(new Date(), getSettings_().TIMEZONE, 'yyyy');
  let next = 1;
  if (sh && sh.getLastRow() >= 2) {
    const ids = sh.getRange(2,2,sh.getLastRow()-1,1).getValues().flat().map(String);
    ids.forEach(id => {
      const m = id.match(new RegExp('^REQ-' + year + '-(\\d+)$'));
      if (m) next = Math.max(next, Number(m[1])+1);
    });
  }
  return 'REQ-' + year + '-' + String(next).padStart(4,'0');
}

function getAllRequisitions_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQ_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,26).getValues()
    .filter(r => r[1])
    .map((r,i) => reqObj_(r,i+2))
    .sort((a,b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
}

function getRequisitionById_(reqId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQ_SHEET);
  if (!sh || sh.getLastRow() < 2 || !String(reqId || '').trim()) return null;
  const match = sh.getRange(2,2,sh.getLastRow()-1,1)
    .createTextFinder(String(reqId))
    .matchCase(true)
    .matchEntireCell(true)
    .findNext();
  if (!match) return null;
  const row = match.getRow();
  return reqObj_(sh.getRange(row,1,1,26).getValues()[0], row);
}

function reqObj_(r,row) {
  // IMPORTANT: google.script.run can silently return null/undefined if complex objects include Date values.
  // Convert every date field to plain text before sending to the browser/admin portal.
  return {
    row: row,
    timestamp: clientSafeDate_(r[0]),
    reqId: String(r[1] || ''),
    staffId: String(r[2] || ''),
    staffName: String(r[3] || ''),
    department: String(r[4] || ''),
    project: String(r[5] || ''),
    priority: String(r[6] || ''),
    neededBefore: clientSafeDate_(r[7]),
    purpose: String(r[8] || ''),
    attachment: String(r[9] || ''),
    status: String(r[10] || ''),
    mdDecisionDate: clientSafeDate_(r[11]),
    mdDecisionBy: String(r[12] || ''),
    mdComment: String(r[13] || ''),
    accountsNotifiedDate: clientSafeDate_(r[14]),
    disbursementStatus: String(r[15] || ''),
    disbursedDate: clientSafeDate_(r[16]),
    disbursedBy: String(r[17] || ''),
    paymentReference: String(r[18] || ''),
    staffNotifiedDate: clientSafeDate_(r[19]),
    token: String(r[20] || ''),
    lastUpdated: clientSafeDate_(r[21]),
    total: Number(r[22] || 0),
    itemsJson: String(r[23] || '[]'),
    itemsSummary: String(r[24] || ''),
    projectId: String(r[25] || '')
  };
}

function summarizeRequisitions_(rows) {
  return {
    pending: rows.filter(r => r.status === 'Pending MD Approval').length,
    awaitingAccounts: rows.filter(r => r.status === 'Approved - Awaiting Accounts').length,
    paid: rows.filter(r => r.status === 'Closed - Paid / Disbursed').length,
    rejected: rows.filter(r => r.status === 'Rejected by MD').length
  };
}

function logReqAudit_(reqId, action, actor, oldStatus, newStatus, notes) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQ_AUDIT_SHEET).appendRow([new Date(), reqId, action, actor, oldStatus, newStatus, notes || '']);
}
