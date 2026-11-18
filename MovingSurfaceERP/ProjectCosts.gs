/**
 * Project Cost Lines: the single source of truth for project spend.
 *
 * Every cost reaches a project through this sheet - typed in directly, or
 * posted automatically when a requisition or expense claim moves through its
 * workflow. Automatic postings are keyed by (Source Module, Source Ref ID) and
 * upserted, so re-running a workflow step can never double-count.
 *
 * Planned Amount feeds the itemized plan; Actual Amount feeds Actual Spend.
 * Lines marked Cancelled are excluded from both.
 */

/*************** COLUMN MAP (1-based, see projectCostHeaders_) ***************/
const COST_COL = {
  lineId: 1, projectId: 2, costDate: 3, category: 4, subCategory: 5, description: 6,
  vendor: 7, quantity: 8, unit: 9, unitCost: 10, plannedAmount: 11, actualAmount: 12,
  sourceModule: 13, sourceRefId: 14, paymentStatus: 15, paidDate: 16, paymentReference: 17,
  recordedBy: 18, notes: 19, lastUpdated: 20
};
const COST_WIDTH = 20;

function projectCostsSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_COSTS_SHEET);
  if (!sheet) throw new Error('Project Cost Lines sheet is missing. Run Setup / Repair ERP Upgrade first.');
  return sheet;
}

function costRowToObject_(row, sheetRow) {
  return {
    row: sheetRow,
    lineId: String(row[COST_COL.lineId - 1] || ''),
    projectId: String(row[COST_COL.projectId - 1] || ''),
    costDate: erpDate_(row[COST_COL.costDate - 1]),
    category: String(row[COST_COL.category - 1] || ''),
    subCategory: String(row[COST_COL.subCategory - 1] || ''),
    description: String(row[COST_COL.description - 1] || ''),
    vendor: String(row[COST_COL.vendor - 1] || ''),
    quantity: Number(row[COST_COL.quantity - 1] || 0),
    unit: String(row[COST_COL.unit - 1] || ''),
    unitCost: Number(row[COST_COL.unitCost - 1] || 0),
    plannedAmount: Number(row[COST_COL.plannedAmount - 1] || 0),
    actualAmount: Number(row[COST_COL.actualAmount - 1] || 0),
    sourceModule: String(row[COST_COL.sourceModule - 1] || 'Manual'),
    sourceRefId: String(row[COST_COL.sourceRefId - 1] || ''),
    paymentStatus: String(row[COST_COL.paymentStatus - 1] || ''),
    paidDate: erpDate_(row[COST_COL.paidDate - 1]),
    paymentReference: String(row[COST_COL.paymentReference - 1] || ''),
    recordedBy: String(row[COST_COL.recordedBy - 1] || ''),
    notes: String(row[COST_COL.notes - 1] || '')
  };
}

function getAllProjectCostLines_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_COSTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, COST_WIDTH).getValues()
    .map((row, index) => costRowToObject_(row, index + 2))
    .filter(line => line.lineId || line.projectId);
}

function getProjectCostLines_(projectId) {
  const wanted = String(projectId || '').trim();
  if (!wanted) return [];
  return getAllProjectCostLines_()
    .filter(line => line.projectId === wanted)
    .sort((a, b) => String(b.costDate).localeCompare(String(a.costDate)) || String(b.lineId).localeCompare(String(a.lineId)));
}

/** Planned and actual totals per project, read in a single pass. */
function computeProjectCostTotals_() {
  const totals = {};
  getAllProjectCostLines_().forEach(function (line) {
    if (!line.projectId) return;
    if (line.paymentStatus === 'Cancelled') return;
    if (!totals[line.projectId]) totals[line.projectId] = {planned: 0, actual: 0};
    totals[line.projectId].planned += line.plannedAmount;
    totals[line.projectId].actual += line.actualAmount;
  });
  return totals;
}

/*************** MANUAL ENTRY ***************/

function erpAddProjectCostLine(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'projects');
    const project = findProjectRow_(p.projectId);
    if (!project) return {success: false, message: 'Select a valid project.'};
    const description = String(p.description || '').trim();
    if (!description) return {success: false, message: 'Description is required.'};

    const quantity = Number(p.quantity || 0) || 0;
    const unitCost = erpNonNegativeAmount_(p.unitCost, 'Unit Cost');
    const paymentStatus = normalizeCostPaymentStatus_(p.paymentStatus || 'Planned');
    const computed = quantity > 0 && unitCost > 0 ? Math.round(quantity * unitCost * 100) / 100 : 0;
    let planned = p.plannedAmount !== undefined && String(p.plannedAmount).trim() !== ''
      ? erpNonNegativeAmount_(p.plannedAmount, 'Planned Amount') : computed;
    let actual = p.actualAmount !== undefined && String(p.actualAmount).trim() !== ''
      ? erpNonNegativeAmount_(p.actualAmount, 'Actual Amount') : 0;

    // A line booked as Paid with no actual entered takes its planned amount, so
    // nothing silently records as paid-but-zero.
    if (paymentStatus === 'Paid' && actual === 0) actual = planned || computed;
    if (!planned && actual) planned = actual;
    if (!planned && !actual) return {success: false, message: 'Enter a planned amount, an actual amount, or a quantity and unit cost.'};

    const sheet = projectCostsSheet_();
    const lineId = erpNextId_(PROJECT_COSTS_SHEET, 'PCL', COST_COL.lineId);
    const now = new Date();
    const row = new Array(COST_WIDTH).fill('');
    row[COST_COL.lineId - 1] = lineId;
    row[COST_COL.projectId - 1] = project.projectId;
    row[COST_COL.costDate - 1] = erpDateInput_(p.costDate, 'Cost Date') || erpToday_();
    row[COST_COL.category - 1] = normalizeProjectCostCategory_(p.category);
    row[COST_COL.subCategory - 1] = String(p.subCategory || '');
    row[COST_COL.description - 1] = description;
    row[COST_COL.vendor - 1] = String(p.vendor || '');
    row[COST_COL.quantity - 1] = quantity;
    row[COST_COL.unit - 1] = String(p.unit || '');
    row[COST_COL.unitCost - 1] = unitCost;
    row[COST_COL.plannedAmount - 1] = planned;
    row[COST_COL.actualAmount - 1] = actual;
    row[COST_COL.sourceModule - 1] = 'Manual';
    row[COST_COL.sourceRefId - 1] = '';
    row[COST_COL.paymentStatus - 1] = paymentStatus;
    row[COST_COL.paidDate - 1] = paymentStatus === 'Paid' ? (erpDateInput_(p.paidDate, 'Paid Date') || erpToday_()) : '';
    row[COST_COL.paymentReference - 1] = String(p.paymentReference || '');
    row[COST_COL.recordedBy - 1] = user.name;
    row[COST_COL.notes - 1] = String(p.notes || '');
    row[COST_COL.lastUpdated - 1] = now;
    sheet.appendRow(row);

    recomputeProjectRollups_(project.projectId);
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'CREATE', 'Project Cost Line', lineId,
      project.projectId + ' / ' + row[COST_COL.category - 1] + ' / planned ' + planned + ' / actual ' + actual);
    return {success: true, message: 'Cost line added: ' + lineId, lineId: lineId};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/** Marks a manual or committed cost line as paid and books its actual amount. */
function erpMarkProjectCostPaid(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'projects');
    const sheet = projectCostsSheet_();
    const line = findProjectCostLine_(p.lineId);
    if (!line) return {success: false, message: 'Cost line not found.'};
    if (line.paymentStatus === 'Paid') return {success: false, message: 'This cost line is already marked paid.'};
    if (line.paymentStatus === 'Cancelled') return {success: false, message: 'This cost line was cancelled.'};

    const actual = p.actualAmount !== undefined && String(p.actualAmount).trim() !== ''
      ? erpNonNegativeAmount_(p.actualAmount, 'Actual Amount')
      : (line.actualAmount || line.plannedAmount);
    if (!actual) return {success: false, message: 'Enter the actual amount paid.'};

    sheet.getRange(line.row, COST_COL.actualAmount).setValue(actual);
    sheet.getRange(line.row, COST_COL.paymentStatus).setValue('Paid');
    sheet.getRange(line.row, COST_COL.paidDate).setValue(erpDateInput_(p.paidDate, 'Paid Date') || erpToday_());
    sheet.getRange(line.row, COST_COL.paymentReference).setValue(String(p.paymentReference || line.paymentReference || ''));
    sheet.getRange(line.row, COST_COL.lastUpdated).setValue(new Date());

    recomputeProjectRollups_(line.projectId);
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Project Cost Line', line.lineId,
      'Marked paid, actual ' + actual);
    return {success: true, message: 'Cost line ' + line.lineId + ' marked paid.'};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cancels a cost line. Cancelled lines stay on the sheet for audit but drop out
 * of every roll-up, so spend history is never silently rewritten.
 */
function erpCancelProjectCostLine(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'projects');
    const sheet = projectCostsSheet_();
    const line = findProjectCostLine_(p.lineId);
    if (!line) return {success: false, message: 'Cost line not found.'};
    if (line.sourceModule !== 'Manual') {
      return {success: false, message: 'This line was posted automatically from ' + line.sourceModule +
        ' ' + line.sourceRefId + '. Cancel it in that module instead.'};
    }
    sheet.getRange(line.row, COST_COL.paymentStatus).setValue('Cancelled');
    sheet.getRange(line.row, COST_COL.notes)
      .setValue(((line.notes ? line.notes + ' | ' : '') + 'Cancelled by ' + user.name + ': ' + String(p.reason || 'no reason given')).slice(0, 2000));
    sheet.getRange(line.row, COST_COL.lastUpdated).setValue(new Date());

    recomputeProjectRollups_(line.projectId);
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Project Cost Line', line.lineId,
      'Cancelled: ' + String(p.reason || ''));
    return {success: true, message: 'Cost line ' + line.lineId + ' cancelled.'};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

function findProjectCostLine_(lineId) {
  const sheet = projectCostsSheet_();
  if (sheet.getLastRow() < 2 || !String(lineId || '').trim()) return null;
  const match = sheet.getRange(2, COST_COL.lineId, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(lineId)).matchCase(true).matchEntireCell(true).findNext();
  if (!match) return null;
  return costRowToObject_(sheet.getRange(match.getRow(), 1, 1, COST_WIDTH).getValues()[0], match.getRow());
}

function normalizeCostPaymentStatus_(value) {
  const supplied = String(value || '').trim();
  const match = projectCostPaymentStatuses_().filter(status => status.toLowerCase() === supplied.toLowerCase());
  return match.length ? match[0] : 'Planned';
}

/*************** AUTOMATIC POSTING FROM OTHER MODULES ***************/

/**
 * Upserts the single cost line that represents a source document.
 *
 * The (Source Module, Source Ref ID) pair is the idempotency key: a requisition
 * that is approved and later paid updates one line through its lifecycle rather
 * than appending a second. Returns a short status string for the caller's
 * message, or '' when there is nothing to post.
 */
function postProjectCostFromSource_(options) {
  const projectId = String(options.projectId || '').trim();
  const sourceModule = String(options.sourceModule || '').trim();
  const sourceRefId = String(options.sourceRefId || '').trim();
  if (!projectId || !sourceModule || !sourceRefId) return '';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_COSTS_SHEET);
  if (!sheet) {
    Logger.log('Project cost line not posted for ' + sourceModule + ' ' + sourceRefId + ': Project Cost Lines sheet is missing.');
    return '';
  }
  if (!findProjectRow_(projectId)) {
    Logger.log('Project cost line not posted for ' + sourceModule + ' ' + sourceRefId + ': project ' + projectId + ' not found.');
    return '';
  }

  const existing = findCostLineBySource_(sourceModule, sourceRefId);
  const planned = erpNonNegativeAmount_(options.plannedAmount, 'Planned Amount');
  const actual = erpNonNegativeAmount_(options.actualAmount, 'Actual Amount');
  const paymentStatus = normalizeCostPaymentStatus_(options.paymentStatus || 'Committed');
  const now = new Date();

  if (existing) {
    sheet.getRange(existing.row, COST_COL.projectId).setValue(projectId);
    sheet.getRange(existing.row, COST_COL.plannedAmount).setValue(planned || existing.plannedAmount);
    sheet.getRange(existing.row, COST_COL.actualAmount).setValue(actual);
    sheet.getRange(existing.row, COST_COL.paymentStatus).setValue(paymentStatus);
    sheet.getRange(existing.row, COST_COL.paidDate)
      .setValue(paymentStatus === 'Paid' ? (options.paidDate || erpToday_()) : '');
    sheet.getRange(existing.row, COST_COL.paymentReference).setValue(String(options.paymentReference || ''));
    sheet.getRange(existing.row, COST_COL.category).setValue(normalizeProjectCostCategory_(options.category));
    sheet.getRange(existing.row, COST_COL.description).setValue(String(options.description || existing.description));
    sheet.getRange(existing.row, COST_COL.lastUpdated).setValue(now);
    recomputeProjectRollups_(projectId);
    return 'Project cost line ' + existing.lineId + ' updated (' + paymentStatus + ').';
  }

  const lineId = erpNextId_(PROJECT_COSTS_SHEET, 'PCL', COST_COL.lineId);
  const row = new Array(COST_WIDTH).fill('');
  row[COST_COL.lineId - 1] = lineId;
  row[COST_COL.projectId - 1] = projectId;
  row[COST_COL.costDate - 1] = options.costDate || erpToday_();
  row[COST_COL.category - 1] = normalizeProjectCostCategory_(options.category);
  row[COST_COL.subCategory - 1] = String(options.subCategory || '');
  row[COST_COL.description - 1] = String(options.description || (sourceModule + ' ' + sourceRefId));
  row[COST_COL.vendor - 1] = String(options.vendor || '');
  row[COST_COL.quantity - 1] = 0;
  row[COST_COL.unit - 1] = '';
  row[COST_COL.unitCost - 1] = 0;
  row[COST_COL.plannedAmount - 1] = planned;
  row[COST_COL.actualAmount - 1] = actual;
  row[COST_COL.sourceModule - 1] = sourceModule;
  row[COST_COL.sourceRefId - 1] = sourceRefId;
  row[COST_COL.paymentStatus - 1] = paymentStatus;
  row[COST_COL.paidDate - 1] = paymentStatus === 'Paid' ? (options.paidDate || erpToday_()) : '';
  row[COST_COL.paymentReference - 1] = String(options.paymentReference || '');
  row[COST_COL.recordedBy - 1] = String(options.recordedBy || 'System');
  row[COST_COL.notes - 1] = 'Posted automatically from ' + sourceModule + ' ' + sourceRefId + '.';
  row[COST_COL.lastUpdated - 1] = now;
  sheet.appendRow(row);

  recomputeProjectRollups_(projectId);
  return 'Project cost line ' + lineId + ' posted to ' + projectId + ' (' + paymentStatus + ').';
}

function findCostLineBySource_(sourceModule, sourceRefId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_COSTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const refs = sheet.getRange(2, COST_COL.sourceRefId, sheet.getLastRow() - 1, 1).getValues();
  const modules = sheet.getRange(2, COST_COL.sourceModule, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < refs.length; i++) {
    if (String(refs[i][0] || '').trim() === String(sourceRefId).trim() &&
        String(modules[i][0] || '').trim() === String(sourceModule).trim()) {
      return costRowToObject_(sheet.getRange(i + 2, 1, 1, COST_WIDTH).getValues()[0], i + 2);
    }
  }
  return null;
}

/** Reads the appended 'Project ID' value from a row of a linked sheet. */
function readProjectIdFromSheetRow_(sheetName, sheetRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheetRow < 2) return '';
  const width = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(header => String(header || '').trim());
  const column = headers.indexOf('Project ID');
  if (column === -1) return '';
  return String(sheet.getRange(sheetRow, column + 1).getValue() || '').trim();
}

function writeProjectIdToSheetRow_(sheetName, sheetRow, projectId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheetRow < 2) return false;
  const width = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(header => String(header || '').trim());
  const column = headers.indexOf('Project ID');
  if (column === -1) return false;
  sheet.getRange(sheetRow, column + 1).setValue(String(projectId || ''));
  return true;
}

/**
 * Derives a cost category for a requisition from its item lines, falling back
 * to materials when no item names a category the project report understands.
 */
function requisitionCostCategory_(itemsJson) {
  let items = [];
  try { items = JSON.parse(String(itemsJson || '[]')); } catch (err) { items = []; }
  if (!Array.isArray(items)) return 'Materials & Consumables';
  for (let i = 0; i < items.length; i++) {
    const category = normalizeProjectCostCategory_(items[i] && items[i].category);
    if (category !== 'Miscellaneous') return category;
  }
  return 'Materials & Consumables';
}
