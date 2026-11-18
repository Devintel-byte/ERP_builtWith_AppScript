/**
 * Projects: delivery and project economics.
 *
 * Three money concepts are kept deliberately separate:
 *   Contract Value     - revenue, what the client pays
 *   Expenditure Budget - the approved cost budget for delivering the project
 *   Actual Spend       - cost actually incurred, DERIVED from Project Cost Lines
 *
 * Actual Spend, Planned Cost, Budget Variance, Project Balance, Margin % and
 * Outstanding Receivable are never typed by a user. They are recomputed from
 * cost lines on every write and by the nightly reconcile.
 */

/*************** COLUMN MAP (1-based, see projectsHeaders_) ***************/
const PRJ_COL = {
  projectId: 1, projectName: 2, clientId: 3, client: 4, leadId: 5, manager: 6,
  startDate: 7, endDate: 8, status: 9, contractValue: 10, budget: 11, plannedCost: 12,
  actualSpend: 13, budgetVariance: 14, balance: 15, marginPct: 16, invoiced: 17,
  received: 18, receivable: 19, location: 20, notes: 21, lastRollup: 22, lastUpdated: 23
};
const PRJ_WIDTH = 23;

function projectsSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET);
  if (!sheet) throw new Error('Projects sheet is missing. Run Setup / Repair ERP Upgrade first.');
  return sheet;
}

function projectRowToObject_(row, sheetRow) {
  const contractValue = Number(row[PRJ_COL.contractValue - 1] || 0);
  const actualSpend = Number(row[PRJ_COL.actualSpend - 1] || 0);
  return {
    row: sheetRow,
    projectId: String(row[PRJ_COL.projectId - 1] || ''),
    projectName: String(row[PRJ_COL.projectName - 1] || ''),
    clientId: String(row[PRJ_COL.clientId - 1] || ''),
    client: String(row[PRJ_COL.client - 1] || ''),
    leadId: String(row[PRJ_COL.leadId - 1] || ''),
    manager: String(row[PRJ_COL.manager - 1] || ''),
    startDate: erpDate_(row[PRJ_COL.startDate - 1]),
    endDate: erpDate_(row[PRJ_COL.endDate - 1]),
    status: String(row[PRJ_COL.status - 1] || ''),
    contractValue: contractValue,
    budget: Number(row[PRJ_COL.budget - 1] || 0),
    plannedCost: Number(row[PRJ_COL.plannedCost - 1] || 0),
    actualSpend: actualSpend,
    budgetVariance: Number(row[PRJ_COL.budgetVariance - 1] || 0),
    balance: Number(row[PRJ_COL.balance - 1] || 0),
    marginPct: Number(row[PRJ_COL.marginPct - 1] || 0),
    invoiced: Number(row[PRJ_COL.invoiced - 1] || 0),
    received: Number(row[PRJ_COL.received - 1] || 0),
    receivable: Number(row[PRJ_COL.receivable - 1] || 0),
    location: String(row[PRJ_COL.location - 1] || ''),
    notes: String(row[PRJ_COL.notes - 1] || ''),
    lastRollup: erpDate_(row[PRJ_COL.lastRollup - 1]),
    budgetUsedPct: Number(row[PRJ_COL.budget - 1] || 0) > 0
      ? Math.round(actualSpend / Number(row[PRJ_COL.budget - 1]) * 1000) / 10
      : 0,
    overBudget: Number(row[PRJ_COL.budget - 1] || 0) > 0 && actualSpend > Number(row[PRJ_COL.budget - 1])
  };
}

function getERPProjects_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, PRJ_WIDTH).getValues()
    .map((row, index) => projectRowToObject_(row, index + 2))
    .filter(project => project.projectId || project.projectName);
}

function findProjectRow_(projectId) {
  const sheet = projectsSheet_();
  if (sheet.getLastRow() < 2 || !String(projectId || '').trim()) return null;
  const match = sheet.getRange(2, PRJ_COL.projectId, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(projectId)).matchCase(true).matchEntireCell(true).findNext();
  if (!match) return null;
  return projectRowToObject_(sheet.getRange(match.getRow(), 1, 1, PRJ_WIDTH).getValues()[0], match.getRow());
}

/** Project options for dropdowns, cheap enough to ship on every portal load. */
function getProjectOptionList_() {
  return getERPProjects_()
    .filter(project => ['Planning', 'Active', 'On Hold'].indexOf(project.status) !== -1)
    .map(project => ({projectId: project.projectId, label: project.projectName + ' - ' + project.projectId}));
}

/*************** ROLL-UPS ***************/

/**
 * Recomputes the derived financial columns for the given project IDs, or for
 * every project when projectIds is omitted. Cost lines are read once per call.
 */
function applyProjectRollups_(projectIds) {
  const sheet = projectsSheet_();
  if (sheet.getLastRow() < 2) return 0;
  const totals = computeProjectCostTotals_();
  const wanted = projectIds && projectIds.length
    ? projectIds.map(id => String(id || '').trim()).filter(Boolean)
    : null;
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, PRJ_WIDTH);
  const values = range.getValues();
  const now = new Date();
  let updated = 0;

  values.forEach(function (row) {
    const projectId = String(row[PRJ_COL.projectId - 1] || '').trim();
    if (!projectId) return;
    if (wanted && wanted.indexOf(projectId) === -1) return;
    const total = totals[projectId] || {planned: 0, actual: 0};
    const contractValue = Number(row[PRJ_COL.contractValue - 1] || 0);
    const budget = Number(row[PRJ_COL.budget - 1] || 0);
    const invoiced = Number(row[PRJ_COL.invoiced - 1] || 0);
    const received = Number(row[PRJ_COL.received - 1] || 0);

    row[PRJ_COL.plannedCost - 1] = total.planned;
    row[PRJ_COL.actualSpend - 1] = total.actual;
    row[PRJ_COL.budgetVariance - 1] = budget - total.actual;
    row[PRJ_COL.balance - 1] = contractValue - total.actual;
    row[PRJ_COL.marginPct - 1] = contractValue > 0
      ? Math.round((contractValue - total.actual) / contractValue * 1000) / 10
      : 0;
    row[PRJ_COL.receivable - 1] = invoiced - received;
    row[PRJ_COL.lastRollup - 1] = now;
    updated++;
  });

  if (updated) range.setValues(values);
  return updated;
}

function recomputeProjectRollups_(projectId) {
  if (!String(projectId || '').trim()) return 0;
  return applyProjectRollups_([projectId]);
}

/**
 * Nightly safety net. Catches projects whose cost lines were edited directly in
 * the spreadsheet, bypassing the portal.
 */
function reconcileAllProjectRollups() {
  const updated = applyProjectRollups_(null);
  recordActivity_('System', 'Project Reconcile', 'SYSTEM', 'UPDATE', 'Projects', '', 'Recomputed rollups for ' + updated + ' project(s)');
  Logger.log('Reconciled project rollups: ' + updated + ' project(s).');
  return updated;
}

/*************** WRITE APIS ***************/

function erpAddProject(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'projects');
    const name = String(p.projectName || '').trim();
    if (!name) return {success: false, message: 'Project name is required.'};
    const status = normalizeProjectStatus_(p.status || 'Planning');
    const contractValue = erpNonNegativeAmount_(p.contractValue, 'Contract Value');
    const budget = erpNonNegativeAmount_(p.budget, 'Expenditure Budget');

    // A project typed against a new company builds the Clients master as it
    // goes. Those fields arrive `client`-prefixed so the project's own Location
    // and Notes are never mistaken for the client's own address and remarks.
    const clientEmail = String(p.clientEmail || '').trim();
    if (clientEmail && !isValidEmail_(clientEmail)) {
      return {success: false, message: 'Enter a valid client email address or leave it blank.'};
    }
    const client = erpResolveClientReference_(p.clientId, p.client,
      clientDetailsFromInput_(p, {accountOwner: p.clientAccountOwner}));
    const sheet = projectsSheet_();
    const id = erpNextId_(PROJECTS_SHEET, 'PRJ', PRJ_COL.projectId);
    const now = new Date();
    const row = new Array(PRJ_WIDTH).fill('');
    row[PRJ_COL.projectId - 1] = id;
    row[PRJ_COL.projectName - 1] = name;
    row[PRJ_COL.clientId - 1] = client.clientId;
    row[PRJ_COL.client - 1] = client.clientName;
    row[PRJ_COL.leadId - 1] = String(p.leadId || '').trim();
    row[PRJ_COL.manager - 1] = String(p.manager || '');
    row[PRJ_COL.startDate - 1] = String(p.startDate || '');
    row[PRJ_COL.endDate - 1] = String(p.endDate || '');
    row[PRJ_COL.status - 1] = status;
    row[PRJ_COL.contractValue - 1] = contractValue;
    row[PRJ_COL.budget - 1] = budget;
    row[PRJ_COL.plannedCost - 1] = 0;
    row[PRJ_COL.actualSpend - 1] = 0;
    row[PRJ_COL.budgetVariance - 1] = budget;
    row[PRJ_COL.balance - 1] = contractValue;
    row[PRJ_COL.marginPct - 1] = contractValue > 0 ? 100 : 0;
    row[PRJ_COL.invoiced - 1] = 0;
    row[PRJ_COL.received - 1] = 0;
    row[PRJ_COL.receivable - 1] = 0;
    row[PRJ_COL.location - 1] = String(p.location || 'Abuja');
    row[PRJ_COL.notes - 1] = String(p.notes || '');
    row[PRJ_COL.lastRollup - 1] = now;
    row[PRJ_COL.lastUpdated - 1] = now;
    sheet.appendRow(row);

    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'CREATE', 'Project', id,
      'Project added; contract ' + contractValue + ', budget ' + budget);
    return {success: true, message: 'Project added: ' + id + '.' + clientResolutionNote_(client), projectId: id};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/** Updates the manually-owned project fields. Derived columns are left alone. */
function erpUpdateProject(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'projects');
    const project = findProjectRow_(p.projectId);
    if (!project) return {success: false, message: 'Project not found.'};
    const sheet = projectsSheet_();
    const changes = [];

    if (p.projectName !== undefined && String(p.projectName).trim()) {
      sheet.getRange(project.row, PRJ_COL.projectName).setValue(String(p.projectName).trim());
      changes.push('name');
    }
    if (p.status !== undefined) {
      sheet.getRange(project.row, PRJ_COL.status).setValue(normalizeProjectStatus_(p.status));
      changes.push('status=' + normalizeProjectStatus_(p.status));
    }
    if (p.manager !== undefined) { sheet.getRange(project.row, PRJ_COL.manager).setValue(String(p.manager)); changes.push('manager'); }
    if (p.startDate !== undefined) { sheet.getRange(project.row, PRJ_COL.startDate).setValue(String(p.startDate)); changes.push('startDate'); }
    if (p.endDate !== undefined) { sheet.getRange(project.row, PRJ_COL.endDate).setValue(String(p.endDate)); changes.push('endDate'); }
    if (p.location !== undefined) { sheet.getRange(project.row, PRJ_COL.location).setValue(String(p.location)); changes.push('location'); }
    if (p.notes !== undefined) { sheet.getRange(project.row, PRJ_COL.notes).setValue(String(p.notes)); changes.push('notes'); }
    if (p.contractValue !== undefined && String(p.contractValue) !== '') {
      const value = erpNonNegativeAmount_(p.contractValue, 'Contract Value');
      sheet.getRange(project.row, PRJ_COL.contractValue).setValue(value);
      changes.push('contractValue=' + value);
    }
    if (p.budget !== undefined && String(p.budget) !== '') {
      const value = erpNonNegativeAmount_(p.budget, 'Expenditure Budget');
      sheet.getRange(project.row, PRJ_COL.budget).setValue(value);
      changes.push('budget=' + value);
    }
    if (p.invoiced !== undefined && String(p.invoiced) !== '') {
      const value = erpNonNegativeAmount_(p.invoiced, 'Invoiced Amount');
      sheet.getRange(project.row, PRJ_COL.invoiced).setValue(value);
      changes.push('invoiced=' + value);
    }
    if (p.received !== undefined && String(p.received) !== '') {
      const value = erpNonNegativeAmount_(p.received, 'Amount Received');
      sheet.getRange(project.row, PRJ_COL.received).setValue(value);
      changes.push('received=' + value);
    }
    if (p.clientId !== undefined && String(p.clientId).trim()) {
      const client = erpResolveClientReference_(p.clientId, p.client);
      sheet.getRange(project.row, PRJ_COL.clientId).setValue(client.clientId);
      sheet.getRange(project.row, PRJ_COL.client).setValue(client.clientName);
      changes.push('client=' + client.clientId);
    }
    if (!changes.length) return {success: false, message: 'Nothing to update.'};

    sheet.getRange(project.row, PRJ_COL.lastUpdated).setValue(new Date());
    recomputeProjectRollups_(project.projectId);
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Project', project.projectId, changes.join(', '));
    return {success: true, message: 'Project ' + project.projectId + ' updated.'};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/*************** PROJECT DETAIL (lazy-loaded) ***************/

/**
 * Everything one project needs, fetched on demand. The portal's project table
 * stays summary-only so cost lines and movement history never ship wholesale
 * to the browser.
 */
function erpGetProjectDetail(p) {
  try {
    const user = getAdminUser_(p && p.adminPin);
    if (!user.valid) return {success: false, message: 'Invalid admin PIN.'};
    requireAdminPageRead_(user, 'projects');
    const project = findProjectRow_(p.projectId);
    if (!project) return {success: false, message: 'Project not found.'};

    const costLines = getProjectCostLines_(project.projectId);
    const byCategory = {};
    costLines.forEach(function (line) {
      if (line.paymentStatus === 'Cancelled') return;
      if (!byCategory[line.category]) byCategory[line.category] = {category: line.category, planned: 0, actual: 0, lines: 0};
      byCategory[line.category].planned += line.plannedAmount;
      byCategory[line.category].actual += line.actualAmount;
      byCategory[line.category].lines++;
    });
    const categoryBreakdown = Object.keys(byCategory)
      .map(key => byCategory[key])
      .sort((a, b) => b.actual - a.actual || b.planned - a.planned);

    return {
      success: true,
      project: project,
      costLines: costLines,
      categoryBreakdown: categoryBreakdown,
      requisitions: getProjectLinkedRequisitions_(project.projectId),
      expenses: getProjectLinkedExpenses_(project.projectId),
      inventoryIssues: getProjectInventoryIssues_(project.projectId),
      assetMovements: getProjectAssetMovements_(project.projectId),
      canWrite: adminCanWritePage_(user, 'projects'),
      categories: projectCostCategories_(),
      paymentStatuses: projectCostPaymentStatuses_()
    };
  } catch (err) {
    return {success: false, message: err.message};
  }
}

/*************** LINKED-MODULE READERS ***************/

/** Rows from a sheet whose appended 'Project ID' column matches. */
function readLinkedByProjectId_(sheetName, projectId, mapper) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const width = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(header => String(header || '').trim());
  const column = headers.indexOf('Project ID');
  if (column === -1) return [];
  const wanted = String(projectId || '').trim();
  if (!wanted) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues()
    .filter(row => String(row[column] || '').trim() === wanted)
    .map(mapper);
}

function getProjectLinkedRequisitions_(projectId) {
  return readLinkedByProjectId_(REQ_SHEET, projectId, row => ({
    reqId: String(row[1] || ''), staffName: String(row[3] || ''), status: String(row[10] || ''),
    total: Number(row[22] || 0), itemsSummary: String(row[24] || ''), timestamp: erpDate_(row[0])
  }));
}

function getProjectLinkedExpenses_(projectId) {
  return readLinkedByProjectId_(EXPENSE_CLAIMS_SHEET, projectId, row => ({
    claimId: String(row[1] || ''), staffName: String(row[3] || ''), category: String(row[6] || ''),
    amount: Number(row[8] || 0), status: String(row[9] || ''), expenseDate: erpDate_(row[5])
  }));
}

function getProjectInventoryIssues_(projectId) {
  return readLinkedByProjectId_(INVENTORY_TX_SHEET, projectId, row => ({
    txId: String(row[1] || ''), itemId: String(row[2] || ''), itemName: String(row[3] || ''),
    type: String(row[4] || ''), quantity: Number(row[5] || 0), timestamp: erpDate_(row[0])
  }));
}

function getProjectAssetMovements_(projectId) {
  return readLinkedByProjectId_(ASSET_MOVEMENTS_SHEET, projectId, row => ({
    movementId: String(row[1] || ''), assetId: String(row[2] || ''), assetName: String(row[3] || ''),
    action: String(row[4] || ''), staffName: String(row[7] || ''), timestamp: erpDate_(row[0])
  }));
}

/*************** PORTFOLIO SUMMARY ***************/

function summarizeProjectPortfolio_(projects) {
  const live = projects.filter(project => ['Planning', 'Active', 'On Hold'].indexOf(project.status) !== -1);
  const contractValue = projects.reduce((sum, project) => sum + project.contractValue, 0);
  const actualSpend = projects.reduce((sum, project) => sum + project.actualSpend, 0);
  return {
    total: projects.length,
    active: projects.filter(project => project.status === 'Active').length,
    live: live.length,
    completed: projects.filter(project => project.status === 'Completed').length,
    contractValue: contractValue,
    budget: projects.reduce((sum, project) => sum + project.budget, 0),
    actualSpend: actualSpend,
    balance: contractValue - actualSpend,
    marginPct: contractValue > 0 ? Math.round((contractValue - actualSpend) / contractValue * 1000) / 10 : 0,
    receivable: projects.reduce((sum, project) => sum + project.receivable, 0),
    overBudget: projects.filter(project => project.overBudget).length
  };
}
