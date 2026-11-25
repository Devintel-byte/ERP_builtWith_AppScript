/**
 * Phase 1 schema migration for the Projects / Leads pipeline.
 *
 * Safe to run repeatedly. Every step checks its own completion first, so a
 * second run is a no-op rather than a duplicate. Sheets whose column order
 * changes are backed up before they are rewritten.
 *
 * Run from the spreadsheet menu: Moving Surface ERP > Migrate Projects / Leads Schema.
 */
function migrateProjectsAndLeadsSchema() {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  const report = [];
  try {
    report.push(migrateCrmClientsToClientsAndLeads_());
    report.push(migrateProjectsSheetColumns_());
    report.push(ensureProjectIdColumns_());
    report.push(migrateCrmPagePermissions_());
    report.push(stripRowPinnedFormulas_());
    report.push(migrateAdminAccountsToStaffPins_());
    recordActivity_('System','Schema Migration','SYSTEM','MIGRATE','ERP Schema','projects-leads',report.join(' | ').slice(0,2000));
  } finally {
    lock.releaseLock();
  }
  const summary = 'Projects / Leads migration complete.\n\n' + report.join('\n');
  Logger.log(summary);
  return summary;
}

/** Copies a sheet to a timestamped hidden backup tab before a destructive rewrite. */
function backupSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(sheetName);
  if (!source || source.getLastRow() < 2) return '';
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Lagos', 'yyyyMMdd-HHmm');
  const backupName = ('Backup ' + sheetName + ' ' + stamp).slice(0, 99);
  if (ss.getSheetByName(backupName)) return backupName;
  const copy = source.copyTo(ss);
  copy.setName(backupName);
  copy.hideSheet();
  return backupName;
}

/** Reads a sheet into objects keyed by its own header text. */
function readSheetAsObjects_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const width = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(header => String(header || '').trim());
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues()
    .filter(row => row.some(cell => String(cell == null ? '' : cell).trim() !== ''))
    .map(row => {
      const record = {};
      headers.forEach((header, index) => { if (header) record[header] = row[index]; });
      return record;
    });
}

/** Rewrites a sheet with new headers, carrying old values across by header name. */
function rewriteSheetWithHeaders_(sheetName, headers, records, rowBuilder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  const width = Math.max(headers.length, sheet.getLastColumn());
  if (sheet.getLastRow() >= 1) {
    sheet.getRange(1, 1, sheet.getLastRow(), width).clearContent().clearDataValidations();
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const rows = records.map(rowBuilder);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  return rows.length;
}

/**
 * Step 1: 'CRM Clients' becomes the 'Clients' master, and any pipeline data it
 * carried (status, opportunity value, next follow-up) becomes a seed Leads row
 * so no field is dropped in the rename.
 */
function migrateCrmClientsToClientsAndLeads_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const legacy = ss.getSheetByName('CRM Clients');
  const existingClients = ss.getSheetByName(CLIENTS_SHEET);
  if (!legacy && existingClients) return 'Clients: already migrated.';
  if (!legacy) return 'Clients: no legacy CRM Clients sheet found; setup will create the new sheets.';

  const records = readSheetAsObjects_('CRM Clients');
  const backup = backupSheet_('CRM Clients');

  // Reuse the legacy tab as the Clients master so its sheet id, and therefore
  // any bookmarked link to it, survives the rename.
  if (existingClients && existingClients.getSheetId() !== legacy.getSheetId()) {
    ss.deleteSheet(legacy);
  } else {
    legacy.setName(CLIENTS_SHEET);
  }

  const clientCount = rewriteSheetWithHeaders_(CLIENTS_SHEET, clientsHeaders_(), records, record => [
    String(record['Client ID'] || ''),
    String(record['Client / Company'] || ''),
    String(record['Contact Person'] || ''),
    '',
    String(record['Email'] || ''),
    String(record['Phone'] || ''),
    String(record['Sector'] || ''),
    String(record['Location'] || ''),
    '',
    normalizeRelationshipStatus_(record['Status']),
    '',
    '',
    '',
    record['Last Contact Date'] || '',
    String(record['Notes'] || ''),
    new Date()
  ]);

  // Carry the pipeline fields into Leads so the rename loses nothing.
  const pipelineRecords = records.filter(record =>
    record['Next Follow-up'] || Number(record['Opportunity Value'] || 0) > 0
  );
  let leadCount = 0;
  if (!ss.getSheetByName(LEADS_SHEET)) {
    leadCount = rewriteSheetWithHeaders_(LEADS_SHEET, leadsHeaders_(), pipelineRecords, (record, index) => {
      const stage = legacyStatusToLeadStage_(record['Status']);
      const estimated = Number(record['Opportunity Value'] || 0);
      const probability = leadStageProbability_(stage);
      return [
        'LEAD-' + String(index + 1).padStart(4, '0'),
        String(record['Client ID'] || ''),
        String(record['Client / Company'] || ''),
        String(record['Notes'] || record['Client / Company'] || 'Migrated opportunity'),
        String(record['Contact Person'] || ''),
        String(record['Email'] || ''),
        String(record['Phone'] || ''),
        'Existing Client',
        stage,
        probability,
        estimated,
        0,
        Math.round(estimated * probability) / 100,
        '',
        record['Last Contact Date'] || '',
        record['Next Follow-up'] || '',
        'Weekly',
        '',
        '',
        '',
        '',
        'Migrated from CRM Clients during the Leads upgrade.',
        new Date(),
        new Date()
      ];
    });
  }
  return 'Clients: ' + clientCount + ' row(s) migrated' + (backup ? ' (backup: ' + backup + ')' : '') +
    '; Leads seeded with ' + leadCount + ' pipeline row(s).';
}

function normalizeRelationshipStatus_(value) {
  const supplied = String(value || '').trim();
  const match = clientRelationshipStatuses_().filter(status => status.toLowerCase() === supplied.toLowerCase());
  if (match.length) return match[0];
  return supplied ? 'Active' : 'Prospect';
}

function legacyStatusToLeadStage_(value) {
  const supplied = String(value || '').trim().toLowerCase();
  if (supplied === 'lead' || supplied === 'prospect') return 'New / Identified';
  if (supplied === 'negotiation') return 'Negotiation';
  if (supplied === 'won' || supplied === 'contracted') return 'Won (Contracted)';
  if (supplied === 'lost') return 'Lost';
  if (supplied === 'dormant' || supplied === 'inactive') return 'Dormant';
  return 'Qualified';
}

/**
 * Step 2: widen Projects from the legacy 14-column layout to the revenue /
 * budget / actual-spend model, mapping old values across by header name.
 */
function migrateProjectsSheetColumns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECTS_SHEET);
  if (!sheet) return 'Projects: sheet missing; setup will create it.';
  const headers = projectsHeaders_();
  const current = sheet.getRange(1, 1, 1, Math.max(headers.length, sheet.getLastColumn())).getValues()[0]
    .map(header => String(header || '').trim());
  if (headers.every((header, index) => current[index] === header)) return 'Projects: already migrated.';

  const records = readSheetAsObjects_(PROJECTS_SHEET);
  const backup = backupSheet_(PROJECTS_SHEET);
  let budgetsCarried = 0;
  const count = rewriteSheetWithHeaders_(PROJECTS_SHEET, headers, records, record => {
    // The legacy 'Budget' column is named as a spend budget, so it migrates to
    // Expenditure Budget. Contract Value is left at zero for deliberate entry
    // rather than guessed at from an ambiguous column.
    const budget = Number(record['Budget'] || 0);
    if (budget > 0) budgetsCarried++;
    return [
      String(record['Project ID'] || ''),
      String(record['Project Name'] || ''),
      '',
      String(record['Client'] || ''),
      '',
      String(record['Project Manager'] || ''),
      record['Start Date'] || '',
      record['End Date'] || '',
      normalizeProjectStatus_(record['Status']),
      0,
      budget,
      0, 0, 0, 0, 0,
      0, 0, 0,
      String(record['Location'] || ''),
      String(record['Notes'] || ''),
      '',
      new Date()
    ];
  });
  return 'Projects: ' + count + ' row(s) migrated to the revenue/budget model' +
    (backup ? ' (backup: ' + backup + ')' : '') + '; ' + budgetsCarried +
    ' legacy Budget value(s) carried into Expenditure Budget. Set Contract Value per project to activate margin reporting.';
}

function normalizeProjectStatus_(value) {
  const supplied = String(value || '').trim();
  if (supplied.toLowerCase() === 'lead') return 'Planning';
  const match = projectStatuses_().filter(status => status.toLowerCase() === supplied.toLowerCase());
  return match.length ? match[0] : 'Planning';
}

/**
 * Step 3: give every cost-bearing module a Project ID key. The column is
 * appended at the end of each sheet so existing column positions - and the
 * hardcoded indices in the working requisition, expense, inventory and asset
 * code - stay valid.
 */
function ensureProjectIdColumns_() {
  const targets = [REQ_SHEET, EXPENSE_CLAIMS_SHEET, INVENTORY_TX_SHEET, ASSET_MOVEMENTS_SHEET, PURCHASE_ORDERS_SHEET];
  const added = [];
  targets.forEach(function (sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return;
    const width = Math.max(1, sheet.getLastColumn());
    const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(header => String(header || '').trim());
    if (headers.indexOf('Project ID') !== -1) return;
    const column = headers.filter(String).length + 1;
    if (sheet.getMaxColumns() < column) sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
    sheet.getRange(1, column).setValue('Project ID');
    added.push(sheetName);
  });
  return added.length
    ? 'Project ID column added to: ' + added.join(', ') + '.'
    : 'Project ID columns: already present on all cost-bearing sheets.';
}

/**
 * Step 4: rewrite the persisted 'crm' page token to 'leads'.
 *
 * Page IDs are stored as literal text in the Admin Users Access Pages and
 * Write Pages cells, and getAdminAccess_ discards any token that is not on its
 * whitelist. Without this step every administrator holding an explicit page
 * list would silently lose the module after the rename.
 */
function migrateCrmPagePermissions_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 'Page permissions: no Admin Users rows to migrate.';
  const range = sheet.getRange(2, 11, sheet.getLastRow() - 1, 2);
  const values = range.getValues();
  let changed = 0;
  const updated = values.map(function (row) {
    return row.map(function (cell) {
      const tokens = String(cell || '').split(',').map(token => token.trim()).filter(Boolean);
      if (!tokens.length) return cell;
      let rowChanged = false;
      const next = [];
      tokens.forEach(function (token) {
        const value = token.toLowerCase() === 'crm' ? 'leads' : token;
        if (value !== token) rowChanged = true;
        if (next.indexOf(value) === -1) next.push(value);
      });
      if (!rowChanged) return cell;
      changed++;
      return next.join(',');
    });
  });
  if (changed) range.setValues(updated);
  return 'Page permissions: ' + changed + ' cell(s) rewritten from the crm token to leads.';
}

/**
 * Step 5: replace row-pinned formulas with plain values.
 *
 * Projects.Balance and the Inventory stock/status columns were written as
 * formulas carrying hardcoded row numbers at append time, so sorting a sheet or
 * deleting a row silently pointed them at the wrong record. Inventory keeps
 * computed values; project balances are recomputed from cost lines instead.
 */
function stripRowPinnedFormulas_() {
  const inventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_ITEMS_SHEET);
  let inventoryFixed = 0;
  if (inventory && inventory.getLastRow() >= 2) {
    const range = inventory.getRange(2, 1, inventory.getLastRow() - 1, 14);
    const formulas = range.getFormulas();
    const values = range.getValues();
    let changed = false;
    formulas.forEach(function (row, index) {
      if (!row[7] && !row[10]) return;
      const current = Number(values[index][4] || 0) + Number(values[index][5] || 0) - Number(values[index][6] || 0);
      values[index][7] = current;
      values[index][10] = current <= Number(values[index][8] || 0) ? 'Low Stock' : 'OK';
      changed = true;
      inventoryFixed++;
    });
    if (changed) range.setValues(values);
  }
  return 'Row-pinned formulas: ' + inventoryFixed + ' inventory row(s) converted to computed values; ' +
    'project balances are now recomputed from Project Cost Lines.';
}


/**
 * Step 6: one PIN per person.
 *
 * Admin Users used to carry its own PIN column, separate from the PIN in Staff
 * List, so a person had two PINs and a reset or recovery only changed one. Sign-in
 * now checks Staff List for every account linked to a staff member.
 *
 *  - Rows with a blank Staff ID are linked to the staff member whose name matches
 *    EXACTLY (case and spacing aside). Ambiguous or missing names are left alone
 *    and listed, because guessing a link would hand someone else's staff identity
 *    to an access account.
 *  - A legacy admin PIN is moved into Staff List only when that person has no
 *    valid PIN there yet. When both exist and differ, Staff List wins and the
 *    conflict is counted: that person now signs in to the ERP with their staff PIN.
 *  - The Admin Users PIN of every linked row is then cleared so the credential is
 *    stored once.
 *
 * Idempotent, and backs both sheets up before writing. The report holds names and
 * counts only, never a PIN.
 */
function migrateAdminAccountsToStaffPins_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const admin = ss.getSheetByName(ADMIN_USERS_SHEET);
  const staffSheet = ss.getSheetByName(STAFF_SHEET);
  if (!admin || admin.getLastRow() < 2 || !staffSheet || staffSheet.getLastRow() < 2) {
    return 'Admin accounts: nothing to link to staff PINs.';
  }
  const adminRange = admin.getRange(2, 1, admin.getLastRow() - 1, 12);
  const rows = adminRange.getValues();
  const staffRows = staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, 18).getValues();

  const byId = Object.create(null);
  const byName = Object.create(null);
  staffRows.forEach(function (row, index) {
    const id = String(row[0] || '').trim();
    if (id) byId[id] = index;
    const name = String(row[1] || '').trim().toLowerCase().replace(/\s+/g, ' ');
    if (name) (byName[name] = byName[name] || []).push(index);
  });

  let linked = 0, moved = 0, conflicts = 0, cleared = 0;
  const unlinked = [], awaitingPin = [], staffPinWrites = [];
  let adminChanged = false;

  rows.forEach(function (row, i) {
    if (!String(row[0] || '').trim() && !String(row[8] || '').trim()) return;
    let id = String(row[8] || '').trim();
    if (!id) {
      const hits = byName[String(row[0] || '').trim().toLowerCase().replace(/\s+/g, ' ')] || [];
      if (hits.length !== 1) {
        unlinked.push(String(row[0] || 'row ' + (i + 2)) + (hits.length ? ' (name matches several staff)' : ' (no matching staff name)'));
        return;
      }
      id = String(staffRows[hits[0]][0]).trim();
      row[8] = id;
      linked++;
      adminChanged = true;
    }
    if (byId[id] === undefined) {
      unlinked.push(String(row[0] || 'row ' + (i + 2)) + ' (Staff ID ' + id + ' not in Staff List)');
      return;
    }
    const staffIndex = byId[id];
    const adminPin = normalizePin_(row[1]);
    let staffPin = normalizePin_(staffRows[staffIndex][4]);
    if (isValidPin_(adminPin)) {
      if (!isValidPin_(staffPin)) {
        staffRows[staffIndex][4] = adminPin;
        staffPinWrites.push({row: staffIndex + 2, pin: adminPin});
        staffPin = adminPin;
        moved++;
      } else if (staffPin !== adminPin) {
        conflicts++;
      }
      row[1] = '';
      cleared++;
      adminChanged = true;
    }
    if (!isValidPin_(staffPin)) awaitingPin.push(String(row[0] || id));
  });

  if (adminChanged) {
    backupSheet_(ADMIN_USERS_SHEET);
    if (staffPinWrites.length) backupSheet_(STAFF_SHEET);
    staffPinWrites.forEach(function (write) {
      const cell = staffSheet.getRange(write.row, 5);
      cell.setNumberFormat('@');
      cell.setValue(write.pin);
    });
    adminRange.setValues(rows);
    invalidatePublicStaffOptionsCache_();
  }
  return 'Admin accounts: ' + linked + ' linked to staff by exact name, ' + moved + ' legacy admin PIN(s) moved to Staff List, ' +
    conflicts + ' conflicting PIN(s) resolved in favour of the Staff List PIN, ' + cleared + ' duplicate admin PIN(s) cleared.' +
    (unlinked.length ? ' Not linked (legacy PIN column still used): ' + unlinked.join('; ') + '.' : '') +
    (awaitingPin.length ? ' No PIN yet, must create one on the attendance page: ' + awaitingPin.join(', ') + '.' : '');
}
