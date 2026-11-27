/**
 * Inventory: quantity-tracked stock that moves in and out through transactions.
 *
 * Two fields are routinely confused, so they are named carefully here:
 *   Unit          - the NOUN a quantity is counted in: pcs, roll, metre, box.
 *   Opening Stock - the starting NUMBER on hand when the item was first recorded.
 *
 * They are not duplicates. Current Stock is always Opening + Stock In - Stock Out,
 * recomputed in script (never a row-pinned formula), so the arithmetic survives a
 * sort or a deleted row.
 */

/*************** COLUMN MAP (1-based) ***************/
const INV_COL = {
  itemId: 1, itemName: 2, category: 3, unit: 4, openingStock: 5, stockIn: 6, stockOut: 7,
  currentStock: 8, reorderLevel: 9, location: 10, stockStatus: 11, preferredVendor: 12,
  notes: 13, lastUpdated: 14
};
const INV_WIDTH = 14;

function inventorySheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_ITEMS_SHEET);
  if (!sheet) throw new Error('Inventory Items sheet is missing. Run Setup / Repair ERP Upgrade first.');
  return sheet;
}

function inventoryRowToObject_(r, sheetRow) {
  const current = Number(r[INV_COL.currentStock - 1] || 0);
  const reorder = Number(r[INV_COL.reorderLevel - 1] || 0);
  return {
    row: sheetRow,
    itemId: String(r[INV_COL.itemId - 1] || ''),
    itemName: String(r[INV_COL.itemName - 1] || ''),
    category: String(r[INV_COL.category - 1] || ''),
    unit: String(r[INV_COL.unit - 1] || ''),
    openingStock: Number(r[INV_COL.openingStock - 1] || 0),
    stockIn: Number(r[INV_COL.stockIn - 1] || 0),
    stockOut: Number(r[INV_COL.stockOut - 1] || 0),
    currentStock: current,
    reorderLevel: reorder,
    location: String(r[INV_COL.location - 1] || ''),
    stockStatus: String(r[INV_COL.stockStatus - 1] || ''),
    preferredVendor: String(r[INV_COL.preferredVendor - 1] || ''),
    notes: String(r[INV_COL.notes - 1] || ''),
    lowStock: current <= reorder
  };
}

function getERPInventoryItems_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_ITEMS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, INV_WIDTH).getValues()
    .map((r, i) => inventoryRowToObject_(r, i + 2))
    .filter(item => item.itemId || item.itemName);
}

function findInventoryRow_(itemId) {
  const sheet = inventorySheet_();
  if (sheet.getLastRow() < 2 || !String(itemId || '').trim()) return null;
  const match = sheet.getRange(2, INV_COL.itemId, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(itemId)).matchCase(true).matchEntireCell(true).findNext();
  if (!match) return null;
  return inventoryRowToObject_(sheet.getRange(match.getRow(), 1, 1, INV_WIDTH).getValues()[0], match.getRow());
}

/** Current stock and its status, derived rather than stored by hand. */
function inventoryStockStatus_(current, reorder) {
  return Number(current || 0) <= Number(reorder || 0) ? 'Low Stock' : 'OK';
}

function recomputeInventoryRow_(sheetRow) {
  const sheet = inventorySheet_();
  const values = sheet.getRange(sheetRow, INV_COL.openingStock, 1, 5).getValues()[0];
  const current = Number(values[0] || 0) + Number(values[1] || 0) - Number(values[2] || 0);
  sheet.getRange(sheetRow, INV_COL.currentStock).setValue(current);
  sheet.getRange(sheetRow, INV_COL.stockStatus).setValue(inventoryStockStatus_(current, values[4]));
  sheet.getRange(sheetRow, INV_COL.lastUpdated).setValue(new Date());
  return current;
}

/** A non-negative whole quantity. Fractional stock is not tracked. */
function inventoryQuantity_(value, label) {
  if (value === undefined || value === null || String(value).trim() === '') return 0;
  const parsed = Number(String(value).replace(/,/g, '').trim());
  if (!isFinite(parsed)) throw new Error(label + ' must be a number.');
  if (parsed < 0) throw new Error(label + ' cannot be negative.');
  return Math.round(parsed);
}

/*************** CREATE ***************/

function erpAddInventoryItem(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'inventory');
    const name = String(p.itemName || '').trim();
    if (!name) return {success: false, message: 'Item name is required.'};
    if (getERPInventoryItems_().some(item => item.itemName.toLowerCase() === name.toLowerCase())) {
      return {success: false, message: 'An inventory item named ' + name + ' already exists. Edit that item instead.'};
    }
    const created = writeInventoryItem_(p, user.name);
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'CREATE', 'Inventory Item', created.itemId,
      name + ' / ' + created.category + ' / opening ' + created.openingStock + ' ' + created.unit);
    return {success: true, message: 'Inventory item added: ' + created.itemId, itemId: created.itemId};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/** Appends one item row. Shared by the single-item form and the CSV import. */
function writeInventoryItem_(p, actorName) {
  const sheet = inventorySheet_();
  const id = erpNextId_(INVENTORY_ITEMS_SHEET, 'INV', INV_COL.itemId);
  const opening = inventoryQuantity_(p.openingStock, 'Opening stock');
  const reorder = inventoryQuantity_(p.reorderLevel, 'Reorder level');
  const category = normalizeInventoryCategory_(p.category);
  const unit = normalizeInventoryUnit_(p.unit);
  const row = new Array(INV_WIDTH).fill('');
  row[INV_COL.itemId - 1] = id;
  row[INV_COL.itemName - 1] = String(p.itemName || '').trim();
  row[INV_COL.category - 1] = category;
  row[INV_COL.unit - 1] = unit;
  row[INV_COL.openingStock - 1] = opening;
  row[INV_COL.stockIn - 1] = 0;
  row[INV_COL.stockOut - 1] = 0;
  row[INV_COL.currentStock - 1] = opening;
  row[INV_COL.reorderLevel - 1] = reorder;
  row[INV_COL.location - 1] = String(p.location || '').trim();
  row[INV_COL.stockStatus - 1] = inventoryStockStatus_(opening, reorder);
  row[INV_COL.preferredVendor - 1] = String(p.preferredVendor || '').trim();
  row[INV_COL.notes - 1] = String(p.notes || '').trim();
  row[INV_COL.lastUpdated - 1] = new Date();
  sheet.appendRow(row);
  return {itemId: id, category: category, unit: unit, openingStock: opening, reorderLevel: reorder};
}

/*************** UPDATE ***************/

/**
 * Edits an item's details. Stock normally moves through transactions, but a
 * correction to the opening balance is allowed here because a typo at creation
 * has no other remedy; the current stock is recomputed either way.
 */
function erpUpdateInventoryItem(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'inventory');
    return applyInventoryItemUpdate_(p, user);
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/**
 * Applies an edit to one item. Takes no lock and performs no authentication, so
 * callers must already hold both; the bulk import relies on that to avoid
 * re-entering the script lock and re-reading Admin Users for every row.
 */
function applyInventoryItemUpdate_(p, user) {
  const item = findInventoryRow_(p.itemId);
  if (!item) return {success: false, message: 'Inventory item not found.'};
  const sheet = inventorySheet_();
  const changes = [];
  if (p.itemName !== undefined) {
    const name = String(p.itemName).trim();
    if (!name) return {success: false, message: 'Item name is required.'};
    if (name.toLowerCase() !== item.itemName.toLowerCase() &&
        getERPInventoryItems_().some(other => other.itemId !== item.itemId &&
          other.itemName.toLowerCase() === name.toLowerCase())) {
      return {success: false, message: 'Another inventory item is already named ' + name + '.'};
    }
    if (name !== item.itemName) {
      sheet.getRange(item.row, INV_COL.itemName).setValue(name);
      changes.push('name ' + item.itemName + ' -> ' + name);
    }
  }
  if (p.category !== undefined) {
    const category = normalizeInventoryCategory_(p.category);
    if (category !== item.category) {
      sheet.getRange(item.row, INV_COL.category).setValue(category);
      changes.push('category=' + category);
    }
  }
  if (p.unit !== undefined) {
    const unit = normalizeInventoryUnit_(p.unit);
    if (unit !== item.unit) {
      sheet.getRange(item.row, INV_COL.unit).setValue(unit);
      changes.push('unit=' + unit);
    }
  }
  if (p.openingStock !== undefined && String(p.openingStock).trim() !== '') {
    const opening = inventoryQuantity_(p.openingStock, 'Opening stock');
    if (opening !== item.openingStock) {
      sheet.getRange(item.row, INV_COL.openingStock).setValue(opening);
      changes.push('openingStock ' + item.openingStock + ' -> ' + opening);
    }
  }
  if (p.reorderLevel !== undefined && String(p.reorderLevel).trim() !== '') {
    const reorder = inventoryQuantity_(p.reorderLevel, 'Reorder level');
    if (reorder !== item.reorderLevel) {
      sheet.getRange(item.row, INV_COL.reorderLevel).setValue(reorder);
      changes.push('reorderLevel=' + reorder);
    }
  }
  if (p.location !== undefined) { sheet.getRange(item.row, INV_COL.location).setValue(String(p.location).trim()); changes.push('location'); }
  if (p.preferredVendor !== undefined) { sheet.getRange(item.row, INV_COL.preferredVendor).setValue(String(p.preferredVendor).trim()); changes.push('vendor'); }
  if (p.notes !== undefined) { sheet.getRange(item.row, INV_COL.notes).setValue(String(p.notes).trim()); changes.push('notes'); }

  if (!changes.length) return {success: false, message: 'Nothing was changed.'};
  const current = recomputeInventoryRow_(item.row);
  recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Inventory Item', item.itemId, changes.join(', '));
  return {success: true, message: 'Inventory item ' + item.itemId + ' updated. Current stock: ' + current + '.'};
}

/*************** DELETE ***************/

/** How many transactions reference an item, so the confirmation can say so. */
function countInventoryTransactions_(itemId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_TX_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const wanted = String(itemId || '').trim();
  if (!wanted) return 0;
  return sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues()
    .filter(r => String(r[0] || '').trim() === wanted).length;
}

/** Everything the browser needs to warn before a delete. */
function erpGetInventoryDeletePreview(p) {
  try {
    const user = getAdminUser_(p && p.adminPin);
    if (!user.valid) return {success: false, message: user.locked ? PIN_LOCKED_MESSAGE : 'Invalid admin PIN.'};
    requireAdminPageWrite_(user, 'inventory');
    const item = findInventoryRow_(p.itemId);
    if (!item) return {success: false, message: 'Inventory item not found.'};
    return {
      success: true,
      item: item,
      transactionCount: countInventoryTransactions_(item.itemId)
    };
  } catch (err) {
    return {success: false, message: err.message};
  }
}

/**
 * Removes an item from the register. Its transaction history is deliberately
 * left in place: those rows record the item ID and name, so the movement log
 * stays readable and auditable after the item itself is retired.
 */
function erpDeleteInventoryItem(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'inventory');
    const item = findInventoryRow_(p.itemId);
    if (!item) return {success: false, message: 'Inventory item not found.'};
    const transactions = countInventoryTransactions_(item.itemId);
    if (transactions > 0 && p.confirmWithHistory !== true) {
      return {success: false, message: item.itemId + ' has ' + transactions +
        ' stock transaction(s). Confirm again to remove the item and keep that history.'};
    }
    inventorySheet_().deleteRow(item.row);
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'DELETE', 'Inventory Item', item.itemId,
      'Removed ' + item.itemName + '; ' + transactions + ' transaction(s) retained in the movement log');
    return {success: true, message: 'Inventory item ' + item.itemId + ' removed.' +
      (transactions ? ' Its ' + transactions + ' transaction(s) remain in the movement log.' : '')};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/*************** BULK IMPORT ***************/

/**
 * Bulk-loads items parsed from a CSV or spreadsheet export.
 *
 * Rows arrive already parsed by the browser. Every row is validated here as
 * well, because the browser is not a trust boundary. An existing item name
 * updates that item rather than creating a duplicate, so re-importing a
 * corrected sheet is safe.
 */
function erpImportInventoryItems(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'inventory');
    const rows = Array.isArray(p.rows) ? p.rows : [];
    if (!rows.length) return {success: false, message: 'No rows were supplied. Choose a CSV file first.'};
    if (rows.length > 500) {
      return {success: false, message: 'Import up to 500 rows at a time. Split the file and run it again.'};
    }

    const existing = {};
    getERPInventoryItems_().forEach(item => { existing[item.itemName.toLowerCase()] = item; });
    const seen = {};
    const created = [], updated = [], skipped = [];

    rows.forEach(function (raw, index) {
      const lineNumber = index + 2; // +1 for the header row, +1 for 1-based lines
      try {
        const name = String(raw.itemName || '').trim();
        if (!name) { skipped.push('Line ' + lineNumber + ': item name is blank.'); return; }
        const key = name.toLowerCase();
        if (seen[key]) { skipped.push('Line ' + lineNumber + ': ' + name + ' appears more than once in the file.'); return; }
        seen[key] = true;

        const match = existing[key];
        if (match) {
          // Only fields the file actually carried a value for are applied, so a
          // CSV that omits a column never silently blanks existing data.
          const update = {itemId: match.itemId, itemName: name};
          ['category','unit','openingStock','reorderLevel','location','preferredVendor','notes']
            .forEach(function (field) {
              if (String(raw[field] == null ? '' : raw[field]).trim() !== '') update[field] = raw[field];
            });
          const result = applyInventoryItemUpdate_(update, user);
          if (result.success) updated.push(match.itemId + ' ' + name);
          else skipped.push('Line ' + lineNumber + ': ' + name + ' unchanged (' + result.message + ')');
          return;
        }
        const row = writeInventoryItem_({
          itemName: name, category: raw.category, unit: raw.unit, openingStock: raw.openingStock,
          reorderLevel: raw.reorderLevel, location: raw.location,
          preferredVendor: raw.preferredVendor, notes: raw.notes
        }, user.name);
        existing[key] = {itemId: row.itemId, itemName: name};
        created.push(row.itemId + ' ' + name);
      } catch (err) {
        skipped.push('Line ' + lineNumber + ': ' + (err.message || err));
      }
    });

    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'IMPORT', 'Inventory Item', '',
      created.length + ' created, ' + updated.length + ' updated, ' + skipped.length + ' skipped');
    const parts = [created.length + ' item(s) created', updated.length + ' updated'];
    if (skipped.length) parts.push(skipped.length + ' skipped');
    return {
      success: true,
      message: 'Import complete: ' + parts.join(', ') + '.',
      created: created, updated: updated, skipped: skipped
    };
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/** The import template, so the browser and the docs cannot drift apart. */
function inventoryImportColumns_() {
  return [
    {key: 'itemName', header: 'Item Name', required: true, note: 'Must be unique. A name already in the register updates that item.'},
    {key: 'category', header: 'Category', required: false, note: 'Must match the category list; anything else becomes Other.'},
    {key: 'unit', header: 'Unit', required: false, note: 'Unit of measure, for example pcs or roll. Defaults to pcs.'},
    {key: 'openingStock', header: 'Opening Stock', required: false, note: 'Whole number on hand now. Defaults to 0.'},
    {key: 'reorderLevel', header: 'Reorder Level', required: false, note: 'Low Stock shows at or below this. Defaults to 0.'},
    {key: 'location', header: 'Location', required: false, note: 'Where the stock is held.'},
    {key: 'preferredVendor', header: 'Preferred Vendor', required: false, note: 'Who you usually buy it from.'},
    {key: 'notes', header: 'Notes', required: false, note: 'Anything else worth recording.'}
  ];
}

/*************** TRANSACTIONS ***************/

function erpPostInventoryTransaction(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'inventory');
    const itemId = String(p.itemId || '').trim();
    const qty = inventoryQuantity_(p.quantity, 'Quantity');
    if (!itemId) return {success: false, message: 'Select an inventory item.'};
    if (qty <= 0) return {success: false, message: 'Quantity must be greater than zero.'};
    const type = matchFromList_(p.type, inventoryTransactionTypes_(), '');
    if (!type) return {success: false, message: 'Select a transaction type.'};

    const item = findInventoryRow_(itemId);
    if (!item) return {success: false, message: 'Inventory item not found.'};

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = inventorySheet_();
    if (type === 'Stock Out' && qty > item.currentStock) {
      return {success: false, message: 'Only ' + item.currentStock + ' ' + item.unit +
        ' of ' + item.itemName + ' are in stock. Record a Stock In first, or reduce the quantity.'};
    }
    if (type === 'Stock In' || type === 'Return') sheet.getRange(item.row, INV_COL.stockIn).setValue(item.stockIn + qty);
    if (type === 'Stock Out') sheet.getRange(item.row, INV_COL.stockOut).setValue(item.stockOut + qty);
    // An adjustment restates the opening balance to make the current stock match
    // a physical count, so the in/out history is left untouched.
    if (type === 'Adjustment') sheet.getRange(item.row, INV_COL.openingStock).setValue(qty - item.stockIn + item.stockOut);
    const current = recomputeInventoryRow_(item.row);

    const tx = ss.getSheetByName(INVENTORY_TX_SHEET);
    const txId = erpNextId_(INVENTORY_TX_SHEET, 'TXN', 2);
    tx.appendRow([new Date(), txId, itemId, item.itemName, type, qty, String(p.project || ''),
      user.name, user.name, String(p.reference || ''), String(p.notes || ''),
      resolveSubmittedProjectId_(p.projectId, p.project)]);
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Inventory Item', itemId,
      type + ' ' + qty + ' ' + item.unit + '; stock now ' + current);
    return {success: true, message: type + ' of ' + qty + ' ' + item.unit + ' posted (' + txId +
      '). ' + item.itemName + ' now at ' + current + ' ' + item.unit + '.'};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}
