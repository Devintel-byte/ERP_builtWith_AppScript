/**
 * Leads (formerly CRM): the client relationship and opportunity pipeline.
 *
 * Clients is the master record for an organization. Leads holds one row per
 * opportunity, so a repeat client can carry several deals at different stages
 * at once. Projects.Lead ID is the authoritative link between a won deal and
 * its delivery record, which is why no project list is denormalized onto the
 * lead row - the link is always read back from Projects.
 */

/*************** COLUMN MAPS (1-based, see clientsHeaders_ / leadsHeaders_) ***************/
const CLIENT_COL = {
  clientId: 1, clientName: 2, contactPerson: 3, roleTitle: 4, email: 5, phone: 6, sector: 7,
  location: 8, website: 9, relationshipStatus: 10, accountOwner: 11, accountOwnerName: 12,
  firstEngaged: 13, lastContact: 14, notes: 15, lastUpdated: 16
};
const CLIENT_WIDTH = 16;

const LEAD_COL = {
  leadId: 1, clientId: 2, clientName: 3, opportunity: 4, contactPerson: 5, email: 6, phone: 7,
  source: 8, stage: 9, probability: 10, estimatedValue: 11, contractValue: 12, weightedValue: 13,
  expectedClose: 14, lastContact: 15, nextFollowup: 16, cadence: 17, owner: 18, ownerName: 19,
  lastReminderSent: 20, lostReason: 21, notes: 22, createdAt: 23, lastUpdated: 24
};
const LEAD_WIDTH = 24;

function clientsSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLIENTS_SHEET);
  if (!sheet) throw new Error('Clients sheet is missing. Run Setup / Repair ERP Upgrade first.');
  return sheet;
}

function leadsSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS_SHEET);
  if (!sheet) throw new Error('Leads sheet is missing. Run Setup / Repair ERP Upgrade first.');
  return sheet;
}

/*************** CLIENTS ***************/

function clientRowToObject_(row, sheetRow) {
  return {
    row: sheetRow,
    clientId: String(row[CLIENT_COL.clientId - 1] || ''),
    clientName: String(row[CLIENT_COL.clientName - 1] || ''),
    contactPerson: String(row[CLIENT_COL.contactPerson - 1] || ''),
    roleTitle: String(row[CLIENT_COL.roleTitle - 1] || ''),
    email: String(row[CLIENT_COL.email - 1] || ''),
    phone: String(row[CLIENT_COL.phone - 1] || ''),
    sector: String(row[CLIENT_COL.sector - 1] || ''),
    location: String(row[CLIENT_COL.location - 1] || ''),
    website: String(row[CLIENT_COL.website - 1] || ''),
    relationshipStatus: String(row[CLIENT_COL.relationshipStatus - 1] || ''),
    accountOwner: String(row[CLIENT_COL.accountOwner - 1] || ''),
    accountOwnerName: String(row[CLIENT_COL.accountOwnerName - 1] || ''),
    firstEngaged: erpDate_(row[CLIENT_COL.firstEngaged - 1]),
    lastContact: erpDate_(row[CLIENT_COL.lastContact - 1]),
    notes: String(row[CLIENT_COL.notes - 1] || '')
  };
}

function getERPClients_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLIENTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, CLIENT_WIDTH).getValues()
    .map((row, index) => clientRowToObject_(row, index + 2))
    .filter(client => client.clientId || client.clientName);
}

function findClientRow_(clientId) {
  const sheet = clientsSheet_();
  if (sheet.getLastRow() < 2 || !String(clientId || '').trim()) return null;
  const match = sheet.getRange(2, CLIENT_COL.clientId, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(clientId)).matchCase(true).matchEntireCell(true).findNext();
  if (!match) return null;
  return clientRowToObject_(sheet.getRange(match.getRow(), 1, 1, CLIENT_WIDTH).getValues()[0], match.getRow());
}

/**
 * The contact detail fields a client record carries, in the order they are
 * written. Projects and Leads both collect these while a new client is being
 * typed, so the Clients master is built from the form that was already open
 * instead of needing a second trip to Add Client.
 */
function clientDetailFields_() {
  return ['contactPerson', 'roleTitle', 'email', 'phone', 'sector', 'location', 'website', 'notes'];
}

/**
 * Reads the client detail fields out of an Add Project / Add Lead payload.
 *
 * They arrive under `client`-prefixed keys (clientSector, clientPhone, ...)
 * precisely because a project's own Location and Notes are not the client's:
 * sharing the plain names would file an event venue as the client's address.
 * `overrides` lets a form reuse a field it already collects under its own name,
 * which is how the lead's Contact person / Email / Phone reach the master.
 */
function clientDetailsFromInput_(p, overrides) {
  const source = p || {};
  const extra = overrides || {};
  const details = {};
  clientDetailFields_().forEach(function (field) {
    const key = 'client' + field.charAt(0).toUpperCase() + field.slice(1);
    const value = extra[field] !== undefined ? extra[field] : source[key];
    details[field] = String(value == null ? '' : value).trim();
  });
  details.accountOwner = String(extra.accountOwner || source.clientAccountOwner || '').trim();
  return details;
}

/**
 * Resolves a client reference to {clientId, clientName, created, enriched}.
 *
 * Accepts an existing Client ID, or a company name which is matched
 * case-insensitively against the master and created if it is new. This is what
 * stops Projects and Leads from drifting apart on spelling.
 *
 * `details` carries the contact fields typed on whichever form is open. A newly
 * created client is written with all of them. An existing client only has its
 * BLANK fields filled in: a lead form must never silently overwrite a phone
 * number or sector somebody curated on the Clients page, so a conflicting value
 * is left alone and the caller is told which fields were filled.
 */
function erpResolveClientReference_(clientId, clientName, details) {
  const supplied = details || {};
  const suppliedId = String(clientId || '').trim();
  if (suppliedId) {
    const existing = findClientRow_(suppliedId);
    if (existing) {
      return {clientId: existing.clientId, clientName: existing.clientName, created: false,
        enriched: fillBlankClientFields_(existing, supplied)};
    }
  }
  const suppliedName = String(clientName || '').trim();
  if (!suppliedName) return {clientId: '', clientName: '', created: false, enriched: []};

  const match = getERPClients_().filter(client =>
    client.clientName.toLowerCase() === suppliedName.toLowerCase());
  if (match.length) {
    return {clientId: match[0].clientId, clientName: match[0].clientName, created: false,
      enriched: fillBlankClientFields_(match[0], supplied)};
  }

  const owner = erpResolveStaffReference_(supplied.accountOwner);
  const sheet = clientsSheet_();
  const id = erpNextId_(CLIENTS_SHEET, 'CL', CLIENT_COL.clientId);
  const row = new Array(CLIENT_WIDTH).fill('');
  row[CLIENT_COL.clientId - 1] = id;
  row[CLIENT_COL.clientName - 1] = suppliedName;
  clientDetailFields_().forEach(function (field) {
    row[CLIENT_COL[field] - 1] = String(supplied[field] || '');
  });
  row[CLIENT_COL.relationshipStatus - 1] = 'Prospect';
  row[CLIENT_COL.accountOwner - 1] = owner.staffId;
  row[CLIENT_COL.accountOwnerName - 1] = owner.name;
  row[CLIENT_COL.firstEngaged - 1] = erpToday_();
  row[CLIENT_COL.lastContact - 1] = erpToday_();
  row[CLIENT_COL.lastUpdated - 1] = new Date();
  sheet.appendRow(row);
  return {clientId: id, clientName: suppliedName, created: true,
    enriched: clientDetailFields_().filter(field => String(supplied[field] || '') !== '')};
}

/**
 * Writes supplied detail fields onto an existing client, but only where the
 * client's own cell is empty. Returns the field names actually filled.
 */
function fillBlankClientFields_(client, supplied) {
  const filled = [];
  clientDetailFields_().forEach(function (field) {
    const value = String(supplied[field] || '').trim();
    if (value && !String(client[field] || '').trim()) filled.push(field);
  });
  if (!filled.length) return [];
  const sheet = clientsSheet_();
  filled.forEach(function (field) {
    sheet.getRange(client.row, CLIENT_COL[field]).setValue(String(supplied[field]).trim());
  });
  sheet.getRange(client.row, CLIENT_COL.lastUpdated).setValue(new Date());
  return filled;
}

/** "contactPerson" -> "contact person", for the message shown after a save. */
function clientFieldLabel_(field) {
  const labels = {contactPerson: 'contact person', roleTitle: 'role', email: 'email', phone: 'phone',
    sector: 'sector', location: 'location', website: 'website', notes: 'notes'};
  return labels[field] || field;
}

/** The trailing sentence telling the user what happened to the Clients master. */
function clientResolutionNote_(client) {
  if (!client || !client.clientId) return '';
  if (client.created) {
    return ' Client ' + client.clientId + ' created' +
      (client.enriched && client.enriched.length
        ? ' with ' + client.enriched.map(clientFieldLabel_).join(', ') + '.' : '.');
  }
  if (client.enriched && client.enriched.length) {
    return ' Client ' + client.clientId + ' updated: ' +
      client.enriched.map(clientFieldLabel_).join(', ') + ' filled in.';
  }
  return '';
}

function erpAddClient(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'leads');
    const name = String(p.clientName || '').trim();
    if (!name) return {success: false, message: 'Client / company name is required.'};
    const email = String(p.email || '').trim();
    if (email && !isValidEmail_(email)) return {success: false, message: 'Enter a valid email address or leave it blank.'};
    if (getERPClients_().some(client => client.clientName.toLowerCase() === name.toLowerCase())) {
      return {success: false, message: 'A client named ' + name + ' already exists.'};
    }

    const sheet = clientsSheet_();
    const id = erpNextId_(CLIENTS_SHEET, 'CL', CLIENT_COL.clientId);
    const owner = erpResolveStaffReference_(p.accountOwner);
    const row = new Array(CLIENT_WIDTH).fill('');
    row[CLIENT_COL.clientId - 1] = id;
    row[CLIENT_COL.clientName - 1] = name;
    row[CLIENT_COL.contactPerson - 1] = String(p.contactPerson || '');
    row[CLIENT_COL.roleTitle - 1] = String(p.roleTitle || '');
    row[CLIENT_COL.email - 1] = email;
    row[CLIENT_COL.phone - 1] = String(p.phone || '');
    row[CLIENT_COL.sector - 1] = String(p.sector || '');
    row[CLIENT_COL.location - 1] = String(p.location || '');
    row[CLIENT_COL.website - 1] = String(p.website || '');
    row[CLIENT_COL.relationshipStatus - 1] = normalizeRelationshipStatus_(p.relationshipStatus || 'Prospect');
    row[CLIENT_COL.accountOwner - 1] = owner.staffId;
    row[CLIENT_COL.accountOwnerName - 1] = owner.name;
    row[CLIENT_COL.firstEngaged - 1] = erpToday_();
    row[CLIENT_COL.notes - 1] = String(p.notes || '');
    row[CLIENT_COL.lastUpdated - 1] = new Date();
    sheet.appendRow(row);

    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'CREATE', 'Client', id, 'Client added: ' + name);
    return {success: true, message: 'Client added: ' + id, clientId: id};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

function erpResolveStaffReference_(staffId) {
  const supplied = String(staffId || '').trim();
  if (!supplied) return {staffId: '', name: ''};
  const staff = getStaffByIdAnyStatus_(supplied);
  if (!staff) return {staffId: '', name: ''};
  return {staffId: staff.id, name: staff.name};
}

/*************** LEADS ***************/

function leadRowToObject_(row, sheetRow) {
  const nextFollowup = erpDate_(row[LEAD_COL.nextFollowup - 1]);
  const stage = String(row[LEAD_COL.stage - 1] || '');
  const daysToFollowup = nextFollowup ? erpDaysBetween_(erpToday_(), nextFollowup) : null;
  return {
    row: sheetRow,
    leadId: String(row[LEAD_COL.leadId - 1] || ''),
    clientId: String(row[LEAD_COL.clientId - 1] || ''),
    clientName: String(row[LEAD_COL.clientName - 1] || ''),
    opportunity: String(row[LEAD_COL.opportunity - 1] || ''),
    contactPerson: String(row[LEAD_COL.contactPerson - 1] || ''),
    email: String(row[LEAD_COL.email - 1] || ''),
    phone: String(row[LEAD_COL.phone - 1] || ''),
    source: String(row[LEAD_COL.source - 1] || ''),
    stage: stage,
    probability: Number(row[LEAD_COL.probability - 1] || 0),
    estimatedValue: Number(row[LEAD_COL.estimatedValue - 1] || 0),
    contractValue: Number(row[LEAD_COL.contractValue - 1] || 0),
    weightedValue: Number(row[LEAD_COL.weightedValue - 1] || 0),
    expectedClose: erpDate_(row[LEAD_COL.expectedClose - 1]),
    lastContact: erpDate_(row[LEAD_COL.lastContact - 1]),
    nextFollowup: nextFollowup,
    cadence: String(row[LEAD_COL.cadence - 1] || 'Weekly'),
    owner: String(row[LEAD_COL.owner - 1] || ''),
    ownerName: String(row[LEAD_COL.ownerName - 1] || ''),
    lastReminderSent: erpDate_(row[LEAD_COL.lastReminderSent - 1]),
    lostReason: String(row[LEAD_COL.lostReason - 1] || ''),
    notes: String(row[LEAD_COL.notes - 1] || ''),
    daysToFollowup: daysToFollowup,
    followupOverdue: daysToFollowup !== null && daysToFollowup < 0 && leadClosedStages_().indexOf(stage) === -1,
    followupDueSoon: daysToFollowup !== null && daysToFollowup >= 0 && daysToFollowup <= 3 &&
      leadClosedStages_().indexOf(stage) === -1,
    isOpen: leadClosedStages_().indexOf(stage) === -1,
    isWon: leadWonStages_().indexOf(stage) !== -1
  };
}

function getERPLeads_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEADS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, LEAD_WIDTH).getValues()
    .map((row, index) => leadRowToObject_(row, index + 2))
    .filter(lead => lead.leadId || lead.clientName);
}

function findLeadRow_(leadId) {
  const sheet = leadsSheet_();
  if (sheet.getLastRow() < 2 || !String(leadId || '').trim()) return null;
  const match = sheet.getRange(2, LEAD_COL.leadId, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(leadId)).matchCase(true).matchEntireCell(true).findNext();
  if (!match) return null;
  return leadRowToObject_(sheet.getRange(match.getRow(), 1, 1, LEAD_WIDTH).getValues()[0], match.getRow());
}

function normalizeLeadStage_(value) {
  const supplied = String(value || '').trim();
  const match = leadStages_().filter(stage => stage.toLowerCase() === supplied.toLowerCase());
  return match.length ? match[0] : 'New / Identified';
}

function normalizeLeadCadence_(value) {
  const supplied = String(value || '').trim();
  const match = leadReminderCadences_().filter(cadence => cadence.toLowerCase() === supplied.toLowerCase());
  return match.length ? match[0] : 'Weekly';
}

function erpAddLead(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'leads');
    const opportunity = String(p.opportunity || '').trim();
    if (!opportunity) return {success: false, message: 'Opportunity name is required.'};
    if (!String(p.clientId || '').trim() && !String(p.clientName || '').trim()) {
      return {success: false, message: 'Select an existing client or enter a company name.'};
    }
    const email = String(p.email || '').trim();
    if (email && !isValidEmail_(email)) return {success: false, message: 'Enter a valid email address or leave it blank.'};

    const stage = normalizeLeadStage_(p.stage || 'New / Identified');
    const probability = p.probability !== undefined && String(p.probability).trim() !== ''
      ? erpPercent_(p.probability, 'Probability %') : leadStageProbability_(stage);
    const estimatedValue = erpNonNegativeAmount_(p.estimatedValue, 'Estimated Value');
    const nextFollowup = erpDateInput_(p.nextFollowup, 'Next Follow-up');
    const cadence = normalizeLeadCadence_(p.cadence || 'Weekly');
    if (cadence !== 'Off' && !nextFollowup) {
      return {success: false, message: 'Set a Next Follow-up date, or set the reminder cadence to Off.'};
    }
    const owner = erpResolveStaffReference_(p.owner);

    // Resolved only once every input has passed validation: this call can
    // create a Clients row, and a lead that fails afterwards must not leave a
    // half-made client behind. The contact details typed here build that
    // master record, so a new company is never entered twice.
    const client = erpResolveClientReference_(p.clientId, p.clientName, clientDetailsFromInput_(p, {
      contactPerson: p.contactPerson, email: email, phone: p.phone, accountOwner: p.owner, notes: ''
    }));
    if (!client.clientId) return {success: false, message: 'Select an existing client or enter a company name.'};

    const sheet = leadsSheet_();
    const id = erpNextId_(LEADS_SHEET, 'LEAD', LEAD_COL.leadId);
    const now = new Date();
    const row = new Array(LEAD_WIDTH).fill('');
    row[LEAD_COL.leadId - 1] = id;
    row[LEAD_COL.clientId - 1] = client.clientId;
    row[LEAD_COL.clientName - 1] = client.clientName;
    row[LEAD_COL.opportunity - 1] = opportunity;
    row[LEAD_COL.contactPerson - 1] = String(p.contactPerson || '');
    row[LEAD_COL.email - 1] = email;
    row[LEAD_COL.phone - 1] = String(p.phone || '');
    row[LEAD_COL.source - 1] = String(p.source || 'Inbound Enquiry');
    row[LEAD_COL.stage - 1] = stage;
    row[LEAD_COL.probability - 1] = probability;
    row[LEAD_COL.estimatedValue - 1] = estimatedValue;
    row[LEAD_COL.contractValue - 1] = erpNonNegativeAmount_(p.contractValue, 'Contract Value');
    row[LEAD_COL.weightedValue - 1] = Math.round(estimatedValue * probability) / 100;
    row[LEAD_COL.expectedClose - 1] = erpDateInput_(p.expectedClose, 'Expected Close Date');
    row[LEAD_COL.lastContact - 1] = erpDateInput_(p.lastContact, 'Last Contact Date') || erpToday_();
    row[LEAD_COL.nextFollowup - 1] = nextFollowup;
    row[LEAD_COL.cadence - 1] = cadence;
    row[LEAD_COL.owner - 1] = owner.staffId;
    row[LEAD_COL.ownerName - 1] = owner.name;
    row[LEAD_COL.lostReason - 1] = '';
    row[LEAD_COL.notes - 1] = String(p.notes || '');
    row[LEAD_COL.createdAt - 1] = now;
    row[LEAD_COL.lastUpdated - 1] = now;
    sheet.appendRow(row);

    logLeadActivity_({
      leadId: id, clientId: client.clientId, activityType: 'Follow-up Logged', direction: 'Internal',
      contactPerson: String(p.contactPerson || ''), outcome: 'Lead created',
      stageBefore: '', stageAfter: stage, nextFollowupSet: nextFollowup,
      loggedBy: user.name, notes: 'Opportunity opened at stage ' + stage
    });
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'CREATE', 'Lead', id,
      client.clientName + ' / ' + opportunity + ' / ' + stage);
    return {success: true, message: 'Lead added: ' + id + '.' + clientResolutionNote_(client), leadId: id};
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/**
 * Moves a lead through the pipeline and records the move in the activity log.
 * Probability follows the stage unless it is set explicitly.
 */
function erpUpdateLeadStage(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'leads');
    const lead = findLeadRow_(p.leadId);
    if (!lead) return {success: false, message: 'Lead not found.'};
    const stage = normalizeLeadStage_(p.stage);
    if (stage === 'Lost' && !String(p.lostReason || '').trim()) {
      return {success: false, message: 'Record a reason when marking a lead Lost.'};
    }
    const probability = p.probability !== undefined && String(p.probability).trim() !== ''
      ? erpPercent_(p.probability, 'Probability %') : leadStageProbability_(stage);
    const nextFollowup = erpDateInput_(p.nextFollowup, 'Next Follow-up');
    const sheet = leadsSheet_();
    const estimatedValue = p.estimatedValue !== undefined && String(p.estimatedValue).trim() !== ''
      ? erpNonNegativeAmount_(p.estimatedValue, 'Estimated Value') : lead.estimatedValue;
    const contractValue = p.contractValue !== undefined && String(p.contractValue).trim() !== ''
      ? erpNonNegativeAmount_(p.contractValue, 'Contract Value') : lead.contractValue;

    sheet.getRange(lead.row, LEAD_COL.stage).setValue(stage);
    sheet.getRange(lead.row, LEAD_COL.probability).setValue(probability);
    sheet.getRange(lead.row, LEAD_COL.estimatedValue).setValue(estimatedValue);
    sheet.getRange(lead.row, LEAD_COL.contractValue).setValue(contractValue);
    sheet.getRange(lead.row, LEAD_COL.weightedValue).setValue(Math.round(estimatedValue * probability) / 100);
    sheet.getRange(lead.row, LEAD_COL.lastContact).setValue(erpToday_());
    if (nextFollowup) sheet.getRange(lead.row, LEAD_COL.nextFollowup).setValue(nextFollowup);
    // A closed lead stops generating reminders, so its stale date is cleared.
    if (leadClosedStages_().indexOf(stage) !== -1) {
      sheet.getRange(lead.row, LEAD_COL.nextFollowup).setValue('');
      sheet.getRange(lead.row, LEAD_COL.cadence).setValue('Off');
    }
    if (p.lostReason !== undefined) sheet.getRange(lead.row, LEAD_COL.lostReason).setValue(String(p.lostReason || ''));
    sheet.getRange(lead.row, LEAD_COL.lastUpdated).setValue(new Date());

    logLeadActivity_({
      leadId: lead.leadId, clientId: lead.clientId, activityType: String(p.activityType || 'Follow-up Logged'),
      direction: String(p.direction || 'Outbound'), contactPerson: lead.contactPerson,
      outcome: String(p.outcome || ('Stage moved to ' + stage)), stageBefore: lead.stage, stageAfter: stage,
      nextFollowupSet: nextFollowup, loggedBy: user.name, notes: String(p.notes || '')
    });
    touchClientLastContact_(lead.clientId);
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Lead', lead.leadId,
      lead.stage + ' -> ' + stage);

    const won = leadWonStages_().indexOf(stage) !== -1;
    return {
      success: true,
      message: 'Lead ' + lead.leadId + ' moved to ' + stage + '.' +
        (won ? ' Use Create project from lead to open its delivery record.' : ''),
      won: won
    };
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

/** Logs a contact touch and reschedules the next follow-up. */
function erpLogLeadActivity(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const user = erpRequireAdmin_(p.adminPin, 'leads');
    const lead = findLeadRow_(p.leadId);
    if (!lead) return {success: false, message: 'Lead not found.'};
    const outcome = String(p.outcome || '').trim();
    if (!outcome) return {success: false, message: 'Record what came out of the contact.'};
    const nextFollowup = erpDateInput_(p.nextFollowup, 'Next Follow-up');

    const sheet = leadsSheet_();
    sheet.getRange(lead.row, LEAD_COL.lastContact).setValue(erpToday_());
    if (nextFollowup) {
      sheet.getRange(lead.row, LEAD_COL.nextFollowup).setValue(nextFollowup);
      // Clearing the reminder stamp lets the scheduler treat the new date fresh.
      sheet.getRange(lead.row, LEAD_COL.lastReminderSent).setValue('');
    }
    sheet.getRange(lead.row, LEAD_COL.lastUpdated).setValue(new Date());

    const activityId = logLeadActivity_({
      leadId: lead.leadId, clientId: lead.clientId,
      activityType: String(p.activityType || 'Follow-up Logged'),
      direction: String(p.direction || 'Outbound'),
      contactPerson: String(p.contactPerson || lead.contactPerson),
      outcome: outcome, stageBefore: lead.stage, stageAfter: lead.stage,
      nextFollowupSet: nextFollowup, loggedBy: user.name, notes: String(p.notes || '')
    });
    touchClientLastContact_(lead.clientId);
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'CREATE', 'Lead Activity', activityId,
      lead.leadId + ': ' + outcome);
    return {
      success: true,
      message: 'Contact logged for ' + lead.leadId + '.' +
        (nextFollowup ? ' Next follow-up set for ' + nextFollowup + '.' : ' No new follow-up date was set.')
    };
  } catch (err) {
    return {success: false, message: err.message};
  } finally {
    lock.releaseLock();
  }
}

function logLeadActivity_(entry) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEAD_ACTIVITY_SHEET);
  if (!sheet) return '';
  const activityId = erpNextId_(LEAD_ACTIVITY_SHEET, 'ACT', 2);
  sheet.appendRow([
    new Date(), activityId, String(entry.leadId || ''), String(entry.clientId || ''),
    String(entry.activityType || ''), String(entry.direction || ''), String(entry.contactPerson || ''),
    String(entry.outcome || '').slice(0, 1000), String(entry.stageBefore || ''), String(entry.stageAfter || ''),
    String(entry.nextFollowupSet || ''), String(entry.loggedBy || ''), String(entry.notes || '').slice(0, 2000)
  ]);
  return activityId;
}

function getLeadActivity_(leadId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEAD_ACTIVITY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const wanted = String(leadId || '').trim();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues()
    .filter(row => !wanted || String(row[2] || '').trim() === wanted)
    .map(row => ({
      timestamp: clientSafeDate_(row[0]), activityId: String(row[1] || ''), leadId: String(row[2] || ''),
      clientId: String(row[3] || ''), activityType: String(row[4] || ''), direction: String(row[5] || ''),
      contactPerson: String(row[6] || ''), outcome: String(row[7] || ''), stageBefore: String(row[8] || ''),
      stageAfter: String(row[9] || ''), nextFollowupSet: erpDate_(row[10]), loggedBy: String(row[11] || ''),
      notes: String(row[12] || '')
    }))
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
}

function touchClientLastContact_(clientId) {
  const client = findClientRow_(clientId);
  if (!client) return;
  const sheet = clientsSheet_();
  sheet.getRange(client.row, CLIENT_COL.lastContact).setValue(erpToday_());
  sheet.getRange(client.row, CLIENT_COL.lastUpdated).setValue(new Date());
}

/*************** LEADS <-> PROJECTS ***************/

/**
 * Opens the delivery record for a won lead, carrying the client, the contract
 * value and the opportunity name across so nothing is retyped.
 */
function erpCreateProjectFromLead(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin, 'leads');
    requireAdminPageWrite_(user, 'projects');
    const lead = findLeadRow_(p.leadId);
    if (!lead) return {success: false, message: 'Lead not found.'};
    if (!lead.isWon) {
      return {success: false, message: 'Move the lead to Won (Contracted) before opening its project.'};
    }
    const existing = getERPProjects_().filter(project => project.leadId === lead.leadId);
    if (existing.length && !p.allowDuplicate) {
      return {success: false, message: 'Project ' + existing[0].projectId + ' is already linked to this lead. ' +
        'Add another only if this lead genuinely produced a second project.'};
    }

    const contractValue = p.contractValue !== undefined && String(p.contractValue).trim() !== ''
      ? erpNonNegativeAmount_(p.contractValue, 'Contract Value')
      : (lead.contractValue || lead.estimatedValue);
    const created = erpAddProject({
      adminPin: p.adminPin,
      projectName: String(p.projectName || lead.opportunity),
      clientId: lead.clientId,
      client: lead.clientName,
      leadId: lead.leadId,
      manager: String(p.manager || lead.ownerName || ''),
      status: 'Planning',
      contractValue: contractValue,
      budget: erpNonNegativeAmount_(p.budget, 'Expenditure Budget'),
      startDate: erpDateInput_(p.startDate, 'Start Date'),
      endDate: erpDateInput_(p.endDate, 'End Date'),
      location: String(p.location || 'Abuja'),
      notes: 'Created from lead ' + lead.leadId + ' (' + lead.clientName + ').'
    });
    if (!created.success) return created;

    // The won deal's contract value is the project's revenue, so keep them equal.
    if (contractValue && contractValue !== lead.contractValue) {
      leadsSheet_().getRange(lead.row, LEAD_COL.contractValue).setValue(contractValue);
    }
    logLeadActivity_({
      leadId: lead.leadId, clientId: lead.clientId, activityType: 'Contract Sent', direction: 'Internal',
      contactPerson: lead.contactPerson, outcome: 'Project ' + created.projectId + ' opened',
      stageBefore: lead.stage, stageAfter: lead.stage, nextFollowupSet: '',
      loggedBy: user.name, notes: 'Delivery record created from this lead.'
    });
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'CREATE', 'Project', created.projectId,
      'Created from lead ' + lead.leadId);
    return {success: true, message: 'Project ' + created.projectId + ' created from ' + lead.leadId + '.',
      projectId: created.projectId};
  } catch (err) {
    return {success: false, message: err.message};
  }
}

/** Attaches an existing project to a lead without re-entering it. */
function erpLinkProjectToLead(p) {
  try {
    const user = erpRequireAdmin_(p.adminPin, 'leads');
    requireAdminPageWrite_(user, 'projects');
    const lead = findLeadRow_(p.leadId);
    if (!lead) return {success: false, message: 'Lead not found.'};
    const project = findProjectRow_(p.projectId);
    if (!project) return {success: false, message: 'Project not found.'};
    if (project.leadId && project.leadId !== lead.leadId) {
      return {success: false, message: 'Project ' + project.projectId + ' is already linked to lead ' + project.leadId + '.'};
    }
    const sheet = projectsSheet_();
    sheet.getRange(project.row, PRJ_COL.leadId).setValue(lead.leadId);
    if (!project.clientId && lead.clientId) {
      sheet.getRange(project.row, PRJ_COL.clientId).setValue(lead.clientId);
      sheet.getRange(project.row, PRJ_COL.client).setValue(lead.clientName);
    }
    sheet.getRange(project.row, PRJ_COL.lastUpdated).setValue(new Date());
    recordActivity_('Admin', user.name, 'ADMIN-ROW-' + user.row, 'UPDATE', 'Project', project.projectId,
      'Linked to lead ' + lead.leadId);
    return {success: true, message: 'Project ' + project.projectId + ' linked to lead ' + lead.leadId + '.'};
  } catch (err) {
    return {success: false, message: err.message};
  }
}

/**
 * The Projects-to-Leads direction: every project belonging to a lead, or to the
 * lead's client, with its revenue, spend and balance rolled up. This is what
 * makes the Leads page able to answer whether a client is profitable, not just
 * what stage the conversation reached.
 */
function getLeadProjectRollup_(lead, projects) {
  const all = projects || getERPProjects_();
  const direct = all.filter(project => project.leadId && project.leadId === lead.leadId);
  const clientWide = lead.clientId
    ? all.filter(project => project.clientId === lead.clientId)
    : all.filter(project => project.client && project.client.toLowerCase() === lead.clientName.toLowerCase());
  const totals = function (rows) {
    const contractValue = rows.reduce((sum, project) => sum + project.contractValue, 0);
    const actualSpend = rows.reduce((sum, project) => sum + project.actualSpend, 0);
    return {
      count: rows.length,
      contractValue: contractValue,
      budget: rows.reduce((sum, project) => sum + project.budget, 0),
      actualSpend: actualSpend,
      balance: contractValue - actualSpend,
      marginPct: contractValue > 0 ? Math.round((contractValue - actualSpend) / contractValue * 1000) / 10 : 0,
      receivable: rows.reduce((sum, project) => sum + project.receivable, 0)
    };
  };
  return {
    linkedProjects: direct.map(project => ({
      projectId: project.projectId, projectName: project.projectName, status: project.status,
      contractValue: project.contractValue, budget: project.budget, actualSpend: project.actualSpend,
      balance: project.balance, marginPct: project.marginPct, receivable: project.receivable
    })),
    linkedTotals: totals(direct),
    clientTotals: totals(clientWide),
    clientProjectCount: clientWide.length
  };
}

/** Lead detail, fetched on demand so activity history never ships wholesale. */
function erpGetLeadDetail(p) {
  try {
    const user = getAdminUser_(p && p.adminPin);
    if (!user.valid) return {success: false, message: 'Invalid admin PIN.'};
    requireAdminPageRead_(user, 'leads');
    const lead = findLeadRow_(p.leadId);
    if (!lead) return {success: false, message: 'Lead not found.'};
    const canReadProjects = adminCanReadPage_(user, 'projects');
    const rollup = canReadProjects ? getLeadProjectRollup_(lead, getERPProjects_()) : null;
    return {
      success: true,
      lead: lead,
      client: findClientRow_(lead.clientId),
      activity: getLeadActivity_(lead.leadId),
      rollup: rollup,
      unlinkedProjects: canReadProjects && adminCanWritePage_(user, 'projects')
        ? getERPProjects_().filter(project => !project.leadId)
            .map(project => ({projectId: project.projectId, projectName: project.projectName}))
        : [],
      canWrite: adminCanWritePage_(user, 'leads'),
      stages: leadStages_(),
      activityTypes: leadActivityTypes_()
    };
  } catch (err) {
    return {success: false, message: err.message};
  }
}

/*************** PIPELINE SUMMARY ***************/

function summarizeLeadPipeline_(leads) {
  const open = leads.filter(lead => lead.isOpen);
  const byStage = {};
  leadStages_().forEach(function (stage) { byStage[stage] = {stage: stage, count: 0, value: 0}; });
  leads.forEach(function (lead) {
    if (!byStage[lead.stage]) byStage[lead.stage] = {stage: lead.stage, count: 0, value: 0};
    byStage[lead.stage].count++;
    byStage[lead.stage].value += lead.estimatedValue;
  });
  return {
    total: leads.length,
    open: open.length,
    won: leads.filter(lead => lead.isWon).length,
    lost: leads.filter(lead => lead.stage === 'Lost').length,
    pipelineValue: open.reduce((sum, lead) => sum + lead.estimatedValue, 0),
    weightedValue: Math.round(open.reduce((sum, lead) => sum + lead.weightedValue, 0) * 100) / 100,
    contractedValue: leads.filter(lead => lead.isWon).reduce((sum, lead) => sum + (lead.contractValue || lead.estimatedValue), 0),
    overdueFollowups: leads.filter(lead => lead.followupOverdue).length,
    dueSoonFollowups: leads.filter(lead => lead.followupDueSoon).length,
    byStage: leadStages_().map(stage => byStage[stage]).filter(entry => entry.count > 0)
  };
}
