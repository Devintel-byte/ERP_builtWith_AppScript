/**
 * Shared schema definitions for the Projects / Leads / Project Cost pipeline.
 * Setup, migration, and the feature modules all read their headers from here so
 * the sheet layout is declared exactly once.
 */

function projectsHeaders_() {
  return [
    'Project ID','Project Name','Client ID','Client','Lead ID','Project Manager','Start Date','End Date','Status',
    'Contract Value','Expenditure Budget','Planned Cost','Actual Spend','Budget Variance','Project Balance','Margin %',
    'Invoiced Amount','Amount Received','Outstanding Receivable','Location','Notes','Last Rollup At','Last Updated'
  ];
}

function clientsHeaders_() {
  return [
    'Client ID','Client / Company','Contact Person','Role / Title','Email','Phone','Sector','Location','Website',
    'Relationship Status','Account Owner','Account Owner Name','First Engaged','Last Contact Date','Notes','Last Updated'
  ];
}

function leadsHeaders_() {
  return [
    'Lead ID','Client ID','Client / Company','Opportunity','Contact Person','Email','Phone','Source','Stage',
    'Probability %','Estimated Value','Contract Value','Weighted Value','Expected Close Date','Last Contact Date',
    'Next Follow-up','Reminder Cadence','Follow-up Owner','Follow-up Owner Name','Last Reminder Sent','Lost Reason',
    'Notes','Created At','Last Updated'
  ];
}

function projectCostHeaders_() {
  return [
    'Line ID','Project ID','Cost Date','Category','Sub-Category','Description','Vendor / Payee','Quantity','Unit',
    'Unit Cost','Planned Amount','Actual Amount','Source Module','Source Ref ID','Payment Status','Paid Date',
    'Payment Reference','Recorded By','Notes','Last Updated'
  ];
}

function leadActivityHeaders_() {
  return [
    'Timestamp','Activity ID','Lead ID','Client ID','Activity Type','Direction','Contact Person','Outcome',
    'Stage Before','Stage After','Next Follow-up Set','Logged By','Notes'
  ];
}

function leadReminderLogHeaders_() {
  return ['Timestamp','Reminder Key','Lead ID','Reminder Type','Due Date','Recipients','Subject','Status','Notes'];
}

/*************** CONTROLLED VOCABULARY ***************/

function projectStatuses_() {
  return ['Planning','Active','On Hold','Completed','Cancelled'];
}

function leadStages_() {
  return ['New / Identified','Qualified','Proposal Sent','Negotiation','Contract Issued','Won (Contracted)','Delivered','Lost','Dormant'];
}

/** Stages that no longer attract follow-up reminders. */
function leadClosedStages_() {
  return ['Delivered','Lost','Dormant'];
}

/** Stages that represent a committed, contracted deal. */
function leadWonStages_() {
  return ['Won (Contracted)','Delivered'];
}

function leadStageProbability_(stage) {
  const map = {
    'New / Identified':10,'Qualified':25,'Proposal Sent':40,'Negotiation':60,
    'Contract Issued':80,'Won (Contracted)':100,'Delivered':100,'Lost':0,'Dormant':0
  };
  const key = String(stage || '').trim();
  return Object.prototype.hasOwnProperty.call(map,key) ? map[key] : 10;
}

function leadSources_() {
  return ['Referral','Inbound Enquiry','Tender / RFP','Existing Client','Event / Exhibition','Cold Outreach','Partner','Other'];
}

function leadReminderCadences_() {
  return ['Weekly','Fortnightly','Monthly','Off'];
}

function clientRelationshipStatuses_() {
  return ['Prospect','Active','Dormant','Blacklisted'];
}

function leadActivityTypes_() {
  return ['Call','Email','Meeting','Site Visit','Proposal Sent','Quote Revised','Contract Sent','Follow-up Logged','Other'];
}

/**
 * Project cost categories. Expense Claims reuse this list so a claim posted
 * against a project lands in a category the project report already understands.
 */
function projectCostCategories_() {
  return [
    'Logistics / Transport','Feeding','Accommodation','Crew & Labour','Per Diem','Equipment Hire',
    'Materials & Consumables','Fabrication / Build','Freight & Clearing','Permits & Levies',
    'Software / Licensing','Contingency','Miscellaneous'
  ];
}

function projectCostPaymentStatuses_() {
  return ['Planned','Committed','Paid','Cancelled'];
}

function projectCostSourceModules_() {
  return ['Manual','Requisition','Expense Claim','Purchase Order'];
}

/** Maps a free-text category onto a known project cost category. */
function normalizeProjectCostCategory_(value) {
  const supplied = String(value || '').trim();
  if (!supplied) return 'Miscellaneous';
  const categories = projectCostCategories_();
  const exact = categories.filter(category => category.toLowerCase() === supplied.toLowerCase());
  if (exact.length) return exact[0];
  const loose = supplied.toLowerCase().replace(/[^a-z]/g,'');
  const fuzzy = categories.filter(category => category.toLowerCase().replace(/[^a-z]/g,'') === loose);
  return fuzzy.length ? fuzzy[0] : 'Miscellaneous';
}

/*************** INVENTORY VOCABULARY ***************/

/**
 * Inventory categories. Kept broad enough that an item always has an obvious
 * home, so the dropdown never pushes people into "Other" by default.
 */
function inventoryCategories_() {
  return [
    'Cable','Connector & Adapter','Display & LED','Audio','Lighting','Power & Distribution',
    'Rigging & Truss','Computing & Storage','Networking','Camera & Optics','Stationery',
    'Cleaning & Consumables','Safety & PPE','Tools & Hardware','Packaging','Other'
  ];
}

/**
 * Units of measure. A unit is the NOUN a quantity is counted in (12 "pcs"),
 * which is a different thing from Opening Stock, the starting NUMBER on hand.
 */
function inventoryUnits_() {
  return ['pcs','set','pair','box','carton','pack','roll','metre','foot','litre','kg','ream'];
}

function inventoryTransactionTypes_() {
  return ['Stock In','Stock Out','Adjustment','Return'];
}

function normalizeInventoryCategory_(value) {
  return matchFromList_(value, inventoryCategories_(), 'Other');
}

function normalizeInventoryUnit_(value) {
  return matchFromList_(value, inventoryUnits_(), 'pcs');
}

/** Case- and punctuation-insensitive match against a controlled list. */
function matchFromList_(value, list, fallback) {
  const supplied = String(value == null ? '' : value).trim();
  if (!supplied) return fallback;
  const exact = list.filter(item => item.toLowerCase() === supplied.toLowerCase());
  if (exact.length) return exact[0];
  const loose = supplied.toLowerCase().replace(/[^a-z0-9]/g, '');
  const fuzzy = list.filter(item => item.toLowerCase().replace(/[^a-z0-9]/g, '') === loose);
  return fuzzy.length ? fuzzy[0] : fallback;
}
