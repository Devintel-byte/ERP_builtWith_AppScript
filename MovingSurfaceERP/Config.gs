/**
 * MOVING SURFACE HR ATTENDANCE + ITEM REQUISITION SYSTEM
 * Fresh-start single-file Google Apps Script.
 *
 * ROUTES AFTER DEPLOYMENT
 * Attendance:       Web App URL
 * Staff Requisition Web App URL + ?req=1
 * Admin/Accounts:  Web App URL + ?admin=1
 *
 * FIRST RUN
 * 1. Paste this entire file into Apps Script.
 * 2. Run setupMovingSurfaceSystem once.
 * 3. Run testAdminLogin.
 * 4. Deploy / redeploy Web App.
 */

/*************** SHEET NAMES ***************/
const SETTINGS_SHEET = 'Settings';
const STAFF_SHEET = 'Staff List';
const LOG_SHEET = 'Attendance Log';
const OVERRIDES_SHEET = 'Admin Overrides';
const LOCATIONS_SHEET = 'Office Locations';
const HOLIDAYS_SHEET = 'Holidays';
const ADMIN_USERS_SHEET = 'Admin Users';
const ADMIN_RESET_LOG_SHEET = 'Admin Reset Log';
const STAFF_RECOVERY_SHEET = 'Staff Recovery Log';
const STAFF_MOVEMENT_LOG_SHEET = 'Staff Movement Log';
const ACCOUNTANT_VIEW_SHEET = 'Accountant View';
const REQ_SETTINGS_SHEET = 'Requisition Settings';
const REQ_SHEET = 'Item Requisitions';
const REQ_AUDIT_SHEET = 'Requisition Audit Log';
const MAIL_QUEUE_SHEET = 'Mail Queue';
const ACTIVITY_LOG_SHEET = 'Activity Log';

/*************** ERP MODULE SHEET NAMES ***************/
const CLIENTS_SHEET = 'Clients';
const LEADS_SHEET = 'Leads';
const LEAD_ACTIVITY_SHEET = 'Lead Activity Log';
const LEAD_REMINDER_LOG_SHEET = 'Lead Reminder Log';
const PROJECTS_SHEET = 'Projects';
const PROJECT_COSTS_SHEET = 'Project Cost Lines';
const EXPENSE_CLAIMS_SHEET = 'Expense Claims';
const INVENTORY_ITEMS_SHEET = 'Inventory Items';
const INVENTORY_TX_SHEET = 'Inventory Transactions';
const ASSETS_SHEET = 'Assets';
const ASSET_MOVEMENTS_SHEET = 'Asset Movements';
const PURCHASE_ORDERS_SHEET = 'Purchase Orders';
const LEAVE_REQUESTS_SHEET = 'Leave Requests';
const ERP_SETTINGS_SHEET = 'ERP Settings';

const SHEET_HEADERS = {
  [SETTINGS_SHEET]: ['Setting','Value','Notes'],
  [STAFF_SHEET]: [
    'Staff ID','Full Name','Department','Work Location','PIN','Device Token Hash','Company Email','Personal Email',
    'Require GPS','Require Device','Monthly Salary','Late Penalty','Leave Allowance','Phone','Notes','Date Joined',
    'Employment Type','Status','Exit Date','Designation','Manager','Bank Name','Account Number','Account Name',
    'Annual Leave Entitlement','Sick Leave Entitlement','Work Days Per Month','QR Code','Can Submit Requisition'
  ],
  [LOG_SHEET]: [
    'Date Key','Staff ID','Staff Name','Work Location','Action','Timestamp','Time','Latitude','Longitude',
    'Distance Metres','Matched Location','GPS Status','Source','Attendance Status','Today Penalty',
    'Month Late Penalty','Month Half-Day Deduction','Month Absence Deduction','Month Total Deduction',
    'Device Status','Device Token Fingerprint','User Agent','Remarks','Checkout Time','Checkout Type'
  ],
  [OVERRIDES_SHEET]: ['Date','Staff ID','Staff Name','Type','Reason','Status','Approved','Notes','Timestamp','Override ID'],
  [LOCATIONS_SHEET]: ['Location Name','Latitude','Longitude','Allowed Radius Metres','Active'],
  [HOLIDAYS_SHEET]: ['Date','Holiday Name','Active'],
  [ADMIN_USERS_SHEET]: ['Name','PIN','Active','Role','Can View','Can Override','Can Reset','Can Onboard','Staff ID','Email','Access Pages','Write Pages'],
  [ADMIN_RESET_LOG_SHEET]: ['Timestamp','Staff ID','Staff Name','Reset Type','PIN Cleared','Device Cleared','Actor','Reason','Notes'],
  [STAFF_RECOVERY_SHEET]: ['Token Hash','Staff ID','Requested At','Expires At','Used At','Status'],
  [STAFF_MOVEMENT_LOG_SHEET]: ['Timestamp','Staff ID','Staff Name','Movement Type','Old Status','New Status','Actor','Notes'],
  [ACCOUNTANT_VIEW_SHEET]: [
    'Month','Staff ID','Staff Name','Department','Designation','Employment Type','Status','Monthly Salary',
    'Required Days','Present Days','Leave Days','Absent Days','Late Count','Half-Day Count','Daily Rate',
    'Late Penalty Total','Half-Day Deduction','Absence Deduction','Total Deduction','Overtime Days','Overtime Pay','Net Pay'
  ],
  [REQ_SETTINGS_SHEET]: ['Setting','Value','Notes'],
  [REQ_SHEET]: [
    'Timestamp','Request ID','Staff ID','Staff Name','Department','Project / Event','Priority','Needed Before',
    'Purpose / Justification','Attachment Link','Status','MD Decision Date','MD Decision By','MD Comment',
    'Accounts Notification Queued Date','Disbursement Status','Disbursed Date','Disbursed By','Payment / Receipt Reference',
    'Staff Notification Queued Date','Approval Token','Last Updated','Total Estimated Cost','Items JSON','Items Summary'
  ],
  [REQ_AUDIT_SHEET]: ['Timestamp','Request ID','Action','Actor','Old Status','New Status','Notes'],
  [MAIL_QUEUE_SHEET]: ['Created At','Status','To','Cc','Subject','HTML Body','Attempts','Last Attempt','Last Error','Related ID'],
  [ACTIVITY_LOG_SHEET]: ['Timestamp','Actor Type','Actor Name','Actor / Staff ID','Action','Entity','Entity ID','Details']
};

/*************** MENU / SETUP ***************/
