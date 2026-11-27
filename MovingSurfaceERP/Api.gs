/**
 * The JSON API used by the pages hosted on our own domain.
 *
 * The Google-hosted pages call the server through google.script.run. Pages served
 * from another domain cannot, so they POST {fn, args} here and receive
 * {ok:true, value} or {ok:false, error}, which the browser bridge turns back into
 * the same success and failure handlers.
 *
 * ONLY the functions listed in apiFunctions_() can be run this way. It is an
 * explicit list on purpose: a request names a key in this map and nothing else, so
 * setup, migration, maintenance and test functions cannot be reached from the web.
 * When a page starts calling a new server function, add it here.
 *
 * Every listed function does its own sign-in check, exactly as it does when called
 * through google.script.run. This file adds no authority of its own.
 */

const API_MAX_BODY_CHARS = 2000000;
const API_MAX_ARGS = 5;

function apiFunctions_() {
  return {
    // Public: no sign-in, or the function checks its own one-time token or PIN.
    completeStaffAccountRecovery: completeStaffAccountRecovery,
    getAdminLoginOptions: getAdminLoginOptions,
    getAttendanceUrl: getAttendanceUrl,
    getPublicAttendanceSettings: getPublicAttendanceSettings,
    getPublicPortalLinks: getPublicPortalLinks,
    getPublicStaffOptions: getPublicStaffOptions,
    getRequisitionReferenceData: getRequisitionReferenceData,
    getStaffRecoveryUrl: getStaffRecoveryUrl,
    recordAttendance: recordAttendance,
    requestStaffAccountRecovery: requestStaffAccountRecovery,
    // Staff portal and requisitions.
    staffPortalGetData: staffPortalGetData,
    staffSubmitExpenseClaim: staffSubmitExpenseClaim,
    staffSubmitLeaveRequest: staffSubmitLeaveRequest,
    submitItemRequisition: submitItemRequisition,
    // Admin ERP.
    adminAddNewStaff: adminAddNewStaff,
    adminCreateStaff: adminCreateStaff,
    adminDecisionRequisition: adminDecisionRequisition,
    adminDeleteStaff: adminDeleteStaff,
    adminGetActivityPage: adminGetActivityPage,
    adminGetAnalytics: adminGetAnalytics,
    adminGetAttendancePage: adminGetAttendancePage,
    adminGetPortalData: adminGetPortalData,
    adminGetStaffReport: adminGetStaffReport,
    adminGetUserManagementData: adminGetUserManagementData,
    adminMarkRequisitionPaid: adminMarkRequisitionPaid,
    adminResetStaff: adminResetStaff,
    adminSaveUserPrivileges: adminSaveUserPrivileges,
    adminSetStaffStatus: adminSetStaffStatus,
    adminUpdateStaff: adminUpdateStaff,
    erpAddAsset: erpAddAsset,
    erpAddClient: erpAddClient,
    erpAddInventoryItem: erpAddInventoryItem,
    erpAddLead: erpAddLead,
    erpAddProject: erpAddProject,
    erpAddProjectCostLine: erpAddProjectCostLine,
    erpAssignAsset: erpAssignAsset,
    erpCancelFieldDuty: erpCancelFieldDuty,
    erpCancelProjectCostLine: erpCancelProjectCostLine,
    erpChangeFieldDutyEnd: erpChangeFieldDutyEnd,
    erpCreateFieldDuty: erpCreateFieldDuty,
    erpCreateProjectFromLead: erpCreateProjectFromLead,
    erpDecideFieldDuty: erpDecideFieldDuty,
    erpDecisionExpenseClaim: erpDecisionExpenseClaim,
    erpDecisionLeaveRequest: erpDecisionLeaveRequest,
    erpDeleteInventoryItem: erpDeleteInventoryItem,
    erpGetInventoryDeletePreview: erpGetInventoryDeletePreview,
    erpGetLeadDetail: erpGetLeadDetail,
    erpGetProjectDetail: erpGetProjectDetail,
    erpImportInventoryItems: erpImportInventoryItems,
    erpLinkProjectToLead: erpLinkProjectToLead,
    erpLogLeadActivity: erpLogLeadActivity,
    erpMarkExpensePaid: erpMarkExpensePaid,
    erpMarkProjectCostPaid: erpMarkProjectCostPaid,
    erpPostInventoryTransaction: erpPostInventoryTransaction,
    erpReturnAsset: erpReturnAsset,
    erpSubmitExpenseClaim: erpSubmitExpenseClaim,
    erpSubmitLeaveRequest: erpSubmitLeaveRequest,
    erpUpdateInventoryItem: erpUpdateInventoryItem,
    erpUpdateLeadStage: erpUpdateLeadStage,
    erpUpdateProject: erpUpdateProject
  };
}

function doPost(e) {
  return ContentService
    .createTextOutput(JSON.stringify(handleApiRequest_(e && e.postData ? e.postData.contents : '')))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Runs one API request body and returns the reply object. Never throws. */
function handleApiRequest_(body) {
  try {
    const text = String(body || '');
    if (!text) return {ok:false, error:'Empty request.'};
    if (text.length > API_MAX_BODY_CHARS) return {ok:false, error:'Request is too large.'};
    let request;
    try { request = JSON.parse(text); } catch (parseError) { return {ok:false, error:'Request is not valid JSON.'}; }
    const name = request && typeof request.fn === 'string' ? request.fn : '';
    const functions = apiFunctions_();
    if (!Object.prototype.hasOwnProperty.call(functions, name)) return {ok:false, error:'Unknown request.'};
    const args = request.args === undefined ? [] : request.args;
    if (!Array.isArray(args) || args.length > API_MAX_ARGS) return {ok:false, error:'Request arguments are not valid.'};
    // A function that returns {success:false} is still a completed call, as it is
    // for google.script.run: only a thrown error goes to the failure handler.
    return {ok:true, value:functions[name].apply(null, args)};
  } catch (err) {
    return {ok:false, error:String((err && err.message) || err)};
  }
}
