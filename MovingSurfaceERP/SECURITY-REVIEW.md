# Security review

**Reviewed:** 2026-09-29  
**Scope:** Moving Surface ERP Apps Script application, with focus on authentication, authorization, access control, and sensitive workflow changes.

## Findings

| # | Severity | File and lines | Vulnerability | Confidence |
|---|---|---|---|---:|
| 1 | HIGH | [StaffPortalApi.gs](./StaffPortalApi.gs#L45), [RequisitionRepository.gs](./RequisitionRepository.gs#L62), [Router.gs](./Router.gs#L9), [Notifications.gs](./Notifications.gs#L41) | Staff portal responses expose the requisition approval token, and the public email-approval route accepts that token as authorization. A requester can use their own disclosed token to approve or reject their pending request. | 10/10 |
| 2 | HIGH | [Attendance.gs](./Attendance.gs#L21), [AdminStaffApi.gs](./AdminStaffApi.gs#L1), [StaffPortalApi.gs](./StaffPortalApi.gs#L14), [Setup.gs](./Setup.gs#L44) | First-time PIN/device enrollment can be claimed by whoever initializes an active staff profile with no credentials. Staff options are public, and self-service PIN setup is enabled by default. | 9/10 |
| 3 | HIGH | [ERPHelpers.gs](./ERPHelpers.gs#L9), [AdminApi.gs](./AdminApi.gs#L159), [Expenses.gs](./Expenses.gs#L27) | Some expense approval and payment endpoints authenticate only that an admin PIN is valid; they do not enforce role capabilities or valid prior workflow states in the server endpoint. | 9/10 |

## Details and recommended remediation

### 1. Requesters can authorize their own requisitions

The staff-facing response includes requisition records returned by `getAllRequisitions_()`, including the stored approval token. The public route in `Router.gs` dispatches `approveReq` requests to the email-approval handler. That handler validates the token and pending status but does not independently establish that the caller is an authorized approver.

**Impact:** A staff member can retrieve the token associated with their request and submit an approval or rejection themselves. An approval can trigger Accounts notifications and may enable downstream disbursement.

**Recommended remediation:**

- Exclude approval tokens from every staff-facing response.
- Treat email links as bearer credentials: do not reveal them to requesters, and use a secure, single-use approval mechanism.
- Require approver authentication and verify the approver’s authorization on the server for every decision.
- Preserve an audit trail for rejected, expired, replayed, and accepted approval attempts.

**Implemented mitigation:** Staff and admin portal response DTOs now omit the internal approval token. The token remains available to server-side email generation and decision processing. The existing email link is still a bearer credential, so this closes the portal-disclosure path but does not independently verify the approver’s identity if the email or link is compromised. The Admin Users records currently lack a stable staff-ID mapping needed to reliably prohibit a linked admin from approving their own request.

### 2. Uninitialized staff profiles can be claimed during first use

New staff profiles can have blank PIN and device-token fields. With self-service PIN setup enabled, the public attendance endpoint permits a caller to choose an initial PIN and register the first device. Public staff options reveal active staff IDs and names. The location coordinates used by the endpoint are supplied by the same client and do not independently prove identity.

**Impact:** An attacker may initialize an unclaimed profile before its intended owner and gain access to that staff member’s attendance or portal account.

**Recommended remediation:**

- Require an out-of-band, one-time onboarding credential or an administrator-issued initial PIN before allowing first use.
- Bind enrollment to a verified staff channel and expire/revoke onboarding credentials after use.
- Do not treat client-supplied GPS coordinates as identity proof.
- Consider limiting the information exposed by public staff-option lists.

### 3. Expense mutations do not enforce authorization and workflow state

The portal calculates role-based capabilities, but the expense approval and payment endpoints accept any valid active admin PIN. The payment endpoint also does not require the claim to be in an approvable state before marking it paid.

**Impact:** A user without the intended approval or disbursement role can call the server functions directly, bypassing UI visibility and changing financial records, including paying unapproved claims.

**Recommended remediation:**

- Enforce role capabilities inside each server-side mutation function; UI checks are not authorization.
- Validate allowed source and destination states for every workflow transition.
- Require approval before payment and prevent duplicate or replayed payment actions.
- Record the actor, prior state, new state, and payment reference in the audit log.

## Review limitations

This is a static code review, not a live penetration test. Google Apps Script deployment settings, spreadsheet sharing, trigger ownership, and Google account configuration were not independently inspected. Review and update those settings during deployment.
