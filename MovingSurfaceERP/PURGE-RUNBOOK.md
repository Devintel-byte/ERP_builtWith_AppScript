# Operational data purge runbook

This procedure clears transactional history from the bound spreadsheet while
keeping the application configuration and staff identity records needed to
continue using the system. It is destructive and cannot be undone from the
spreadsheet after the version history or backups expire.

## Before you start

1. Get approval from the workbook owner and confirm the records have reached
   the organization's approved retention/deletion date.
2. Pause attendance and ERP use during the purge. In Apps Script **Triggers**,
   temporarily remove the one-minute `processPendingMailQueue` trigger. Remove
   only the trigger; keep the `processPendingMailQueue` function in
   `Notifications.gs`. After verification, run the public
   `installMailQueueWorker` function from the Apps Script editor to recreate
   the trigger if it is missing. Ensure no administrator or scheduled job
   writes records while sheets are being cleared.
3. Make a restricted, dated copy of the workbook. If records must be retained
   for payroll, legal, finance, or security review, export them to an approved
   restricted archive before deletion. In particular, do not destroy the only
   copy of access, PIN/device reset, recovery, or administrative audit history
   unless the retention owner has approved it.
4. Confirm the target workbook and its spreadsheet ID with a second
   administrator. Do not run a purge against a test or production workbook
   without explicitly identifying which one is in scope.

## Keep these sheets

Keep all rows, including headers, on the following sheets:

| Sheet | Why it is retained |
| --- | --- |
| `Settings` | Attendance, timezone, overtime, GPS, and PIN/device policy |
| `Staff List` | Staff profiles, PINs, device-token hashes, and company email addresses |
| `Admin Users` | Admin login PINs and roles |
| `Office Locations` | Approved attendance locations |
| `Holidays` | Calendar reference data |
| `Requisition Settings` | Requisition routing, notification, and currency settings |
| `ERP Settings` | ERP configuration |
| `Assets` | Asset register/master records |
| `Inventory Items` | Inventory balances and item master records |
| `Projects` | Project master records |
| `CRM Clients` | Client master records |
| `Vendors` | Vendor master records |
| `ERP Dashboard` | Dashboard formulas and labels |

The protected sheets contain sensitive information. Restrict workbook access
and do not include PINs, device hashes, bank details, or staff contact details
in purge reports or email notifications.

## Clear these operational records

For every listed sheet, retain row 1 (the header) and clear the contents of
rows 2 through the last populated row. Use **Edit → Delete values** or
**Clear contents** on that data range; do not delete the sheet, its header,
formatting, validation, or formulas outside the data range.

### Attendance, access, and audit history

- `Attendance Log`
- `Admin Overrides`
- `Admin Reset Log`
- `Staff Recovery Log`
- `Staff Movement Log`
- `Activity Log`
- `Accountant View` — derived payroll snapshot; rebuild it after the purge

These include security and payroll evidence, not just convenience logs. Archive
them first when required by the retention policy.

### Requisitions, notices, and other transactions

- `Item Requisitions`
- `Requisition Audit Log`
- `Mail Queue` — remove queued and failed notices as well as sent rows, so an
  old queued email is not sent after the purge
- `Expense Claims`
- `Leave Requests`
- `Asset Movements`
- `Inventory Transactions`
- `Purchase Orders`
- `ERP Notifications`

If modules have added sheets not listed here, classify each before clearing it:
transaction/history sheets are purge candidates; configuration and master
data are not. Never assume a newly added sheet is safe to erase.

## Google Sheets procedure

1. Open the confirmed workbook and inspect **all sheet tabs** against the keep
   and clear lists above.
2. For each clear sheet that has data rows, select from `A2` through the last
   used column and last used row, then clear **contents only**. Preserve row 1.
   Avoid deleting rows or sheets because that can alter formulas, validation,
   references, or setup assumptions.
3. For `Mail Queue`, ensure no `PENDING` or `PROCESSING` row remains. For mail that
   has already been sent, clearing the queue only removes this workbook's
   record; it cannot recall a message already delivered.
4. Rebuild the derived payroll snapshot by running the spreadsheet menu item
   **Sync Accountant View Rows** (or the equivalent server-side
   `syncAccountantViewRows` function) after the attendance rows are cleared.
5. Recalculate and inspect the `ERP Dashboard`; confirm that formula cells
   remain intact and counts reflect the now-empty transaction sheets.
6. Have a second administrator verify that the retained settings, staff,
   login, location, and master data remain present and that the cleared sheets
   retain their headers.
7. Resume normal operations and record the approver, workbook identifier,
   date/time, sheet names cleared, archive location, and verifier in the
   organization's controlled change record.

## Important limits and follow-up

- Clearing sheet rows does not delete files linked from requisitions, expense
  claims, or other records in Google Drive. Review those attachments under
  their own ownership and retention rules; remove only files exclusively in
  scope and approved for deletion.
- Clearing `Staff Recovery Log` invalidates outstanding recovery tokens;
  clearing `Mail Queue` can cancel pending notices. Tell affected users to
  request a new link or notification if necessary.
- `installMailQueueWorker` is the editor-runnable installer; it creates a
  one-minute trigger that calls `processPendingMailQueue`. If the worker
  function itself was deleted from the project, restore it from
  `Notifications.gs` before running the installer. Reinstalling the trigger
  does not restore deleted source code.
- Do not run `setupMovingSurfaceERPUpgrade` merely to rebuild the dashboard
  after a purge. Setup seeds example asset, inventory, project, and CRM rows
  when those master sheets are empty. Use the specific accountant sync action
  instead, and review any other setup function before running it.
- This spreadsheet procedure does not erase Google Drive version history,
  downloaded exports, sent mail, Apps Script execution logs, or external
  backups. Handle those separately under the organization's retention and
  deletion policy.
