# Moving Surface ERP Apps Script Project

This is the modular Apps Script replacement for the legacy `MSERP.txt` source. Keep the existing spreadsheet, add these files to the bound Apps Script project, and deploy the project as a web app.

## Web app routes

| Page | URL |
| --- | --- |
| Attendance (default) | Web app URL |
| Staff portal | Web app URL plus `?staff=1` |
| Requisition | Web app URL plus `?req=1` |
| Admin ERP | Web app URL plus `?admin=1` |
| Staff account recovery | Web app URL plus `?recover=1` |

If multiple route flags are supplied, the resolver uses the fixed precedence `admin`, `req`, `staff`, then attendance. Route links are composed centrally in `Router.gs`.
Incoming query parameter names are case-insensitive and rendered with lowercase names in the browser URL. Parameter values (including recovery tokens) retain their original case.

## Feature boundaries

- `Config.gs`, `Setup.gs`, `ERPSetup.gs`: schema, defaults, and first-run setup
- `Router.gs`, `Menu.gs`: route dispatch, canonical links, and spreadsheet menus
- `Attendance.gs`, `Auth.gs`, `Payroll*.gs`: attendance, PIN, and payroll rules
- `StaffRecovery.gs`, `RecoveryView.html`: one-time PIN and device recovery links
- `Requisitions.gs`, `RequisitionRepository.gs`, `Notifications.gs`: requisition workflow, persistence, and email delivery
- `ErpSchema.gs`: sheet headers, column vocabularies, and the project cost categories
- `ErpMigration.gs`: the one-time Projects / Leads schema migration
- `Projects.gs`, `ProjectCosts.gs`: project economics and the project cost spine
- `Leads.gs`, `LeadReminders.gs`: client master, opportunity pipeline, and follow-up reminders
- `AdminApi.gs`, `StaffPortalApi.gs`: browser-facing server APIs
- `SheetRepository.gs`, `ERPHelpers.gs`, `Utils.gs`: spreadsheet access and shared helpers
- `Index.html`, `Styles.html`, and the `*View.html` files: shared shell, design system, and feature screens

Keep sheet reads and business rules in server-side files. UI files should call server APIs with `google.script.run` and render returned values.

## First run and deployment

1. Create or open the bound Apps Script project for the workbook and add the modular files.
2. Run `setupMovingSurfaceERPUpgrade` once and approve the requested Google permissions. It creates/repairs core and ERP sheets (including the Activity Log), runs the Projects / Leads schema migration, normalizes existing PIN cells to text, syncs the accountant view, and installs the one-minute email queue trigger plus the daily lead-reminder and project-reconcile triggers.
3. Map each enabled admin account to a Staff ID in `Admin Users` (sample accounts are pre-linked) and review the role, page permissions, and notification email columns below. Admin Users holds no PIN: each person creates their own on the attendance page. See *One PIN for every page*.
4. Review recipient addresses in `Requisition Settings`.
5. Run `testRoutingAndPinRules`, `testAdminLogin`, `testERPPortalData`, `testRecoveryEmailMatchRules`, `testAttendanceTransitionRules`, `testRequisitionClientDto`, `testAdminAccessRules`, `testAdminCredentialResolution`, `testUserPermissionMatrixValidation`, `testRequisitionNotificationClassification`, `testProjectsAndLeadsPipeline`, `testInventoryModule`, `testClientDetailCapture`, `testDeviceMismatchPolicy`, `testUserAgentClassification`, `testSinglePinCredentials`, `testStartingPinRules`, `testLocationToleranceRules`, `testFieldDutyExcusedDays`, `testFieldDutyDateAndSettingParsing`, `testFieldDutyApprovalRules`, `testHostedApiRules`, and `testPinAttemptLimit`.
6. Deploy or update the web app and test each route using the exact query strings above.
7. Add and verify the staff company email addresses used for account recovery. Recovery links are one-use, expire after 30 minutes, and reset the PIN; the person chooses whether the browser they are on also becomes their attendance device. Run setup again after schema changes to create the Staff Recovery Log and add the overtime settings.
8. After upgrading the attendance schema, run `setupMovingSurfaceERPUpgrade` once. `Attendance Log` gains explicit `Checkout Time` and `Checkout Type` columns, and existing regular/overtime checkout rows are backfilled from their recorded event time and action.

## Admin roles and page permissions

Super Admin and Admin can open the Users page in the admin portal to create or modify an access account for a staff member. The page provides an explicit per-page Read/Write matrix and account capabilities. Write access always includes Read. Dashboard Read is required to sign in. There is no separate admin PIN: a person signs in to the ERP with the same PIN they use for attendance, and an account for someone who has not created one yet tells you to ask them to do so first. PINs are never returned to the browser or included in audit details.

Role labels use a standard dropdown (MD/CEO, SuperAdmin, Admin/HR, Accountant, Technical, Driver, Motion Artist, 3D Artist, Head of Operations, IoT Engineer, Lead Software Engineer, and Janitor). Choose **Add a role label…** to enter another label; it is added to future dropdowns after you save it on an account. Existing role labels in `Admin Users` are also retained in the list.

The `Admin Users` sheet adds four columns after its existing role and capability fields:

| Column | Purpose |
| --- | --- |
| Staff ID | Links the PIN row to an active staff account and appears in the admin login selector. |
| Email | Receives requisition alerts for the user's role. |
| Access Pages | Optional comma-separated page IDs that narrow a non-senior role's default read access. |
| Write Pages | Optional comma-separated page IDs that narrow a non-senior role's default write access. |

Run setup after upgrading, then fill Staff ID and Email for each enabled account. Login options are limited to active Admin Users rows whose Staff ID matches an active staff profile. Until mapped, an account cannot log in through the selector. The selected Staff ID is part of the credential, so two people choosing the same four digits cannot be confused. Do not store credentials in documentation or source code.

For existing accounts with blank page lists, these role defaults provide the initial access:

- Super Admin, Admin, and HR roles can read and write all pages.
- Technical can read and write all pages except Staff.
- Accountant can read all pages except Staff and write Requisitions and Expense Claims.
- Other roles are read-only by default and cannot access Staff.
- Inventory is view-only for Accountant and other non-technical roles; Technical and the senior admin roles can maintain it.

Page IDs are `dashboard`, `attendance`, `staff`, `requisitions`, `inventory`, `assets`, `expenses`, `leave`, `projects`, `leads`, `reports`, and `activity`. The `crm` page was renamed to `leads`; `migrateProjectsAndLeadsSchema` rewrites the stored token so existing page permissions survive the rename. When an administrator saves a person's permission matrix, Access Pages and Write Pages become that person's explicit policy, replacing the role defaults. This lets Admin/Super Admin grant or remove individual page access independently of the role label. Only Super Admin and Admin can open or call the user-management functions; HR and other roles cannot. Admin and Super Admin retain Users access and Dashboard read access so the user manager cannot accidentally lock itself out; other pages can be individually restricted. The Users page is not assignable as a page permission. `Can View=NO` denies portal access. Pages a user cannot view are removed from the navigation and search results, and server-side API checks remain authoritative.

The Users page supports account changes even for legacy Admin Users rows that have not yet been linked to a Staff ID; choose Modify access and link the existing row to the correct staff member. Duplicate staff-to-account mappings are rejected. Account changes are written to the Activity Log without recording PIN values.

Requisition alerts classify each requested line by exact case-insensitive match to an Inventory Items name or ID. Inventory requests notify Super Admin, Admin (including combined Admin/HR aliases), and Technical roles; purchase lines notify Super Admin, Admin (including combined Admin/HR aliases), and Accountant roles. Mixed requests notify the union. Populate the Email column to activate role recipients. The configured MD recipient continues to receive approval/rejection links; role notifications are informational and contain no privileged one-click decision link. A valid deployed admin route is required for its dashboard link.

Inventory is quantity-tracked stock that moves in and out through transactions (for example, paper, toner, or consumable supplies). Assets are durable, individually identifiable property tracked by serial/asset ID, assignment, condition, and return (for example, a specific laptop or vehicle). An inventory item can become an asset only when the business process issues and registers a particular unit as individually tracked property.

Deploy the web app to execute as the workbook owner and choose an access audience that includes staff who need to open attendance and recovery (for example, anyone with the link for a public staff-facing app, or signed-in users in the organization's Workspace domain). A deployment restricted to the owner or administrators redirects emailed staff recovery links to Google sign-in/access-denied pages. Redeploy after changing routes or views; new recovery emails use the current deployment URL.

## Overtime settings and device recovery

The `Settings` sheet contains `CLOSE_TIME` (default `17:00`), `OVERTIME_GRACE_MINUTES` (default `15`), and optional `OVERTIME_PAY_RATE` (amount per overtime day; `0` disables overtime pay). An overtime checkout is accepted starting at closing time plus the grace period in the configured timezone and is counted once per staff member per date. The dashboard and analytics read live overtime totals. Run `Sync Accountant View Rows` from the spreadsheet menu after attendance or pay-rate changes to refresh the payroll snapshot.

Staff account recovery is initiated from the Attendance, Staff Portal, or Requisition page. The request must match the active staff ID and company email on file. A mismatch is rejected with an instruction to correct the details, and no email is sent. For a valid request, the message is sent only to the company email stored on the staff profile, never to an arbitrary address supplied in the form. Only a hash of the random recovery token is stored in `Staff Recovery Log`; the link is emailed, expires after 30 minutes, and can be used once to set a new PIN, and, when the person leaves "Also make this browser my attendance device" ticked, to replace the registered device token. Device tokens are stored and audited as SHA-256 fingerprints, not bearer values. Running setup migrates existing device tokens and old attendance-log token values to fingerprints. Only attendance checks the registered device; the portal, requisitions and admin ERP do not. Recovery depends on Apps Script email delivery and a current web app deployment; staff who cannot access their company email must contact Admin.

With a web app deployed as **Execute as: Me** and **Who has access: Anyone**, the recovery link is a short-lived bearer link: opening it does not verify which Google account is signed in. If opening the emailed link works only when a particular Google account is active, verify that the email contains the current public `/exec` deployment URL (not a development `/dev` URL or an older restricted deployment) and test that exact URL in a private browser window. See `PURGE-RUNBOOK.md` for a controlled process to clear operational records while retaining core settings and staff data.

Staff profiles can be disabled/reactivated without losing history. Permanent deletion is blocked while attendance, requisition, expense, leave, or asset records reference a staff ID; disable the profile instead to preserve operational and audit history.

Setup is an administrative operation, not a request-time operation. Do not call setup/repair functions from web logins or normal form submissions.

## Projects, Leads, and project economics

### The money model

Three amounts per project are kept deliberately separate, because conflating them is what makes a margin figure meaningless:

| Column | Meaning | Who writes it |
| --- | --- | --- |
| `Contract Value` | Revenue. What the client pays. | Entered, or carried from the won lead |
| `Expenditure Budget` | The approved cost budget for delivering the project. | Entered |
| `Planned Cost` | The itemized plan: sum of planned cost lines. | Derived |
| `Actual Spend` | Cost actually incurred. | Derived |
| `Budget Variance` | `Expenditure Budget` less `Actual Spend`. | Derived |
| `Project Balance` | `Contract Value` less `Actual Spend`: what the project returns. | Derived |
| `Margin %` | `Project Balance` as a percentage of `Contract Value`. | Derived |
| `Invoiced Amount`, `Amount Received` | Cash in. | Entered |
| `Outstanding Receivable` | `Invoiced Amount` less `Amount Received`. | Derived |

**Derived columns are never typed.** They are recomputed from `Project Cost Lines` on every write, and again by the nightly `reconcileAllProjectRollups` trigger, which catches rows edited directly in the spreadsheet. `Planned Cost` is shown next to `Expenditure Budget` so an itemization that has drifted from the approved budget is visible rather than hidden.

The legacy `Budget` column migrated into `Expenditure Budget`, because that is what its name described. `Contract Value` starts at zero on migrated rows and must be set per project before margin reporting means anything.

### Project Cost Lines

`Project Cost Lines` is the single source of truth for project spend: logistics, feeding, accommodation, crew, per diem, equipment hire, materials, fabrication, freight, permits, licensing, contingency. A line carries a `Planned Amount` (the plan) and an `Actual Amount` (what moved). Lines with `Payment Status` of `Cancelled` stay on the sheet for audit and drop out of every roll-up, so spend history is never silently rewritten.

Costs reach a project two ways:

- **Typed in** on the project detail view, with an optional quantity and unit cost.
- **Posted automatically** when a requisition or expense claim for that project moves through its workflow. Approval posts the line as `Committed` (planned, not yet spent); payment updates the same line to `Paid` and books the actual amount.

Automatic postings are keyed on `(Source Module, Source Ref ID)` and upserted. Re-running a workflow step, retrying after a failure, or a double click can never double-count: one source document produces exactly one cost line that walks through its own lifecycle.

Inventory issues and asset movements are shown on the project detail view as movement, not cost: those transactions carry quantities, not prices, so inventing a cost from them would corrupt the spend figure.

### Project keys on the cost-bearing sheets

`Item Requisitions`, `Expense Claims`, `Inventory Transactions`, `Asset Movements`, and `Purchase Orders` each gained a `Project ID` column, appended at the end of the sheet so every existing column position stays where the working code expects it. The free-text `Project / Event` column is kept as the human label.

When a staff member submits a requisition or claim from the staff portal, the free-text project they type is matched against project names and IDs, so the existing habit still produces a usable key. Admins pick the project from a dropdown instead.

Expense claim categories now draw from the same list as project cost categories, so a claim posted against a project lands in a category the project report already understands.

### Leads

`CRM Clients` became two sheets:

- **`Clients`** is the master record for an organization: contact, email, phone, sector, relationship status, account owner. Projects and leads both point at a `Client ID`, so a client is spelled one way everywhere.
- **`Leads`** holds one row per *opportunity*. A repeat client can carry several deals at different stages at once, which a single row per client could not represent.

Stages are `New / Identified`, `Qualified`, `Proposal Sent`, `Negotiation`, `Contract Issued`, `Won (Contracted)`, `Delivered`, `Lost`, `Dormant`. Each carries a default probability, so `Estimated Value x Probability` gives a weighted pipeline forecast rather than a total that treats a cold enquiry like a signed contract.

`Lead Activity Log` records every contact: call, email, meeting, site visit, proposal sent, quote revised, contract sent. This is the per-client follow-up history, not just a status cell.

### How Projects and Leads connect

- **Leads to Projects:** a lead at `Won (Contracted)` offers **Create project from lead**, which carries the client, the opportunity name and the contract value into a new project. Nothing is retyped, so names cannot drift. A second project from the same lead is refused unless explicitly requested.
- **Projects to Leads:** the lead detail view reads `Projects.Lead ID` back and shows every linked project with its revenue, spend, balance, margin and receivable, plus a client-wide lifetime total. This is what lets the Leads page answer whether a client is actually profitable, not only what stage the conversation reached.
- **Link existing project** attaches a project that already exists to a lead without re-entry.

The link is held only on `Projects.Lead ID`. No project list is denormalized onto the lead row, so the two can never disagree.

### Follow-up reminders

One daily trigger, `processLeadFollowUpReminders`, runs at 07:00 in the configured timezone and drives three behaviours:

| Reminder | When | Who gets it |
| --- | --- | --- |
| **Weekly digest** | On `LEAD_WEEKLY_DIGEST_DAY` (default Monday) | **One email per follow-up owner**, listing their leads due inside the horizon plus everything overdue |
| **Final reminder** | Exactly `LEAD_FINAL_REMINDER_DAYS` (default 3) before the follow-up date | The follow-up owner, copied to the MD and senior admin roles |
| **Overdue** | A follow-up date passes | No separate email. It escalates inside the next digest and surfaces as a dashboard hotspot |

The digest is deliberately one email per owner rather than one per lead: a reminder system that sends twenty mails a week gets filtered to trash by week three. Overdue leads apply pressure through the digest and the dashboard instead of a fourth email.

Every send is keyed in `Lead Reminder Log` (for example `LEAD-0007|FINAL|2026-11-05`). A retry, a manual re-run from the menu, or a double trigger cannot deliver the same reminder twice. Reminders are queued into the existing `Mail Queue` and delivered by the one-minute worker, so the sweep never waits on mail delivery.

Recipients resolve from the lead's `Follow-up Owner` (a Staff ID) to that person's `Admin Users` email, falling back to their company email and then to the senior admin roles, so a lead is never left unwatched because its owner field is blank. A lead with no owner is reported in an "unassigned" digest that asks for one to be set.

Closing a lead (`Lost`, `Dormant`, `Delivered`) clears its follow-up date and switches its cadence to `Off`, so a dead deal cannot keep generating reminders. A won deal still gets reminders: delivery and payment both need chasing.

### Reminder settings

These live in `ERP Settings` and are upserted by setup without overwriting a value you have already configured:

| Setting | Default | Purpose |
| --- | --- | --- |
| `LEAD_REMINDERS_ENABLED` | `YES` | Set to `NO` to suspend all follow-up reminders |
| `LEAD_WEEKLY_DIGEST_DAY` | `Monday` | Weekday the per-owner digest is sent |
| `LEAD_REMINDER_HORIZON_DAYS` | `28` | How far ahead the digest looks |
| `LEAD_FINAL_REMINDER_DAYS` | `3` | Days before the due date that the final reminder fires |

### Running the migration

`setupMovingSurfaceERPUpgrade` runs it automatically, before headers are enforced. It is also available on its own from the spreadsheet menu as **Migrate Projects / Leads Schema**, and is safe to run repeatedly: every step checks its own completion first.

The migration order matters. Enforcing the new `Projects` header row first would relabel the legacy columns in place and strand the data sitting under them, so the migration runs before `setupERPModuleSheets_`.

What it does:

1. Renames `CRM Clients` to `Clients` and remaps its columns by header name. Its pipeline fields (status, opportunity value, next follow-up) become seed `Leads` rows, so the rename loses nothing.
2. Widens `Projects` to the revenue / budget / actual-spend model, mapping old values across **by header name, not by position**.
3. Appends a `Project ID` column to the five cost-bearing sheets.
4. Rewrites the stored `crm` page token to `leads` in `Admin Users`.
5. Replaces row-pinned formulas with computed values.

Sheets whose column order changes are copied to a hidden, timestamped `Backup <sheet> <stamp>` tab before they are rewritten. Review and delete those backups once you are satisfied with the result.

**Why step 4 is not cosmetic:** page IDs are stored as literal text in the `Admin Users` `Access Pages` and `Write Pages` cells, and `getAdminAccess_` discards any token not on its whitelist. Renaming the page without rewriting those cells would silently remove the module from every administrator holding an explicit page list, with no error shown.

### Row-pinned formulas

`Projects.Balance` and the `Inventory Items` stock and status columns were previously written as formulas carrying hardcoded row numbers at append time (`=H2-J2`, `=E2+F2-G2`). Sorting the sheet or deleting a row left those formulas pointing at the wrong record. Both now store computed values written by the script, and the migration converts any existing formula rows.

### Performance

The admin portal ships project and lead summaries on sign-in, but **not** cost lines or lead activity history: those are the two tables that grow fastest, and loading them wholesale would eventually exceed the Apps Script payload and execution limits. They load on demand through `erpGetProjectDetail` and `erpGetLeadDetail` when a row is opened.

Project roll-ups are materialized onto the `Projects` row when a cost line is written, not recomputed on every read, so opening the portal does not rescan the cost sheet.

## Inventory

### Unit and Opening Stock are different things

These two were easy to confuse because the old screen displayed neither:

| Field | What it is | Example |
| --- | --- | --- |
| **Unit** | The *noun* a quantity is counted in. A label, never a number. | `pcs`, `roll`, `metre`, `box` |
| **Opening Stock** | The starting *number* on hand when the item was first recorded. | `12` |

They are not duplicates, and neither is redundant: without a unit, a current stock of `10` does not say ten of what. The table now shows the full arithmetic so the relationship is visible rather than implied:

```
Opening + Stock In - Stock Out = Current Stock
```

Current Stock is always recomputed in script, never stored by hand and never a row-pinned formula. **Opening Stock seeds the balance once; day-to-day movement goes through stock transactions**, so the history stays intact. Editing Opening Stock is still allowed, because a typo at creation has no other remedy, and the current stock is recomputed when you do.

`Low Stock` shows when current stock is **at or below** the reorder level.

Transaction types:

| Type | Effect |
| --- | --- |
| `Stock In`, `Return` | Adds to Stock In |
| `Stock Out` | Adds to Stock Out, and is refused if it would drive the balance negative |
| `Adjustment` | Restates the opening balance so current stock matches a physical count, leaving the in/out history untouched |

### Category and Unit are chosen from a list

Both are dropdowns in the portal, and the same vocabularies are applied as sheet data validation so a row edited directly in the spreadsheet cannot drift. An unrecognized category becomes `Other`; an unrecognized unit becomes `pcs`. Matching ignores case, spacing and punctuation, so `display & led` and `Display&LED` both resolve to `Display & LED` — which is what makes a messy CSV import land in the right place.

### Editing and removing items

Every row carries **Edit** and **Delete**.

- **Edit** opens a modal prefilled with the item's current details, and shows how its current stock is derived so the opening-stock field is unambiguous.
- **Delete** always goes through a confirmation first. The confirmation names the item, its current stock, and how many stock transactions reference it. Removing an item **keeps its transaction history**: those rows record the item ID and name, so the movement log stays readable and auditable after the item itself is retired. An item with history needs that explicit confirmation before the server will act on it.

### Bulk import

The inventory page takes a **CSV** upload for bulk loading. Export your spreadsheet as CSV first — `.xlsx` is not parsed, deliberately, because doing so would mean either a third-party library loaded from a CDN or a round trip through Drive, and neither is worth it for a file format every spreadsheet tool can already export.

The file is parsed in the browser (quoted fields, embedded commas and newlines, escaped quotes, CRLF and a leading BOM are all handled; tab-separated files are detected automatically), previewed before anything is sent, and then validated again on the server — the browser is not a trust boundary.

| Column | Required | Notes |
| --- | --- | --- |
| `Item Name` | **Yes** | Must be unique. A name already in the register **updates** that item instead of duplicating it. |
| `Category` | No | From the category list; anything else becomes `Other`. |
| `Unit` | No | Unit of measure. Defaults to `pcs`. |
| `Opening Stock` | No | Whole number on hand now. Defaults to `0`. |
| `Reorder Level` | No | `Low Stock` shows at or below this. Defaults to `0`. |
| `Location` | No | Where the stock is held. |
| `Preferred Vendor` | No | Usual supplier. |
| `Notes` | No | Anything else worth recording. |

Headings are matched loosely (case, spacing and punctuation are ignored) and extra columns are ignored. A column the file leaves empty is **not** applied, so a partial CSV never silently blanks data you already have. Up to 500 rows per import.

Because an existing name updates rather than duplicates, **re-importing a corrected file is safe**: the same file imported repeatedly leaves one copy of each item with the same balance. Bad rows are reported individually with their line number and the rest of the file still imports. A **Download template** button produces a CSV with exactly these headings.

The column reference shown at the bottom of the inventory page is generated from `inventoryImportColumns_()`, the same list the importer reads, so the documentation cannot drift from the behaviour.

### A new client is captured once, from whichever form is open

Typing a company nobody has traded with before into **Add Project** or **Add
Lead** used to create a Clients row holding nothing but the name, and somebody
then had to open the Clients page and enter the contact all over again.

Both forms now reveal a **New client details** block as soon as a new company
name is typed, and the Clients master is built in the same save:

| Form | Fields that reach the client record |
| --- | --- |
| Add Lead | Contact person, Email and Phone (the lead's own fields), plus Role / title, Sector, Location and Website |
| Add Project | Contact person, Role / title, Email, Phone, Sector, Location and Website |

A client created this way starts as a **Prospect**, and the lead's follow-up
owner becomes its account owner.

Two rules keep the data honest:

- The fields travel to the server under `client`-prefixed names
  (`clientLocation`, `clientSector`, ...). A project's own Location is the
  venue and its Notes are about the job, so sharing the plain names would file
  an event venue as the company's registered address.
- **An existing client is never overwritten from these forms.** The server
  fills in only the fields the Clients record has left blank and leaves any
  curated value alone, then reports what it changed: *"Lead added: LEAD-0004.
  Client CL-0002 updated: email, location filled in."*

Picking an existing client from the dropdown instead hides the block, locks the
new-name box so the two cannot disagree, shows what is already on file, and
offers the known contact person, email and phone to the lead form — without
overwriting anything already typed there.

The client is resolved only after every other input has passed validation, so a
lead rejected for a missing follow-up date leaves no half-made client behind.

## Device registration and iPhones

**Only marking attendance** requires a registered device, and only **Check In** requires the office location, unless the person is on approved field duty (see *Field duty*). Check Out and Overtime Check Out work from anywhere, including with location denied or unavailable, because people forget to sign out before leaving; the position is still recorded when the phone supplies one, and the log says `Location Not Provided` when it does not. The staff portal, requisitions and the admin ERP sign in with the Staff ID and PIN alone, from any device and any place. The attendance lock is one registered browser per person: the browser holds a random token in its site storage and the sheet holds only the SHA-256 hash of it. It is what stops one person clocking in for a colleague.

### Why iPhones keep losing it

The token lives only in the browser's storage, and iPhone browsers are the least willing to keep it:

1. **The page runs in a cross-site frame.** Google serves every Apps Script web app inside a frame from `googleusercontent.com`. Safari treats storage in that position as third-party: it is partitioned, restricted by Intelligent Tracking Prevention, and can be discarded when the site has not been interacted with directly.
2. **Links opened from WhatsApp, Gmail, Instagram and similar open in an in-app browser.** That browser has its own storage, separate from Safari, and often wipes it. A person who registers in Safari and opens tomorrow's link from WhatsApp is, to the server, on an unregistered device.
3. **Safari, Chrome on iPhone and a Home Screen shortcut each have separate storage.** Registering in one and checking in from another looks identical to borrowing a colleague's phone.
4. **Private tabs** keep nothing.
5. **Failure was silent.** A blocked `localStorage` call used to throw in the middle of a check-in, and a token the browser failed to save was assumed saved.

The server cannot tell "this iPhone lost its storage" from "this is somebody else's phone": both arrive with no token. That is why this cannot be fixed by code in the page alone.

### What the application does now

- Every page reads and writes the token through one wrapper (`window.deviceTokenStore` in `Index.html`) that never throws, reads back what it wrote, and asks the browser to treat storage as persistent.
- The attendance page tells people **before** they type anything when the browser is blocking storage, or when their account is registered to a device but this browser has no token, and says what to do.
- A check-in that succeeds but could not be saved warns the person that tomorrow may be rejected.
- Rejections are split into **no token on this browser** and **a different device is registered**, each with its own guidance, and both point at the self-service *Lost your device? Recover your account* link before Admin.
- **Rejected attempts are now logged** to the Activity Log as `ATTENDANCE REJECTED - DEVICE`, with the cause and a browser label (`iOS Safari`, `iOS (in-app browser or home-screen app)`, `iOS (Chrome/Firefox/Edge)`, `Android`). Filter that action to see whether the complaints are really iPhones, really in-app browsers, or really a handful of people.

### The `DEVICE_MISMATCH_POLICY` setting

| Value | Behaviour |
| --- | --- |
| `BLOCK` (default) | An unrecognised browser is rejected. Nothing about the lock changes. |
| `REBIND_AT_OFFICE` | An unrecognised browser is re-registered and the check-in goes through, **only** when the PIN is correct and GPS puts the person inside an approved office location. |

Anything else is treated as `BLOCK`. Rebinding is recorded as `Device Re-registered (...)` in the attendance log and as `DEVICE REBOUND AT OFFICE` in the Activity Log.

**This is a policy decision, so it is off by default.** It means someone physically at the office who knows a colleague's PIN could register their own phone as the colleague's device. The location condition keeps the lock meaningful for anyone clocking in from elsewhere, and every rebind is logged, but it is a real loosening. Turn it on if the daily rejections cost more than that risk. **Replace every placeholder PIN first** (see *One PIN for every page*): with this policy on, an account still holding its placeholder can be claimed at the office without knowing any PIN.

It applies to attendance only, because nothing else checks a device. Requisitions, expense claims and leave are protected by the PIN alone, with no device or location condition.

Run **Setup / Repair ERP Upgrade** once to add the setting row to an existing workbook.

### If it still happens

- Ask the person to open the link in **Safari or Chrome in a normal tab**, never from inside WhatsApp or Gmail, and to use the same browser every day.
- Better still, have them **Add to Home Screen** from Safari and always use that icon. Installed web apps get their own storage, which iOS does not expire the way it does for ordinary sites.
- The durable fix is to host the front end on your own domain (for example under `movingsurface.com.ng`) so the token is first-party storage, with Apps Script behind it as the API. That removes the cross-site frame entirely and is the only change that addresses the root cause.

## Interface behaviour

### A reload keeps you where you are

Every save reloads the portal's data. That reload used to re-render all pages and then force the dashboard back into view, so any action bounced the user out of the page they were working on. The portal now tracks the active page and restores it after a reload, keeps the scroll position, and only falls back to the dashboard when the page no longer exists or the user's permissions no longer include it. A fresh sign-in still starts on the dashboard.

Page filters are cleared when you *navigate* to a different page, not when the data reloads underneath you.

### Icons

Icons are inline SVG from a small registry in `Index.html`; there is no icon library loaded from a CDN. Two things used to leave blank gaps:

- A screen rebuilt its markup with `innerHTML` and forgot to call `refreshIcons` afterwards, or refreshed only part of what it had just written.
- An icon name that was not in the registry rendered as nothing at all, silently.

A `MutationObserver` now upgrades any `i[data-lucide]` as it enters the document, so **no caller has to remember to refresh icons**, and changing an element's `data-lucide` re-renders it. An unregistered name renders a visible placeholder and warns to the browser console instead of leaving a hole. Rendering is idempotent and terminates: calling `refreshIcons` by hand on top of the observer is harmless.

To add an icon, add its SVG path data to the `paths` map in `Index.html` and use `<i data-lucide="your-name"></i>` anywhere.

### Pointer behaviour

Plain text is not interactive, so it keeps the ordinary arrow cursor. An I-beam
over a table cell or a heading reads as "you can type here" when you cannot, and
a hand pointer reads as "click me". The page therefore defaults to `cursor:
default`, and only the things that really are typeable or clickable opt back in:
`input` and `textarea` ask for the caret, and buttons, links, selects and
checkboxes ask for the pointer.

Icons and the SVG inside a button are excluded from pointer events entirely, so
an icon can no longer swallow the click meant for the button around it.

Text is still selectable. Switching pointer events off across whole blocks of
text would also make it impossible to select and copy a client's email or an
item ID out of a table, which is worth more than the hover cue costs.

`runstyles.js` in the test harness enforces this: a new `cursor: text` or
`cursor: pointer` rule on anything that is not a field or a control fails.

### Placeholders

Form fields show real placeholders rather than prefilled example values that had to be deleted first. Dropdowns open on a prompt (`Select a category…`) that is `disabled` and carries an empty value, so it is visible but can never be submitted. Forms that edit an existing record still prefill the record's current values, which is the one case where a prefilled value is correct.

### Delete confirmations

Deletions are routed through `window.confirmAppAction`, which opens a modal naming what will be removed and what will be kept. The server independently requires confirmation for a delete that would discard history, so the guard does not depend on the browser.

Edit screens use `window.openFormModal`, a reusable modal that takes form markup, returns the values the caller asks for, and can refuse a save with an inline message.

## Check-in location

Check In must happen at an approved office (`Office Locations` sheet); Check Out and Overtime Check Out do not need a location at all.

The test is deliberately not a hard line. A phone indoors often reports a position hundreds of metres out, and says how uncertain it is. A check-in passes when:

```
distance to the nearest office  <=  office radius  +  LOCATION_TOLERANCE_METRES  +  phone-reported accuracy (capped)
```

| Setting | Default | Meaning |
| --- | --- | --- |
| `LOCATION_TOLERANCE_METRES` | `150` | Flat extra distance for ordinary GPS drift. `0` makes the radius exact. |
| `LOCATION_MAX_ACCURACY_ALLOWANCE_METRES` | `500` | Cap on how much of the phone's own uncertainty is allowed, so a deliberately vague reading cannot stretch an office across the city. |

The page also stops taking the first reading a phone returns, which indoors is usually its worst. It listens for up to eight seconds, submits as soon as a reading is accurate to 40 m, and otherwise submits the best one it saw.

### Tracing a failure

Every refused check-in is written to the Activity Log as `ATTENDANCE REJECTED - LOCATION`, with the nearest office, the distance, the radius and tolerance, the accuracy the phone reported, and the kind of browser. The person is shown the same numbers. Read them like this:

- **Distance a few hundred metres, accuracy large:** a weak fix. Precise location and Wi-Fi on, near a window, try again. The tolerance settings are the lever if it is common.
- **Distance similar for everyone, always wrong by the same amount:** the office coordinates are wrong. Fix the `Office Locations` row.
- **"N active office row(s) have unreadable coordinates":** an office row holds text, an out-of-range value or swapped latitude and longitude. Such a row used to be skipped silently, so that office never matched anyone. A decimal comma (`9,0765`) is now read correctly.
- **"Your phone did not share a location":** permission is off for the site, usually after an iPhone privacy prompt was dismissed.

Accepted-on-tolerance check-ins are noted in the attendance log's Remarks, so you can see how often the leniency is being used.

## Field duty

Staff sometimes travel for installation or event jobs, within the country or abroad, for days or weeks, and do not mark attendance while away. Without a record of that, payroll counts every working day they were away as an absence. **Field duty** is that record.

An **approved** field duty does three things:

1. **Its working days are excused in payroll.** They count as worked, never as absent, and are paid in full.
2. **The person shows as "on field duty"**, with the site and the date they are due back, instead of appearing in "Not signed in".
3. **Checking in is optional and is not tied to the office.** If they do check in, no office location is needed, and the check-in is recorded as `On Field Duty`: never Late, Half Day or Absent, because those rules are Lagos office hours. The registered device is still required, and being on duty never counts as being at the office for `REBIND_AT_OFFICE`. Checking out has never needed a location.

A **pending, rejected or cancelled** duty excuses nothing.

### Recording and approving

The **Field Duty** page records who, the first and last day (both included), the site or city, the reason, and optionally the project.

| Who | May |
| --- | --- |
| Admin, Super Admin | Approve any field duty. They always have the page, even if their access list predates it. |
| The project's manager | Approve a duty that names their project. |
| Anyone else with the page | Record a duty, which waits as *Pending Approval*. |

- Someone who is allowed to approve a duty approves it as they record it.
- **Nobody approves their own duty**, except a Super Admin, so a workbook with a single Super Admin cannot get stuck.
- **A project manager needs an access account with the Field Duty page.** Grant it on the Users page. Their approval rights come from being the project's manager, not from the page.
- **The manager link is a name match.** Projects store the manager as text, so the person's staff name must equal it (case, spacing and punctuation are ignored). If it matches nobody, or two people, nobody gets approval through it and an Admin must approve. A typo fails safe.
- Rejecting needs a reason. A decided duty is final.
- Duties for the same person may not overlap, and one duty covers at most 366 days.
- **Backdating is allowed**, for a trip nobody filed. It needs a reason like any other, is flagged `BACKDATED` in the Activity Log, and excuses the days already gone once approved.
- A **started** duty cannot be cancelled, only shortened or extended by an approver with *Change end*. Cancelling would un-excuse days already worked and possibly already paid. Before it starts, the requester or an approver can cancel it.

### How payroll counts a month

For each working day in the month (the weekdays in the `WORKDAYS` setting), exactly one of these applies, in this order:

1. **The person checked in**: present. Nothing is excused, so a check-in on a duty day is never counted twice.
2. **Approved field duty**: a field duty day.
3. **Approved paid leave**: a leave day.
4. **A public holiday**: a holiday day.

```
expected days = the smaller of (required days, working days in this month)
absent days   = expected days - present - field duty days - leave days - holiday days   (never below zero)
```

**Why the cap.** A staff member's required days are a fixed figure (22 by default), but a month often has fewer working days: November 2026 has 21. Excused days can never exceed the month's real working days, so without the cap a person on field duty for the **whole** month would still be one day short and lose a day's pay, and so would someone who attended every working day. With it, a full month on field duty is full pay, and so is full attendance. The daily rate stays on the required days, so salaries do not move. This also removes the one-or-two-day phantom absence full-attendance staff were already getting in short months, which is a deduction change in staff's favour for everyone, not only people on duty. If the month's working days cannot be determined the required days are used, as before.

`WORKDAYS` is a single global setting, so a staff member whose Work Days Per Month assumes a six-day week is capped at the global week's working days. Set `WORKDAYS` to match, or leave such staff for a per-person setting.

Weekends are never counted. Leave is counted in **working days**, not the calendar days on the leave request, so a leave that spans a weekend does not excuse the weekend. Ranges are clipped to the month.

**Leave and holidays are part of this change, and they change deductions.** Before it, payroll fixed leave at zero, so approved leave was deducted as absence, and holidays were never excused. Now:

- **Approved leave is excused**, except leave of type **Unpaid**, which stays a deduction. Pending and rejected leave excuses nothing.
- **A holiday on a working day is excused.** A row in `Holidays` counts unless its `Active` cell says `NO`, so a holiday with a blank flag is not turned into a deduction.
- Anyone with approved leave this month will see their deduction fall as soon as this is deployed.

The Accountant View gains two columns at the end, **Field Duty Days** and **Holiday Days**. `Leave Days` now shows real values.

Note that the monthly figure follows the existing design: required days are the whole month, so a duty later in the month is excused in advance, in step with the required days it offsets.

### Setup

Run **Setup / Repair ERP Upgrade** once. It creates the `Field Duty` sheet and adds the two Accountant View columns. Nothing else needs migrating.

## Hosting on our own domain

The pages can be served from `mserp.movingsurface.com.ng` while Apps Script keeps doing the work. The Google web app address stays working, which is the way back.

| File | Job |
| --- | --- |
| `Api.gs` | `doPost` JSON API. Runs only the functions listed in `apiFunctions_()`. **Add a new server function there when a page starts calling it**, or the hosted page gets "Unknown request". |
| `Security.gs` | Wrong-PIN limit, and the owner-only guard on setup and maintenance functions. |
| `../web/bridge.js` | Replaces `google.script.run` on the hosted pages. Never retries. |
| `../web/build.js` | Builds `web/dist/` from these sources. Re-run it, and re-upload, after changing any page or `Styles.html`. |

Go live:

1. Copy the `.gs` and `.html` files into Apps Script, including the new `Api.gs` and `Security.gs`.
2. Run **Setup / Repair ERP Upgrade** from the editor and approve the new permission it asks for (reading your email address, used by the owner-only guard). Then run `testHostedApiRules` and `testPinAttemptLimit`.
3. **Deploy > Manage deployments**, edit the existing deployment, choose **New version**. Execute as **Me**, access **Anyone**.
4. `node web/build.js --api <your /exec address>`
5. Upload everything inside `web/dist/`, including the hidden `.htaccess`, to the subdomain's document root. To try it first, upload to a `/new/` subfolder; the pages use relative paths.
6. Test every page. Then set `PUBLIC_BASE_URL` in the `Settings` sheet to `https://mserp.movingsurface.com.ng`, so recovery emails, dashboard links and page links use the domain. Email approval links stay on Google.
7. Deal with device registrations: they belong to the old address, so set `DEVICE_MISMATCH_POLICY` to `REBIND_AT_OFFICE` for one week (replace all placeholder PINs first), then back to `BLOCK`.

**PIN limit.** Five wrong PINs for a Staff ID within five minutes lock that ID for up to five minutes on attendance, the staff portal and the admin ERP together. Account recovery or an admin reset lifts it. Anyone can trigger it against a colleague's ID, which is why it is short. Guesses sent at exactly the same moment can slip a few extra tries past it.

**Owner-only functions.** Setup, migration and maintenance functions run only from the editor or the spreadsheet menu. Time-driven trigger handlers are not guarded.

**Analytics.** Paste one snippet into `web/analytics.html` and rebuild. It is left off the recovery page.

## One PIN for every page

A person has one four-digit PIN. It lives in **Staff List** (column E) and is the credential for attendance, the staff portal, requisitions and the admin ERP.

- **Admin sign-in checks the Staff List PIN.** An `Admin Users` row linked to a Staff ID is matched by that Staff ID and checked against that person's staff PIN. The `Admin Users` PIN column is read only for legacy rows with no Staff ID, so the two stores cannot drift apart.
- **Creating, recovering or resetting a PIN changes it everywhere at once.** An admin reset that clears the PIN also stops that person signing in to the ERP until they choose a new one on the attendance page.
- **Placeholder PINs are replaced, not kept.** The three sample accounts (MS-001, MS-002, MS-003) shipped with placeholder PINs. Until replaced they still sign in, but attendance treats such an account as having no PIN: it asks the person to create one, that choice replaces the placeholder in Staff List (and so on every page), and the placeholder can never be chosen again, by creation or by recovery. The new PIN is written only when the whole check-in succeeds, after the device, location and attendance-state checks, so a rejected request can never set someone's PIN. A fresh workbook seeds those accounts with no PIN at all.
- **A new workbook has a bootstrap step.** With no PIN seeded, the first Super Admin creates theirs on the attendance page, from the office location, before the admin ERP will accept them. Whoever opens that page first for a given account claims it, so do this promptly after setup.
- **Test helpers use no real PINs.** `testAdminLogin`, `testAdminPortalData` and `testERPPortalData` check the sign-in rules on synthetic data, and nothing writes a PIN to the log.

### Migrating an existing workbook

Run **Setup / Repair ERP Upgrade** (or *Migrate Projects / Leads Schema*). Step 6 is idempotent and backs up `Admin Users` (and `Staff List` if it writes to it) first. It:

1. Links each `Admin Users` row with a blank Staff ID to the staff member whose name matches **exactly**. Ambiguous or unknown names are left unlinked and listed in the report, and those rows keep using their own PIN column so nobody is locked out.
2. Moves a legacy admin PIN into Staff List only when that person has no valid PIN there yet.
3. Where both PINs exist and differ, **keeps the Staff List PIN**. That person now types their staff PIN into the admin ERP, so tell them before you run it. The report counts these cases.
4. Clears the PIN column on every linked row so the credential is stored once.

The report lists names and counts only, never a PIN. The hidden `Backup Admin Users ...` sheet it leaves behind does contain the old PINs, so delete it once you are satisfied.

## PIN rules

PINs are four decimal digits and are stored as text. Leading zeroes are significant: `0008` is valid, while `8` is not a valid submitted PIN. Existing one-to-four digit numeric values are normalized by the setup migration; enter the resulting four digits when logging in.

## Performance and email delivery

- Public staff options are cached briefly; PIN verification always reads current staff data.
- Staff-by-ID lookups and same-day duplicate attendance checks use targeted searches instead of loading the full staff roster or attendance log.
- The admin payroll view aggregates attendance in one pass rather than rescanning it separately for every staff member.
- Routine web requests no longer repair/format all spreadsheet tabs. Run setup after schema changes.
- Notification messages are recorded in `Mail Queue` and sent by the time-driven `processPendingMailQueue` worker. The public `installMailQueueWorker` function can recreate its one-minute trigger from the Apps Script editor. The web request no longer waits for `MailApp.sendEmail`; email delivery is asynchronous and may take up to about a minute plus Google's delivery time. Inspect `Mail Queue` for `FAILED` rows and the last error.
- Google Apps Script has cold starts, spreadsheet/network latency, and location-provider delays. Those platform services cannot be made millisecond-fast or guaranteed to return in milliseconds. The UI reuses a recent location fix and caps the GPS wait at ten seconds.

The web UI uses an in-project responsive design system and inline SVG icons rather than runtime CDN dependencies, avoiding additional external page-load requests. Homepage hotspots, login/action toasts, and the 10-row requisition pager use local data and built-in browser APIs; they add no polling or client-library dependencies. The public Agent Skills marketplace offers frontend-design guidance for the development agent; those skills do not run inside a deployed Apps Script web app and are not app plugins/connectors.
