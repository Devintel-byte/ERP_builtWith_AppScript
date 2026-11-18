/**
 * Follow-up reminder engine for the Leads pipeline.
 *
 * One daily trigger drives three kinds of reminder:
 *
 *   Weekly digest  - on the configured weekday, ONE email per follow-up owner
 *                    listing their leads due inside the horizon plus anything
 *                    already overdue. One digest, not one mail per lead, so the
 *                    reminders stay readable instead of being filtered away.
 *   Final reminder - fires when a follow-up is exactly LEAD_FINAL_REMINDER_DAYS
 *                    away (default 3). Individual, urgent, copied to management.
 *   Overdue        - no separate email. Overdue leads escalate inside the next
 *                    digest and surface as a dashboard hotspot.
 *
 * Every send is keyed in Lead Reminder Log, so a retry, a manual re-run or a
 * double trigger can never deliver the same reminder twice.
 */

function getLeadReminderSettings_() {
  const raw = getKeyValueSheet_(ERP_SETTINGS_SHEET);
  const horizon = Number(raw.LEAD_REMINDER_HORIZON_DAYS || 28);
  const finalDays = Number(raw.LEAD_FINAL_REMINDER_DAYS || 3);
  return {
    HORIZON_DAYS: horizon > 0 ? Math.floor(horizon) : 28,
    FINAL_DAYS: finalDays > 0 ? Math.floor(finalDays) : 3,
    DIGEST_WEEKDAY: normalizeWeekday_(raw.LEAD_WEEKLY_DIGEST_DAY || 'Monday'),
    REMINDERS_ENABLED: String(raw.LEAD_REMINDERS_ENABLED || 'YES').toUpperCase() !== 'NO'
  };
}

function normalizeWeekday_(value) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const supplied = String(value || '').trim().toLowerCase();
  for (let i = 0; i < days.length; i++) {
    if (days[i].toLowerCase() === supplied) return i;
  }
  return 1;
}

/*************** IDEMPOTENCY ***************/

function leadReminderLogSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEAD_REMINDER_LOG_SHEET);
  if (!sheet) throw new Error('Lead Reminder Log sheet is missing. Run Setup / Repair ERP Upgrade first.');
  return sheet;
}

/** Every reminder key already recorded, as a lookup object. */
function loadSentReminderKeys_() {
  const sheet = leadReminderLogSheet_();
  const sent = Object.create(null);
  if (sheet.getLastRow() < 2) return sent;
  sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().forEach(function (row) {
    const key = String(row[0] || '').trim();
    if (key) sent[key] = true;
  });
  return sent;
}

function recordReminderSent_(entry) {
  leadReminderLogSheet_().appendRow([
    new Date(), String(entry.key || ''), String(entry.leadId || ''), String(entry.type || ''),
    String(entry.dueDate || ''), String(entry.recipients || ''), String(entry.subject || '').slice(0, 500),
    String(entry.status || 'QUEUED'), String(entry.notes || '').slice(0, 1000)
  ]);
}

/*************** RECIPIENTS ***************/

/**
 * Resolves a lead's reminder recipients: the follow-up owner's portal email
 * first, falling back to the senior admin roles when no owner is set, so a lead
 * never goes unwatched because its owner field is blank.
 */
function getLeadOwnerEmail_(lead) {
  const ownerId = String(lead.owner || '').trim();
  if (!ownerId) return '';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET);
  if (sheet && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(12, sheet.getLastColumn())).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][8] || '').trim() === ownerId && String(rows[i][2] || '').toUpperCase() === 'YES') {
        const email = String(rows[i][9] || '').trim();
        if (email && isValidEmail_(email)) return email;
      }
    }
  }
  const staff = getStaffByIdAnyStatus_(ownerId);
  const companyEmail = staff ? String(staff.companyEmail || '').trim() : '';
  return companyEmail && isValidEmail_(companyEmail) ? companyEmail : '';
}

function getLeadManagementEmails_() {
  const settings = getReqSettings_();
  return uniqueEmails_(
    getAdminRoleEmails_(['superadmin', 'admin', 'adminhr', 'administrator']).concat([settings.MD_EMAIL])
  );
}

/*************** THE DAILY JOB ***************/

/**
 * Daily reminder sweep. Installed by installLeadReminderWorker_ to run early in
 * the configured timezone; also safe to run by hand from the menu.
 */
function processLeadFollowUpReminders() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('Lead reminder sweep skipped: another run holds the lock.');
    return {skipped: true};
  }
  try {
    const settings = getLeadReminderSettings_();
    if (!settings.REMINDERS_ENABLED) {
      Logger.log('Lead reminders are disabled via ERP Settings LEAD_REMINDERS_ENABLED.');
      return {disabled: true};
    }
    const today = erpToday_();
    const sent = loadSentReminderKeys_();
    const leads = getERPLeads_().filter(function (lead) {
      return lead.leadId &&
        leadClosedStages_().indexOf(lead.stage) === -1 &&
        normalizeLeadCadence_(lead.cadence) !== 'Off' &&
        lead.nextFollowup;
    });

    const result = {date: today, finalReminders: 0, digests: 0, overdue: 0, leadsConsidered: leads.length};
    const buckets = {};

    leads.forEach(function (lead) {
      const days = erpDaysBetween_(today, lead.nextFollowup);
      if (days === null) return;
      if (days < 0) result.overdue++;

      // Final reminder: exactly N days out.
      if (days === settings.FINAL_DAYS) {
        const key = lead.leadId + '|FINAL|' + lead.nextFollowup;
        if (!sent[key] && sendLeadFinalReminder_(lead, days, key)) {
          sent[key] = true;
          result.finalReminders++;
          stampLeadReminder_(lead, today);
        }
      }

      // Collect for the weekly digest: overdue, or due inside the horizon.
      if (days <= settings.HORIZON_DAYS) {
        const email = getLeadOwnerEmail_(lead);
        const bucketKey = email || 'UNASSIGNED';
        if (!buckets[bucketKey]) buckets[bucketKey] = {email: email, ownerName: lead.ownerName, leads: []};
        buckets[bucketKey].leads.push({lead: lead, days: days});
      }
    });

    const todayWeekday = new Date(today + 'T00:00:00').getDay();
    if (todayWeekday === settings.DIGEST_WEEKDAY) {
      Object.keys(buckets).forEach(function (bucketKey) {
        const bucket = buckets[bucketKey];
        const key = 'DIGEST|' + bucketKey + '|' + today;
        if (sent[key]) return;
        if (sendLeadWeeklyDigest_(bucket, today, key)) {
          sent[key] = true;
          result.digests++;
        }
      });
    }

    Logger.log('Lead reminder sweep ' + today + ': ' + JSON.stringify(result));
    if (result.finalReminders || result.digests) {
      recordActivity_('System', 'Lead Reminders', 'SYSTEM', 'NOTIFY', 'Leads', '',
        result.finalReminders + ' final reminder(s), ' + result.digests + ' digest(s), ' + result.overdue + ' overdue');
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

function stampLeadReminder_(lead, today) {
  try {
    leadsSheet_().getRange(lead.row, LEAD_COL.lastReminderSent).setValue(today);
  } catch (err) {
    Logger.log('Could not stamp Last Reminder Sent for ' + lead.leadId + ': ' + (err.message || err));
  }
}

/*************** FINAL (T-MINUS) REMINDER ***************/

function sendLeadFinalReminder_(lead, days, key) {
  const ownerEmail = getLeadOwnerEmail_(lead);
  const management = getLeadManagementEmails_();
  const to = ownerEmail ? [ownerEmail] : management;
  const cc = ownerEmail ? management.filter(email => email.toLowerCase() !== ownerEmail.toLowerCase()) : [];
  if (!to.length) {
    Logger.log('Final reminder not queued for ' + lead.leadId + ': no recipient could be resolved.');
    recordReminderSent_({key: key, leadId: lead.leadId, type: 'FINAL', dueDate: lead.nextFollowup,
      recipients: '', subject: '', status: 'NO_RECIPIENT',
      notes: 'Set a Follow-up Owner with a portal email, or an MD email in Requisition Settings.'});
    return false;
  }

  const subject = 'Follow-up due in ' + days + ' day(s) - ' + lead.clientName + ' (' + lead.leadId + ')';
  const sections = emailInfoTable_([
    ['Lead', lead.leadId],
    ['Client', lead.clientName],
    ['Opportunity', lead.opportunity || '-'],
    ['Stage', lead.stage],
    ['Follow-up due', lead.nextFollowup + ' (' + days + ' day(s) away)'],
    ['Contact', (lead.contactPerson || '-') + (lead.email ? ' / ' + lead.email : '') + (lead.phone ? ' / ' + lead.phone : '')],
    ['Opportunity value', 'NGN ' + formatMoney_(lead.estimatedValue)],
    ['Expected close', lead.expectedClose || '-'],
    ['Owner', lead.ownerName || 'Unassigned'],
    ['Last contact', lead.lastContact || 'No contact logged']
  ]) +
    '<div style="margin:18px 0;padding:13px 15px;border-left:4px solid #bd7517;background:#fff7e8;color:#875313">' +
    '<b>Final reminder</b><br>This is the last scheduled reminder before the follow-up date. ' +
    'Log the contact on the Leads page so the next follow-up is scheduled.</div>' +
    (lead.notes ? emailSection_('Notes', emailParagraph_(lead.notes)) : '');

  const dashboardUrl = getWebAppRouteUrl_('admin');
  const body = sections + (dashboardUrl
    ? '<div style="margin-top:22px">' + emailButton_(dashboardUrl, 'Open the Leads page', '#087f68') + '</div>'
    : '');

  const queued = tryQueueEmail_({
    to: to.join(','), cc: cc.join(','), subject: subject,
    htmlBody: emailFrame_('Follow-up due soon', lead.clientName + ' - ' + lead.stage, body),
    relatedId: lead.leadId
  });
  recordReminderSent_({
    key: key, leadId: lead.leadId, type: 'FINAL', dueDate: lead.nextFollowup,
    recipients: to.concat(cc).join(','), subject: subject,
    status: queued ? 'QUEUED' : 'QUEUE_FAILED', notes: days + ' day(s) before due date'
  });
  return queued;
}

/*************** WEEKLY DIGEST ***************/

function sendLeadWeeklyDigest_(bucket, today, key) {
  const management = getLeadManagementEmails_();
  const to = bucket.email ? [bucket.email] : management;
  if (!to.length) {
    Logger.log('Weekly lead digest not queued: no recipient for bucket ' + (bucket.ownerName || 'unassigned'));
    recordReminderSent_({key: key, leadId: '', type: 'WEEKLY', dueDate: today, recipients: '', subject: '',
      status: 'NO_RECIPIENT', notes: 'No owner email and no management email configured.'});
    return false;
  }

  const rows = bucket.leads.slice().sort((a, b) => a.days - b.days);
  const overdue = rows.filter(entry => entry.days < 0);
  const thisWeek = rows.filter(entry => entry.days >= 0 && entry.days <= 7);
  const later = rows.filter(entry => entry.days > 7);
  const pipelineValue = rows.reduce((sum, entry) => sum + entry.lead.estimatedValue, 0);

  const subject = 'Lead follow-ups - ' + overdue.length + ' overdue, ' + thisWeek.length + ' due this week';
  let body = emailInfoTable_([
    ['Owner', bucket.ownerName || (bucket.email ? bucket.email : 'Unassigned leads')],
    ['Overdue', String(overdue.length)],
    ['Due this week', String(thisWeek.length)],
    ['Due later', String(later.length)],
    ['Pipeline value in view', 'NGN ' + formatMoney_(pipelineValue)]
  ]);

  if (overdue.length) {
    body += '<div style="margin:18px 0;padding:13px 15px;border-left:4px solid #b83b43;background:#fdecec;color:#8d2b32">' +
      '<b>' + overdue.length + ' follow-up(s) are past their date.</b><br>' +
      'These were scheduled and not logged. Clear them first.</div>';
    body += emailSection_('Overdue', leadDigestTable_(overdue));
  }
  if (thisWeek.length) body += emailSection_('Due this week', leadDigestTable_(thisWeek));
  if (later.length) body += emailSection_('Coming up', leadDigestTable_(later));
  if (!bucket.email) {
    body += '<div style="margin:18px 0;padding:13px 15px;border-left:4px solid #2864c5;background:#eef3fc;color:#1d4a94">' +
      '<b>These leads have no follow-up owner.</b><br>' +
      'Assign an owner on the Leads page so the reminders reach the right person.</div>';
  }

  const dashboardUrl = getWebAppRouteUrl_('admin');
  if (dashboardUrl) body += '<div style="margin-top:22px">' + emailButton_(dashboardUrl, 'Open the Leads page', '#087f68') + '</div>';

  const queued = tryQueueEmail_({
    to: to.join(','), cc: '', subject: subject,
    htmlBody: emailFrame_('Weekly follow-up digest', 'Week of ' + today, body),
    relatedId: 'LEAD-DIGEST-' + today
  });
  recordReminderSent_({
    key: key, leadId: rows.map(entry => entry.lead.leadId).join(' ').slice(0, 400), type: 'WEEKLY',
    dueDate: today, recipients: to.join(','), subject: subject,
    status: queued ? 'QUEUED' : 'QUEUE_FAILED',
    notes: rows.length + ' lead(s); ' + overdue.length + ' overdue'
  });
  return queued;
}

function leadDigestTable_(entries) {
  const header = '<tr>' +
    ['Due', 'Lead', 'Client', 'Opportunity', 'Stage', 'Value (NGN)'].map(label =>
      '<th style="text-align:left;padding:7px 9px;background:#f1f5f4;border-bottom:1px solid #dbe3e1;font-size:12px">' +
      escapeHtml_(label) + '</th>').join('') + '</tr>';
  const rows = entries.map(function (entry) {
    const lead = entry.lead;
    const dueLabel = entry.days < 0
      ? lead.nextFollowup + ' (' + Math.abs(entry.days) + 'd overdue)'
      : lead.nextFollowup + ' (in ' + entry.days + 'd)';
    const colour = entry.days < 0 ? '#b83b43' : entry.days <= 3 ? '#bd7517' : '#334';
    return '<tr>' +
      '<td style="padding:7px 9px;border-bottom:1px solid #eef1f0;font-size:12px;color:' + colour + ';font-weight:700">' + escapeHtml_(dueLabel) + '</td>' +
      '<td style="padding:7px 9px;border-bottom:1px solid #eef1f0;font-size:12px">' + escapeHtml_(lead.leadId) + '</td>' +
      '<td style="padding:7px 9px;border-bottom:1px solid #eef1f0;font-size:12px">' + escapeHtml_(lead.clientName) + '</td>' +
      '<td style="padding:7px 9px;border-bottom:1px solid #eef1f0;font-size:12px">' + escapeHtml_(lead.opportunity) + '</td>' +
      '<td style="padding:7px 9px;border-bottom:1px solid #eef1f0;font-size:12px">' + escapeHtml_(lead.stage) + '</td>' +
      '<td style="padding:7px 9px;border-bottom:1px solid #eef1f0;font-size:12px;text-align:right">' + escapeHtml_(formatMoney_(lead.estimatedValue)) + '</td>' +
      '</tr>';
  }).join('');
  return '<table style="width:100%;border-collapse:collapse;margin-top:6px">' + header + rows + '</table>';
}

/*************** TRIGGER MANAGEMENT ***************/

function installLeadReminderWorker_() {
  const exists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === 'processLeadFollowUpReminders';
  });
  if (!exists) {
    ScriptApp.newTrigger('processLeadFollowUpReminders')
      .timeBased().atHour(7).everyDays(1).inTimezone(erpTimezone_()).create();
  }
}

function installProjectReconcileWorker_() {
  const exists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === 'reconcileAllProjectRollups';
  });
  if (!exists) {
    ScriptApp.newTrigger('reconcileAllProjectRollups')
      .timeBased().atHour(2).everyDays(1).inTimezone(erpTimezone_()).create();
  }
}

/** Public entry points so the triggers can be recreated from the editor. */
function installLeadReminderWorker() {
  installLeadReminderWorker_();
  installProjectReconcileWorker_();
  Logger.log('Lead reminder and project reconcile triggers are installed.');
}

/** Menu helper: runs the sweep now and reports what it did. */
function runLeadReminderSweepNow() {
  const result = processLeadFollowUpReminders();
  SpreadsheetApp.getUi().alert('Lead follow-up reminders\n\n' + JSON.stringify(result, null, 2) +
    '\n\nQueued mail is delivered by the one-minute Mail Queue worker.');
}
