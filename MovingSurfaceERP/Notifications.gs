function sendRequisitionToMd_(reqId) {
  const settings = getReqSettings_();
  const row = getRequisitionById_(reqId);
  if (!row) return false;
  const itemRows = parseRequisitionItems_(row);
  const inventoryItems = getERPInventoryItems_();
  const requestType = classifyRequisitionItems_(itemRows,inventoryItems);
  const requestsInventory = requestType.inventory;
  const requestsPurchase = requestType.purchase;
  const roleEmails = [];
  if (requestsInventory) roleEmails.push.apply(roleEmails,getAdminRoleEmails_(['superadmin','admin','adminhr','administrator','technical']));
  if (requestsPurchase) roleEmails.push.apply(roleEmails,getAdminRoleEmails_(['superadmin','admin','adminhr','administrator','accountant']));
  const managerEmails = uniqueEmails_([settings.MD_EMAIL]);
  const notificationEmails = uniqueEmails_(roleEmails).filter(email =>
    !managerEmails.some(managerEmail => managerEmail.toLowerCase() === email.toLowerCase())
  );
  const highValue = Number(row.total || 0) >= Number(settings.HIGH_VALUE_THRESHOLD || 500000);
  const subject = (highValue ? 'HIGH VALUE - ' : '') + 'New Item Requisition - ' + reqId;
  let sections = emailInfoTable_([
    ['Request ID',reqId],['Requested by',row.staffName+' ('+row.staffId+')'],
    ['Department',row.department],['Project / event',row.project || '-'],
    ['Priority',row.priority],['Needed before',row.neededBefore || '-'],
    ['Estimated total','NGN '+formatMoney_(row.total)],
    ['Request type',requestsInventory && requestsPurchase ? 'Inventory and purchase items' :
      requestsInventory ? 'Inventory item' : requestsPurchase ? 'Purchase item' : 'No item lines']
  ]);
  sections += emailSection_('Requested items',emailItemsTable_(itemRows));
  sections += emailSection_('Purpose',emailParagraph_(row.purpose));
  const attachmentUrl = safeEmailUrl_(row.attachment);
  sections += emailSection_('Attachment',attachmentUrl
    ? '<a href="'+escapeHtml_(attachmentUrl)+'" style="color:#087f68;font-weight:700">Open attachment</a>'
    : '<span style="color:#687874">No attachment supplied</span>');
  if (highValue) sections += '<div style="margin:18px 0;padding:13px 15px;border-left:4px solid #bd7517;background:#fff7e8;color:#875313"><b>High-value request</b><br>This request is above the configured approval threshold.</div>';

  let approvalQueued = false;
  const approveUrl = getWebAppRouteUrl_('attendance', {approveReq:reqId,decision:'Approved',token:row.token});
  const rejectUrl = getWebAppRouteUrl_('attendance', {approveReq:reqId,decision:'Rejected',token:row.token});
  if (!managerEmails.length) {
    Logger.log('Requisition approval notification not queued: MD_EMAIL is not configured for ' + reqId);
  } else if (!approveUrl || !rejectUrl) {
    Logger.log('Requisition approval notification not queued: deploy the web app before accepting requests (' + reqId + ')');
  } else {
    const approvalBody = sections+'<div style="margin-top:22px">'+emailButton_(approveUrl,'Approve request','#087f68')+
      '&nbsp;&nbsp;'+emailButton_(rejectUrl,'Reject request','#b83b43')+'</div>';
    approvalQueued = tryQueueEmail_({
      to:managerEmails.join(','),
      cc:settings.HR_EMAIL || '',
      subject:subject,
      htmlBody:emailFrame_('New item requisition','Approval requested',approvalBody),
      relatedId:reqId
    });
  }
  const adminDashboardUrl = getWebAppRouteUrl_('admin');
  const notificationContent = sections+(adminDashboardUrl
    ? '<div style="margin-top:22px">'+emailButton_(adminDashboardUrl,'Open admin dashboard','#2864c5')+'</div>'
    : '');
  const notificationQueued = notificationEmails.length ? tryQueueEmail_({
    to:notificationEmails.join(','),
    subject:'Requisition alert - '+reqId,
    htmlBody:emailFrame_('New item requisition','For your awareness',notificationContent),
    relatedId:reqId
  }) : false;
  return approvalQueued || notificationQueued;
}

function handleEmailApprovalPage_(params) {
  const result = processEmailRequisitionDecision_(params.approvereq || params.approveReq, params.decision, params.token);
  return HtmlService.createHtmlOutput(`<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1"><style>body{font-family:Arial;background:#f5f7fb;padding:20px}.card{max-width:560px;margin:40px auto;background:#fff;padding:24px;border-radius:16px;box-shadow:0 4px 20px rgba(0,0,0,.08)}.ok{color:#166534}.bad{color:#991b1b}</style></head><body><div class="card"><h2 class="${result.success ? 'ok' : 'bad'}">${result.success ? 'Request Updated' : 'Action Failed'}</h2><p>${escapeHtml_(result.message)}</p></div></body></html>`).setTitle('Requisition Decision');
}

function processEmailRequisitionDecision_(reqId, decision, token) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const row = getRequisitionById_(reqId);
    if (!row) return {success:false, message:'Request not found.'};
    if (String(row.token) !== String(token)) return {success:false, message:'Invalid or expired approval link.'};
    if (row.status !== 'Pending MD Approval') return {success:false, message:'This request has already been processed. Current status: '+row.status};
    if (decision !== 'Approved' && decision !== 'Rejected') return {success:false, message:'Invalid decision.'};

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQ_SHEET);
    const now = new Date();
    const newStatus = decision === 'Approved' ? 'Approved - Awaiting Accounts' : 'Rejected by MD';

    sh.getRange(row.row, 11).setValue(newStatus);
    sh.getRange(row.row, 12).setValue(now);
    sh.getRange(row.row, 13).setValue('MD Email Link');
    sh.getRange(row.row, 22).setValue(now);
    logReqAudit_(row.reqId, decision, 'MD Email Link', row.status, newStatus, 'Email approval link used');
    recordActivity_('Approver','MD Email Link','EMAIL-APPROVAL','UPDATE','Requisition',row.reqId,row.status+' -> '+newStatus);

    if (decision === 'Approved') {
      const accountsQueued = sendApprovedRequisitionToAccounts_(row.reqId);
      const staffQueued = notifyStaffRequisitionStatus_(row.reqId, newStatus);
      if (accountsQueued) sh.getRange(row.row, 15).setValue(now);
      if (staffQueued) sh.getRange(row.row, 20).setValue(now);
      return {success:true, message:'Request '+row.reqId+' approved.'+(accountsQueued ? ' Accounts notification queued.' : ' Accounts notification could not be queued; check Mail Queue.')+(staffQueued ? ' Staff notification queued.' : '')};
    }

    const staffQueued = notifyStaffRequisitionStatus_(row.reqId, newStatus);
    if (staffQueued) sh.getRange(row.row, 20).setValue(now);
    return {success:true, message:'Request '+row.reqId+' rejected.'+(staffQueued ? ' Staff notification queued.' : '')};
  } finally {
    lock.releaseLock();
  }
}

function sendApprovedRequisitionToAccounts_(reqId) {
  const settings = getReqSettings_();
  const row = getRequisitionById_(reqId);
  if (!row) return false;
  if (!settings.ACCOUNTS_EMAIL) {
    Logger.log('Accounts notification not queued: ACCOUNTS_EMAIL is not configured for ' + reqId);
    return false;
  }
  const dashboardUrl = getWebAppRouteUrl_('admin');
  const sections = emailInfoTable_([
    ['Request ID',reqId],['Requested by',row.staffName+' ('+row.staffId+')'],
    ['Department',row.department],['Estimated total','NGN '+formatMoney_(row.total)]
  ])+emailSection_('Requested items',emailItemsTable_(parseRequisitionItems_(row)))+
    emailSection_('Purpose',emailParagraph_(row.purpose))+
    '<div style="margin-top:22px">'+emailButton_(dashboardUrl,'Open accounts dashboard','#2864c5')+'</div>';
  const body = emailFrame_('Requisition approved','Ready for accounts processing',sections);
  return tryQueueEmail_({
    to:settings.ACCOUNTS_EMAIL,
    cc:settings.HR_EMAIL || '',
    subject:'Approved Requisition for Disbursement - '+reqId,
    htmlBody:body,
    relatedId:reqId
  });
}

function notifyStaffRequisitionStatus_(reqId, status) {
  const settings = getReqSettings_();
  if (String(settings.ALLOW_STAFF_EMAIL_COPY || 'YES').toUpperCase() !== 'YES') return false;
  const row = getRequisitionById_(reqId);
  if (!row) return false;
  const staff = getStaffByIdAnyStatus_(row.staffId);
  if (!staff) return false;
  const email = String(staff.companyEmail || staff.personalEmail || '').trim();
  if (!email) return false;
  return tryQueueEmail_({
    to: email,
    subject: 'Requisition Update - '+reqId,
    htmlBody: emailFrame_('Requisition update','Status changed',emailInfoTable_([
      ['Request ID',reqId],['Status',status],['Estimated total','NGN '+formatMoney_(row.total)]
    ])+emailSection_('Requested items',emailItemsTable_(parseRequisitionItems_(row)))),
    relatedId:reqId
  });
}

function getAdminRoleEmails_(roles) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_USERS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const accepted = Object.create(null);
  roles.forEach(role => { accepted[normalizeAdminRole_(role)] = true; });
  return sheet.getRange(2,1,sheet.getLastRow()-1,Math.max(12,sheet.getLastColumn())).getValues()
    .filter(row => String(row[2]).toUpperCase() === 'YES' && accepted[normalizeAdminRole_(row[3])])
    .map(row => String(row[9] || '').trim())
    .filter(isValidEmail_);
}

function uniqueEmails_(emails) {
  const found = Object.create(null);
  return emails.map(email => String(email || '').trim()).filter(email => {
    const key = email.toLowerCase();
    if (!email || found[key]) return false;
    found[key] = true;
    return true;
  });
}

function parseRequisitionItems_(row) {
  try {
    const items = JSON.parse(String(row.itemsJson || '[]'));
    return Array.isArray(items) ? items : [];
  } catch (err) {
    return String(row.itemsSummary || '').split('\n').filter(Boolean).map(line => ({
      itemName:line,quantity:'',unitCost:'',total:''
    }));
  }
}

function inventoryContainsItem_(inventory,itemName) {
  const query = String(itemName || '').trim().toLowerCase();
  return Boolean(query && inventory.some(item =>
    String(item.itemName || '').trim().toLowerCase() === query ||
    String(item.itemId || '').trim().toLowerCase() === query
  ));
}

function classifyRequisitionItems_(items,inventory) {
  return {
    inventory:items.some(item => inventoryContainsItem_(inventory,item.itemName)),
    purchase:items.some(item => !inventoryContainsItem_(inventory,item.itemName))
  };
}

function safeEmailUrl_(value) {
  const url = String(value || '').trim();
  return /^https?:\/\/[^\s]+$/i.test(url) ? url : '';
}

function emailFrame_(title,subtitle,content) {
  return '<div style="margin:0;padding:28px 12px;background:#f2f6f4;font-family:Arial,Helvetica,sans-serif;color:#17211f">'+
    '<div style="max-width:620px;margin:0 auto;overflow:hidden;border:1px solid #dce6e2;border-radius:14px;background:#fff">'+
    '<div style="padding:20px 25px;background:#12332b;color:#fff"><div style="font-size:13px;font-weight:700;letter-spacing:.04em">MOVING SURFACE</div>'+
    '<h1 style="margin:12px 0 4px;font-size:22px;line-height:1.25">'+escapeHtml_(title)+'</h1>'+
    '<div style="color:#c3d8d0;font-size:14px">'+escapeHtml_(subtitle)+'</div></div>'+
    '<div style="padding:24px 25px">'+content+'</div>'+
    '<div style="padding:14px 25px;border-top:1px solid #dce6e2;background:#f8fbf9;color:#687874;font-size:12px">Automated notice from Moving Surface ERP. Do not reply to this message.</div>'+
    '</div></div>';
}

function emailInfoTable_(rows) {
  return '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:14px">'+rows.map((row,index)=>
    '<tr style="background:'+(index%2?'#f8fbf9':'#fff')+'"><th align="left" style="width:38%;padding:10px 12px;border-bottom:1px solid #e5ece8;color:#687874;font-weight:600">'+escapeHtml_(row[0])+'</th>'+
    '<td style="padding:10px 12px;border-bottom:1px solid #e5ece8;color:#17211f">'+escapeHtml_(row[1])+'</td></tr>'
  ).join('')+'</table>';
}

function emailItemsTable_(items) {
  if (!items.length) return '<p style="margin:0;color:#687874">No item details were available.</p>';
  return '<table role="presentation" style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr>'+
    ['Item','Qty','Unit cost','Line total'].map(label=>'<th align="left" style="padding:9px 8px;border-bottom:2px solid #dce6e2;color:#687874">'+label+'</th>').join('')+
    '</tr></thead><tbody>'+items.map((item,index)=>
      '<tr style="background:'+(index%2?'#f8fbf9':'#fff')+'"><td style="padding:9px 8px;border-bottom:1px solid #e5ece8">'+escapeHtml_(item.itemName || '')+'</td>'+
      '<td style="padding:9px 8px;border-bottom:1px solid #e5ece8">'+escapeHtml_(item.quantity || '')+'</td>'+
      '<td style="padding:9px 8px;border-bottom:1px solid #e5ece8">NGN '+formatMoney_(item.unitCost || 0)+'</td>'+
      '<td style="padding:9px 8px;border-bottom:1px solid #e5ece8">NGN '+formatMoney_(item.total || (Number(item.quantity || 0)*Number(item.unitCost || 0)))+'</td></tr>'
    ).join('')+'</tbody></table>';
}

function emailSection_(title,content) {
  return '<section style="margin-top:20px"><h2 style="margin:0 0 9px;color:#17211f;font-size:15px">'+escapeHtml_(title)+'</h2>'+content+'</section>';
}

function emailParagraph_(value) {
  return '<p style="margin:0;white-space:pre-wrap;line-height:1.55;color:#344440">'+escapeHtml_(value || '-')+'</p>';
}

function emailButton_(url,label,color) {
  return '<a href="'+escapeHtml_(url)+'" style="display:inline-block;padding:11px 16px;border-radius:7px;background:'+color+';color:#fff;text-decoration:none;font-size:14px;font-weight:700">'+escapeHtml_(label)+'</a>';
}

function tryQueueEmail_(message) {
  try {
    enqueueEmail_(message);
    return true;
  } catch (err) {
    Logger.log('Unable to queue email for ' + (message.relatedId || 'unknown item') + ': ' + (err.stack || err.message));
    return false;
  }
}

function enqueueEmail_(message) {
  const to = String(message.to || '').trim();
  if (!to) throw new Error('An email recipient is required.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIL_QUEUE_SHEET);
  if (!sheet) throw new Error('Mail Queue sheet is missing. Run Setup / Repair Core System.');
  sheet.appendRow([
    new Date(),'PENDING',to,String(message.cc || ''),String(message.subject || ''),
    String(message.htmlBody || ''),0,'','',String(message.relatedId || '')
  ]);
}

function installMailQueueWorker_() {
  const exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === 'processPendingMailQueue';
  });
  if (!exists) {
    ScriptApp.newTrigger('processPendingMailQueue').timeBased().everyMinutes(1).create();
  }
}

function installMailQueueWorker() {
  installMailQueueWorker_();
}

function processPendingMailQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIL_QUEUE_SHEET);
  if (!sheet) {
    lock.releaseLock();
    throw new Error('Mail Queue sheet is missing. Run Setup / Repair Core System.');
  }

  const jobs = [];
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const rows = sheet.getRange(2,1,lastRow-1,10).getValues();
    const now = new Date();
    const claimValues = [];
    rows.forEach(function(row, index) {
      if (jobs.length >= 10) return;
      const status = String(row[1] || '');
      const lastAttempt = row[7] instanceof Date ? row[7] : (row[7] ? new Date(row[7]) : null);
      const stale = status === 'PROCESSING' && lastAttempt && now.getTime() - lastAttempt.getTime() > 15 * 60 * 1000;
      if (status !== 'PENDING' && !stale) return;
      const sheetRow = index + 2;
      const attempts = Number(row[6] || 0);
      claimValues.push({row:sheetRow,attempts:attempts});
      jobs.push({row:sheetRow,to:String(row[2] || ''),cc:String(row[3] || ''),subject:String(row[4] || ''),htmlBody:String(row[5] || ''),attempts:attempts,relatedId:String(row[9] || '')});
    });
    claimValues.forEach(function(claim) {
      sheet.getRange(claim.row,2).setValue('PROCESSING');
      sheet.getRange(claim.row,8).setValue(now);
    });
  } finally {
    lock.releaseLock();
  }

  const results = [];
  jobs.forEach(function(job) {
    let status = 'SENT';
    let errorText = '';
    const attempts = job.attempts + 1;
    try {
      MailApp.sendEmail({
        to:job.to,
        cc:job.cc,
        subject:job.subject,
        htmlBody:job.htmlBody
      });
    } catch (err) {
      status = attempts >= 5 ? 'FAILED' : 'PENDING';
      errorText = String(err.stack || err.message || err).slice(0,2000);
      Logger.log('Mail delivery failed for ' + job.relatedId + ' (attempt ' + attempts + '): ' + errorText);
    }
    results.push({row:job.row,status:status,attempts:attempts,errorText:errorText});
  });

  if (results.length) {
    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(2,2,lastRow-1,8).getValues();
    results.forEach(function(result) {
      values[result.row - 2][0] = result.status;
      values[result.row - 2][5] = result.attempts;
      values[result.row - 2][6] = new Date();
      values[result.row - 2][7] = result.errorText;
    });
    sheet.getRange(2,2,values.length,8).setValues(values);
  }
}
