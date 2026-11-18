const STAFF_RECOVERY_LIFETIME_MINUTES_ = 30;
const STAFF_RECOVERY_REQUEST_INTERVAL_MINUTES_ = 15;

function requestStaffAccountRecovery(payload) {
  const response = {success:true,message:'Request received. If your Staff ID and company email match an active profile, check your company inbox. If no link arrives, contact Admin.'};
  const failure = {success:false,message:'We could not process the request. Check your details or contact Admin.'};
  const staffId = String(payload && payload.staffId || '').trim();
  const email = String(payload && payload.email || '').trim().toLowerCase();
  if (!staffId || !email || email.length > 254) return failure;

  const cache = CacheService.getScriptCache();
  const throttleKey = 'STAFF_RECOVERY_REQUEST_' + sha256Hex_(staffId.toLowerCase()+'|'+email);
  if (cache.get(throttleKey)) {
    recordActivity_('Public','Unverified','','RECOVERY REQUEST THROTTLED','Staff Recovery','', 'A recovery request was rate limited.');
    return {success:false,message:'Please wait before submitting another recovery request.'};
  }
  cache.put(throttleKey,'1',60);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const staff = getStaffById_(staffId);
    if (!recoveryEmailMatches_(staff,email)) {
      recordActivity_('Public','Unverified','','RECOVERY REQUEST REJECTED','Staff Recovery','', 'The submitted recovery details did not match an active staff profile.');
      return {
        success:false,
        message:'We could not verify this Staff ID and company email combination. Check the Staff ID and registered company email; no recovery email was sent.'
      };
    }

    const recoverySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_RECOVERY_SHEET);
    if (!recoverySheet) throw new Error('Run setupMovingSurfaceERPUpgrade to create the Staff Recovery Log sheet.');
    const existing = recoverySheet.getLastRow() > 1
      ? recoverySheet.getRange(2,1,recoverySheet.getLastRow()-1,6).getValues()
      : [];
    const now = new Date();
    const recent = existing.some(row => {
      const requestedAt = row[2] instanceof Date ? row[2] : new Date(row[2]);
      return String(row[1]) === staff.id && !isNaN(requestedAt.getTime()) &&
        now.getTime()-requestedAt.getTime() < STAFF_RECOVERY_REQUEST_INTERVAL_MINUTES_*60000;
    });
    if (recent) {
      recordActivity_('Staff',staff.name,staff.id,'RECOVERY REQUEST THROTTLED','Staff Recovery','', 'A recovery link was requested too recently.');
      return response;
    }

    revokeStaffRecoveryRequests_(staff.id);
    const token = Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
    const expiresAt = new Date(now.getTime()+STAFF_RECOVERY_LIFETIME_MINUTES_*60000);
    const recoveryRow = recoverySheet.getLastRow()+1;
    recoverySheet.appendRow([sha256Hex_(token),staff.id,now,expiresAt,'','Pending']);
    const recoveryUrl = getWebAppRouteUrl_('recovery',{token:token});
    if (!recoveryUrl) throw new Error('The web app must be deployed before account recovery can be used.');

    const subject = 'Moving Surface account recovery';
    const body = 'A request was made to reset the PIN and registered device for your staff account. '+
      'Use this one-time link within '+STAFF_RECOVERY_LIFETIME_MINUTES_+' minutes:\n\n'+recoveryUrl+
      '\n\nIf you did not request this, contact your administrator. Do not forward this link.';
    const htmlBody = '<p>A request was made to reset the PIN and registered device for your staff account.</p>'+
      '<p><a href="'+escapeHtml_(recoveryUrl)+'">Reset PIN and register this device</a></p>'+
      '<p>This one-time link expires in '+STAFF_RECOVERY_LIFETIME_MINUTES_+' minutes. If you did not request this, contact your administrator. Do not forward the link.</p>';
    try {
      MailApp.sendEmail({to:staff.companyEmail,subject:subject,body:body,htmlBody:htmlBody});
    } catch (mailError) {
      recoverySheet.getRange(recoveryRow,6).setValue('Delivery Failed');
      recordActivity_('System','Recovery service',staff.id,'RECOVERY EMAIL FAILED','Staff Recovery','', 'Email delivery failed; no recovery token was included in the audit record.');
      return response;
    }
    recoverySheet.getRange(recoveryRow,6).setValue('Sent');
    recordActivity_('Staff',staff.name,staff.id,'RECOVERY LINK SENT','Staff Recovery','', 'One-time PIN and device recovery link emailed; expires in 30 minutes.');
    return response;
  } catch (err) {
    recordActivity_('System','Recovery service','','RECOVERY REQUEST FAILED','Staff Recovery','',String(err.message || err).slice(0,300));
    return failure;
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function recoveryEmailMatches_(staff,email) {
  const registeredEmail = String(staff && staff.companyEmail || '').trim().toLowerCase();
  const suppliedEmail = String(email || '').trim().toLowerCase();
  return Boolean(staff && registeredEmail && suppliedEmail && registeredEmail === suppliedEmail);
}

function completeStaffAccountRecovery(payload) {
  const pin = String(payload && payload.pin || '').trim();
  const confirmPin = String(payload && payload.confirmPin || '').trim();
  const token = String(payload && payload.token || '').trim();
  if (!/^[0-9]{4}$/.test(pin) || pin !== confirmPin) {
    return {success:false,message:'Enter the same four-digit PIN in both fields.'};
  }
  if (!/^[a-f0-9]{64}$/i.test(token)) {
    return {success:false,message:'This recovery link is invalid or has expired. Request a new link.'};
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_RECOVERY_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return {success:false,message:'This recovery link is invalid or has expired. Request a new link.'};
    const values = sheet.getRange(2,1,sheet.getLastRow()-1,6).getValues();
    const tokenHash = sha256Hex_(token);
    let recoveryRow = 0;
    for (let i=values.length-1;i>=0;i--) {
      if (String(values[i][0]) === tokenHash) {
        recoveryRow = i+2;
        break;
      }
    }
    if (!recoveryRow) return {success:false,message:'This recovery link is invalid or has expired. Request a new link.'};
    const row = values[recoveryRow-2];
    const expiresAt = row[3] instanceof Date ? row[3] : new Date(row[3]);
    if (String(row[5]) !== 'Sent' || isNaN(expiresAt.getTime()) || expiresAt.getTime() <= Date.now()) {
      if (String(row[5]) === 'Sent') sheet.getRange(recoveryRow,6).setValue('Expired');
      recordActivity_('Public','Unverified','','RECOVERY LINK REJECTED','Staff Recovery','', 'An invalid, expired, or already used recovery link was submitted.');
      return {success:false,message:'This recovery link is invalid or has expired. Request a new link.'};
    }
    const staff = getStaffById_(row[1]);
    if (!staff) {
      sheet.getRange(recoveryRow,6).setValue('Revoked');
      return {success:false,message:'This recovery link is invalid or has expired. Request a new link.'};
    }

    const newDeviceToken = staff.id+'-'+Utilities.getUuid();
    updateStaffPin_(staff.row,pin);
    updateStaffDeviceToken_(staff.row,newDeviceToken);
    sheet.getRange(recoveryRow,5,1,2).setValues([[new Date(),'Completed']]);
    revokeStaffRecoveryRequests_(staff.id);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_RESET_LOG_SHEET)
      .appendRow([new Date(),staff.id,staff.name,'PIN + New Device','YES','YES','Self-service Recovery','Lost device','One-time email recovery link']);
    recordActivity_('Staff',staff.name,staff.id,'CREDENTIAL AND DEVICE RECOVERY','Staff',staff.id,'PIN reset and registered device replaced using a one-time email link.');
    invalidatePublicStaffOptionsCache_();
    return {success:true,staffId:staff.id,deviceToken:newDeviceToken,message:'Your PIN and registered device have been reset. This browser is now your registered device.'};
  } catch (err) {
    recordActivity_('System','Recovery service','','RECOVERY COMPLETION FAILED','Staff Recovery','',String(err.message || err).slice(0,300));
    return {success:false,message:'We could not complete account recovery. Request a new link or contact Admin.'};
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function revokeStaffRecoveryRequests_(staffId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_RECOVERY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  const range = sheet.getRange(2,1,sheet.getLastRow()-1,6);
  const rows = range.getValues();
  rows.forEach((row,index) => {
    if (String(row[1]) === String(staffId) && ['Pending','Sent'].includes(String(row[5]))) {
      sheet.getRange(index+2,6).setValue('Revoked');
    }
  });
}

function sha256Hex_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(value),Utilities.Charset.UTF_8)
    .map(byte => ('0'+(byte&0xff).toString(16)).slice(-2)).join('');
}
