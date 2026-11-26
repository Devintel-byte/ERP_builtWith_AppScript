function recordAttendance(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const settings = getSettings_();
    const staff = getStaffById_(payload.staffId);
    if (!staff) return {success:false, message:'Staff not found or inactive.'};

    const now = new Date();
    const dateKey = Utilities.formatDate(now, settings.TIMEZONE, 'yyyy-MM-dd');
    const timeText = Utilities.formatDate(now, settings.TIMEZONE, 'HH:mm:ss');
    const action = String(payload.action || '');
    if (!['Check In','Check Out','Check Out Overtime'].includes(action)) {
      return {success:false, message:'Select a valid attendance action.'};
    }
    if (action === 'Check Out Overtime') {
      if (!isOvertimeCheckoutAllowed_(now,settings)) {
        return {success:false, message:'Overtime checkout is available after '+settings.CLOSE_TIME+' plus '+settings.OVERTIME_GRACE_MINUTES+' minutes. Please use Check out for now.'};
      }
    }

    // Only Check In has to happen at an approved location. Checking out, including
    // overtime checkout, works from anywhere: people forget to sign out before they
    // leave the office and must not be stuck unable to once they have gone. The
    // position is still recorded when the phone supplies one.
    const latitude = attendanceCoordinate_(payload.latitude);
    const longitude = attendanceCoordinate_(payload.longitude);
    const hasFix = latitude !== null && longitude !== null;
    const accuracy = attendanceCoordinate_(payload.accuracy);
    let locInfo = getMatchedLocation_(hasFix ? latitude : NaN, hasFix ? longitude : NaN, settings, accuracy);
    // On an approved field duty a check-in is optional and is not tied to the office:
    // the person may be in another city or country. The device check still applies,
    // and being on duty never counts as being at the office for re-registering a device.
    const fieldDuty = action === 'Check In' ? getApprovedFieldDutyOn_(staff.id, dateKey) : null;
    const locationRequired = action === 'Check In' && !fieldDuty;
    if (locationRequired && String(staff.requireGps).toUpperCase() === 'YES' && String(settings.REQUIRE_GPS).toUpperCase() === 'YES' && !locInfo.allowed) {
      // Every refusal is kept, with the numbers behind it. "I was in the office and it
      // said I was not" cannot be traced without them: the log shows whether the phone
      // reported a weak fix, whether the office coordinates are wrong, or whether
      // an office row could not be read at all.
      recordActivity_('Staff',staff.name,staff.id,'ATTENDANCE REJECTED - LOCATION','Attendance',dateKey,
        (hasFix ? 'nearest '+locInfo.name+' at '+Math.round(locInfo.distance)+'m; radius '+locInfo.radius+'m + tolerance '+Math.round(locInfo.tolerance)+'m' : 'no location received')+
        '; phone accuracy '+(accuracy === null ? 'not reported' : '+/-'+Math.round(accuracy)+'m')+
        (locInfo.invalidRows ? '; '+locInfo.invalidRows+' active office row(s) have unreadable coordinates' : '')+
        '; '+classifyUserAgent_(payload.userAgent)+'.');
      return {success:false, message:locationRejectionMessage_(hasFix,locInfo,accuracy)};
    }

    let deviceStatus = 'Device Verified';
    let tokenToReturn = '';
    let pinWasCreated = false;
    let pinToCreate = '';

    const submittedPin = String(payload.pin == null ? '' : payload.pin).trim();
    const confirmedPin = String(payload.confirmPin == null ? '' : payload.confirmPin).trim();
    const storedPin = normalizePin_(staff.pin);
    if (!isValidPin_(submittedPin)) return {success:false, message:'PIN must be exactly 4 digits.'};
    // A sample account still holding the placeholder PIN it shipped with is treated
    // as having no PIN yet, so the person's own choice replaces it everywhere.
    const selfPinAllowed = String(settings.ALLOW_SELF_PIN).toUpperCase() === 'YES';
    const mustCreatePin = !staff.pin || (selfPinAllowed && isStartingPin_(staff.id, staff.pin));
    if (mustCreatePin) {
      if (!selfPinAllowed) return {success:false, message:'PIN is not set. Contact Admin.'};
      if (!isValidPin_(confirmedPin) || submittedPin !== confirmedPin) return {success:false, message:'PIN confirmation failed. Enter the same four digits in both fields.'};
      if (isStartingPin_(staff.id, submittedPin)) return {success:false, message:'Choose your own four-digit PIN. The starting PIN that came with your account cannot be kept.'};
      // Only remembered here. It is written after the device and transition checks
      // below, because a request they reject must not be able to set anyone's PIN:
      // the admin ERP no longer checks a device, so a PIN is worth taking.
      pinToCreate = submittedPin;
    } else {
      if (String(settings.REQUIRE_PIN).toUpperCase() === 'YES' && (!isValidPin_(submittedPin) || submittedPin !== storedPin)) {
        return {success:false, message:'Incorrect PIN. Attendance rejected.'};
      }
    }

    const requireDevice = String(staff.requireDevice).toUpperCase() === 'YES' && String(settings.REQUIRE_DEVICE_TOKEN).toUpperCase() === 'YES';
    if (requireDevice) {
      const gpsEnforced = String(staff.requireGps).toUpperCase() === 'YES' && String(settings.REQUIRE_GPS).toUpperCase() === 'YES';
      const deviceCheck = resolveDeviceCheck_(staff.deviceToken, payload.deviceToken, {
        autoRegister: String(settings.AUTO_REGISTER_DEVICE).toUpperCase() === 'YES',
        policy: settings.DEVICE_MISMATCH_POLICY,
        insideApprovedLocation: gpsEnforced && locInfo.allowed === true
      });
      if (deviceCheck.outcome === 'register' || deviceCheck.outcome === 'rebind') {
        const newToken = staff.id + '-' + Utilities.getUuid();
        updateStaffDeviceToken_(staff.row, newToken);
        staff.deviceToken = sha256Hex_(newToken);
        tokenToReturn = newToken;
        deviceStatus = deviceCheck.outcome === 'register' ? 'New Device Registered'
          : (deviceCheck.issue === 'missing' ? 'Device Re-registered (no registration on this browser)'
            : 'Device Replaced (previous registration did not match)');
        if (deviceCheck.outcome === 'rebind') {
          recordActivity_('Staff',staff.name,staff.id,'DEVICE REBOUND AT OFFICE','Staff',staff.id,
            deviceCheck.issue+' token; '+classifyUserAgent_(payload.userAgent)+'; inside '+locInfo.name+' ('+Math.round(locInfo.distance)+'m).');
        }
      } else if (deviceCheck.outcome === 'reject') {
        // Rejections used to leave no trace, so nobody could tell whether the
        // complaints were mostly iPhones, mostly one person, or mostly in-app browsers.
        recordActivity_('Staff',staff.name,staff.id,'ATTENDANCE REJECTED - DEVICE','Attendance',dateKey,
          deviceCheck.issue+' token; '+classifyUserAgent_(payload.userAgent)+'; policy '+settings.DEVICE_MISMATCH_POLICY+'.');
        return {success:false, deviceIssue:deviceCheck.issue, message:deviceRejectionMessage_(deviceCheck.issue)};
      }
    }
    if (pinWasCreated || deviceStatus === 'New Device Registered') invalidatePublicStaffOptionsCache_();

    const attendanceState = getAttendanceStateForDay_(dateKey,staff.id,settings.TIMEZONE);
    const transitionError = getAttendanceTransitionError_(action,attendanceState);
    if (transitionError) {
      return {success:false,message:transitionError};
    }
    if (pinToCreate) {
      updateStaffPin_(staff.row, pinToCreate);
      staff.pin = pinToCreate;
      pinWasCreated = true;
      invalidatePublicStaffOptionsCache_();
    }

    // Lateness is measured against Lagos office hours, which do not apply on field duty.
    const attendanceStatus = fieldDuty ? 'On Field Duty' : getAttendanceStatus_(now, action, settings);
    const penalty = getPenaltySummary_(staff, attendanceStatus, action, now, settings);

    const checkoutType=attendanceCheckoutType_(action,'');
    const checkoutTime=checkoutType ? timeText : '';
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET).appendRow([
      dateKey, staff.id, staff.name, staff.location, action, now, timeText,
      hasFix ? latitude : '', hasFix ? longitude : '', hasFix ? Math.round(locInfo.distance) : '',
      hasFix ? locInfo.name : 'Location not provided',
      !hasFix ? 'Location Not Provided' : locInfo.allowed ? 'Inside Approved Location' : 'Outside Approved Location',
      'QR/Web App', attendanceStatus, penalty.todayPenalty, penalty.monthLatePenalty,
      penalty.monthHalfDayDeduction, penalty.monthAbsenceDeduction, penalty.monthTotalDeduction,
      deviceStatus, deviceTokenFingerprint_(staff.deviceToken), payload.userAgent || '', (pinWasCreated ? 'PIN created. ' : '') + (fieldDuty ? 'Field duty '+fieldDuty.dutyId+' ('+fieldDuty.site+'). ' : '') + (accuracy !== null ? 'GPS accuracy: '+Math.round(accuracy)+'m' : '') + (hasFix && locInfo.allowed && !locInfo.withinRadius ? ' Accepted inside tolerance, '+Math.round(locInfo.distance-locInfo.radius)+'m beyond the radius.' : ''),
      checkoutTime,checkoutType
    ]);
    recordActivity_('Staff',staff.name,staff.id,'ATTENDANCE '+action.toUpperCase(),'Attendance',dateKey,attendanceStatus+' at '+timeText+' from '+locInfo.name);

    const responseMessages=buildAttendanceResponseMessages_(
      action,staff.name,timeText,attendanceStatus,hasFix ? locInfo.name : 'Location not provided',deviceStatus,hasFix ? locInfo.distance : null,pinWasCreated,penalty
    );
    if (fieldDuty) {
      const dutyLine = 'Field duty: '+fieldDuty.site+' (until '+fieldDuty.endDate+'). Office hours and location do not apply.';
      responseMessages.additionalMessage += '\n'+dutyLine;
      responseMessages.message += '\n'+dutyLine;
    }
    return {
      success:true,alert:['Late','Half Day','Absent'].includes(attendanceStatus),
      deviceToken:tokenToReturn,message:responseMessages.message,
      attendanceResult:{
        primaryMessage:responseMessages.primaryMessage,
        additionalMessage:responseMessages.additionalMessage
      }
    };
  } catch (err) {
    return {success:false, message:'Error: '+(err.stack || err.message)};
  } finally {
    lock.releaseLock();
  }
}

/**
 * What a person is told when check-in is refused for location. It states the
 * numbers and, when the phone itself said its fix was weak, says that first,
 * because that is the commonest reason someone standing in the office fails.
 */
function locationRejectionMessage_(hasFix,locInfo,accuracy) {
  if (!hasFix) {
    return 'Attendance rejected. Your phone did not share a location. Allow precise location for this site in your browser settings and try again.';
  }
  const lines = [
    'Attendance rejected. You appear to be outside the approved office area.',
    '',
    'Nearest: '+locInfo.name,
    'Distance: '+Math.round(locInfo.distance)+'m',
    'Allowed: '+locInfo.radius+'m plus '+Math.round(locInfo.tolerance)+'m for GPS error'
  ];
  if (accuracy !== null && accuracy > 150) {
    lines.push('','Your phone reported a weak location fix (+/-'+Math.round(accuracy)+'m). Turn on precise location and Wi-Fi, move near a window or step outside, wait a few seconds, then try again.');
  } else if (locInfo.distance < locInfo.radius + locInfo.tolerance + 1000) {
    lines.push('','If you are in the office, try again in a moment. If it keeps happening, tell Admin the distance shown above so the office location can be checked.');
  } else {
    lines.push('','If you are at the office and this distance is large, tell Admin: the office coordinates may be wrong.');
  }
  if (locInfo.invalidRows) lines.push('','Note: '+locInfo.invalidRows+' office location(s) in the settings sheet could not be read.');
  return lines.join('\n');
}

/** A usable latitude/longitude, or null when the phone sent nothing (blank is not zero). */
function attendanceCoordinate_(value) {
  if (value === null || value === undefined || String(value).trim() === '') return null;
  const number = Number(value);
  return isFinite(number) ? number : null;
}

function buildAttendanceResponseMessages_(action,staffName,timeText,status,location,deviceStatus,distance,pinWasCreated,penalty) {
  const primaryMessage=action+' successful\n\n'+staffName+'\n'+timeText;
  const additionalLines=[
    'Status: '+status,
    'Location: '+location,
    'Device: '+deviceStatus
  ];
  if (distance !== null && distance !== undefined) additionalLines.push('Distance: '+Math.round(Number(distance)||0)+'m');
  if (pinWasCreated) additionalLines.push('','PIN created successfully. Keep it secure.');
  if (action==='Check In') {
    additionalLines.push(
      '',
      'Today Penalty: NGN '+formatMoney_(penalty.todayPenalty),
      'Month Late Penalty: NGN '+formatMoney_(penalty.monthLatePenalty),
      'Month Half-Day Deduction: NGN '+formatMoney_(penalty.monthHalfDayDeduction),
      'Month Absence Deduction: NGN '+formatMoney_(penalty.monthAbsenceDeduction),
      'Current Month Total Deduction: NGN '+formatMoney_(penalty.monthTotalDeduction)
    );
  }
  const additionalMessage=additionalLines.join('\n');
  return {
    primaryMessage:primaryMessage,
    additionalMessage:additionalMessage,
    message:primaryMessage+'\n\n'+additionalMessage
  };
}

function getAttendanceTransitionError_(action,state) {
  if (action === 'Check In') {
    if (state.checkedIn) return 'You already submitted Check In today.';
    if (state.checkedOut) return 'You have already checked out today. Check in is no longer available.';
    return '';
  }
  if (!state.checkedIn) return 'Check in before checking out today.';
  if (state.checkedOut) return 'You already checked out today.';
  return '';
}

function isOvertimeCheckoutAllowed_(now,settings) {
  const currentMinutes = toMinutes_(Utilities.formatDate(now,settings.TIMEZONE,'HH:mm'));
  return isOvertimeCheckoutAllowedAtMinute_(currentMinutes,settings);
}

function isOvertimeCheckoutAllowedAtMinute_(currentMinutes,settings) {
  return Number(currentMinutes) >= toMinutes_(settings.CLOSE_TIME)+Number(settings.OVERTIME_GRACE_MINUTES);
}

function getPublicAttendanceSettings() {
  const settings = getSettings_();
  const startMinutes = toMinutes_(settings.CLOSE_TIME)+settings.OVERTIME_GRACE_MINUTES;
  const hour24 = Math.floor(startMinutes/60);
  const minute = startMinutes%60;
  const hour12 = hour24%12 || 12;
  return {
    overtimeStart:hour12+':'+String(minute).padStart(2,'0')+' '+(hour24<12?'am':'pm')
  };
}

/*************** STAFF REQUISITION PAGE ***************/

/**
 * What a staff member is told when their browser is not the registered one.
 * "Missing" and "mismatch" need different advice: a browser that sends nothing
 * has usually lost its storage, while a wrong token means another device is
 * registered. Both point at the self-service recovery link before Admin.
 */
function deviceRejectionMessage_(issue) {
  if (issue === 'unregistered') return 'No device registered. Contact Admin.';
  if (issue === 'missing') {
    return 'Attendance rejected. This browser has no saved device registration, but a device is already registered to your account.\n\n' +
      'This usually means the link was opened in a different browser, an in-app browser (WhatsApp, Gmail), a private tab, or iPhone Safari cleared its site data.\n\n' +
      'Open the attendance link in the browser you registered with, or tap "Lost your device? Recover your account" to register this one.';
  }
  return 'Attendance rejected. A different device is registered to your account.\n\n' +
    'If you changed phone or browser, tap "Lost your device? Recover your account" below to register this one. Admin can also reset it.';
}
