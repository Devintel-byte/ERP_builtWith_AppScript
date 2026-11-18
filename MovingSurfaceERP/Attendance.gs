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

    let locInfo = getMatchedLocation_(Number(payload.latitude), Number(payload.longitude), settings);
    if (String(staff.requireGps).toUpperCase() === 'YES' && String(settings.REQUIRE_GPS).toUpperCase() === 'YES' && !locInfo.allowed) {
      return {success:false, message:'Attendance rejected. You are outside approved office locations.\n\nNearest: '+locInfo.name+'\nDistance: '+Math.round(locInfo.distance)+'m\nAllowed radius: '+locInfo.radius+'m'};
    }

    let deviceStatus = 'Device Verified';
    let tokenToReturn = '';
    let pinWasCreated = false;

    const submittedPin = String(payload.pin == null ? '' : payload.pin).trim();
    const confirmedPin = String(payload.confirmPin == null ? '' : payload.confirmPin).trim();
    const storedPin = normalizePin_(staff.pin);
    if (!isValidPin_(submittedPin)) return {success:false, message:'PIN must be exactly 4 digits.'};
    if (!staff.pin) {
      if (String(settings.ALLOW_SELF_PIN).toUpperCase() !== 'YES') return {success:false, message:'PIN is not set. Contact Admin.'};
      if (!isValidPin_(confirmedPin) || submittedPin !== confirmedPin) return {success:false, message:'PIN confirmation failed. Enter the same four digits in both fields.'};
      updateStaffPin_(staff.row, submittedPin);
      staff.pin = submittedPin;
      pinWasCreated = true;
    } else {
      if (String(settings.REQUIRE_PIN).toUpperCase() === 'YES' && (!isValidPin_(submittedPin) || submittedPin !== storedPin)) {
        return {success:false, message:'Incorrect PIN. Attendance rejected.'};
      }
    }

    const requireDevice = String(staff.requireDevice).toUpperCase() === 'YES' && String(settings.REQUIRE_DEVICE_TOKEN).toUpperCase() === 'YES';
    if (requireDevice) {
      if (!staff.deviceToken) {
        if (String(settings.AUTO_REGISTER_DEVICE).toUpperCase() !== 'YES') return {success:false, message:'No device registered. Contact Admin.'};
        const newToken = staff.id + '-' + Utilities.getUuid();
        updateStaffDeviceToken_(staff.row, newToken);
        staff.deviceToken = sha256Hex_(newToken);
        tokenToReturn = newToken;
        deviceStatus = 'New Device Registered';
      } else if (!deviceTokenMatches_(staff.deviceToken,payload.deviceToken)) {
        return {success:false, message:'Attendance rejected. This staff is already registered to another device.\n\nIf you changed phone/browser, contact Admin to reset your device.'};
      }
    }
    if (pinWasCreated || deviceStatus === 'New Device Registered') invalidatePublicStaffOptionsCache_();

    const attendanceState = getAttendanceStateForDay_(dateKey,staff.id,settings.TIMEZONE);
    const transitionError = getAttendanceTransitionError_(action,attendanceState);
    if (transitionError) {
      return {success:false,message:transitionError};
    }

    const attendanceStatus = getAttendanceStatus_(now, action, settings);
    const penalty = getPenaltySummary_(staff, attendanceStatus, action, now, settings);

    const checkoutType=attendanceCheckoutType_(action,'');
    const checkoutTime=checkoutType ? timeText : '';
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET).appendRow([
      dateKey, staff.id, staff.name, staff.location, action, now, timeText,
      Number(payload.latitude), Number(payload.longitude), Math.round(locInfo.distance),
      locInfo.name, locInfo.allowed ? 'Inside Approved Location' : 'Outside Approved Location',
      'QR/Web App', attendanceStatus, penalty.todayPenalty, penalty.monthLatePenalty,
      penalty.monthHalfDayDeduction, penalty.monthAbsenceDeduction, penalty.monthTotalDeduction,
      deviceStatus, deviceTokenFingerprint_(staff.deviceToken), payload.userAgent || '', (pinWasCreated ? 'PIN created. ' : '') + (payload.accuracy ? 'GPS accuracy: '+Math.round(payload.accuracy)+'m' : ''),
      checkoutTime,checkoutType
    ]);
    recordActivity_('Staff',staff.name,staff.id,'ATTENDANCE '+action.toUpperCase(),'Attendance',dateKey,attendanceStatus+' at '+timeText+' from '+locInfo.name);

    const responseMessages=buildAttendanceResponseMessages_(
      action,staff.name,timeText,attendanceStatus,locInfo.name,deviceStatus,locInfo.distance,pinWasCreated,penalty
    );
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

function buildAttendanceResponseMessages_(action,staffName,timeText,status,location,deviceStatus,distance,pinWasCreated,penalty) {
  const primaryMessage=action+' successful\n\n'+staffName+'\n'+timeText;
  const additionalLines=[
    'Status: '+status,
    'Location: '+location,
    'Device: '+deviceStatus,
    'Distance: '+Math.round(Number(distance)||0)+'m'
  ];
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
