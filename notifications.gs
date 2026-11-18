// ============================================================================
// NOTIFICATIONS.GS — Email Queue, Alerts, Principal's Digest & Scheduled Triggers
// ============================================================================

/**
 * Helper to send emails based on the current Email Mode setting.
 * Handles "Live", "Redirect", and "Off" modes.
 */
function sendEmailHelper(to, subject, body, options, optionalSettings) {
  var settings = optionalSettings || getSettings();
  var mode = settings["Email Mode"] || "Live";
  var redirectEmail = settings["Redirect Email"] || "";
  var senderName = settings["Email Sender Name"] || "Cathedral Sub App";
  var replyToEmail = settings["Reply To Email"] || "";

  if (!options) {
    options = {};
  }
  options.name = senderName;
  if (replyToEmail && replyToEmail.trim() !== "") {
    options.replyTo = replyToEmail.trim();
  }

  if (mode === "Off") {
    console.log("Email sending is turned Off. Suppressed email to: " + to);
    return "SUPPRESSED";
  }

  if (mode === "Redirect" && redirectEmail) {
    console.log("Email mode is Redirect. Redirecting email originally intended for: " + to + " to: " + redirectEmail);
    // Suppress CCs and BCCs so only the redirect email gets the message
    if (options) {
      if (options.cc) {
        body += "\n\n[Original CC: " + options.cc + "]";
        if (options.htmlBody) options.htmlBody += "<p><em>[Original CC: " + options.cc + "]</em></p>";
        delete options.cc;
      }
      if (options.bcc) {
        delete options.bcc;
      }
    }
    to = redirectEmail;
    subject = "[REDIRECTED] " + subject;
  }

  GmailApp.sendEmail(to, subject, body, options);
  return "SENT";
}

/**
 * Enqueues an email to be sent later by a background trigger.
 * Stores email details in the "Email Queue" sheet.
 */
function enqueueEmail(to, subject, body, options) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Email Queue");

    if (!sheet) {
      sheet = ss.insertSheet("Email Queue");
      sheet.appendRow(["Timestamp", "To", "Subject", "Body", "Options", "Status"]);
    }

    // Look up the name from the Staff Roster
    var recipientName = to; // Default to email
    try {
      var rosterData = getRosterDataCached(ss);
      var targetEmail = String(to).toLowerCase().trim();

      for (var i = 1; i < rosterData.length; i++) {
      if (!rosterData[i] || rosterData[i].length < 3) continue;
      if (rosterData[i] && String(rosterData[i][1]).toLowerCase().trim() === targetEmail) {
          recipientName = String(rosterData[i][0]).trim();
          break;
        }
      }
    } catch(e) {
      // Ignore roster lookup errors, fallback to email
    }

    // Reformat name if it contains a comma (e.g., "Last, First")
    var formattedName = recipientName;
    if (formattedName.indexOf(",") > -1) {
      var parts = formattedName.split(",");
      if (parts.length > 1) formattedName = parts[1].trim() + " " + parts[0].trim();
    }

    var plainGreeting = "Dear " + formattedName + ",\n\n";
    body = plainGreeting + body;

    if (options && options.htmlBody) {
      var htmlGreeting = "<p>Dear " + formattedName + ",</p>";
      options.htmlBody = htmlGreeting + options.htmlBody;
    }

    var timestamp = new Date();
    var optionsStr = options ? JSON.stringify(options) : "{}";

    sheet.appendRow([timestamp, to, subject, body, optionsStr, "Pending"]);
  } catch (e) {
    console.error("Failed to enqueue email: " + e.message);
    // Fallback to sending synchronously if queue fails
    sendEmailHelper(to, subject, body, options);
  }
}

/**
 * Processes the email queue. Runs periodically via a time-driven trigger.
 */
function processEmailQueue() {
  var cache = (typeof CacheService !== "undefined" && CacheService && typeof CacheService.getScriptCache === "function")
    ? CacheService.getScriptCache()
    : null;
  var mutexKey = "process_email_queue_mutex";
  if (cache) {
    try {
      if (cache.get(mutexKey)) {
        console.warn("processEmailQueue already running; skipping overlapping trigger.");
        return;
      }
      cache.put(mutexKey, "1", 120);
    } catch (cErr) {}
  }

  try {
    var ss = getSS();
    try {
      checkAndSendUnfilledCoverageReminders(null, ss);
    } catch (remErr) {
      console.error("Error in checkAndSendUnfilledCoverageReminders: " + remErr.message);
    }

    var sheet = ss.getSheetByName("Email Queue");
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return; // Only headers

    var hasPendingEmails = false;
    var needsScheduleLookup = false;
    for (var i = 1; i < data.length; i++) {
      if (!data[i]) continue;
      if (data[i][5] === "Pending") {
        hasPendingEmails = true;
        var body = String(data[i][3]);
        var optionsStr = String(data[i][4]);
        if (body.indexOf("{{ROOM|") !== -1 || body.indexOf("{{COURSE|") !== -1 ||
            body.indexOf("{{ROOM_URL|") !== -1 || body.indexOf("{{COURSE_URL|") !== -1 ||
            optionsStr.indexOf("{{ROOM|") !== -1 || optionsStr.indexOf("{{COURSE|") !== -1 ||
            optionsStr.indexOf("{{ROOM_URL|") !== -1 || optionsStr.indexOf("{{COURSE_URL|") !== -1) {
          needsScheduleLookup = true;
          break;
        }
      }
    }

    if (!hasPendingEmails) return;

    var statuses = [];
    var settings = getSettings(ss);

    var scheduleLookup = null;
    var snapshottedLookup = null;
    if (needsScheduleLookup) {
      try {
        var v2Compat = (typeof readV2AbsenceDataCompat === "function") ? readV2AbsenceDataCompat(ss) : null;
        if (v2Compat && v2Compat.snapshottedScheduleLookup && Object.keys(v2Compat.snapshottedScheduleLookup).length > 0) {
          snapshottedLookup = v2Compat.snapshottedScheduleLookup;
        }
      } catch (v2Err) {}
    }

    var getFallbackScheduleLookup = function() {
      if (scheduleLookup !== null) return scheduleLookup;
      try {
        var scheduleData = getMasterScheduleData();
        scheduleLookup = buildScheduleLookup(scheduleData, settings["Term ID"] || "3601");
      } catch (err) {
        console.error("Failed to load master schedule for email queue processing: " + err.message);
        scheduleLookup = {};
      }
      return scheduleLookup;
    };

    for (let i = 0; i < data.length; i++) {
        statuses.push([data[i][5]]);
    }

    // Helper to replace placeholders using regex (checks snapshotted Coverage_Assignments first, then falls back to Master Schedule Cache)
    var replacePlaceholders = function(text, targetIsoDate) {
      if (!text || typeof text !== 'string') return text;
      var termForDate = (targetIsoDate && typeof resolveTermIdForDate === 'function')
        ? String(resolveTermIdForDate(targetIsoDate, settings) || "").toLowerCase().trim()
        : "";
      return text.replace(/\{\{(ROOM|COURSE|ROOM_URL|COURSE_URL)\|([^}]+)\}\}/g, function(match, type, joinKey) {
        var val = "No Class Assigned";
        var normKey = String(joinKey || "").toLowerCase().trim();
        var entry = null;
        if (snapshottedLookup && snapshottedLookup[normKey]) {
          entry = snapshottedLookup[normKey];
        }
        if (!entry) {
          var masterLookup = getFallbackScheduleLookup();
          if (masterLookup) {
            if (termForDate && masterLookup[normKey + "-" + termForDate]) {
              entry = masterLookup[normKey + "-" + termForDate];
            } else {
              entry = masterLookup[normKey];
            }
          }
        }
        if (entry) {
           var isRoom = (type === 'ROOM' || type === 'ROOM_URL');
           val = isRoom ? (entry.room || val) : (entry.course || val);
        }
        if (type === 'ROOM_URL' || type === 'COURSE_URL') {
           return encodeURIComponent(val);
        }
        return val;
      });
    }

    var MAX_EMAILS_PER_RUN = 25;
    var processedCount = 0;

    for (var j = 1; j < data.length; j++) {
    if (!data[j]) continue;
      if (data[j][5] === "Pending") {
        if (processedCount >= MAX_EMAILS_PER_RUN) {
          break; // Leave remaining Pending emails for the next 1-minute trigger run
        }
        processedCount++;
        var to = data[j][1];
        var subject = data[j][2];
        let body = data[j][3];
        let optionsStr = data[j][4];
        var options = {};

        try {
          options = JSON.parse(optionsStr);
        } catch (e) {
          console.error("Failed to parse options for queued email: " + e.message);
        }

        var calSync = options && options.calendarSync ? options.calendarSync : null;
        var targetIsoDate = (calSync && calSync.details && calSync.details.isoDate) ? calSync.details.isoDate : "";
        if (options && options.calendarSync) {
          delete options.calendarSync;
        }

        if (needsScheduleLookup) {
          body = replacePlaceholders(String(body), targetIsoDate);
          if (options.htmlBody) {
            options.htmlBody = replacePlaceholders(String(options.htmlBody), targetIsoDate);
          }
          if (calSync && calSync.details) {
            if (calSync.details.room) calSync.details.room = replacePlaceholders(String(calSync.details.room), targetIsoDate);
            if (calSync.details.course) calSync.details.course = replacePlaceholders(String(calSync.details.course), targetIsoDate);
          }
        }

        if (calSync && calSync.subEmail && calSync.details) {
          try {
            syncSubDutyCalendarEvent(calSync.subEmail, calSync.type, calSync.details, settings);
          } catch (calErr) {
            console.warn("Auto calendar sync error for " + calSync.subEmail + ": " + calErr.message);
          }
        }

        try {
          var result = sendEmailHelper(to, subject, body, options, settings);
          if (result === "SUPPRESSED") {
            statuses[j][0] = "Suppressed (Off)";
          } else {
            statuses[j][0] = "Sent";
          }
        } catch (e) {
          console.error("Failed to send queued email to " + to + ": " + e.message);
          statuses[j][0] = "Failed: " + e.message;
        }
      }
    }

    // Acquire ScriptLock ONLY for the brief sheet mutation phase (updating statuses & deleting sent rows)
    var lock = LockService.getScriptLock();
    var hasLock = lock.tryLock(30000);
    if (!hasLock) {
      console.warn("Could not obtain lock for processEmailQueue sheet cleanup");
    }
    try {
      // Batch update statuses
      sheet.getRange(1, 6, statuses.length, 1).setValues(statuses);

      // Cleanup completed (Sent or Suppressed) emails while keeping Failed rows for inspection
      var rowsToDelete = [];
      for (let k = statuses.length - 1; k >= 1; k--) {
        var status = String(statuses[k][0] || "");
        if (status === "Sent" || status === "Suppressed (Off)") {
           rowsToDelete.push(k + 1);
        }
      }

      // Delete from bottom up in contiguous batches
      if (rowsToDelete.length > 0) {
        var startRow = rowsToDelete[0];
        var numRows = 1;

        for (let j = 1; j < rowsToDelete.length; j++) {
          if (rowsToDelete[j] === startRow - numRows) {
            numRows++;
          } else {
            sheet.deleteRows(startRow - numRows + 1, numRows);
            startRow = rowsToDelete[j];
            numRows = 1;
          }
        }
        sheet.deleteRows(startRow - numRows + 1, numRows);
      }
    } finally {
      if (hasLock) lock.releaseLock();
    }

  } catch (e) {
    console.error("Error in processEmailQueue: " + e.message);
  } finally {
    if (cache) {
      try { cache.remove(mutexKey); } catch (cErr2) {}
    }
  }
}

/**
 * Returns the status of the Email Queue for the Admin Settings / Data Management view.
 */
function getEmailQueueStatus(clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var sheet = ss.getSheetByName("Email Queue");
    if (!sheet) {
      return { pendingCount: 0, failedCount: 0, items: [] };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { pendingCount: 0, failedCount: 0, items: [] };
    }

    var pendingCount = 0;
    var failedCount = 0;
    var items = [];
    var tz = Session.getScriptTimeZone();

    for (var i = 1; i < data.length; i++) {
      if (!data[i]) continue;
      var status = String(data[i][5] || "").trim();
      var isPending = (status === "Pending");
      var isFailed = (status.indexOf("Failed") === 0);

      if (isPending) pendingCount++;
      if (isFailed) failedCount++;

      if ((isPending || isFailed) && items.length < 50) {
        var tsRaw = data[i][0];
        var tsFormatted = (tsRaw instanceof Date && !isNaN(tsRaw.getTime()))
          ? Utilities.formatDate(tsRaw, tz, "yyyy-MM-dd HH:mm:ss")
          : String(tsRaw || "");
        items.push({
          rowNumber: i + 1,
          timestamp: tsFormatted,
          to: String(data[i][1] || ""),
          subject: String(data[i][2] || ""),
          status: status
        });
      }
    }

    return {
      pendingCount: pendingCount,
      failedCount: failedCount,
      items: items
    };
  } catch (err) {
    notifyAdminOfError("getEmailQueueStatus", err);
    throw new Error("Failed to load email queue status: " + err.message);
  }
}

/**
 * Resets all Failed rows in the Email Queue back to Pending and processes the queue immediately.
 */
function retryFailedEmails(clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    notifyAdminOfError("retryFailedEmails_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  var retriedCount = 0;
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var sheet = ss.getSheetByName("Email Queue");
    if (!sheet) {
      return { success: true, retriedCount: 0, status: { pendingCount: 0, failedCount: 0, items: [] } };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, retriedCount: 0, status: { pendingCount: 0, failedCount: 0, items: [] } };
    }

    var statuses = [];
    for (var i = 0; i < data.length; i++) {
      var st = data[i] ? String(data[i][5] || "") : "";
      if (i > 0 && st.indexOf("Failed") === 0) {
        statuses.push(["Pending"]);
        retriedCount++;
      } else {
        statuses.push([st]);
      }
    }

    if (retriedCount > 0) {
      sheet.getRange(1, 6, statuses.length, 1).setValues(statuses);
      logAuditAction("EMAIL_QUEUE_RETRY", "Email Queue", "Reset " + retriedCount + " failed email(s) to Pending", clientEmail);
    }
  } catch (err) {
    notifyAdminOfError("retryFailedEmails", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }

  if (retriedCount > 0) {
    processEmailQueue();
  }
  return {
    success: true,
    retriedCount: retriedCount,
    status: getEmailQueueStatus(clientEmail)
  };
}

/**
 * Deletes all Failed rows from the Email Queue sheet.
 */
function clearFailedEmails(clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    notifyAdminOfError("clearFailedEmails_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var sheet = ss.getSheetByName("Email Queue");
    if (!sheet) {
      return { success: true, deletedCount: 0, status: { pendingCount: 0, failedCount: 0, items: [] } };
    }

    var data = sheet.getDataRange().getValues();
    var rowsToDelete = [];
    for (var k = data.length - 1; k >= 1; k--) {
      if (!data[k]) continue;
      var status = String(data[k][5] || "");
      if (status.indexOf("Failed") === 0) {
        rowsToDelete.push(k + 1);
      }
    }

    for (var d = 0; d < rowsToDelete.length; d++) {
      sheet.deleteRow(rowsToDelete[d]);
    }

    if (rowsToDelete.length > 0) {
      logAuditAction("EMAIL_QUEUE_CLEAR_FAILED", "Email Queue", "Cleared " + rowsToDelete.length + " failed email(s)", clientEmail);
    }

    return {
      success: true,
      deletedCount: rowsToDelete.length
    };
  } catch (err) {
    notifyAdminOfError("clearFailedEmails", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Triggers processEmailQueue immediately from the Settings UI and returns updated status.
 */
function runEmailQueueNow(clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    processEmailQueue();
    return {
      success: true,
      status: getEmailQueueStatus(clientEmail)
    };
  } catch (err) {
    notifyAdminOfError("runEmailQueueNow", err);
    return { success: false, error: err.message };
  }
}

/**
 * Sets up a time-driven trigger to process the email queue every 1 minute.
 * Note: This function is meant to be run manually from the Apps Script editor
 * exactly once during the initial setup of the application.
 */
function setupEmailQueueTrigger() {
  // First, remove any existing triggers for this function to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processEmailQueue') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create a new trigger to run every 1 minute
  ScriptApp.newTrigger('processEmailQueue')
    .timeBased()
    .everyMinutes(1)
    .create();

  console.log("Email Queue trigger created successfully to run every 1 minute.");
}

/**
 * Sets up a time-driven trigger to send the Principal's Digest email every Friday at 12 PM.
 * Note: This function is meant to be run manually from the Apps Script editor
 * exactly once during the initial setup of the application.
 */
function setupPrincipalsDigestTrigger() {
  // First, remove any existing triggers for this function to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runPrincipalsDigestWeekly') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create a new trigger to run every Friday at 12 PM
  ScriptApp.newTrigger('runPrincipalsDigestWeekly')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(12)
    .create();

  console.log("Principal's Digest trigger created successfully to run every Friday around 12 PM.");
}

function getCoordinatorEmail(ss) {
  ss = ss || getSS();
  var data = getRosterDataCached(ss);
  if (!data || data.length === 0) return null;

  for (var i = 1; i < data.length; i++) {
  if (!data[i]) continue;
    // Column 2 is Role, Column 1 is Email
    var roles = String(data[i][2] || "").toLowerCase().split(",").map(function(r) { return r.trim(); });
    if (roles.indexOf("sub coordinator") !== -1) return String(data[i][1]).trim();
  }
  return null;
}

/**
 * Helper to check if a substitute (identified by email or name) has the out-of-building
 * "Substitute" role in the Staff Roster.
 */
function isExternalSubstituteByEmailOrName_(ss, emailOrName, optionalRosterData) {
  if (!emailOrName) return false;
  var target = String(emailOrName).trim().toLowerCase();
  if (!target || target === "no sub needed") return false;
  var cleanNameTarget = target.replace(/\s+-\s+.*$/, "").trim();

  var rosterData = optionalRosterData || getRosterDataCached(ss || getSS());
  if (!rosterData || rosterData.length === 0) return false;

  for (var i = 1; i < rosterData.length; i++) {
    if (!rosterData[i]) continue;
    var rName = String(rosterData[i][0] || "").trim().toLowerCase();
    var rEmail = String(rosterData[i][1] || "").trim().toLowerCase();
    if ((rEmail && rEmail === target) || (rName && (rName === target || rName === cleanNameTarget))) {
      var roles = String(rosterData[i][2] || "").toLowerCase().split(",").map(function(r) { return r.trim(); });
      return roles.indexOf("substitute") !== -1;
    }
  }
  return false;
}

/**
 * Helper to check if an absence date is within 48 hours of the given timestamp (default: now).
 * Uses the earliest requested period start time on the absence date when available (falling back
 * to 07:50 AM on the absence date). Returns true if the absence is in the future (or earlier today)
 * and starts within 48 hours.
 */
function isAbsenceWithin48Hours_(absenceDateVal, periodsRaw, nowOverride, optionalSS) {
  try {
    var now = (nowOverride instanceof Date && !isNaN(nowOverride.getTime())) ? nowOverride : new Date();
    var ss = optionalSS || getSS();
    var isoDate = typeof getStandardizedDateString === "function"
      ? getStandardizedDateString(absenceDateVal)
      : "";
    if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) {
      var parsed = typeof parseLocalDateSafe === "function" ? parseLocalDateSafe(absenceDateVal) : new Date(absenceDateVal);
      if (!parsed || isNaN(parsed.getTime())) return false;
      var tz = (typeof Session !== "undefined" && Session && typeof Session.getScriptTimeZone === "function")
        ? Session.getScriptTimeZone()
        : "America/Indiana/Indianapolis";
      isoDate = Utilities.formatDate(parsed, tz, "yyyy-MM-dd");
    }

    var dParts = isoDate.split("-");
    var yr = parseInt(dParts[0], 10);
    var mo = parseInt(dParts[1], 10) - 1;
    var dy = parseInt(dParts[2], 10);

    // Default start of school day: 7:50 AM on the absence date
    var targetStartObj = new Date(yr, mo, dy, 7, 50, 0);
    var endOfAbsenceDay = new Date(yr, mo, dy, 23, 59, 59);

    if (periodsRaw && typeof getPeriodTimeRangeForDate === "function") {
      var settings = typeof getSettings === "function" ? getSettings(ss) : {};
      var dateColors = typeof getDateColorsCached === "function" ? getDateColorsCached(ss) : null;
      var normPeriods = (typeof normalizePeriodsListValue === "function" ? normalizePeriodsListValue(periodsRaw) : String(periodsRaw))
        .split(",")
        .map(function(p) { return parsePeriod(p); })
        .filter(function(p) { return Boolean(p); });

      var earliestStart = null;
      for (var i = 0; i < normPeriods.length; i++) {
        var pRange = getPeriodTimeRangeForDate(isoDate, normPeriods[i], dateColors, settings);
        if (pRange && pRange.startDateObj instanceof Date && !isNaN(pRange.startDateObj.getTime())) {
          if (!earliestStart || pRange.startDateObj.getTime() < earliestStart.getTime()) {
            earliestStart = pRange.startDateObj;
          }
        }
      }
      if (earliestStart) {
        targetStartObj = earliestStart;
      }
    }

    // If the absence day has already completely passed before today, do not treat as an upcoming 48h cancellation
    if (now.getTime() > endOfAbsenceDay.getTime()) {
      return false;
    }

    var diffHours = (targetStartObj.getTime() - now.getTime()) / (1000 * 60 * 60);
    return diffHours <= 48;
  } catch (e) {
    console.warn("Error in isAbsenceWithin48Hours_: " + e.message);
    return false;
  }
}

/**
 * Helper to get a list of emails by role.
 */
function getEmailsByRole(sheetSS, roleStr) {
  var ss = sheetSS || getSS();
  var data = getRosterDataCached(ss);
  if (!data || data.length === 0) return [];

  var emails = [];
  var targetRole = roleStr.toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
  if (!data[i]) continue;
    var roles = String(data[i][2]).toLowerCase().split(",").map(function(r) { return r.trim(); });
    if (roles.indexOf(targetRole) !== -1) {
       if (data[i]) emails.push(String(data[i][1]).trim());
    }
  }
  return emails;
}

/**
 * Helper to send an email for urgent coverage requests.
 */
function sendUrgentCoverageEmail(ss, teacherName, formData, instructions) {
  var coordinatorEmail = getCoordinatorEmail(ss);
  if (coordinatorEmail) {
    var subject = "URGENT COVERAGE NEEDED: " + teacherName;
    var settings = getSettings();
    var appUrl = settings["App URL"] || DEFAULT_APP_URL;

    var displayPeriods = formatPeriodsListDisplay(formData.periods);
    var body = "An urgent absence request has been submitted requiring immediate attention.\n\n" +
               "Teacher: " + teacherName + "\n" +
               "Date Needed: " + formData.date + "\n" +
               "Periods: " + displayPeriods + "\n" +
               "Reason: " + formData.reason + "\n\n" +
               "Instructions: " + (instructions ? instructions : "None") + "\n\n" +
               "Please log into the Cathedral Sub App to assign a sub: " + appUrl;

    var htmlBody = "<p>An urgent absence request has been submitted requiring immediate attention.</p>" +
                   "<ul>" +
                   "<li><strong>Teacher:</strong> " + teacherName + "</li>" +
                   "<li><strong>Date Needed:</strong> " + formData.date + "</li>" +
                   "<li><strong>Periods:</strong> " + displayPeriods + "</li>" +
                   "<li><strong>Reason:</strong> " + formData.reason + "</li>" +
                   "</ul>" +
                   "<p><strong>Instructions:</strong> " + (instructions ? instructions : "None") + "</p>" +
                   "<p>Please log into the <a href='" + appUrl + "'>Cathedral Sub App</a> to assign a sub.</p>";

    enqueueEmail(coordinatorEmail, subject, body, { htmlBody: htmlBody });
  }
}

var _inMemoryUnfilledReminderSentMap = {};

/**
 * Checks whether the 15-minute unfilled coverage reminder has already been processed
 * for a given `(isoDate, period)` pair today.
 */
function hasPeriodReminderBeenProcessed_(isoDate, period, cache) {
  var pKey = String(period);
  var cacheKey = "UNFILLED_REMINDER_" + isoDate + "_P" + pKey;
  if (_inMemoryUnfilledReminderSentMap[cacheKey]) return true;

  if (cache && typeof cache.get === "function") {
    try {
      if (cache.get(cacheKey)) {
        _inMemoryUnfilledReminderSentMap[cacheKey] = true;
        return true;
      }
    } catch (e) {}
  }

  if (typeof PropertiesService !== "undefined" && PropertiesService && typeof PropertiesService.getScriptProperties === "function") {
    try {
      var props = PropertiesService.getScriptProperties();
      if (props && typeof props.getProperty === "function") {
        var raw = props.getProperty("UNFILLED_REMINDER_SENT_MAP");
        if (raw) {
          var parsed = JSON.parse(raw);
          if (parsed && parsed.date === isoDate && parsed.periods && parsed.periods[pKey]) {
            _inMemoryUnfilledReminderSentMap[cacheKey] = true;
            return true;
          }
        }
      }
    } catch (e2) {}
  }
  return false;
}

/**
 * Marks a `(isoDate, period)` pair as processed so its reminder is not sent again today.
 */
function markPeriodReminderProcessed_(isoDate, period, cache) {
  var pKey = String(period);
  var cacheKey = "UNFILLED_REMINDER_" + isoDate + "_P" + pKey;
  _inMemoryUnfilledReminderSentMap[cacheKey] = true;

  if (cache && typeof cache.put === "function") {
    try {
      cache.put(cacheKey, "1", 21600); // 6 hours
    } catch (e) {}
  }

  if (typeof PropertiesService !== "undefined" && PropertiesService && typeof PropertiesService.getScriptProperties === "function") {
    try {
      var props = PropertiesService.getScriptProperties();
      if (props && typeof props.getProperty === "function" && typeof props.setProperty === "function") {
        var raw = props.getProperty("UNFILLED_REMINDER_SENT_MAP");
        var parsed = null;
        if (raw) {
          try { parsed = JSON.parse(raw); } catch (je) { parsed = null; }
        }
        if (!parsed || parsed.date !== isoDate || typeof parsed.periods !== "object" || !parsed.periods) {
          parsed = { date: isoDate, periods: {} };
        }
        parsed.periods[pKey] = true;
        props.setProperty("UNFILLED_REMINDER_SENT_MAP", JSON.stringify(parsed));
      }
    } catch (e2) {}
  }
}

/**
 * Clears in-memory deduplication state for unfilled coverage reminders (used in unit tests).
 */
function resetUnfilledCoverageReminderState_() {
  _inMemoryUnfilledReminderSentMap = {};
}

/**
 * Checks whether any class period is starting within the configured lead time (default 15 minutes)
 * and sends a consolidated reminder email to all Sub Coordinators (falling back to Admins) if
 * there are still unfilled sub coverages for that period.
 *
 * Designed to run inside the 1-minute `processEmailQueue` trigger with zero Google Sheet reads
 * on 99% of ticks (only reads sheet data at the exact minute a period crosses the lead-time threshold).
 */
function checkAndSendUnfilledCoverageReminders(nowDateOverride, optionalSS) {
  var now = (nowDateOverride instanceof Date && !isNaN(nowDateOverride.getTime()))
    ? nowDateOverride
    : new Date();

  // Fast exit outside school days (Mon-Fri) or outside school hours (6 AM - 4 PM)
  var dayOfWeek = now.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return { checked: false, reason: "weekend" };
  var hour = now.getHours();
  if (hour < 6 || hour >= 16) return { checked: false, reason: "outside_school_hours" };

  var ss = optionalSS || getSS();
  var settings = getSettings(ss);
  var leadMinutes = parseInt(
    settings["Unfilled Coverage Reminder Lead Time (Minutes)"] !== undefined
      ? settings["Unfilled Coverage Reminder Lead Time (Minutes)"]
      : "15",
    10
  );
  if (isNaN(leadMinutes) || leadMinutes <= 0) {
    return { checked: false, reason: "disabled" };
  }

  var tz = (typeof Session !== "undefined" && Session && typeof Session.getScriptTimeZone === "function")
    ? Session.getScriptTimeZone()
    : "America/Indiana/Indianapolis";
  var todayIso = (typeof Utilities !== "undefined" && Utilities && typeof Utilities.formatDate === "function")
    ? Utilities.formatDate(now, tz, "yyyy-MM-dd")
    : (now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0") + "-" + String(now.getDate()).padStart(2, "0"));

  var dateColors = (typeof getDateColorsCached === "function") ? getDateColorsCached(ss) : {};
  var todayColor = String((dateColors && dateColors[todayIso]) || "").trim();
  if (todayColor.toLowerCase() === "no school") {
    return { checked: false, reason: "no_school" };
  }

  var cache = (typeof CacheService !== "undefined" && CacheService && typeof CacheService.getScriptCache === "function")
    ? CacheService.getScriptCache()
    : null;

  var candidatePeriods = ['0', '1', '10', '2', '3', '4', '5', '6', '7', '8'];
  var duePeriods = [];

  for (var cIdx = 0; cIdx < candidatePeriods.length; cIdx++) {
    var p = candidatePeriods[cIdx];
    var pRange = (typeof getPeriodTimeRangeForDate === "function")
      ? getPeriodTimeRangeForDate(todayIso, p, dateColors, settings)
      : null;
    if (!pRange || !pRange.startDateObj) continue;

    var diffMinutes = (pRange.startDateObj.getTime() - now.getTime()) / 60000;
    if (diffMinutes > 0 && diffMinutes <= leadMinutes) {
      if (!hasPeriodReminderBeenProcessed_(todayIso, p, cache)) {
        duePeriods.push({
          period: p,
          periodRange: pRange,
          diffMinutes: diffMinutes
        });
      }
    }
  }

  if (duePeriods.length === 0) {
    return { checked: true, duePeriods: 0, emailsSent: 0 };
  }

  // Mark due periods processed immediately so overlapping triggers or empty-unfilled checks don't re-read sheets every minute
  for (var mIdx = 0; mIdx < duePeriods.length; mIdx++) {
    markPeriodReminderProcessed_(todayIso, duePeriods[mIdx].period, cache);
  }

  // Load absence data and roster data now that at least one period has crossed the reminder threshold
  var v2Compat = (typeof readV2AbsenceDataCompat === "function") ? readV2AbsenceDataCompat(ss) : null;
  var absenceData = null;
  var snapshottedScheduleLookup = null;
  if (v2Compat && v2Compat.absenceData) {
    absenceData = v2Compat.absenceData;
    snapshottedScheduleLookup = v2Compat.snapshottedScheduleLookup || null;
  } else {
    var mainSheet = ss.getSheetByName("Absence Requests");
    if (!mainSheet) return { checked: true, duePeriods: duePeriods.length, emailsSent: 0 };
    absenceData = mainSheet.getDataRange().getValues();
  }

  var rosterData = (typeof getRosterDataCached === "function") ? getRosterDataCached(ss) : [];
  var nameLookup = (typeof buildNameLookup === "function") ? buildNameLookup(rosterData) : {};
  var fallbackScheduleLookup = null;
  var getFallbackSchedule = function() {
    if (fallbackScheduleLookup !== null) return fallbackScheduleLookup;
    try {
      var schedData = (typeof getMasterScheduleData === "function") ? getMasterScheduleData() : [];
      fallbackScheduleLookup = (typeof buildScheduleLookup === "function")
        ? buildScheduleLookup(schedData, settings["Term ID"] || "3601")
        : {};
    } catch (sErr) {
      fallbackScheduleLookup = {};
    }
    return fallbackScheduleLookup;
  };

  var formatTeacherDisplayName = function(rawName) {
    var str = String(rawName || "").trim();
    if (str.indexOf(",") !== -1) {
      var parts = str.split(",");
      if (parts.length > 1) return parts[1].trim() + " " + parts[0].trim();
    }
    return str;
  };

  // Collect all active absences for today
  var todayActiveRows = [];
  for (var i = 1; i < absenceData.length; i++) {
    var row = absenceData[i];
    if (!row || row.length < 5) continue;
    var status = String(row[19] || "Active").trim();
    if (status.toLowerCase() === "canceled" || status.toLowerCase() === "cancelled") continue;

    var rowIsoDate = (typeof getStandardizedDateString === "function")
      ? getStandardizedDateString(row[3])
      : String(row[3] || "").trim();
    if (rowIsoDate === todayIso) {
      todayActiveRows.push(row);
    }
  }

  var recipients = getEmailsByRole(ss, "sub coordinator");
  if (!recipients || recipients.length === 0) {
    recipients = getEmailsByRole(ss, "admin");
  }
  // Deduplicate and clean recipient emails
  var cleanRecipients = [];
  if (recipients && recipients.length > 0) {
    for (var rIdx = 0; rIdx < recipients.length; rIdx++) {
      var em = String(recipients[rIdx] || "").trim();
      if (em && cleanRecipients.indexOf(em) === -1) {
        cleanRecipients.push(em);
      }
    }
  }

  if (cleanRecipients.length === 0) {
    return { checked: true, duePeriods: duePeriods.length, emailsSent: 0, reason: "no_recipients" };
  }

  var appUrl = settings["App URL"] || (typeof DEFAULT_APP_URL !== "undefined" ? DEFAULT_APP_URL : "");
  var emailsSent = 0;

  for (var dIdx = 0; dIdx < duePeriods.length; dIdx++) {
    var dueObj = duePeriods[dIdx];
    var targetP = String(dueObj.period);
    var pRangeObj = dueObj.periodRange;
    var minsLeft = Math.max(1, Math.round(dueObj.diffMinutes));

    var unfilledList = [];
    var absentEmailsInPeriod = {};
    var busySubNamesInPeriod = {};

    for (var tIdx = 0; tIdx < todayActiveRows.length; tIdx++) {
      var aRow = todayActiveRows[tIdx];
      var tEmail = String(aRow[2] || "").toLowerCase().trim();
      var reqPeriods = String(aRow[4] || "")
        .split(",")
        .map(function(tok) { return parsePeriod(tok); })
        .filter(function(tok) { return Boolean(tok); });

      if (reqPeriods.indexOf(targetP) === -1) continue;

      if (tEmail) absentEmailsInPeriod[tEmail] = true;

      var subColIdx = getSubColumnIndex(targetP) - 1;
      var assignedSub = (subColIdx >= 0 && subColIdx < aRow.length)
        ? String(aRow[subColIdx] || "").trim()
        : "";

      if (assignedSub !== "") {
        if (assignedSub.toLowerCase() !== "no sub needed") {
          busySubNamesInPeriod[assignedSub.toLowerCase()] = true;
        }
        continue; // Already covered or marked No Sub Needed
      }

      // Unfilled coverage in targetP!
      var rawTeacherName = nameLookup[tEmail] || tEmail;
      var displayTeacherName = formatTeacherDisplayName(rawTeacherName);

      // Resolve Course & Room
      var schedEntry = null;
      if (snapshottedScheduleLookup && typeof lookupTeacherScheduleInfo === "function") {
        schedEntry = lookupTeacherScheduleInfo(snapshottedScheduleLookup, tEmail, targetP, todayIso, settings);
      }
      if (!schedEntry || (!schedEntry.course && !schedEntry.room) ||
          (schedEntry.course === "No Class Assigned" && schedEntry.room === "No Class Assigned")) {
        var masterSched = getFallbackSchedule();
        if (masterSched && typeof lookupTeacherScheduleInfo === "function") {
          var fallbackEntry = lookupTeacherScheduleInfo(masterSched, tEmail, targetP, todayIso, settings);
          if (fallbackEntry) schedEntry = fallbackEntry;
        }
      }

      var courseStr = (schedEntry && schedEntry.course) ? String(schedEntry.course) : "Class Coverage Needed";
      var roomStr = (schedEntry && schedEntry.room) ? String(schedEntry.room) : "TBD";

      var rawInstructions = String(aRow[8] || "").trim();
      var cleanInstructions = rawInstructions.replace(/^\[HR Docs Provided\]\s*/i, "").trim();
      var subPlanUrl = (typeof extractSubPlanUrl === "function") ? extractSubPlanUrl(aRow) : String(aRow[20] || "").trim();

      unfilledList.push({
        absenceId: String(aRow[0] || ""),
        teacherEmail: tEmail,
        teacherName: displayTeacherName,
        course: courseStr,
        room: roomStr,
        reason: String(aRow[5] || "").trim(),
        instructions: cleanInstructions,
        subPlanUrl: subPlanUrl
      });
    }

    if (unfilledList.length === 0) {
      continue;
    }

    // Find available staff who have Sub Duty in targetP today
    var availableDutyTeachers = [];
    for (var sIdx = 1; sIdx < rosterData.length; sIdx++) {
      var sRow = rosterData[sIdx];
      if (!sRow || sRow.length < 4) continue;
      var sNameRaw = String(sRow[0] || "").trim();
      var sEmail = String(sRow[1] || "").toLowerCase().trim();
      var sDutyRaw = String(sRow[3] || "").trim();
      if (!sNameRaw || !sDutyRaw) continue;

      var resolvedDuty = (typeof resolveStaffDutyForDate === "function")
        ? resolveStaffDutyForDate(sDutyRaw, todayIso, settings)
        : sDutyRaw;
      if (!resolvedDuty) continue;

      var dutyPeriods = String(resolvedDuty)
        .split(",")
        .map(function(d) { return parsePeriod(d); })
        .filter(function(d) { return Boolean(d); });

      if (dutyPeriods.indexOf(targetP) === -1) continue;

      // Exclude if absent during this period or already assigned to cover another class this period
      if (sEmail && absentEmailsInPeriod[sEmail]) continue;
      if (busySubNamesInPeriod[sNameRaw.toLowerCase()]) continue;
      var formattedStaffName = formatTeacherDisplayName(sNameRaw);
      if (busySubNamesInPeriod[formattedStaffName.toLowerCase()]) continue;

      availableDutyTeachers.push(formattedStaffName);
    }

    var periodLabel = formatPeriodLabelDisplay(targetP);
    var timeRangeDisplay = (pRangeObj && pRangeObj.displayRange) ? pRangeObj.displayRange : "";
    var dayColorDisplay = (pRangeObj && pRangeObj.dayColor) ? (pRangeObj.dayColor + " Day") : "";
    var subject = "⚠️ Unfilled Sub Coverage Starting in " + minsLeft + " Min: " + periodLabel +
                  (timeRangeDisplay ? " (" + timeRangeDisplay + ")" : "");

    var plainLines = [
      "Unfilled Sub Coverage Alert — " + periodLabel + (timeRangeDisplay ? " (" + timeRangeDisplay + ")" : "") + (dayColorDisplay ? " [" + dayColorDisplay + "]" : ""),
      "",
      "The following " + (unfilledList.length === 1 ? "class is" : (unfilledList.length + " classes are")) + " scheduled to begin in approximately " + minsLeft + " minutes and still need a substitute:",
      ""
    ];

    for (var uIdx = 0; uIdx < unfilledList.length; uIdx++) {
      var item = unfilledList[uIdx];
      plainLines.push(
        (uIdx + 1) + ". " + item.teacherName + " — " + item.course + " (Room: " + item.room + ")" +
        (item.instructions ? "\n   Instructions: " + item.instructions : "") +
        (item.subPlanUrl ? "\n   Sub Plan: " + item.subPlanUrl : "")
      );
    }

    plainLines.push("");
    if (availableDutyTeachers.length > 0) {
      plainLines.push("Available Teachers with " + periodLabel + " Sub Duty Today:");
      plainLines.push("- " + availableDutyTeachers.join("\n- "));
    } else {
      plainLines.push("Available Teachers with " + periodLabel + " Sub Duty Today: None available (all duty teachers are already assigned or absent)");
    }

    if (appUrl) {
      plainLines.push("");
      plainLines.push("Assign a substitute in the Cathedral Sub App: " + appUrl);
    }

    var plainBody = plainLines.join("\n");

    var htmlRows = "";
    for (var hIdx = 0; hIdx < unfilledList.length; hIdx++) {
      var u = unfilledList[hIdx];
      var instrHtml = u.instructions ? escapeHtmlServer_(u.instructions) : "<span style='color:#9ca3af;font-style:italic;'>None</span>";
      if (u.subPlanUrl) {
        instrHtml += "<div style='margin-top:4px;'><a href='" + escapeHtmlServer_(u.subPlanUrl) + "' target='_blank' style='color:#00843D;font-weight:bold;text-decoration:underline;'>📄 View Sub Plan</a></div>";
      }
      htmlRows += "<tr>" +
        "<td style='padding:10px 12px;border-bottom:1px solid #e5e7eb;font-weight:bold;color:#111827;'>" + escapeHtmlServer_(u.teacherName) + "</td>" +
        "<td style='padding:10px 12px;border-bottom:1px solid #e5e7eb;color:#374151;'>" + escapeHtmlServer_(u.course) + "</td>" +
        "<td style='padding:10px 12px;border-bottom:1px solid #e5e7eb;color:#374151;font-weight:600;'>" + escapeHtmlServer_(u.room) + "</td>" +
        "<td style='padding:10px 12px;border-bottom:1px solid #e5e7eb;color:#4b5563;font-size:13px;'>" + instrHtml + "</td>" +
      "</tr>";
    }

    var dutyBadgesHtml = "";
    if (availableDutyTeachers.length > 0) {
      for (var dtIdx = 0; dtIdx < availableDutyTeachers.length; dtIdx++) {
        dutyBadgesHtml += "<span style='display:inline-block;background-color:#ecfdf5;color:#065f46;border:1px solid #a7f3d0;border-radius:9999px;padding:4px 12px;font-size:13px;font-weight:600;margin:3px 6px 3px 0;'>" +
          escapeHtmlServer_(availableDutyTeachers[dtIdx]) + "</span>";
      }
    } else {
      dutyBadgesHtml = "<p style='margin:0;color:#6b7280;font-style:italic;font-size:13px;'>None available — all teachers with " + escapeHtmlServer_(periodLabel) + " Sub Duty are already covering a class or absent.</p>";
    }

    var htmlBody =
      "<div style='font-family:Arial,sans-serif;max-width:680px;color:#1f2937;'>" +
        "<div style='background-color:#fef2f2;border-left:4px solid #dc2626;padding:14px 16px;border-radius:6px;margin-bottom:16px;'>" +
          "<h2 style='margin:0 0 4px 0;font-size:18px;color:#991b1b;'>⚠️ Unfilled Sub Coverage Starting in " + minsLeft + " Minutes</h2>" +
          "<p style='margin:0;font-size:14px;color:#7f1d1d;'>" +
            "<strong>" + escapeHtmlServer_(periodLabel) + "</strong>" +
            (timeRangeDisplay ? " (" + escapeHtmlServer_(timeRangeDisplay) + ")" : "") +
            (dayColorDisplay ? " • " + escapeHtmlServer_(dayColorDisplay) : "") +
          "</p>" +
        "</div>" +
        "<p style='font-size:14px;margin:0 0 12px 0;'>The following " +
          (unfilledList.length === 1 ? "class still needs" : (unfilledList.length + " classes still need")) +
          " a substitute:</p>" +
        "<table style='width:100%;border-collapse:collapse;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;margin-bottom:18px;font-size:14px;'>" +
          "<thead>" +
            "<tr style='background-color:#002147;color:#ffffff;text-align:left;'>" +
              "<th style='padding:10px 12px;'>Absent Teacher</th>" +
              "<th style='padding:10px 12px;'>Course</th>" +
              "<th style='padding:10px 12px;'>Room</th>" +
              "<th style='padding:10px 12px;'>Instructions / Sub Plan</th>" +
            "</tr>" +
          "</thead>" +
          "<tbody>" + htmlRows + "</tbody>" +
        "</table>" +
        "<div style='background-color:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:12px 16px;margin-bottom:18px;'>" +
          "<h3 style='margin:0 0 8px 0;font-size:14px;color:#002147;'>Available Teachers with " + escapeHtmlServer_(periodLabel) + " Sub Duty Today:</h3>" +
          "<div>" + dutyBadgesHtml + "</div>" +
        "</div>" +
        (appUrl
          ? "<p style='margin:18px 0;'><a href='" + escapeHtmlServer_(appUrl) + "' target='_blank' style='display:inline-block;background-color:#002147;color:#ffffff;font-weight:bold;text-decoration:none;padding:10px 20px;border-radius:6px;font-size:14px;'>Open Cathedral Sub App to Assign Coverage</a></p>"
          : "") +
      "</div>";

    var recipientCsv = cleanRecipients.join(",");
    sendEmailHelper(recipientCsv, subject, plainBody, { htmlBody: htmlBody }, settings);
    emailsSent++;

    if (typeof logAuditAction === "function") {
      try {
        logAuditAction(
          "UNFILLED_REMINDER_SENT",
          todayIso + "_P" + targetP,
          "Sent 15-min unfilled reminder for " + periodLabel + " (" + unfilledList.length + " unfilled) to " + recipientCsv,
          "System"
        );
      } catch (audErr) {}
    }
  }

  return {
    checked: true,
    duePeriods: duePeriods.length,
    emailsSent: emailsSent
  };
}

/**
 * Escapes HTML special characters for server-side email templates.
 */
function escapeHtmlServer_(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Encodes a string for a URL query parameter while preserving deferred
 * schedule placeholders ({{ROOM|...}} / {{COURSE|...}}) as URL-encoded
 * placeholder tokens ({{ROOM_URL|...}} / {{COURSE_URL|...}}) for processEmailQueue.
 */
function encodeCalendarParamWithPlaceholders(str) {
  var raw = String(str || "");
  var parts = raw.split(/(\{\{(?:ROOM|COURSE)\|[^}]+\}\})/g);
  var out = "";
  for (var i = 0; i < parts.length; i++) {
    var seg = parts[i];
    if (!seg) continue;
    var m = seg.match(/^\{\{(ROOM|COURSE)\|([^}]+)\}\}$/);
    if (m) {
      out += "{{" + m[1] + "_URL|" + m[2] + "}}";
    } else {
      out += encodeURIComponent(seg);
    }
  }
  return out;
}

/**
 * Builds a Google Calendar template URL for a sub coverage assignment.
 * Uses exact period start/end clock times (`YYYYMMDDTHHMM00/YYYYMMDDTHHMM00`) when available,
 * falling back to an all-day event (`YYYYMMDD/YYYYMMDD`) if period times are not applicable.
 */
function buildSubDutyGoogleCalendarUrl(details, appUrl, optionalSettings) {
  if (!details) return "";

  var isoDate = details.isoDate || (typeof getStandardizedDateString === "function" ? getStandardizedDateString(details.date) : "");
  var datesParam = "";
  var timeSummaryLine = "";

  var periodRange = (isoDate && details.period && typeof getPeriodTimeRangeForDate === "function")
    ? getPeriodTimeRangeForDate(isoDate, details.period, null, optionalSettings)
    : null;

  if (periodRange && periodRange.startLocalStr && periodRange.endLocalStr) {
    datesParam = "&dates=" + periodRange.startLocalStr + "/" + periodRange.endLocalStr;
    timeSummaryLine = "Time: " + periodRange.startTime + " - " + periodRange.endTime + (periodRange.dayColor ? " (" + periodRange.dayColor + " Day)" : "");
  } else if (isoDate && /^\d{4}-\d{2}-\d{2}$/.test(isoDate)) {
    var dParts = isoDate.split("-");
    var yr = parseInt(dParts[0], 10);
    var mo = parseInt(dParts[1], 10) - 1;
    var dy = parseInt(dParts[2], 10);
    var startObj = new Date(Date.UTC(yr, mo, dy));
    var endObj = new Date(Date.UTC(yr, mo, dy + 1));
    var pad2 = function(n) { return n < 10 ? "0" + n : String(n); };
    var startYMD = String(startObj.getUTCFullYear()) + pad2(startObj.getUTCMonth() + 1) + pad2(startObj.getUTCDate());
    var endYMD = String(endObj.getUTCFullYear()) + pad2(endObj.getUTCMonth() + 1) + pad2(endObj.getUTCDate());
    datesParam = "&dates=" + startYMD + "/" + endYMD;
  }

  var displayPeriodLabel = details.period ? formatPeriodLabelDisplay(details.period) : "Sub Coverage";
  var title = "Sub Coverage: " + displayPeriodLabel + " - " + (details.teacherName || "Teacher");
  var locationStr = details.room ? "Room " + details.room + " - Cathedral High School" : "Cathedral High School";
  var descLines = [
    "Teacher to Cover: " + (details.teacherName || ""),
    details.period ? "Period: " + formatPeriodDisplay(details.period) : "",
    timeSummaryLine,
    details.course ? "Course: " + details.course : "",
    details.room ? "Room: " + details.room : "",
    "Special Instructions: " + (details.instructions ? details.instructions : "None provided"),
    details.subPlanUrl ? "Sub Plan: " + details.subPlanUrl : "",
    "Take Attendance: https://powerschool.cathedral-irish.org/subs/pw.html",
    appUrl ? "Cathedral Sub App: " + appUrl : ""
  ].filter(function(line) { return Boolean(line); }).join("\n");

  var url = "https://calendar.google.com/calendar/render?action=TEMPLATE" +
            "&text=" + encodeCalendarParamWithPlaceholders(title) +
            datesParam +
            "&location=" + encodeCalendarParamWithPlaceholders(locationStr) +
            "&details=" + encodeCalendarParamWithPlaceholders(descLines);
  return url;
}

/**
 * Automatically creates, updates, or deletes a Google Calendar event on a substitute's
 * calendar when coverage is Assigned, Modified, or Canceled.
 */
function syncSubDutyCalendarEvent(subEmail, type, details, optionalSettings) {
  if (!subEmail || !details) return { synced: false, reason: "Missing subEmail or details" };
  if (typeof CalendarApp === "undefined" || !CalendarApp) {
    return { synced: false, reason: "CalendarApp not available" };
  }

  var settings = optionalSettings || (typeof getSettings === "function" ? getSettings() : {});
  var autoSyncSetting = String(settings["Auto Sync Google Calendar"] !== undefined ? settings["Auto Sync Google Calendar"] : "true").toLowerCase().trim();
  if (autoSyncSetting === "false" || autoSyncSetting === "off" || autoSyncSetting === "0") {
    return { synced: false, reason: "Auto Sync Google Calendar disabled in Settings" };
  }

  var isoDate = details.isoDate || (typeof getStandardizedDateString === "function" ? getStandardizedDateString(details.date) : "");
  if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) {
    return { synced: false, reason: "Invalid or missing isoDate" };
  }

  var dParts = isoDate.split("-");
  var yr = parseInt(dParts[0], 10);
  var mo = parseInt(dParts[1], 10) - 1;
  var dy = parseInt(dParts[2], 10);
  var dayStart = new Date(yr, mo, dy, 0, 0, 0);
  var dayEnd = new Date(yr, mo, dy, 23, 59, 59);

  var normPeriod = details.period !== undefined ? parsePeriod(details.period) : "";
  var absId = details.id || details.absenceId || "";
  var tagKey = (absId ? String(absId) : ("date_" + isoDate)) + "_P" + normPeriod;
  var idMarker = "[SubAppID: " + tagKey + "]";

  var cal = null;
  var useGuestFallback = false;
  try {
    cal = CalendarApp.getCalendarById(subEmail);
  } catch (e) {
    cal = null;
  }
  if (!cal && typeof CalendarApp.getDefaultCalendar === "function") {
    try {
      cal = CalendarApp.getDefaultCalendar();
      useGuestFallback = true;
    } catch (e2) {
      cal = null;
    }
  }
  if (!cal) return { synced: false, reason: "Target calendar not accessible" };

  // Find any existing matching event on that day
  var existingEvent = null;
  try {
    var events = cal.getEvents(dayStart, dayEnd);
    for (var i = 0; i < events.length; i++) {
      var ev = events[i];
      var evTag = (typeof ev.getTag === "function") ? ev.getTag("subapp_key") : "";
      var evDesc = (typeof ev.getDescription === "function") ? String(ev.getDescription() || "") : "";
      var evTitle = (typeof ev.getTitle === "function") ? String(ev.getTitle() || "") : "";
      if (evTag === tagKey || evDesc.indexOf(idMarker) !== -1 ||
          (!absId && evTitle.indexOf("Sub Coverage: " + formatPeriodLabelDisplay(details.period) + " - " + (details.teacherName || "")) === 0)) {
        existingEvent = ev;
        break;
      }
    }
  } catch (findErr) {
    console.warn("Could not search existing calendar events: " + findErr.message);
  }

  if (type === "Canceled") {
    if (existingEvent && typeof existingEvent.deleteEvent === "function") {
      existingEvent.deleteEvent();
      return { synced: true, action: "deleted", tagKey: tagKey };
    }
    return { synced: false, action: "none_to_delete", tagKey: tagKey };
  }

  var periodRange = (details.period && typeof getPeriodTimeRangeForDate === "function")
    ? getPeriodTimeRangeForDate(isoDate, details.period, null, settings)
    : null;

  var displayPeriodLabel = details.period ? formatPeriodLabelDisplay(details.period) : "Sub Coverage";
  var hasRealCourse = details.course && details.course !== "No Class Assigned" && details.course.indexOf("{{") === -1;
  var hasRealRoom = details.room && details.room !== "No Class Assigned" && details.room.indexOf("{{") === -1;
  var title = "Sub Coverage: " + displayPeriodLabel + " - " + (details.teacherName || "Teacher") +
              (hasRealCourse ? " (" + details.course + ")" : "");
  var locationStr = hasRealRoom ? ("Room " + details.room + " - Cathedral High School") : "Cathedral High School";
  var appUrl = settings["App URL"] || DEFAULT_APP_URL;

  var descLines = [
    "Teacher to Cover: " + (details.teacherName || ""),
    details.period ? "Period: " + formatPeriodDisplay(details.period) : "",
    (periodRange && periodRange.startTime && periodRange.endTime)
      ? ("Time: " + periodRange.startTime + " - " + periodRange.endTime + (periodRange.dayColor ? " (" + periodRange.dayColor + " Day)" : ""))
      : "",
    details.course ? "Course: " + details.course : "",
    details.room ? "Room: " + details.room : "",
    "Special Instructions: " + (details.instructions ? details.instructions : "None provided"),
    details.subPlanUrl ? "Sub Plan: " + details.subPlanUrl : "",
    "Take Attendance: https://powerschool.cathedral-irish.org/subs/pw.html",
    appUrl ? "Cathedral Sub App: " + appUrl : "",
    idMarker
  ].filter(function(l) { return Boolean(l); }).join("\n");

  if (existingEvent) {
    if (typeof existingEvent.setTitle === "function") existingEvent.setTitle(title);
    if (typeof existingEvent.setLocation === "function") existingEvent.setLocation(locationStr);
    if (typeof existingEvent.setDescription === "function") existingEvent.setDescription(descLines);
    if (periodRange && periodRange.startDateObj && periodRange.endDateObj && typeof existingEvent.setTime === "function") {
      existingEvent.setTime(periodRange.startDateObj, periodRange.endDateObj);
    }
    if (typeof existingEvent.setTag === "function") existingEvent.setTag("subapp_key", tagKey);
    return { synced: true, action: "updated", tagKey: tagKey };
  }

  var eventOptions = {
    location: locationStr,
    description: descLines
  };
  if (useGuestFallback) {
    eventOptions.guests = subEmail;
    eventOptions.sendInvites = false;
  }

  var createdEvent = null;
  if (periodRange && periodRange.startDateObj && periodRange.endDateObj && typeof cal.createEvent === "function") {
    createdEvent = cal.createEvent(title, periodRange.startDateObj, periodRange.endDateObj, eventOptions);
  } else if (typeof cal.createAllDayEvent === "function") {
    createdEvent = cal.createAllDayEvent(title, dayStart, eventOptions);
  }

  if (createdEvent && typeof createdEvent.setTag === "function") {
    createdEvent.setTag("subapp_key", tagKey);
  }

  return { synced: Boolean(createdEvent), action: "created", tagKey: tagKey };
}

/**
 * Sends a notification email to a substitute.
 */
function sendSubNotification(subEmail, type, details) {
  if (!subEmail) return;

  var subject = "";
  var body = "";
  var htmlBody = "";

  var settings = getSettings();
  var appUrl = settings["App URL"] || DEFAULT_APP_URL;

  var pTimeObj = (details.period && (details.isoDate || details.date) && typeof getPeriodTimeRangeForDate === "function")
    ? getPeriodTimeRangeForDate(details.isoDate || details.date, details.period, null, settings)
    : null;
  var pTimeSuffix = (pTimeObj && pTimeObj.displayRange) ? (" (" + pTimeObj.displayRange + ")") : "";
  var displayPeriod = details.period ? (formatPeriodDisplay(details.period) + pTimeSuffix) : "";
  var displayPeriodLabel = details.period ? (formatPeriodLabelDisplay(details.period) + pTimeSuffix) : "";

  var detailsText = "Date: " + details.date + "\n";
  if (details.period) detailsText += "Period: " + displayPeriod + "\n";
  detailsText += "Teacher: " + details.teacherName + "\n";
  if (details.period) {
    detailsText += "Room: " + details.room + "\n";
    detailsText += "Course: " + details.course + "\n";
  }
  detailsText += "\nSpecial Instructions:\n" + (details.instructions ? details.instructions : "None provided");

  var detailsHtml = "<ul>" +
                    "<li><strong>Date:</strong> " + details.date + "</li>" +
                    (details.period ? "<li><strong>Period:</strong> " + displayPeriod + "</li>" : "") +
                    "<li><strong>Teacher:</strong> " + details.teacherName + "</li>" +
                    (details.period ? "<li><strong>Room:</strong> " + details.room + "</li>" : "") +
                    (details.period ? "<li><strong>Course:</strong> " + details.course + "</li>" : "") +
                    "</ul>" +
                    "<p><strong>Special Instructions:</strong><br>" + (details.instructions ? details.instructions : "None provided") + "</p>";

  var gcalUrl = buildSubDutyGoogleCalendarUrl(details, appUrl, settings);
  var calendarText = gcalUrl ? ("\n\nAdd to Google Calendar: " + gcalUrl) : "";
  var calendarHtml = gcalUrl
    ? "<p style='margin: 16px 0;'><a href='" + gcalUrl + "' target='_blank' style='display: inline-block; background-color: #00843D; color: #ffffff; font-weight: bold; text-decoration: none; padding: 10px 18px; border-radius: 6px;'>📅 Add to Google Calendar</a></p>"
    : "";

  var attendanceText = "\n\nPlease remember to take attendance here: https://powerschool.cathedral-irish.org/subs/pw.html";
  var attendanceHtml = "<p><strong>Please remember to take attendance here: <a href='https://powerschool.cathedral-irish.org/subs/pw.html'>https://powerschool.cathedral-irish.org/subs/pw.html</a></strong></p>";

  if (type === 'Assigned') {
    subject = "Coverage Assignment: " + details.date + (details.period ? " " + displayPeriodLabel : "");
    body = "You have been assigned to cover a class.\n\n" + detailsText + calendarText + attendanceText + "\n\nPlease check the Cathedral Sub App for more information: " + appUrl;
    htmlBody = "<p>You have been assigned to cover a class.</p>" + detailsHtml + calendarHtml + attendanceHtml + "<p>Please check the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
  } else if (type === 'Canceled') {
    subject = "CANCELED - Coverage Assignment: " + details.date + (details.period ? " " + displayPeriodLabel : "");
    body = "Your assigned coverage has been CANCELED. You are no longer needed for this assignment.\n\n" + detailsText + "\n\nGo to the Cathedral Sub App for more information: " + appUrl;
    htmlBody = "<p>Your assigned coverage has been CANCELED. You are no longer needed for this assignment.</p>" + detailsHtml + "<p>Go to the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
  } else if (type === 'Modified') {
    subject = "UPDATED - Coverage Assignment: " + details.date + (details.period ? " " + displayPeriodLabel : "");
    body = "There has been an update to your assigned coverage.\n\nUpdated Details:\n" + detailsText + calendarText + attendanceText + "\n\nPlease check the Cathedral Sub App for more information: " + appUrl;
    htmlBody = "<p>There has been an update to your assigned coverage.</p><h3>Updated Details:</h3>" + detailsHtml + calendarHtml + attendanceHtml + "<p>Please check the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
  }

  try {
    var emailOptions = {
      htmlBody: htmlBody,
      calendarSync: {
        subEmail: subEmail,
        type: type,
        details: details
      }
    };

    // Anytime an out-of-building substitute ("Substitute" role) has their coverage canceled,
    // also notify the Sub Coordinator(s) on the cancellation email.
    if (type === 'Canceled' && typeof isExternalSubstituteByEmailOrName_ === 'function') {
      var ssForSub = getSS();
      if (isExternalSubstituteByEmailOrName_(ssForSub, subEmail)) {
        var coordEmails = getEmailsByRole(ssForSub, "sub coordinator");
        var targetSubLower = String(subEmail).trim().toLowerCase();
        var filteredCoords = [];
        for (var cIdx = 0; cIdx < coordEmails.length; cIdx++) {
          var ce = String(coordEmails[cIdx] || "").trim();
          if (ce && ce.toLowerCase() !== targetSubLower && filteredCoords.indexOf(ce) === -1) {
            filteredCoords.push(ce);
          }
        }
        if (filteredCoords.length > 0) {
          emailOptions.cc = filteredCoords.join(",");
        }
      }
    }

    enqueueEmail(subEmail, subject, body, emailOptions);
  } catch (e) {
    console.error("Failed to enqueue email to " + subEmail + ": " + e.message);
  }
}

function generatePrincipalsDigestHTML(dateObj) {
  var ss = getSS();
  var settings = getSettings();
  var appUrl = settings["App URL"] || DEFAULT_APP_URL;

  // Resolve target dates
  var refDate = dateObj ? new Date(dateObj) : new Date();
  if (typeof dateObj === 'string' && dateObj.indexOf('-') > -1) {
    var refParts = dateObj.split('-');
    if (refParts.length === 3) {
      refDate = new Date(parseInt(refParts[0], 10), parseInt(refParts[1], 10) - 1, parseInt(refParts[2], 10), 12, 0, 0);
    }
  }
  
  // Calculate Monday to Friday of the CURRENT week (assuming refDate is Friday)
  var currentWeekMonday = new Date(refDate);
  currentWeekMonday.setDate(refDate.getDate() - (refDate.getDay() === 0 ? 6 : refDate.getDay() - 1));
  currentWeekMonday.setHours(0, 0, 0, 0);
  
  var currentWeekFriday = new Date(currentWeekMonday);
  currentWeekFriday.setDate(currentWeekMonday.getDate() + 4);
  currentWeekFriday.setHours(23, 59, 59, 999);

  // Calculate Monday to Friday of the NEXT week
  var nextWeekMonday = new Date(currentWeekMonday);
  nextWeekMonday.setDate(currentWeekMonday.getDate() + 7);
  var nextWeekFriday = new Date(nextWeekMonday);
  nextWeekFriday.setDate(nextWeekMonday.getDate() + 4);
  nextWeekFriday.setHours(23, 59, 59, 999);

  // Extract HR rates
  var specialRatesArray = [];
  try {
      specialRatesArray = JSON.parse(settings["Special Pay Rates"] || "[]");
  } catch(e) {}

  var hrRates = {
    green: parseFloat(settings["Green Day Pay Rate"] || "10"),
    blueGold: parseFloat(settings["Blue/Gold Day Pay Rate"] || "20"),
    special: specialRatesArray
  };

  // Get data
  var dateColors = getDateColorsCached(ss);

  var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
  var rosterData = rosterSheet.getDataRange().getValues();
  var nameLookup = {};
  var dutyLookup = {};
  for (var roIdx = 1; roIdx < rosterData.length; roIdx++) {
    if (!rosterData[roIdx] || rosterData[roIdx].length < 4) continue;
    var e = String(rosterData[roIdx][1]).toLowerCase().trim();
    var nm = String(rosterData[roIdx][0]).trim();
    if (e) nameLookup[e] = nm;
    dutyLookup[nm] = String(rosterData[roIdx][3] || "").trim();
  }

  var v2CompatDigest = (typeof readV2AbsenceDataCompat === "function") ? readV2AbsenceDataCompat(ss) : null;
  var absenceSheet = !v2CompatDigest ? getSheetOrThrow(ss, "Absence Requests") : null;
  var absenceData = v2CompatDigest ? v2CompatDigest.absenceData : absenceSheet.getDataRange().getValues();

  var currentWeekAbsences = {};
  var currentWeekCoverage = {};
  var nextWeekAbsences = {};

  for (var aIdx = 1; aIdx < absenceData.length; aIdx++) {
    var row = absenceData[aIdx];
    if (!row || row.length < 4) continue;
    var status = String(row[19] || "").trim();
    if (status.toLowerCase() !== "active") continue; // Only active requests

    var dateVal = row[3];
    var dateObjRow;
    if (dateVal instanceof Date) {
      dateObjRow = new Date(dateVal);
    } else {
      let dStr = String(dateVal).trim();
      var parts = dStr.split("-");
      if (parts.length === 3) {
        dateObjRow = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
      } else {
        dateObjRow = new Date(dStr);
      }
    }
    if (!dateObjRow || isNaN(dateObjRow.getTime())) continue;

    var teacherEmail = String(row[2]).toLowerCase().trim();
    var teacherName = nameLookup[teacherEmail] || teacherEmail;

    var isCurrentWeek = dateObjRow >= currentWeekMonday && dateObjRow <= currentWeekFriday;
    var isNextWeek = dateObjRow >= nextWeekMonday && dateObjRow <= nextWeekFriday;

    if (!isCurrentWeek && !isNextWeek) continue;

    var duration = String(row[6]).trim();
    var periodsStr = String(row[4]).trim();
    var periods = periodsStr ? periodsStr.split(',').map(function(p){return p.trim();}) : [];

    var daysAbsent = 0;
    if (duration.toLowerCase().includes("full")) daysAbsent = 1;
    else if (duration.toLowerCase().includes("half")) daysAbsent = 0.5;

    var targetAbsences = isCurrentWeek ? currentWeekAbsences : nextWeekAbsences;
    if (!targetAbsences[teacherName]) {
      targetAbsences[teacherName] = { days: 0, periods: 0 };
    }
    targetAbsences[teacherName].days += daysAbsent;
    targetAbsences[teacherName].periods += periods.length;

    if (isCurrentWeek) {
      var dateFormatted = Utilities.formatDate(dateObjRow, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var dayColor = dateColors[dateFormatted] || "Green";

      for (var j = 0; j < periods.length; j++) {
        var p = periods[j];
        if (!p) continue;
        var subColIdx = getSubColumnIndex(p);
        if (subColIdx > 0 && subColIdx <= row.length) {
          var assignedSubRaw = row[subColIdx - 1];
          if (assignedSubRaw && String(assignedSubRaw).trim() !== "") {
            // Trim suffix if exists (e.g. 'Name - Duty')
            var assignedSub = String(assignedSubRaw).trim().replace(/\s+-\s+.*$/, "");
            
            // Check if sub is dedicated substitute or teacher. Calculate pay only for teachers (not role substitute).
            // (Assuming teachers are handled with hrRates based on day color. We need to know if the person is a sub or teacher)
            var subEmailLookup = "";
            for (var r = 1; r < rosterData.length; r++) {
                if (String(rosterData[r][0]).trim() === assignedSub) {
                    if (!rosterData[r]) continue;
                    subEmailLookup = String(rosterData[r][1]).toLowerCase().trim();
                    break;
                }
            }
            
            var isSubstituteRole = false;
            var rawTeacherDuty = dutyLookup[assignedSub];
            var teacherDuty = (typeof resolveStaffDutyForDate === "function")
              ? resolveStaffDutyForDate(rawTeacherDuty, dateFormatted, settings)
              : rawTeacherDuty;
            var teacherDutyArr = teacherDuty ? String(teacherDuty).split(',').map(function(d){return parsePeriod(d);}) : [];
            var isDuty = teacherDutyArr.includes(parsePeriod(p));

            if (subEmailLookup) {
                for (var rsIdx = 1; rsIdx < rosterData.length; rsIdx++) {
                    if (rosterData[rsIdx] && String(rosterData[rsIdx][1]).toLowerCase().trim() === subEmailLookup) {
                         var roles = String(rosterData[rsIdx][2]).toLowerCase();
                         if (roles.includes("substitute")) {
                             isSubstituteRole = true;
                         }
                         break;
                    }
                }
            }

            if (!currentWeekCoverage[assignedSub]) {
              currentWeekCoverage[assignedSub] = { periods: 0, pay: 0 };
            }
            currentWeekCoverage[assignedSub].periods += 1;

            if (!isSubstituteRole) {
               if (!isDuty) {
                   var specialPayRate = null;
                   for (var sIdx = 0; sIdx < hrRates.special.length; sIdx++) {
                       if (hrRates.special[sIdx].date === dateFormatted) {
                           specialPayRate = parseFloat(hrRates.special[sIdx].rate || 0);
                           break;
                       }
                   }

                   if (specialPayRate !== null) {
                       currentWeekCoverage[assignedSub].pay += specialPayRate;
                   } else if (dayColor.toLowerCase() === "blue" || dayColor.toLowerCase() === "gold") {
                       currentWeekCoverage[assignedSub].pay += hrRates.blueGold;
                   } else {
                       currentWeekCoverage[assignedSub].pay += hrRates.green;
                   }
               }
            }
          }
        }
      }
    }
  }

  // Format HTML
  var html = "<div style='font-family: sans-serif; color: #333;'>";
  html += "<p>Hello Principals,</p>";
  html += "<p>Here is your weekly digest of substitute coverage and absences.</p>";

  // SECTION 1
  html += "<h3 style='color: #002147; border-bottom: 1px solid #ccc; padding-bottom: 5px;'>Absences This Week (" + Utilities.formatDate(currentWeekMonday, Session.getScriptTimeZone(), "MMM d") + " - " + Utilities.formatDate(currentWeekFriday, Session.getScriptTimeZone(), "MMM d") + ")</h3>";
  var sec1 = false;
  html += "<ul>";
  for (var cwName in currentWeekAbsences) {
    if (currentWeekAbsences[cwName].days > 0 || currentWeekAbsences[cwName].periods > 0) {
      sec1 = true;
      let text = [];
      if (currentWeekAbsences[cwName].days > 0) text.push(currentWeekAbsences[cwName].days + " days");
      if (currentWeekAbsences[cwName].periods > 0) text.push(currentWeekAbsences[cwName].periods + " periods");
      html += "<li><strong>" + cwName + "</strong>: " + text.join(" and ") + "</li>";
    }
  }
  if (!sec1) html += "<li>No absences this week.</li>";
  html += "</ul>";

  // SECTION 2
  html += "<h3 style='color: #002147; border-bottom: 1px solid #ccc; padding-bottom: 5px;'>Coverage This Week</h3>";
  var sec2 = false;
  html += "<ul>";
  for (var ccName in currentWeekCoverage) {
    if (currentWeekCoverage[ccName].periods > 0) {
      sec2 = true;
      var pText = currentWeekCoverage[ccName].periods + (currentWeekCoverage[ccName].periods === 1 ? " class" : " classes");
      var payText = currentWeekCoverage[ccName].pay > 0 ? " ($" + currentWeekCoverage[ccName].pay.toFixed(2) + " extra pay)" : "";
      html += "<li><strong>" + ccName + "</strong>: covered " + pText + payText + "</li>";
    }
  }
  if (!sec2) html += "<li>No classes covered this week.</li>";
  html += "</ul>";

  // SECTION 3
  html += "<h3 style='color: #002147; border-bottom: 1px solid #ccc; padding-bottom: 5px;'>Upcoming Absences Next Week (" + Utilities.formatDate(nextWeekMonday, Session.getScriptTimeZone(), "MMM d") + " - " + Utilities.formatDate(nextWeekFriday, Session.getScriptTimeZone(), "MMM d") + ")</h3>";
  var sec3 = false;
  html += "<ul>";
  for (var nwName in nextWeekAbsences) {
    if (nextWeekAbsences[nwName].days > 0 || nextWeekAbsences[nwName].periods > 0) {
      sec3 = true;
      let text = [];
      if (nextWeekAbsences[nwName].days > 0) text.push(nextWeekAbsences[nwName].days + " days");
      if (nextWeekAbsences[nwName].periods > 0) text.push(nextWeekAbsences[nwName].periods + " periods");
      html += "<li><strong>" + nwName + "</strong>: " + text.join(" and ") + "</li>";
    }
  }
  if (!sec3) html += "<li>No upcoming absences next week.</li>";
  html += "</ul>";

  html += "<p style='margin-top: 20px;'>For more information, go to the <a href='" + appUrl + "'>Cathedral Sub App</a>.</p>";
  html += "<p>Best,<br>Cathedral Sub App</p>";
  html += "</div>";

  return html;
}

function runPrincipalsDigestWeekly() {
  sendPrincipalsDigest(new Date());
}

function sendPrincipalsDigest(dateObj) {
  var htmlBody = generatePrincipalsDigestHTML(dateObj);
  var ss = getSS();
  var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
  var rosterData = rosterSheet.getDataRange().getValues();
  var principals = [];
  
  for (var i = 1; i < rosterData.length; i++) {
    if (!rosterData[i]) continue;
    if (!rosterData[i]) continue;
    var email = String(rosterData[i][1]).toLowerCase().trim();
    var roles = String(rosterData[i][2]).toLowerCase();
    if (roles.includes("principal") && email) {
      if (principals.indexOf(email) === -1) principals.push(email);
    }
  }
  
  if (principals.length === 0) {
    console.log("No principals found in Staff Roster to send digest to.");
    return;
  }
  
  var settings = getSettings();
  var options = { htmlBody: htmlBody };
  if (settings["Email Sender Name"]) {
      options.name = settings["Email Sender Name"];
  }
  if (settings["Reply To Email"]) {
      options.replyTo = settings["Reply To Email"];
  }

  var subject = "Weekly Principal's Digest - " + Utilities.formatDate(dateObj ? new Date(dateObj) : new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");

  for (var j = 0; j < principals.length; j++) {
    sendEmailHelper(principals[j], subject, "", options, settings);
  }
  console.log("Sent Principal's Digest to " + principals.join(", "));
}

function generatePrincipalsDigestPreview(dateStr, clientEmail) {
  var ss = getSS();
  var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
  var roleLower = String(user.role || "").toLowerCase().trim();
  var settings = getSettings(ss);
  var rolePermsStr = settings["RolePermissions"] || "{}";
  var isAuthorized = (
    roleLower === "admin" ||
    roleLower === "principal" ||
    hasPermission(roleLower, "Settings", rolePermsStr) ||
    hasPermission(roleLower, "Admin Dashboard", rolePermsStr)
  );
  if (!isAuthorized) {
    throw new Error("Unauthorized to preview Principal's Digest.");
  }
  var dateObj = dateStr ? (typeof parseLocalDateSafe === "function" ? parseLocalDateSafe(dateStr) : new Date(dateStr)) : new Date();
  return generatePrincipalsDigestHTML(dateObj);
}

function sendDailySubFeedbackRequests() {
  try {
    var ss = getSS();
    var v2CompatFb = (typeof readV2AbsenceDataCompat === "function") ? readV2AbsenceDataCompat(ss) : null;
    var mainSheet = !v2CompatFb ? getSheetOrThrow(ss, "Absence Requests") : null;
    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var settings = getSettings(ss);

    var data = v2CompatFb ? v2CompatFb.absenceData : mainSheet.getDataRange().getValues();
    var rosterData = rosterSheet.getDataRange().getValues();
    var appUrl = settings["App URL"] || DEFAULT_APP_URL;

    var subEmailLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      if (!rosterData[r]) continue;
      if (!rosterData[r]) continue;
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }
    var nameLookup = buildNameLookup(rosterData);

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var requestsToSend = {}; // Keyed by substitute email

    for (var i = 1; i < data.length; i++) {
      if (!data[i] || data[i].length < 4) continue;
      var row = data[i];
      var status = String(row[19] || 'Active');
      if (status === 'Canceled') continue;

      var dateVal = row[3];
      if (!dateVal) continue;

      var rowDate = typeof parseLocalDateSafe === "function" ? parseLocalDateSafe(dateVal) : new Date(dateVal);
      rowDate.setHours(0, 0, 0, 0);

      if (rowDate.getTime() === today.getTime()) {
        var rowTeacherEmail = String(row[2]).toLowerCase();
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
        if (teacherName.includes(",")) {
          var parts = teacherName.split(",");
          if (parts.length > 1) teacherName = parts[1].trim() + " " + parts[0].trim();
        }

        var periodsRequested = String(row[4]).split(",").map(function(p) { return parsePeriod(p); });
        var rowId = String(row[0]);
        var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', '10'];

        for (var pIdx = 0; pIdx < allPeriods.length; pIdx++) {
          var p = allPeriods[pIdx];
          if (periodsRequested.indexOf(String(p)) !== -1) {
             var subColumnIndex = getSubColumnIndex(p) - 1;
             var assignedSub = String(row[subColumnIndex] || "").trim();

             if (assignedSub !== "") {
                var subEmail = subEmailLookup[assignedSub];
                if (subEmail) {
                    subEmail = subEmail.toLowerCase();
                    if (!requestsToSend[subEmail]) {
                        requestsToSend[subEmail] = { name: assignedSub, jobs: [] };
                    }
                    requestsToSend[subEmail].jobs.push({
                       id: rowId,
                       teacherName: teacherName,
                       period: p
                    });
                }
             }
          }
        }
      }
    }

    // Process sending emails
    for (var email in requestsToSend) {
        var subData = requestsToSend[email];
        var subject = "Leave Feedback for Today's Classes";

        var bodyText = "Hello " + subData.name + ",\n\nThank you for covering classes today. Please leave feedback for your classes:\n\n";
        var bodyHtml = "<p>Hello " + subData.name + ",</p><p>Thank you for covering classes today. Please leave feedback for your classes:</p><ul>";

        for (var j = 0; j < subData.jobs.length; j++) {
            var job = subData.jobs[j];
            var jobPeriodLabel = formatPeriodLabelDisplay(job.period);
            var jobLink = appUrl + "?feedbackAbsenceId=" + encodeURIComponent(job.id) + "&feedbackPeriod=" + encodeURIComponent(job.period);
            bodyText += "- " + jobPeriodLabel + " (" + job.teacherName + "): " + jobLink + "\n";
            bodyHtml += "<li>" + jobPeriodLabel + " (" + job.teacherName + "): <a href='" + jobLink + "'>Leave Feedback</a></li>";
        }

        bodyText += "\nThank you!";
        bodyHtml += "</ul><p>Thank you!</p>";

        enqueueEmail(email, subject, bodyText, {htmlBody: bodyHtml});
    }

  } catch (e) {
    console.error("Error in sendDailySubFeedbackRequests: " + e.message);
  }
}


function setupSubFeedbackTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runSubFeedbackRequests' || triggers[i].getHandlerFunction() === 'runTeacherFeedbackConsolidated') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create daily triggers that internally check for weekdays
  ScriptApp.newTrigger('runSubFeedbackRequests')
    .timeBased()
    .everyDays(1)
    .atHour(15) // 3 PM
    .create();
}

function runSubFeedbackRequests() {
    var today = new Date().getDay();
    if (today > 0 && today < 6) { // Monday (1) to Friday (5)
        sendDailySubFeedbackRequests();
    }
}

