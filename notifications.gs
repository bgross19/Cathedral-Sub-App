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
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.warn("Could not obtain lock for processEmailQueue");
    return;
  }

  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("Email Queue");
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return; // Only headers

    var statuses = [];
    var needsScheduleLookup = false;
    var settings = getSettings(ss);

    // Check if any pending emails need schedule lookups
    for (var i = 1; i < data.length; i++) {
    if (!data[i]) continue;
      if (data[i][5] === "Pending") {
        var body = String(data[i][3]);
        var optionsStr = String(data[i][4]);
        if (body.indexOf("{{ROOM|") !== -1 || body.indexOf("{{COURSE|") !== -1 ||
            optionsStr.indexOf("{{ROOM|") !== -1 || optionsStr.indexOf("{{COURSE|") !== -1) {
          needsScheduleLookup = true;
          break;
        }
      }
    }

    var scheduleLookup = null;
    if (needsScheduleLookup) {
      try {
        var scheduleData = getMasterScheduleData();
        scheduleLookup = buildScheduleLookup(scheduleData);
      } catch (err) {
        console.error("Failed to load master schedule for email queue processing: " + err.message);
      }
    }

    for (let i = 0; i < data.length; i++) {
        statuses.push([data[i][5]]);
    }

    // Helper to replace placeholders using regex
    var replacePlaceholders = function(text) {
      if (!text || typeof text !== 'string') return text;
      return text.replace(/\{\{(ROOM|COURSE)\|([^}]+)\}\}/g, function(match, type, joinKey) {
        var val = "No Class Assigned";
        if (scheduleLookup && scheduleLookup[joinKey]) {
           val = type === 'ROOM' ? (scheduleLookup[joinKey].room || val) : (scheduleLookup[joinKey].course || val);
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

        if (needsScheduleLookup) {
          body = replacePlaceholders(String(body));
          if (options.htmlBody) {
            options.htmlBody = replacePlaceholders(String(options.htmlBody));
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

  } catch (e) {
    console.error("Error in processEmailQueue: " + e.message);
  } finally {
    lock.releaseLock();
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
    lock.waitLock(10000);
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
    lock.waitLock(10000);
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
    if (data[i] && String(data[i][2]).toLowerCase().trim() === "sub coordinator") return String(data[i][1]).trim();
  }
  return null;
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

  var displayPeriod = details.period ? formatPeriodDisplay(details.period) : "";
  var displayPeriodLabel = details.period ? formatPeriodLabelDisplay(details.period) : "";

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

  var attendanceText = "\n\nPlease remember to take attendance here: https://powerschool.cathedral-irish.org/subs/pw.html";
  var attendanceHtml = "<p><strong>Please remember to take attendance here: <a href='https://powerschool.cathedral-irish.org/subs/pw.html'>https://powerschool.cathedral-irish.org/subs/pw.html</a></strong></p>";

  if (type === 'Assigned') {
    subject = "Coverage Assignment: " + details.date + (details.period ? " " + displayPeriodLabel : "");
    body = "You have been assigned to cover a class.\n\n" + detailsText + attendanceText + "\n\nPlease check the Cathedral Sub App for more information: " + appUrl;
    htmlBody = "<p>You have been assigned to cover a class.</p>" + detailsHtml + attendanceHtml + "<p>Please check the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
  } else if (type === 'Canceled') {
    subject = "CANCELED - Coverage Assignment: " + details.date + (details.period ? " " + displayPeriodLabel : "");
    body = "Your assigned coverage has been CANCELED. You are no longer needed for this assignment.\n\n" + detailsText + "\n\nGo to the Cathedral Sub App for more information: " + appUrl;
    htmlBody = "<p>Your assigned coverage has been CANCELED. You are no longer needed for this assignment.</p>" + detailsHtml + "<p>Go to the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
  } else if (type === 'Modified') {
    subject = "UPDATED - Coverage Assignment: " + details.date + (details.period ? " " + displayPeriodLabel : "");
    body = "There has been an update to your assigned coverage.\n\nUpdated Details:\n" + detailsText + attendanceText + "\n\nPlease check the Cathedral Sub App for more information: " + appUrl;
    htmlBody = "<p>There has been an update to your assigned coverage.</p><h3>Updated Details:</h3>" + detailsHtml + attendanceHtml + "<p>Please check the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
  }

  try {
    enqueueEmail(subEmail, subject, body, { htmlBody: htmlBody });
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

  var absenceSheet = getSheetOrThrow(ss, "Absence Requests");
  var absenceData = absenceSheet.getDataRange().getValues();

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
            var teacherDuty = dutyLookup[assignedSub];
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
  var dateObj = dateStr ? new Date(dateStr) : new Date();
  return generatePrincipalsDigestHTML(dateObj);
}

function sendDailySubFeedbackRequests() {
  try {
    var ss = getSS();
    var mainSheet = getSheetOrThrow(ss, "Absence Requests");
    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var settings = getSettings(ss);

    var data = mainSheet.getDataRange().getValues();
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

      var rowDate = new Date(dateVal);
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

