// ============================================================================
// ABSENCES.GS — Absence Request CRUD, Sub Assignments, and Sub Feedback
// ============================================================================

/**
 * Helper to calculate if an absence request is urgent based on submission time.
 */
function calculateIsUrgentByTime(absenceDateStr, timestamp, ss) {
  try {
    var tz = Session.getScriptTimeZone();

    // Parse formData.date (format YYYY-MM-DD)
    var absParts = absenceDateStr.split("-");
    var absenceDate = new Date(parseInt(absParts[0], 10), parseInt(absParts[1], 10) - 1, parseInt(absParts[2], 10), 12, 0, 0);

    // Get current date components in the script timezone
    var nowTzDateStr = Utilities.formatDate(timestamp, tz, "yyyy-MM-dd");
    var nowTzHourStr = Utilities.formatDate(timestamp, tz, "HH");
    var nowTzDayStr = Utilities.formatDate(timestamp, tz, "E"); // Mon, Tue, Wed, Thu, Fri, Sat, Sun

    var nowParts = nowTzDateStr.split("-");
    var currentLocalDate = new Date(parseInt(nowParts[0], 10), parseInt(nowParts[1], 10) - 1, parseInt(nowParts[2], 10), 12, 0, 0);
    var currentHour = parseInt(nowTzHourStr, 10);

    var settings = getSettings(ss);
    var cutoffHour = parseInt(settings["Urgency Cutoff Time"] || "15", 10);

    var diffTime = absenceDate.getTime() - currentLocalDate.getTime();
    var diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));

    if (diffDays <= 0) {
      return true; // Same day or past
    } else if (diffDays === 1) {
      if (currentHour >= cutoffHour || nowTzDayStr === "Sun") {
        return true; // Next day after cutoff hour, or Sun for Mon
      }
    } else if (diffDays === 2) {
      if (nowTzDayStr === "Sat") {
        return true; // Sat for Mon
      }
    } else if (diffDays === 3) {
      if (nowTzDayStr === "Fri" && currentHour >= cutoffHour) {
        return true; // Fri after cutoff hour for Mon
      }
    }
    return false;
  } catch (e) {
    console.error("Error calculating urgency: " + e.message);
    return false;
  }
}

function submitMultipleAbsenceRequests(requestsToSubmit, clientEmail) {
  if (!requestsToSubmit || requestsToSubmit.length === 0) {
    return { success: true, count: 0 };
  }
  if (requestsToSubmit.length === 1) {
    var singleRes = submitAbsence(requestsToSubmit[0], typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    return singleRes.success ? { success: true, count: 1 } : singleRes;
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    notifyAdminOfError("submitMultipleAbsenceRequests", e);
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var mainSheet = getSheetOrThrow(ss, "Absence Requests");
    var timestamp = new Date();
    var submitterEmail = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    var submitterData = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    var rosterData = getRosterDataCached(ss);
    var settings = getSettings(ss);
    var appUrl = settings["App URL"] || DEFAULT_APP_URL;
    var reasons = [];
    try {
      reasons = JSON.parse(settings["Absence Reasons"] || "[]");
    } catch (e) {}

    var hrEmailsCache = null;
    var prinEmailsCache = null;
    var rowsToAppend = [];
    var postWriteActions = [];
    var failedCount = 0;
    var lastError = "";

    for (var idx = 0; idx < requestsToSubmit.length; idx++) {
      var formData = requestsToSubmit[idx];
      var targetEmail = submitterEmail;
      var teacherName = submitterData.name;
      var instructions = formData.specialInstructions || "";

      if (formData.onBehalfTeacher && String(formData.onBehalfTeacher).trim() !== "") {
        try {
          assertPermission(submitterData, "Add Request on Behalf", "Unauthorized to submit absence requests on behalf of another teacher.");
        } catch (permErr) {
          failedCount++;
          lastError = permErr.message + " (on date " + formData.date + ")";
          continue;
        }
        var foundEmail = null;
        var targetName = formData.onBehalfTeacher.trim().toLowerCase();
        for (var r = 1; r < rosterData.length; r++) {
          if (rosterData[r] && (String(rosterData[r][0]).toLowerCase().indexOf(targetName) !== -1 || String(rosterData[r][0]).toLowerCase() === targetName)) {
            foundEmail = rosterData[r][1];
            teacherName = rosterData[r][0];
            break;
          }
        }
        if (foundEmail) {
          targetEmail = foundEmail;
          instructions = (instructions ? instructions + "\n" : "") + "(Submitted by " + submitterData.name + ")";
        } else {
          failedCount++;
          lastError = "Teacher '" + formData.onBehalfTeacher + "' not found in Staff Roster. (on date " + formData.date + ")";
          continue;
        }
      }

      var urgencyFormatted = formData.urgency === 'Urgent' ? 'Urgent (Less than 24 hr notice)' : 'Standard (Advanced Notice)';
      if (formData.hrConfirmed) instructions = "[HR Docs Provided] " + instructions;

      var uniqueId = Utilities.getUuid();
      var subPlanUrl = formData.subPlanUrl ? String(formData.subPlanUrl).trim() : "";

      rowsToAppend.push([
        uniqueId, timestamp, targetEmail, formData.date, "'" + formData.periods,
        formData.reason, formData.duration, urgencyFormatted, instructions,
        "", "", "", "", "", "", "", "", "", "", "Active", subPlanUrl, "[]"
      ]);

      postWriteActions.push({
        uniqueId: uniqueId,
        formData: formData,
        teacherName: teacherName,
        instructions: instructions,
        urgencyFormatted: urgencyFormatted
      });
    }

    if (rowsToAppend.length > 0) {
      var startRow = mainSheet.getLastRow() + 1;
      mainSheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

      for (var a = 0; a < postWriteActions.length; a++) {
        var act = postWriteActions[a];
        var fData = act.formData;
        var isMarkedUrgent = act.urgencyFormatted === 'Urgent (Less than 24 hr notice)';
        var isUrgentByTime = calculateIsUrgentByTime(fData.date, timestamp, ss);
        if (isMarkedUrgent || isUrgentByTime) {
          sendUrgentCoverageEmail(ss, act.teacherName, fData, act.instructions);
        }

        var hrRequired = false;
        var principalRequired = false;
        for (var rsnIdx = 0; rsnIdx < reasons.length; rsnIdx++) {
          if (reasons[rsnIdx].reason === fData.reason) {
            hrRequired = reasons[rsnIdx].hrRequired === true;
            principalRequired = reasons[rsnIdx].principalRequired === true;
            break;
          }
        }

        var displayPeriods = formatPeriodsListDisplay(fData.periods);

        if (hrRequired) {
          if (!hrEmailsCache) hrEmailsCache = getEmailsByRole(ss, "hr");
          for (var h = 0; h < hrEmailsCache.length; h++) {
            var hrSubject = act.teacherName + " " + fData.reason + " Absence Request";
            var hrBody = "An absence request has been submitted requiring attention.\n\n" +
                         "Teacher: " + act.teacherName + "\n" +
                         "Date Needed: " + fData.date + "\n" +
                         "Periods: " + displayPeriods + "\n" +
                         "Reason: " + fData.reason + "\n\n" +
                         "Instructions: " + (act.instructions ? act.instructions : "None") + "\n\n" +
                         "Please log into the Cathedral Sub App for more information: " + appUrl;
            var hrHtmlBody = "<p>An absence request has been submitted requiring attention.</p>" +
                             "<ul>" +
                             "<li><strong>Teacher:</strong> " + act.teacherName + "</li>" +
                             "<li><strong>Date Needed:</strong> " + fData.date + "</li>" +
                             "<li><strong>Periods:</strong> " + displayPeriods + "</li>" +
                             "<li><strong>Reason:</strong> " + fData.reason + "</li>" +
                             "</ul>" +
                             "<p><strong>Instructions:</strong> " + (act.instructions ? act.instructions : "None") + "</p>" +
                             "<p>Please log into the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
            enqueueEmail(hrEmailsCache[h], hrSubject, hrBody, { htmlBody: hrHtmlBody });
          }
        }

        if (principalRequired) {
          if (!prinEmailsCache) prinEmailsCache = getEmailsByRole(ss, "principal");
          for (var p = 0; p < prinEmailsCache.length; p++) {
            var prinSubject = act.teacherName + " " + fData.reason + " Absence Request";
            var prinBody = "An absence request has been submitted requiring attention.\n\n" +
                           "Teacher: " + act.teacherName + "\n" +
                           "Date Needed: " + fData.date + "\n" +
                           "Periods: " + displayPeriods + "\n" +
                           "Reason: " + fData.reason + "\n\n" +
                           "Instructions: " + (act.instructions ? act.instructions : "None") + "\n\n" +
                           "Please log into the Cathedral Sub App for more information: " + appUrl;
            var prinHtmlBody = "<p>An absence request has been submitted requiring attention.</p>" +
                               "<ul>" +
                               "<li><strong>Teacher:</strong> " + act.teacherName + "</li>" +
                               "<li><strong>Date Needed:</strong> " + fData.date + "</li>" +
                               "<li><strong>Periods:</strong> " + displayPeriods + "</li>" +
                               "<li><strong>Reason:</strong> " + fData.reason + "</li>" +
                               "</ul>" +
                               "<p><strong>Instructions:</strong> " + (act.instructions ? act.instructions : "None") + "</p>" +
                               "<p>Please log into the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
            enqueueEmail(prinEmailsCache[p], prinSubject, prinBody, { htmlBody: prinHtmlBody });
          }
        }

        var confSubject = "New Absence Request Confirmation";
        var confBody = "Your absence request has been successfully submitted.\n\n" +
                       "Details:\n" +
                       "Date: " + fData.date + "\n" +
                       "Periods: " + displayPeriods + "\n" +
                       "Reason: " + fData.reason + "\n" +
                       "Duration: " + fData.duration + "\n" +
                       "Instructions: " + (act.instructions ? act.instructions : "None") + "\n\n" +
                       "Return to Cathedral Sub App: " + appUrl;
        var confHtmlBody = "<p>Your absence request has been successfully submitted.</p>" +
                           "<h3>Details:</h3>" +
                           "<ul>" +
                           "<li><strong>Date:</strong> " + fData.date + "</li>" +
                           "<li><strong>Periods:</strong> " + displayPeriods + "</li>" +
                           "<li><strong>Reason:</strong> " + fData.reason + "</li>" +
                           "<li><strong>Duration:</strong> " + fData.duration + "</li>" +
                           "</ul>" +
                           "<p><strong>Instructions:</strong> " + (act.instructions ? act.instructions : "None") + "</p>" +
                           "<p><a href='" + appUrl + "'>Return to Cathedral Sub App</a></p>";
        enqueueEmail(submitterEmail, confSubject, confBody, { htmlBody: confHtmlBody });
      }

      logAuditAction(
        "ABSENCE_SUBMITTED_BATCH",
        postWriteActions.map(function(a) { return a.uniqueId; }).join(","),
        "Submitted " + rowsToAppend.length + " absence requests (" + postWriteActions.map(function(a) { return a.formData.date; }).join(", ") + ")",
        clientEmail
      );
    }

    if (failedCount > 0) {
      if (rowsToAppend.length > 0) {
        return { success: false, error: "Partial success. " + rowsToAppend.length + " saved, but " + failedCount + " failed. Last error: " + lastError };
      }
      return { success: false, error: lastError };
    }
    return { success: true, count: rowsToAppend.length };
  } catch (err) {
    notifyAdminOfError("submitMultipleAbsenceRequests", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function submitAbsence(formData, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("submitAbsence", e);
    return {
      success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var mainSheet = getSheetOrThrow(ss, "Absence Requests");
    
    var timestamp = new Date();
    var submitterEmail = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    var submitterData = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);

    var targetEmail = submitterEmail;
    var teacherName = submitterData.name;

    var instructions = formData.specialInstructions || "";

    // Handle "On Behalf" submissions
    if (formData.onBehalfTeacher && String(formData.onBehalfTeacher).trim() !== "") {
      assertPermission(submitterData, "Add Request on Behalf", "Unauthorized to submit absence requests on behalf of another teacher.");

      // Find the teacher's email based on the provided name
      var rosterData = getRosterDataCached(ss);

      var foundEmail = null;
      var targetName = formData.onBehalfTeacher.trim().toLowerCase();

      for (var i = 1; i < rosterData.length; i++) {
        if (String(rosterData[i][0]).toLowerCase().indexOf(targetName) !== -1 || String(rosterData[i][0]).toLowerCase() === targetName) {
           foundEmail = rosterData[i][1];
           teacherName = rosterData[i][0]; // Proper casing
           break;
        }
      }

      if (foundEmail) {
        targetEmail = foundEmail;
        instructions = (instructions ? instructions + "\n" : "") + "(Submitted by " + submitterData.name + ")";
      } else {
        return { success: false, error: "Teacher '" + formData.onBehalfTeacher + "' not found in Staff Roster." };
      }
    }

    var urgencyFormatted = formData.urgency === 'Urgent' ? 'Urgent (Less than 24 hr notice)' : 'Standard (Advanced Notice)';
    if (formData.hrConfirmed) instructions = "[HR Docs Provided] " + instructions;

    var uniqueId = Utilities.getUuid();
    var subPlanUrl = formData.subPlanUrl ? String(formData.subPlanUrl).trim() : "";

    var newRow = [
      uniqueId, timestamp, targetEmail, formData.date, "'" + formData.periods,
      formData.reason, formData.duration, urgencyFormatted, instructions,
      "", "", "", "", "", "", "", "", "", "", "Active", subPlanUrl, "[]"
    ];
    mainSheet.appendRow(newRow);

    var isMarkedUrgent = urgencyFormatted === 'Urgent (Less than 24 hr notice)';
    var isUrgentByTime = calculateIsUrgentByTime(formData.date, timestamp, ss);

    var shouldSendUrgentEmail = isMarkedUrgent || isUrgentByTime;

    if (shouldSendUrgentEmail) {
      sendUrgentCoverageEmail(ss, teacherName, formData, instructions);
    }
    
    // Check Settings to see if HR or Principal needs to be notified
    var settings = getSettings();
    var appUrl = settings["App URL"] || DEFAULT_APP_URL;
    var reasons = [];
    try {
        reasons = JSON.parse(settings["Absence Reasons"] || "[]");
    } catch(e) {
        // Ignore parse error
    }

    var hrRequired = false;
    var principalRequired = false;
    for (var rsnIdx = 0; rsnIdx < reasons.length; rsnIdx++) {
        if (reasons[rsnIdx].reason === formData.reason) {
            hrRequired = reasons[rsnIdx].hrRequired === true;
            principalRequired = reasons[rsnIdx].principalRequired === true;
            break;
        }
    }

    var displayPeriods = formatPeriodsListDisplay(formData.periods);

    if (hrRequired) {
        var hrEmails = getEmailsByRole(ss, "hr");
        for (var h = 0; h < hrEmails.length; h++) {
            var hrEmail = hrEmails[h];
            var hrSubject = teacherName + " " + formData.reason + " Absence Request";
            var hrBody = "An absence request has been submitted requiring attention.\n\n" +
                         "Teacher: " + teacherName + "\n" +
                         "Date Needed: " + formData.date + "\n" +
                         "Periods: " + displayPeriods + "\n" +
                         "Reason: " + formData.reason + "\n\n" +
                         "Instructions: " + (instructions ? instructions : "None") + "\n\n" +
                         "Please log into the Cathedral Sub App for more information: " + appUrl;

            var hrHtmlBody = "<p>An absence request has been submitted requiring attention.</p>" +
                             "<ul>" +
                             "<li><strong>Teacher:</strong> " + teacherName + "</li>" +
                             "<li><strong>Date Needed:</strong> " + formData.date + "</li>" +
                             "<li><strong>Periods:</strong> " + displayPeriods + "</li>" +
                             "<li><strong>Reason:</strong> " + formData.reason + "</li>" +
                             "</ul>" +
                             "<p><strong>Instructions:</strong> " + (instructions ? instructions : "None") + "</p>" +
                             "<p>Please log into the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";

            enqueueEmail(hrEmail, hrSubject, hrBody, { htmlBody: hrHtmlBody });
        }
    }

    if (principalRequired) {
        var prinEmails = getEmailsByRole(ss, "principal");
        for (var p = 0; p < prinEmails.length; p++) {
            var prinEmail = prinEmails[p];
            var prinSubject = teacherName + " " + formData.reason + " Absence Request";
            var prinBody = "An absence request has been submitted requiring attention.\n\n" +
                           "Teacher: " + teacherName + "\n" +
                           "Date Needed: " + formData.date + "\n" +
                           "Periods: " + displayPeriods + "\n" +
                           "Reason: " + formData.reason + "\n\n" +
                           "Instructions: " + (instructions ? instructions : "None") + "\n\n" +
                           "Please log into the Cathedral Sub App for more information: " + appUrl;

            var prinHtmlBody = "<p>An absence request has been submitted requiring attention.</p>" +
                               "<ul>" +
                               "<li><strong>Teacher:</strong> " + teacherName + "</li>" +
                               "<li><strong>Date Needed:</strong> " + formData.date + "</li>" +
                               "<li><strong>Periods:</strong> " + displayPeriods + "</li>" +
                               "<li><strong>Reason:</strong> " + formData.reason + "</li>" +
                               "</ul>" +
                               "<p><strong>Instructions:</strong> " + (instructions ? instructions : "None") + "</p>" +
                               "<p>Please log into the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";

            enqueueEmail(prinEmail, prinSubject, prinBody, { htmlBody: prinHtmlBody });
        }
    }

    logAuditAction("ABSENCE_SUBMITTED", uniqueId, "Requested coverage for " + formData.date + " (Periods: " + formData.periods + ")");
    // SEND CONFIRMATION EMAIL TO SUBMITTER

    var confSubject = "New Absence Request Confirmation";
    var confBody = "Your absence request has been successfully submitted.\n\n" +
                   "Details:\n" +
                   "Date: " + formData.date + "\n" +
                   "Periods: " + displayPeriods + "\n" +
                   "Reason: " + formData.reason + "\n" +
                   "Duration: " + formData.duration + "\n" +
                   "Instructions: " + (instructions ? instructions : "None") + "\n\n" +
                   "Return to Cathedral Sub App: " + appUrl;

    var confHtmlBody = "<p>Your absence request has been successfully submitted.</p>" +
                       "<h3>Details:</h3>" +
                       "<ul>" +
                       "<li><strong>Date:</strong> " + formData.date + "</li>" +
                       "<li><strong>Periods:</strong> " + displayPeriods + "</li>" +
                       "<li><strong>Reason:</strong> " + formData.reason + "</li>" +
                       "<li><strong>Duration:</strong> " + formData.duration + "</li>" +
                       "</ul>" +
                       "<p><strong>Instructions:</strong> " + (instructions ? instructions : "None") + "</p>" +
                       "<p><a href='" + appUrl + "'>Return to Cathedral Sub App</a></p>";

    enqueueEmail(submitterEmail, confSubject, confBody, { htmlBody: confHtmlBody });

        return {
      success: true };
  } catch (err) {
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cancels a single assigned sub duty by the sub themselves.
 */
function cancelMySubDuty(absenceId, period, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("cancelMySubDuty", e);
    return {
      success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var userEmail = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined).toLowerCase();
    var userData = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    var userName = String(userData.name).trim();

    var data = sheet.getDataRange().getValues();
    var targetUserName = userName.toLowerCase();
    for (var i = 1; i < data.length; i++) {
    if (!data[i]) continue;
      if (String(data[i][0]) === String(absenceId)) {
        var subColumnIndex = getSubColumnIndex(period);
        var assignedSub = String(data[i][subColumnIndex - 1] || "").trim();

        if (assignedSub.toLowerCase() === targetUserName) {
          var coordinatorEmail = getCoordinatorEmail(ss);

          var rosterData = getRosterDataCached(ss);
          var nameLookup = buildNameLookup(rosterData);
          var scheduleLookup = null; // Deferred to email queue to save time
          var details = getAbsenceDetailsLocal(data[i], period, scheduleLookup, nameLookup);

          sheet.getRange(i + 1, subColumnIndex).setValue("");
          logAuditAction("SUB_DUTY_CANCELLED", absenceId, "Cancelled coverage for period " + period, clientEmail);

          if (details && typeof syncSubDutyCalendarEvent === "function") {
            try {
              syncSubDutyCalendarEvent(userEmail, "Canceled", details);
            } catch (calErr) {
              console.warn("Calendar removal on cancelMySubDuty failed: " + calErr.message);
            }
          }

          if (coordinatorEmail && details) {
            var subject = "SUB CANCELLATION: " + userName + " cancelled coverage";
            var settings = getSettings();
            var appUrl = settings["App URL"] || DEFAULT_APP_URL;

            var displayCancelPeriod = formatPeriodDisplay(details.period);
            var body = userName + " has cancelled their assigned coverage.\n\n" +
                       "Date: " + details.date + "\n" +
                       "Period: " + displayCancelPeriod + "\n" +
                       "Teacher to Cover: " + details.teacherName + "\n" +
                       "Room: " + details.room + "\n" +
                       "Course: " + details.course + "\n\n" +
                       "This period is now UNFILLED. Please log into the Cathedral Sub App to reassign a sub: " + appUrl;

            var htmlBody = "<p>" + userName + " has cancelled their assigned coverage.</p>" +
                           "<ul>" +
                           "<li><strong>Date:</strong> " + details.date + "</li>" +
                           "<li><strong>Period:</strong> " + displayCancelPeriod + "</li>" +
                           "<li><strong>Teacher to Cover:</strong> " + details.teacherName + "</li>" +
                           "<li><strong>Room:</strong> " + details.room + "</li>" +
                           "<li><strong>Course:</strong> " + details.course + "</li>" +
                           "</ul>" +
                           "<p>This period is now UNFILLED. Please log into the <a href='" + appUrl + "'>Cathedral Sub App</a> to reassign a sub.</p>";

            enqueueEmail(coordinatorEmail, subject, body, { cc: userEmail, htmlBody: htmlBody });
          }

              return {
      success: true };
        } else {
          throw new Error("You are not currently assigned to this period.");
        }
      }
    }
    throw new Error("Absence ID not found.");
  } catch (err) {
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cancels an entire absence request.
 */
/**
 * Local variant of getAbsenceDetails avoiding spreadsheet lookups.
 */
function getAbsenceDetailsLocal(row, period, scheduleLookup, nameLookup) {
  var teacherEmail = String(row[2]);
  var teacherName = nameLookup[teacherEmail.toLowerCase()] || teacherEmail;
  var dateVal = row[3];
  var formattedDate = dateVal;
  var isoDate = typeof getStandardizedDateString === "function" ? getStandardizedDateString(dateVal) : "";
  if (dateVal instanceof Date) {
    formattedDate = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MMM d, yyyy");
  } else {
    try {
      formattedDate = Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "MMM d, yyyy");
    } catch(e) {
      console.error("Error formatting date: " + e.message);
      // ignore
    }
  }
  var instructions = String(row[8] || "");
  var subPlanUrl = typeof extractSubPlanUrl === "function" ? extractSubPlanUrl(row) : "";

  var roomStr = "No Class Assigned";
  var courseStr = "No Class Assigned";

  if (period) {
    var joinKey = teacherEmail.toLowerCase() + "-" + getScheduleJoinPeriod(period);
    if (scheduleLookup) {
      var scheduleInfo = (typeof lookupTeacherScheduleInfo === "function")
        ? lookupTeacherScheduleInfo(scheduleLookup, teacherEmail, period, isoDate)
        : scheduleLookup[joinKey];
      if (scheduleInfo) {
        roomStr = scheduleInfo.room || roomStr;
        courseStr = scheduleInfo.course || courseStr;
      }
    } else {
      // Defer to background process to save execution time
      roomStr = "{{ROOM|" + joinKey + "}}";
      courseStr = "{{COURSE|" + joinKey + "}}";
    }
  }

  return {
    id: String(row[0] || ""),
    teacherName: teacherName,
    date: formattedDate,
    isoDate: isoDate,
    period: period,
    room: roomStr,
    course: courseStr,
    instructions: instructions,
    subPlanUrl: subPlanUrl
  };
}

function cancelAbsence(absenceId, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("cancelAbsence", e);
    return {
      success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var rosterData = getRosterDataCached(ss);

    var subEmailLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      if (!rosterData[r]) continue;
      if (!rosterData[r]) continue;
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

    var scheduleLookup = null; // Deferred to email queue to save time
    var nameLookup = buildNameLookup(rosterData);

    var data = sheet.getDataRange().getValues();
    var targetIndex = -1;
    for (var i = 1; i < data.length; i++) {
    if (!data[i]) continue;
      if (String(data[i][0]) === String(absenceId)) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex !== -1) {
      i = targetIndex;
      var currentUserEmail = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined).toLowerCase();
      var teacherEmail = String(data[i][2]).toLowerCase();

      if (currentUserEmail !== teacherEmail) {
        var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
        assertPermission(user, "Admin Dashboard", "Unauthorized to cancel this absence.");
      }

      sheet.getRange(i + 1, 20).setValue("Canceled");
      logAuditAction("ABSENCE_CANCELLED", absenceId, "Cancelled entire absence request", clientEmail);

      var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', '10'];
      for (var pIdx = 0; pIdx < allPeriods.length; pIdx++) {
        var p = allPeriods[pIdx];
        var subIndex = getSubColumnIndex(p) - 1;
        var subName = String(data[i][subIndex] || "").trim();
        if (subName) {
          var email = subEmailLookup[subName];
          if (email) {
            var details = getAbsenceDetailsLocal(data[i], p, scheduleLookup, nameLookup);
            if (details) sendSubNotification(email, "Canceled", details);
          }
        }
      }

      var rawDate = data[i][3];
      var formattedDateForEmail = rawDate;
      if (rawDate instanceof Date) {
          formattedDateForEmail = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MMM d, yyyy");
      } else {
          try {
              formattedDateForEmail = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "MMM d, yyyy");
          } catch(e) { console.error("Error formatting date: " + e.message); }
      }

      var canceledPeriodsDisplay = formatPeriodsListDisplay(data[i][4]);
      var teacherSubject = "Absence Request Canceled";
      var teacherBody = "Your absence request for " + formattedDateForEmail + " has been canceled.\n\n" +
                        "Reason: " + data[i][5] + "\n" +
                        "Periods: " + canceledPeriodsDisplay + "\n\n" +
                        "If you have questions, please contact the sub coordinator.";

      var teacherHtml = "<p>Your absence request for <strong>" + formattedDateForEmail + "</strong> has been canceled.</p>" +
                        "<ul><li><strong>Reason:</strong> " + data[i][5] + "</li>" +
                        "<li><strong>Periods:</strong> " + canceledPeriodsDisplay + "</li></ul>" +
                        "<p>If you have questions, please contact the sub coordinator.</p>";

      enqueueEmail(teacherEmail, teacherSubject, teacherBody, {htmlBody: teacherHtml});

          return {
      success: true };
    }

    throw new Error("Absence ID not found.");
  } catch (err) {
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates an absence request.
 */
function updateAbsence(absenceId, formData, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("updateAbsence", e);
    return {
      success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var rosterData = getRosterDataCached(ss);

    var subEmailLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      if (!rosterData[r]) continue;
      if (!rosterData[r]) continue;
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

    var scheduleLookup = null; // Deferred to email queue to save time
    var nameLookup = buildNameLookup(rosterData);

    var data = sheet.getDataRange().getValues();
    var targetIndex = -1;
    for (var i = 1; i < data.length; i++) {
    if (!data[i]) continue;
      if (String(data[i][0]) === String(absenceId)) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex !== -1) {
      i = targetIndex;
      var currentUserEmail = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined).toLowerCase();
      var teacherEmail = String(data[i][2]).toLowerCase();

      if (currentUserEmail !== teacherEmail) {
        var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
        assertPermission(user, "Admin Dashboard", "Unauthorized to modify this absence.");
      }

      var oldPeriods = String(data[i][4]).split(",").map(function(p){return parsePeriod(p);});
      var newPeriods = String(formData.periods).split(",").map(function(p){return parsePeriod(p);});
      var oldDateRaw = data[i][3];
      var oldDateFormatted = "";
      if (oldDateRaw instanceof Date) {
        oldDateFormatted = Utilities.formatDate(oldDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        try {
           oldDateFormatted = Utilities.formatDate(new Date(oldDateRaw), Session.getScriptTimeZone(), "yyyy-MM-dd");
        } catch(e) {
           console.error("Error formatting date: " + e.message);
           oldDateFormatted = String(oldDateRaw);
        }
      }
      var newDate = String(formData.date);

      var dateChanged = oldDateFormatted !== newDate;

      var urgencyFormatted = formData.urgency === 'Urgent' ? 'Urgent (Less than 24 hr notice)' : 'Standard (Advanced Notice)';
      var instructions = formData.specialInstructions;
      if (formData.hrConfirmed) instructions = "[HR Docs Provided] " + instructions;
      var oldSubPlanUrl = extractSubPlanUrl(data[i]);
      var newSubPlanUrl = formData.subPlanUrl !== undefined ? String(formData.subPlanUrl || "").trim() : oldSubPlanUrl;

      // Identify sub clears
      var subClearUpdates = [];
      var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', '10'];
      for (var pIdx = 0; pIdx < allPeriods.length; pIdx++) {
        var p = allPeriods[pIdx];
        var subIndex = getSubColumnIndex(p) - 1;
        var subName = String(data[i][subIndex] || "").trim();

        if (subName) {
          var email = subEmailLookup[subName];
          var isPeriodStillNeeded = newPeriods.indexOf(String(p)) !== -1;

          if (email) {
             if (dateChanged || !isPeriodStillNeeded) {
               var cancelDetails = getAbsenceDetailsLocal(data[i], p, scheduleLookup, nameLookup);
               if (cancelDetails) {
                  cancelDetails.date = Utilities.formatDate(new Date(oldDateRaw), Session.getScriptTimeZone(), "MMM d, yyyy");
                  sendSubNotification(email, "Canceled", cancelDetails);
               }
               subClearUpdates.push(subIndex);
             } else {
               var oldReason = String(data[i][5]);
               var oldDuration = String(data[i][6]);
               var oldUrgency = String(data[i][7]);
               var oldInstructions = String(data[i][8]);

               var detailsChanged = (oldReason !== formData.reason) ||
                                    (oldDuration !== formData.duration) ||
                                    (oldUrgency !== urgencyFormatted) ||
                                    (oldInstructions !== instructions) ||
                                    (oldSubPlanUrl !== newSubPlanUrl);

               if (detailsChanged) {
                  var modDetails = getAbsenceDetailsLocal(data[i], p, scheduleLookup, nameLookup);
                  if(modDetails) {
                    modDetails.instructions = instructions;
                    modDetails.subPlanUrl = newSubPlanUrl;
                    sendSubNotification(email, "Modified", modDetails);
                  }
               }
             }
          } else if (dateChanged || !isPeriodStillNeeded) {
             subClearUpdates.push(subIndex);
          }
        }
      }

      // Build a contiguous array for columns 4 through 17 (index 3 to 16)
      // D is 4 (Date), E is 5 (Periods), F is 6 (Reason), G is 7 (Duration), H is 8 (Urgency), I is 9 (Instructions)
      // J to S are 10 to 19 (Sub 1 to Advisory)
      var updatedRowBlock = [
        formData.date, "'" + formData.periods, formData.reason,
        formData.duration, urgencyFormatted, instructions
      ];

      // Append the existing sub data to the block
      for (var sIdx = 9; sIdx <= 18; sIdx++) {
         updatedRowBlock.push(data[i][sIdx]);
      }

      // Clear the necessary sub periods in the block
      for (var s = 0; s < subClearUpdates.length; s++) {
         // subIndex is between 9 and 18.
         // updatedRowBlock[0] maps to column 4.
         // updatedRowBlock[subIndex - 3] maps to column subIndex + 1.
         updatedRowBlock[subClearUpdates[s] - 3] = "";
      }

      // Single API call to update both the basic info and all sub clearing at once
      sheet.getRange(i + 1, 4, 1, 16).setValues([updatedRowBlock]);
      if (formData.subPlanUrl !== undefined) {
        sheet.getRange(i + 1, 21).setValue(newSubPlanUrl);
      }
      logAuditAction("ABSENCE_UPDATED", absenceId, "Updated absence details (Date: " + formData.date + ", Periods: " + formData.periods + ")", clientEmail);


      var rawDate = data[i][3];
      var formattedDateForEmail = rawDate;
      if (rawDate instanceof Date) {
          formattedDateForEmail = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MMM d, yyyy");
      } else {
          try {
              formattedDateForEmail = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "MMM d, yyyy");
          } catch(e) { console.error("Error formatting date: " + e.message); }
      }

      var updatedPeriodsDisplay = formatPeriodsListDisplay(formData.periods);
      var teacherSubject = "Absence Request Updated";
      var teacherBody = "Your absence request for " + formattedDateForEmail + " has been updated.\n\n" +
                        "Updated Details:\n" +
                        "Date: " + formData.date + "\n" +
                        "Periods: " + updatedPeriodsDisplay + "\n" +
                        "Reason: " + formData.reason + "\n" +
                        "Duration: " + formData.duration + "\n\n" +
                        "If you have questions, please contact the sub coordinator.";

      var teacherHtml = "<p>Your absence request for <strong>" + formattedDateForEmail + "</strong> has been updated.</p>" +
                        "<ul><li><strong>Date:</strong> " + formData.date + "</li>" +
                        "<li><strong>Periods:</strong> " + updatedPeriodsDisplay + "</li>" +
                        "<li><strong>Reason:</strong> " + formData.reason + "</li>" +
                        "<li><strong>Duration:</strong> " + formData.duration + "</li></ul>" +
                        "<p>If you have questions, please contact the sub coordinator.</p>";

      enqueueEmail(teacherEmail, teacherSubject, teacherBody, {htmlBody: teacherHtml});

          return {
      success: true };
    }

    throw new Error("Absence ID not found.");
  } catch (err) {
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Helper to get teacher name from email.
 */
/**
 * Helper to get full absence details.
 */
// Global cache for getAbsenceDetails to improve performance
var _cachedNameLookup = null;
var _cachedScheduleLookup = null;

function getAbsenceDetails(absenceId, period, optionalData, clientEmail) {
  var ss = getSS();
  var data;

  if (optionalData && Array.isArray(optionalData) && optionalData.length > 0) {
    data = optionalData;
  } else {
    var sheet = getSheetOrThrow(ss, "Absence Requests");
    if (!sheet) return null;
    data = sheet.getDataRange().getValues();
  }

  if (!_cachedNameLookup) {
    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    if (rosterSheet) {
      var rosterData = rosterSheet.getDataRange().getValues();
      _cachedNameLookup = buildNameLookup(rosterData);
    } else {
      _cachedNameLookup = {};
    }
  }

  if (!_cachedScheduleLookup) {
    var scheduleData = getMasterScheduleData();
    if (scheduleData.length > 0) {
      _cachedScheduleLookup = buildScheduleLookup(scheduleData);
    } else {
      _cachedScheduleLookup = {};
    }
  }

  for (var i = 1; i < data.length; i++) {
  if (!data[i]) continue;
    if (String(data[i][0]) === String(absenceId)) {
      var details = getAbsenceDetailsLocal(data[i], period, _cachedScheduleLookup, _cachedNameLookup);
      details.rowIndex = i + 1;
      return details;
    }
  }
  return null;
}


/**
 * Assigns a substitute to a specific period for an absence request.
 */
function assignSubToPeriod(absenceId, period, subName, forceOverride, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("assignSubToPeriod", e);
    return {
      success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");
    var callerUser = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    var callerRoleLower = String(callerUser.role || "").toLowerCase().trim();
    var settingsForPerm = getSettings(ss);
    var rolePermsStr = settingsForPerm["RolePermissions"] || "{}";
    var canManageAllSubs = (
      callerRoleLower === "admin" ||
      callerRoleLower === "sub coordinator" ||
      hasPermission(callerRoleLower, "Admin Dashboard", rolePermsStr) ||
      hasPermission(callerRoleLower, "Today at a Glance", rolePermsStr)
    );

    var rosterData = getRosterDataCached(ss);

    var subEmailLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      if (!rosterData[r]) continue;
      if (!rosterData[r]) continue;
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

    var nameLookup = buildNameLookup(rosterData);

    // Re-read data under lock
    var data = sheet.getDataRange().getValues();

    var targetIndex = -1;
    for (var i = 1; i < data.length; i++) {
    if (!data[i]) continue;
      if (String(data[i][0]) === String(absenceId)) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex !== -1) {
      i = targetIndex;
      var periodsRequested = String(data[i][4]).split(",").map(function(p) { return parsePeriod(p); });

      if (periodsRequested.indexOf(parsePeriod(period)) === -1) {
          throw new Error("Period " + period + " was not requested for this absence.");
      }

      var subColumnIndex = getSubColumnIndex(period); // 1-based index for Apps Script Ranges: Col J is 10 (Period 1 Sub)

      var existingSub = String(data[i][subColumnIndex - 1] || "").trim(); // array is 0-indexed, so subColumnIndex - 1
      var newSub = String(subName || "").trim();

      // Normalize "No Sub Needed" early for consistent authorization & comparison checks
      if (newSub.toLowerCase() === "no sub needed") {
          newSub = "No Sub Needed";
      }

      if (!canManageAllSubs) {
        var callerCleanName = String(callerUser.name || "").trim().toLowerCase();
        var isSelfSignup = (newSub !== "" && newSub !== "No Sub Needed" && newSub.toLowerCase() === callerCleanName && existingSub === "");
        if (!isSelfSignup) {
          throw new Error("Unauthorized to assign or modify substitute coverage for other users.");
        }
      }

      if (existingSub === newSub) {
             return {
      success: true }; // No change
      }

      // Double check for race condition
      if (existingSub !== "" && newSub !== "") {
        throw new Error("Sorry, this job was just filled by someone else!");
      }

      // Get role of the assigned person
      var newSubRole = "";
      if (newSub !== "" && newSub !== "No Sub Needed") {
        for (var roIdx = 1; roIdx < rosterData.length; roIdx++) {
          if (String(rosterData[roIdx][0]).trim() === newSub) {
            newSubRole = String(rosterData[roIdx][2]).trim(); // Role is at index 2
            break;
          }
        }
      }

      // Check if the new sub is available based on Dates and SubstituteAvailability or Master Schedule
      if (newSub !== "" && newSub !== "No Sub Needed" && !forceOverride) {
        let newSubEmail = (subEmailLookup[newSub] || "").toLowerCase();
        if (newSubEmail !== "") {
           var isSubstitute = newSubRole.indexOf("Substitute") !== -1;
           let targetDateRaw = data[i][3];
           let targetDateStr = (targetDateRaw instanceof Date) ? Utilities.formatDate(targetDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(targetDateRaw).trim();

           if (isSubstitute) {
             // For substitutes, check SubstituteAvailability sheet
             var allSubAvail = getAllSubstituteAvailability();
             var subAvail = allSubAvail[newSubEmail] || {};
             var availStatus = subAvail[targetDateStr] || "Not Available";

             if (availStatus !== "Available" && availStatus !== "AM Only" && availStatus !== "PM Only") {
                 throw new Error(JSON.stringify({type: "AVAILABILITY_ERROR", message: "Sub not listed as available, proceed?"}));
             }

             // Fetch day color from cached Dates lookup
             var dayColor = "Green"; // default
             try {
               var cachedDateColors = getDateColorsCached(ss);
               if (cachedDateColors && cachedDateColors[targetDateStr]) {
                 dayColor = String(cachedDateColors[targetDateStr]).trim();
               }
             } catch(e) {
               console.error("Failed to fetch day color: " + e.message);
             }

             var p = parseInt(period);
             if (availStatus === "AM Only") {
                 if ((dayColor === "Green" && p > 4) || (dayColor === "Blue" && p > 2) || (dayColor === "Gold" && p > 6)) {
                     throw new Error(JSON.stringify({type: "AVAILABILITY_ERROR", message: "Sub not listed as available, proceed?"}));
                 }
             }
             if (availStatus === "PM Only") {
                 if ((dayColor === "Green" && p <= 4) || (dayColor === "Blue" && p <= 2) || (dayColor === "Gold" && p <= 6)) {
                     throw new Error(JSON.stringify({type: "AVAILABILITY_ERROR", message: "Sub not listed as available, proceed?"}));
                 }
             }
           } else {
             // For teachers/others, check master schedule for the date's semester term
             var scheduleData = getMasterScheduleData();
             var teacherSchedule = [];
             var targetTermForDate = (typeof resolveTermIdForDate === "function")
               ? String(resolveTermIdForDate(targetDateStr, settingsForPerm) || "").trim()
               : "";
             var fullYearTermForDate = (!isNaN(parseInt(targetTermForDate, 10)) && parseInt(targetTermForDate, 10) >= 100)
               ? String(Math.floor(parseInt(targetTermForDate, 10) / 100) * 100)
               : "";
             if (scheduleData && scheduleData.length > 0) {
               var headers = scheduleData[0];
               var emailIdx = headers.indexOf("EMAIL_ADDR");
               var periodIdx = headers.indexOf("PERIOD");
               var termColIdx = headers.indexOf("TERM");
               if (emailIdx > -1 && periodIdx > -1) {
                 for (var s = 1; s < scheduleData.length; s++) {
                   if (!scheduleData[s] || scheduleData[s].length <= Math.max(emailIdx, periodIdx)) continue;
                   if (String(scheduleData[s][emailIdx]).toLowerCase().trim() === newSubEmail) {
                     var rowTermVal = (termColIdx > -1 && scheduleData[s][termColIdx] !== undefined)
                       ? String(scheduleData[s][termColIdx]).trim()
                       : "";
                     if (rowTermVal && targetTermForDate && rowTermVal !== targetTermForDate && rowTermVal !== fullYearTermForDate) {
                       continue;
                     }
                     var pVal = String(scheduleData[s][periodIdx]).trim();
                     var joinP = getScheduleJoinPeriod(pVal);
                     if (teacherSchedule.indexOf(joinP) === -1) {
                       teacherSchedule.push(joinP);
                     }
                   }
                 }
               }
             }
             var requestedJoinPeriod = getScheduleJoinPeriod(period);
             if (teacherSchedule.indexOf(requestedJoinPeriod) !== -1) {
                 // The teacher has a class during this period in this semester
                 throw new Error(JSON.stringify({type: "AVAILABILITY_ERROR", message: "Sub not listed as available, proceed?"}));
             }
           }
        }
      }

      // Check if the new sub is absent for a full day on the same date
      if (newSub !== "" && newSub !== "No Sub Needed") {
        let targetDateRaw = data[i][3]; // Date object or string
        let targetDateStr = (targetDateRaw instanceof Date) ? Utilities.formatDate(targetDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(targetDateRaw).trim();
        let newSubEmail = (subEmailLookup[newSub] || "").toLowerCase();

        if (newSubEmail !== "") {
          for (var j = 1; j < data.length; j++) {
          if (!data[j]) continue;
            var rowEmail = String(data[j][2] || "").toLowerCase();
            if (rowEmail === "") continue;

            var rowStatus = String(data[j][19] || "Active");
            var rowDuration = String(data[j][6] || "").trim();

            if (rowEmail === newSubEmail && rowStatus !== "Canceled" && rowDuration === "Full Day") {
              let rowDateRaw = data[j][3];
              var rowDateStr = (rowDateRaw instanceof Date) ? Utilities.formatDate(rowDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(rowDateRaw).trim();

              if (rowDateStr === targetDateStr) {
                var currentUserEmail = (getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined) || "").toLowerCase();
                if (currentUserEmail === newSubEmail) {
                  throw new Error("Cannot sign up due to your own absence.");
                } else {
                  throw new Error("Cannot assign " + newSub + " as a substitute because they have an absence request on this date.");
                }
              }
            }
          }
        }
      }

      // Cancel existing sub if there is one (and we are clearing it)
      if (existingSub) {
         var existingEmail = subEmailLookup[existingSub];
         if (existingEmail) {
            let scheduleLookup = null; // Deferred to email queue to save time
            let details = getAbsenceDetailsLocal(data[i], period, scheduleLookup, nameLookup);
            sendSubNotification(existingEmail, 'Canceled', details);
         }
      }

      // Write the subname
      sheet.getRange(i + 1, subColumnIndex).setValue(newSub);
      logAuditAction("SUB_ASSIGNED", absenceId, "Assigned " + (newSub || "NO ONE") + " to period " + period, clientEmail);

      // Notify new sub if there is one
      if (newSub && newSub !== "No Sub Needed") {
         var newEmail = subEmailLookup[newSub];
         if (newEmail) {
            let scheduleLookup = null; // Deferred to email queue to save time
            let details = getAbsenceDetailsLocal(data[i], period, scheduleLookup, nameLookup);
            sendSubNotification(newEmail, 'Assigned', details);
         }
      }

          return {
      success: true };
    }

    throw new Error("Absence Request ID not found.");
  } catch (err) {
    if (err.message && err.message.indexOf("AVAILABILITY_ERROR") !== -1) {
       throw err; // Re-throw to allow frontend withFailureHandler to catch it for the override confirmation
    }
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function saveSubFeedback(absenceId, period, rating, note, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("saveSubFeedback", e);
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var email = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    var targetEmail = String(email).toLowerCase();

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    var subName = "Unknown Sub";
    for (var r = 1; r < rosterData.length; r++) {
       if (rosterData[r] && String(rosterData[r][1]).toLowerCase().trim() === targetEmail) {
           subName = String(rosterData[r][0]).trim();
           break;
       }
    }

    var data = sheet.getDataRange().getValues();
    var targetIndex = -1;
    for (var i = 1; i < data.length; i++) {
    if (!data[i]) continue;
      if (String(data[i][0]) === String(absenceId)) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex === -1) {
       throw new Error("Could not find absence request with ID " + absenceId);
    }

    var row = data[targetIndex];
    let subFeedbackRaw = extractSubFeedbackRaw(row);
    var subFeedbackParsed = [];
    try {
      subFeedbackParsed = JSON.parse(subFeedbackRaw);
    } catch(e) {
      // Ignore parse error
    }

    var normalizedPeriod = parsePeriod(period);
    var existingIndex = -1;
    for (var j = 0; j < subFeedbackParsed.length; j++) {
       if (parsePeriod(subFeedbackParsed[j].period) === normalizedPeriod && String(subFeedbackParsed[j].subName).toLowerCase() === subName.toLowerCase()) {
           existingIndex = j;
           break;
       }
    }

    var feedbackObj = {
        subName: subName,
        period: normalizedPeriod,
        rating: parseInt(rating) || 0,
        note: String(note || "").trim()
    };

    if (existingIndex !== -1) {
       subFeedbackParsed[existingIndex] = feedbackObj;
    } else {
       subFeedbackParsed.push(feedbackObj);
    }

    // Write Sub Feedback to Column 22 (index 21), keeping Column 21 (index 20) for Sub Plan URL
    sheet.getRange(targetIndex + 1, 22).setValue(JSON.stringify(subFeedbackParsed));
    if (row.length > 20 && String(row[20] || "").trim().indexOf("[") === 0) {
       sheet.getRange(targetIndex + 1, 21).setValue("");
    }

    var teacherEmail = String(row[2]).toLowerCase().trim();
    var absenceDateVal = row[3];
    var absenceDateStr = (absenceDateVal instanceof Date) ? Utilities.formatDate(absenceDateVal, Session.getScriptTimeZone(), "M/d/yyyy") : String(absenceDateVal);

    var nameLookup = buildNameLookup(rosterData);
    var teacherName = nameLookup[teacherEmail] || teacherEmail;
    if (teacherName.includes(",")) {
        var parts = teacherName.split(",");
        if (parts.length > 1) teacherName = parts[1].trim() + " " + parts[0].trim();
    }

    var displayFeedbackPeriod = formatPeriodDisplay(period);
    var displayFeedbackPeriodLabel = formatPeriodLabelDisplay(period);
    var subject = "Substitute Feedback Received - " + displayFeedbackPeriodLabel;
    var parsedRating = parseInt(rating) || 0;
    var ratingStars = "★".repeat(parsedRating) + "☆".repeat(Math.max(0, 5 - parsedRating));

    var bodyText = "Hello " + teacherName + ",\n\n";
    bodyText += subName + " just submitted feedback for covering your class on " + absenceDateStr + ":\n\n";
    bodyText += "- Period: " + displayFeedbackPeriod + "\n";
    bodyText += "- Rating: " + rating + "/5\n";
    bodyText += "- Notes: " + (note || "No notes") + "\n\n";
    bodyText += "Thank you!";

    var bodyHtml = "<p>Hello " + teacherName + ",</p>";
    bodyHtml += "<p>" + subName + " just submitted feedback for covering your class on " + absenceDateStr + ":</p>";
    bodyHtml += "<ul>";
    bodyHtml += "<li><strong>Period:</strong> " + displayFeedbackPeriod + "</li>";
    bodyHtml += "<li><strong>Rating:</strong> " + ratingStars + " (" + rating + "/5)</li>";
    bodyHtml += "<li><strong>Notes:</strong> " + (note || "<em>No notes provided</em>") + "</li>";
    bodyHtml += "</ul><p>Thank you!</p>";

    enqueueEmail(teacherEmail, subject, bodyText, {htmlBody: bodyHtml});

    logAuditAction("SUB_FEEDBACK_SAVED", absenceId, "Saved Sub Feedback - Period: " + period + ", Rating: " + rating);

    return { success: true };

  } catch (err) {
    notifyAdminOfError("saveSubFeedback", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

