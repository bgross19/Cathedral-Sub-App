const fs = require('fs');
let code = fs.readFileSync('code.gs', 'utf8');

function replaceFunction(codeStr, funcName, newImpl) {
    const regex = new RegExp(`function\\s+${funcName}\\s*\\([^)]*\\)\\s*\\{`, 'g');
    const match = regex.exec(codeStr);
    if (!match) return codeStr;

    let startIdx = match.index;
    let braceCount = 0;
    let i = startIdx;
    let foundStartBrace = false;

    while (i < codeStr.length) {
        if (codeStr[i] === '{') {
            braceCount++;
            foundStartBrace = true;
        } else if (codeStr[i] === '}') {
            braceCount--;
        }

        if (foundStartBrace && braceCount === 0) {
            break;
        }
        i++;
    }

    let endIdx = i + 1;
    let oldFunc = codeStr.substring(startIdx, endIdx);
    return codeStr.replace(oldFunc, newImpl);
}

// 1. updateAbsence
let updateAbsenceNew = `function updateAbsence(absenceId, formData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];

    var subEmailLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

    var scheduleLookup = buildScheduleLookup(scheduleData);
    var nameLookup = buildNameLookup(rosterData);

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        var currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();
        var teacherEmail = String(data[i][2]).toLowerCase();

        if (currentUserEmail !== teacherEmail) {
          var user = getUserData(ss);
          assertRole(user, ["admin", "sub coordinator"], "Unauthorized to modify this absence.");
        }

        var oldPeriods = String(data[i][4]).split(",").map(function(p){return p.trim()});
        var newPeriods = String(formData.periods).split(",").map(function(p){return p.trim()});
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

        // Update basic info
        sheet.getRange(i + 1, 4, 1, 6).setValues([[
          formData.date, "'" + formData.periods, formData.reason,
          formData.duration, urgencyFormatted, instructions
        ]]);

        // Notify subs
        for (var p = 1; p <= 8; p++) {
          var subIndex = 8 + p;
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
                 sheet.getRange(i + 1, subIndex + 1).setValue("");
               } else {
                 var oldReason = String(data[i][5]);
                 var oldDuration = String(data[i][6]);
                 var oldUrgency = String(data[i][7]);
                 var oldInstructions = String(data[i][8]);

                 var detailsChanged = (oldReason !== formData.reason) ||
                                      (oldDuration !== formData.duration) ||
                                      (oldUrgency !== urgencyFormatted) ||
                                      (oldInstructions !== instructions);

                 if (detailsChanged) {
                    var modDetails = getAbsenceDetailsLocal(data[i], p, scheduleLookup, nameLookup);
                    if(modDetails) {
                      modDetails.instructions = instructions;
                      sendSubNotification(email, "Modified", modDetails);
                    }
                 }
               }
            } else if (dateChanged || !isPeriodStillNeeded) {
               sheet.getRange(i + 1, subIndex + 1).setValue("");
            }
          }
        }

        if (currentUserEmail !== teacherEmail) {
           var rawDate = data[i][3];
           var formattedDateForEmail = rawDate;
           if (rawDate instanceof Date) {
               formattedDateForEmail = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MMM d, yyyy");
           } else {
               try {
                   formattedDateForEmail = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "MMM d, yyyy");
               } catch(e) { console.error("Error formatting date: " + e.message); }
           }

           var teacherSubject = "Absence Request Updated by Administrator";
           var teacherBody = "Your absence request for " + formattedDateForEmail + " has been updated by an administrator.\\n\\n" +
                             "Updated Details:\\n" +
                             "Date: " + formData.date + "\\n" +
                             "Periods: " + formData.periods + "\\n" +
                             "Reason: " + formData.reason + "\\n" +
                             "Duration: " + formData.duration + "\\n\\n" +
                             "If you have questions, please contact the sub coordinator.";

           var teacherHtml = "<p>Your absence request for <strong>" + formattedDateForEmail + "</strong> has been updated by an administrator.</p>" +
                             "<ul><li><strong>Date:</strong> " + formData.date + "</li>" +
                             "<li><strong>Periods:</strong> " + formData.periods + "</li>" +
                             "<li><strong>Reason:</strong> " + formData.reason + "</li>" +
                             "<li><strong>Duration:</strong> " + formData.duration + "</li></ul>" +
                             "<p>If you have questions, please contact the sub coordinator.</p>";

           sendEmailHelper(teacherEmail, teacherSubject, teacherBody, {htmlBody: teacherHtml});
        }

        return { success: true };
      }
    }
    throw new Error("Absence ID not found.");
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}`;

// 2. cancelAbsence
let cancelAbsenceNew = `function cancelAbsence(absenceId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];

    var subEmailLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

    var scheduleLookup = buildScheduleLookup(scheduleData);
    var nameLookup = buildNameLookup(rosterData);

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        var currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();
        var teacherEmail = String(data[i][2]).toLowerCase();

        if (currentUserEmail !== teacherEmail) {
          var user = getUserData(ss);
          assertRole(user, ["admin", "sub coordinator"], "Unauthorized to cancel this absence.");
        }

        sheet.getRange(i + 1, 18).setValue("Canceled");

        for (var p = 1; p <= 8; p++) {
          var subIndex = 8 + p;
          var subName = String(data[i][subIndex] || "").trim();
          if (subName) {
            var email = subEmailLookup[subName];
            if (email) {
              var details = getAbsenceDetailsLocal(data[i], p, scheduleLookup, nameLookup);
              if (details) sendSubNotification(email, "Canceled", details);
            }
          }
        }

        if (currentUserEmail !== teacherEmail) {
           var rawDate = data[i][3];
           var formattedDateForEmail = rawDate;
           if (rawDate instanceof Date) {
               formattedDateForEmail = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MMM d, yyyy");
           } else {
               try {
                   formattedDateForEmail = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "MMM d, yyyy");
               } catch(e) { console.error("Error formatting date: " + e.message); }
           }

           var teacherSubject = "Absence Request Canceled by Administrator";
           var teacherBody = "Your absence request for " + formattedDateForEmail + " has been canceled by an administrator.\\n\\n" +
                             "Reason: " + data[i][5] + "\\n" +
                             "Periods: " + data[i][4] + "\\n\\n" +
                             "If you have questions, please contact the sub coordinator.";

           var teacherHtml = "<p>Your absence request for <strong>" + formattedDateForEmail + "</strong> has been canceled by an administrator.</p>" +
                             "<ul><li><strong>Reason:</strong> " + data[i][5] + "</li>" +
                             "<li><strong>Periods:</strong> " + data[i][4] + "</li></ul>" +
                             "<p>If you have questions, please contact the sub coordinator.</p>";

           sendEmailHelper(teacherEmail, teacherSubject, teacherBody, {htmlBody: teacherHtml});
        }

        return { success: true };
      }
    }
    throw new Error("Absence ID not found.");
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}`;

// 3. cancelMySubDuty
let cancelMySubDutyNew = `function cancelMySubDuty(absenceId, period) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var userEmail = Session.getActiveUser().getEmail().toLowerCase();
    var userData = getUserData(ss);
    var userName = String(userData.name).trim();

    var data = sheet.getDataRange().getValues();
    var targetUserName = userName.toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        var subColumnIndex = 10 + parseInt(period) - 1;
        var assignedSub = String(sheet.getRange(i + 1, subColumnIndex).getValue() || "").trim();

        if (assignedSub.toLowerCase() === targetUserName) {
          var coordinatorEmail = getCoordinatorEmail(ss);
          var details = getAbsenceDetails(absenceId, period);

          sheet.getRange(i + 1, subColumnIndex).setValue("");

          if (coordinatorEmail && details) {
            var subject = "SUB CANCELLATION: " + userName + " cancelled coverage";
            var settings = getSettings();
            var appUrl = settings["App URL"] || "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec";

            var body = userName + " has cancelled their assigned coverage.\\n\\n" +
                       "Date: " + details.date + "\\n" +
                       "Period: " + details.period + "\\n" +
                       "Teacher to Cover: " + details.teacherName + "\\n" +
                       "Room: " + details.room + "\\n" +
                       "Course: " + details.course + "\\n\\n" +
                       "This period is now UNFILLED. Please log into the Cathedral Sub App to reassign a sub: " + appUrl;

            var htmlBody = "<p>" + userName + " has cancelled their assigned coverage.</p>" +
                           "<ul>" +
                           "<li><strong>Date:</strong> " + details.date + "</li>" +
                           "<li><strong>Period:</strong> " + details.period + "</li>" +
                           "<li><strong>Teacher to Cover:</strong> " + details.teacherName + "</li>" +
                           "<li><strong>Room:</strong> " + details.room + "</li>" +
                           "<li><strong>Course:</strong> " + details.course + "</li>" +
                           "</ul>" +
                           "<p>This period is now UNFILLED. Please log into the <a href='" + appUrl + "'>Cathedral Sub App</a> to reassign a sub.</p>";

            sendEmailHelper(coordinatorEmail, subject, body, { cc: userEmail, htmlBody: htmlBody });
          }

          return { success: true };
        } else {
          throw new Error("You are not currently assigned to this period.");
        }
      }
    }
    throw new Error("Absence ID not found.");
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}`;

// 4. assignSubToPeriod
let assignSubToPeriodNew = `function assignSubToPeriod(absenceId, period, subName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];

    var subEmailLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

    var scheduleLookup = buildScheduleLookup(scheduleData);
    var nameLookup = buildNameLookup(rosterData);

    // Re-read data under lock
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        var periodsRequested = String(data[i][4]).split(",").map(function(p) { return p.trim(); });

        if (periodsRequested.indexOf(String(period)) === -1) {
            throw new Error("Period " + period + " was not requested for this absence.");
        }

        var subColumnIndex = 10 + parseInt(period) - 1; // 1-based index for Apps Script Ranges: Col J is 10 (Period 1 Sub)

        var existingSub = String(data[i][subColumnIndex - 1] || "").trim(); // array is 0-indexed, so subColumnIndex - 1
        var newSub = String(subName || "").trim();

        if (existingSub === newSub) {
           return { success: true }; // No change
        }

        // Double check for race condition
        if (existingSub !== "" && newSub !== "") {
          throw new Error("Sorry, this job was just filled by someone else!");
        }

        // Get details for email
        var details = getAbsenceDetailsLocal(data[i], period, scheduleLookup, nameLookup);

        // Cancel existing sub if there is one (and we are clearing it)
        if (existingSub && details) {
           var existingEmail = subEmailLookup[existingSub];
           if (existingEmail) {
              sendSubNotification(existingEmail, 'Canceled', details);
           }
        }

        // Write the subname
        sheet.getRange(i + 1, subColumnIndex).setValue(newSub);

        // Notify new sub if there is one
        if (newSub && details) {
           var newEmail = subEmailLookup[newSub];
           if (newEmail) {
              sendSubNotification(newEmail, 'Assigned', details);
           }
        }

        return { success: true };
      }
    }

    throw new Error("Absence Request ID not found.");
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}`;

// 5. submitAbsence
let submitAbsenceNew = `function submitAbsence(formData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var mainSheet = ss.getSheetByName("Absence Requests");

    var urgencyFormatted = formData.urgency === 'Urgent' ? 'Urgent (Less than 24 hr notice)' : 'Standard (Advanced Notice)';
    var instructions = formData.specialInstructions;
    if (formData.hrConfirmed) instructions = "[HR Docs Provided] " + instructions;

    var timestamp = new Date();
    var email = Session.getActiveUser().getEmail();
    var uniqueId = Utilities.getUuid();

    var newRow = [
      uniqueId, timestamp, email, formData.date, "'" + formData.periods,
      formData.reason, formData.duration, urgencyFormatted, instructions,
      "", "", "", "", "", "", "", "", "Active"
    ];
    mainSheet.appendRow(newRow);

    var teacherName = getUserData(ss).name;

    var isMarkedUrgent = urgencyFormatted === 'Urgent (Less than 24 hr notice)';
    var isUrgentByTime = calculateIsUrgentByTime(formData.date, timestamp, ss);

    var shouldSendUrgentEmail = isMarkedUrgent || isUrgentByTime;

    if (shouldSendUrgentEmail) {
      sendUrgentCoverageEmail(ss, teacherName, formData, instructions);
    }

    // SEND CONFIRMATION EMAIL TO SUBMITTER
    var settings = getSettings();
    var appUrl = settings["App URL"] || "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec";

    var confSubject = "New Absence Request Confirmation";
    var confBody = "Your absence request has been successfully submitted.\\n\\n" +
                   "Details:\\n" +
                   "Date: " + formData.date + "\\n" +
                   "Periods: " + formData.periods + "\\n" +
                   "Reason: " + formData.reason + "\\n" +
                   "Duration: " + formData.duration + "\\n" +
                   "Instructions: " + (instructions ? instructions : "None") + "\\n\\n" +
                   "Return to Cathedral Sub App: " + appUrl;

    var confHtmlBody = "<p>Your absence request has been successfully submitted.</p>" +
                       "<h3>Details:</h3>" +
                       "<ul>" +
                       "<li><strong>Date:</strong> " + formData.date + "</li>" +
                       "<li><strong>Periods:</strong> " + formData.periods + "</li>" +
                       "<li><strong>Reason:</strong> " + formData.reason + "</li>" +
                       "<li><strong>Duration:</strong> " + formData.duration + "</li>" +
                       "</ul>" +
                       "<p><strong>Instructions:</strong> " + (instructions ? instructions : "None") + "</p>" +
                       "<p><a href='" + appUrl + "'>Return to Cathedral Sub App</a></p>";

    sendEmailHelper(email, confSubject, confBody, { htmlBody: confHtmlBody });

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}`;

code = replaceFunction(code, 'updateAbsence', updateAbsenceNew);
code = replaceFunction(code, 'cancelAbsence', cancelAbsenceNew);
code = replaceFunction(code, 'cancelMySubDuty', cancelMySubDutyNew);
code = replaceFunction(code, 'assignSubToPeriod', assignSubToPeriodNew);
code = replaceFunction(code, 'submitAbsence', submitAbsenceNew);

fs.writeFileSync('code.gs', code);
console.log("Patched code.gs successfully.");
