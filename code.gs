/**
 * Sets up the database headers in the Google Sheet.
 * Run this function once from the Apps Script editor.
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Absence Requests");
  if (!sheet) {
    sheet = ss.insertSheet("Absence Requests");
  }

  var headers = [
    "ID", "Timestamp", "Email", "Date", "Periods", "Reason", "Duration",
    "Urgency", "Instructions", "Period 1 Sub", "Period 2 Sub", "Period 3 Sub",
    "Period 4 Sub", "Period 5 Sub", "Period 6 Sub", "Period 7 Sub", "Period 8 Sub", "Status"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Optionally delete Split Responses sheet as it's no longer used
  var splitSheet = ss.getSheetByName("Split Responses");
  if (splitSheet) {
    // ss.deleteSheet(splitSheet); // Uncomment to delete automatically, or delete manually.
  }
}

/**
 * Serves the web app.
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = Session.getActiveUser().getEmail();
  
  var htmlOutput = template.evaluate();
  htmlOutput.setTitle('Cathedral Sub Coverage');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  return htmlOutput;
}

/**
 * Grabs the user's name and role on startup.
 */
function getUserData(ss) {
  var email = Session.getActiveUser().getEmail();
  var ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  
  var rosterSheet = ss.getSheetByName("Staff Roster");
  var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
  var name = "Teacher"; 
  
  for (var i = 1; i < rosterData.length; i++) {
    if (String(rosterData[i][1]).toLowerCase() === String(email).toLowerCase()) { 
      name = rosterData[i][0]; 
      break;
    }
  }
  
  var roleSheet = ss.getSheetByName("User Roles");
  var roleData = roleSheet ? roleSheet.getDataRange().getValues() : [];
  var role = "User"; 
  
  for (var j = 1; j < roleData.length; j++) {
    if (String(roleData[j][0]).toLowerCase() === String(email).toLowerCase()) { 
      role = roleData[j][1]; 
      break;
    }
  }
  
  return { name: String(name), role: String(role), email: String(email) };
}

/**
 * Fetches the logged-in user's upcoming absences.
 */
function getMyAbsences() {
  try {
    var email = Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Absence Requests");
    
    if (!sheet) return []; 
    
    var data = sheet.getDataRange().getValues();
    var myAbsences = [];
    
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][17] || 'Active');
      if (status === 'Canceled') continue;

      if (String(data[i][2]).toLowerCase() === String(email).toLowerCase()) { // Email is index 2 now
        var dateVal = data[i][3]; // Date is index 3 now
        if (!dateVal) continue; 

        var formattedDate = "Unknown Date";
        if (dateVal instanceof Date) {
           formattedDate = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MMM d, yyyy");
        } else {
           formattedDate = new Date(dateVal).toLocaleDateString();
        }

        var urgencyStr = String(data[i][7] || ''); // Urgency is index 7 now

        var yyyymmdd = "";
        if (dateVal instanceof Date) {
            yyyymmdd = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
            try {
                yyyymmdd = Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "yyyy-MM-dd");
            } catch(e) {
                yyyymmdd = String(dateVal); // fallback
            }
        }

        // Forced strings to ensure perfect serialization
        myAbsences.push({
          id: String(data[i][0]),
          date: String(formattedDate),
          rawDateString: String(dateVal), // Original raw val
          formDateString: String(yyyymmdd), // Safe formatting for input type="date"
          periods: String(data[i][4]), // Periods is index 4 now
          reason: String(data[i][5]),  // Reason is index 5 now
          urgency: urgencyStr.includes('Urgent') ? 'Urgent' : 'Standard',
          duration: String(data[i][6]),
          instructions: String(data[i][8])
        });
      }
    }
    
    return myAbsences.reverse(); 
  } catch (err) {
    throw new Error("MyAbsences Error: " + err.message);
  }
}

/**
 * Fetches the sub duties assigned to the logged-in user over the next calendar week.
 */
function getMySubDuties() {
  try {
    var userEmail = Session.getActiveUser().getEmail().toLowerCase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userData = getUserData(ss);
    var userName = String(userData.name).trim().toLowerCase();

    var mainSheet = ss.getSheetByName("Absence Requests");
    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");

    if (!mainSheet) return [];

    var data = mainSheet.getDataRange().getValues();
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    // Master Schedule data mapping
    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];
    var scheduleLookup = {};
    if (scheduleData.length > 0) {
      var headers = scheduleData[0];
      var joinIdx = headers.indexOf("EMAIL_PERIOD_JOIN");
      var roomIdx = headers.indexOf("ROOM");
      var courseIdx = headers.indexOf("COURSE_NAMES");

      if (joinIdx > -1) {
        for (var s = 1; s < scheduleData.length; s++) {
          var joinKey = String(scheduleData[s][joinIdx]).toLowerCase();
          var room = roomIdx > -1 ? scheduleData[s][roomIdx] : "";
          var course = courseIdx > -1 ? scheduleData[s][courseIdx] : "";
          scheduleLookup[joinKey] = { room: room, course: course };
        }
      }
    }

    var nameLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      var rosterEmail = String(rosterData[r][1]).toLowerCase();
      nameLookup[rosterEmail] = String(rosterData[r][0]);
    }

    var myDuties = [];
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var targetEnd = new Date(today);
    targetEnd.setDate(today.getDate() + 6); // Up to next week (7 days total including today)
    targetEnd.setHours(23, 59, 59, 999);

    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][17] || 'Active');
      if (status === 'Canceled') continue;

      var dateString = data[i][3];
      if (!dateString) continue;

      var rowDate = new Date(dateString);
      if (isNaN(rowDate.getTime())) continue;

      if (rowDate >= today && rowDate <= targetEnd) {
        var teacherEmail = String(data[i][2]).toLowerCase();
        var teacherName = nameLookup[teacherEmail] || teacherEmail;

        if (teacherName.includes(",")) {
          var parts = teacherName.split(",");
          teacherName = parts[1].trim() + " " + parts[0].trim();
        }

        var periodsRequested = String(data[i][4]).split(",").map(function(p) { return p.trim(); });
        var rowId = String(data[i][0]);
        var formattedDate = String(Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MMM d, yyyy"));
        var rawDate = Number(rowDate.getTime());

        var reason = String(data[i][5]);
        var duration = String(data[i][6]);
        var instructions = String(data[i][8]);

        for (var p = 1; p <= 8; p++) {
          if (periodsRequested.indexOf(String(p)) !== -1) {
            var subColumnIndex = 8 + p; // 9 for P1, 10 for P2, etc.
            var assignedSub = String(data[i][subColumnIndex] || "").trim().toLowerCase();

            // If the assigned sub matches the user's name
            if (assignedSub === userName) {
              var joinKey = teacherEmail + "-" + p;
              var scheduleInfo = scheduleLookup[joinKey];
              var roomStr = scheduleInfo && scheduleInfo.room ? String(scheduleInfo.room) : "No Class Assigned";
              var courseStr = scheduleInfo && scheduleInfo.course ? String(scheduleInfo.course) : "No Class Assigned";

              myDuties.push({
                id: rowId,
                teacherName: String(teacherName),
                teacherEmail: String(teacherEmail),
                date: formattedDate,
                period: String(p),
                rawDate: rawDate,
                room: roomStr,
                course: courseStr,
                reason: reason,
                duration: duration,
                instructions: instructions
              });
            }
          }
        }
      }
    }

    myDuties.sort(function(a, b) {
      if (a.rawDate === b.rawDate) {
        return parseInt(a.period) - parseInt(b.period);
      }
      return a.rawDate - b.rawDate;
    });

    return myDuties;

  } catch (err) {
    throw new Error("MySubDuties Error: " + err.message);
  }
}

/**
 * Fetches unfilled sub requests for the next 2 days (or through Monday if weekend) for the Admin Dashboard.
 */
function getQuickCoverData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("Absence Requests");
    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");

    if (!mainSheet) return [];

    var data = mainSheet.getDataRange().getValues();
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    // Master Schedule data mapping
    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];
    var scheduleLookup = {};
    if (scheduleData.length > 0) {
      var headers = scheduleData[0];
      var joinIdx = headers.indexOf("EMAIL_PERIOD_JOIN");
      var roomIdx = headers.indexOf("ROOM");
      var courseIdx = headers.indexOf("COURSE_NAMES");

      if (joinIdx > -1) {
        for (var s = 1; s < scheduleData.length; s++) {
          var joinKey = String(scheduleData[s][joinIdx]).toLowerCase();
          var room = roomIdx > -1 ? scheduleData[s][roomIdx] : "";
          var course = courseIdx > -1 ? scheduleData[s][courseIdx] : "";
          scheduleLookup[joinKey] = { room: room, course: course };
        }
      }
    }

    var nameLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      var rosterEmail = String(rosterData[r][1]).toLowerCase();
      nameLookup[rosterEmail] = String(rosterData[r][0]); 
    }

    var unfilled = [];
    var today = new Date();
    today.setHours(0, 0, 0, 0); 

    var targetEnd = new Date(today);
    var dayOfWeek = today.getDay(); // 0 is Sunday, 5 is Friday, 6 is Saturday

    var daysToAdd = 1;
    if (dayOfWeek === 5) { // Friday
      daysToAdd = 3; // +3 to reach Monday
    } else if (dayOfWeek === 6) { // Saturday
      daysToAdd = 2; // +2 to reach Monday
    }

    targetEnd.setDate(today.getDate() + daysToAdd);
    targetEnd.setHours(23, 59, 59, 999);

    // Indices in new format:
    // 0:ID, 1:Timestamp, 2:Email, 3:Date, 4:Periods, 5:Reason, 6:Duration, 7:Urgency, 8:Instructions
    // 9:P1 Sub, 10:P2 Sub, ..., 16:P8 Sub, 17:Status
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][17] || 'Active');
      if (status === 'Canceled') continue;

      var dateString = data[i][3];
      if (!dateString) continue; 
      
      var rowDate = new Date(dateString);
      if (isNaN(rowDate.getTime())) continue; 

      if (rowDate >= today && rowDate <= targetEnd) {
        var teacherEmail = String(data[i][2]).toLowerCase();
        var teacherName = nameLookup[teacherEmail] || teacherEmail; 

        if (teacherName.includes(",")) {
          var parts = teacherName.split(",");
          teacherName = parts[1].trim() + " " + parts[0].trim();
        }

        var periodsRequested = String(data[i][4]).split(",").map(function(p) { return p.trim(); });
        var rowId = String(data[i][0]);
        var formattedDate = String(Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MMM d, yyyy"));
        var rawDate = Number(rowDate.getTime());

        var reason = String(data[i][5]);
        var duration = String(data[i][6]);
        var instructions = String(data[i][8]);

        for (var p = 1; p <= 8; p++) {
          if (periodsRequested.indexOf(String(p)) !== -1) {
            var subColumnIndex = 8 + p; // 9 for P1, 10 for P2, etc.
            var assignedSub = data[i][subColumnIndex];

            if (!assignedSub || String(assignedSub).trim() === "") {

              var joinKey = teacherEmail + "-" + p;
              var scheduleInfo = scheduleLookup[joinKey];
              var roomStr = scheduleInfo && scheduleInfo.room ? String(scheduleInfo.room) : "No Class Assigned";
              var courseStr = scheduleInfo && scheduleInfo.course ? String(scheduleInfo.course) : "No Class Assigned";

              unfilled.push({
                id: rowId,
                teacherName: String(teacherName),
                teacherEmail: String(teacherEmail),
                date: formattedDate,
                period: String(p),
                rawDate: rawDate,
                room: roomStr,
                course: courseStr,
                reason: reason,
                duration: duration,
                instructions: instructions
              });
            }
          }
        }
      }
    }

    unfilled.sort(function(a, b) {
      if (a.rawDate === b.rawDate) {
        return parseInt(a.period) - parseInt(b.period);
      }
      return a.rawDate - b.rawDate;
    });

    return unfilled;
    
  } catch (err) {
    // If we throw this, the frontend failure handler will catch it and show you the error!
    throw new Error("Backend Error: " + err.message);
  }
}

/**
 * Fetches the list of staff names from the Staff Roster.
 */
function getStaffList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Staff Roster");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var staffNames = [];

    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][0]).trim();
      if (name) {
        staffNames.push(name);
      }
    }
    return staffNames.sort();
  } catch (err) {
    return [];
  }
}

function getCoordinatorEmail(ss) {
  var ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var roleSheet = ss.getSheetByName("User Roles");
  if (!roleSheet) return null;
  
  var data = roleSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === "sub coordinator") return String(data[i][0]); 
  }
  return null;
}

function submitAbsence(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("Absence Requests"); 
    
    var urgencyFormatted = formData.urgency === 'Urgent' ? 'Urgent (Less than 24 hr notice)' : 'Standard (Advanced Notice)';
    var instructions = formData.specialInstructions;
    if (formData.hrConfirmed) instructions = "[HR Docs Provided] " + instructions;

    var timestamp = new Date();
    var email = Session.getActiveUser().getEmail();
    var uniqueId = Utilities.getUuid();

    // "ID", "Timestamp", "Email", "Date", "Periods", "Reason", "Duration",
    // "Urgency", "Instructions", "Period 1 Sub", "Period 2 Sub", "Period 3 Sub",
    // "Period 4 Sub", "Period 5 Sub", "Period 6 Sub", "Period 7 Sub", "Period 8 Sub", "Status"
    var newRow = [
      uniqueId, timestamp, email, formData.date, "'" + formData.periods,
      formData.reason, formData.duration, urgencyFormatted, instructions,
      "", "", "", "", "", "", "", "", "Active"
    ];
    mainSheet.appendRow(newRow);

    var teacherName = getUserData(ss).name;
    if (urgencyFormatted === 'Urgent (Less than 24 hr notice)') {
      var coordinatorEmail = getCoordinatorEmail(ss);
      if (coordinatorEmail) {
        var subject = "URGENT COVERAGE NEEDED: " + teacherName;
        var body = "An urgent absence request has been submitted requiring immediate attention.\n\n" +
                   "Teacher: " + teacherName + "\n" +
                   "Date Needed: " + formData.date + "\n" +
                   "Periods: " + formData.periods + "\n" +
                   "Reason: " + formData.reason + "\n\n" +
                   "Instructions: " + (instructions ? instructions : "None") + "\n\n" +
                   "Please log into the Coverage Portal to assign a sub.";
        GmailApp.sendEmail(coordinatorEmail, subject, body);
      }
    }
    
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Cancels a single assigned sub duty by the sub themselves.
 */
function cancelMySubDuty(absenceId, period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Absence Requests");
    if (!sheet) throw new Error("Absence Requests sheet not found.");

    var userEmail = Session.getActiveUser().getEmail().toLowerCase();
    var userData = getUserData(ss);
    var userName = String(userData.name).trim();

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        var subColumnIndex = 10 + parseInt(period) - 1; // 1-based col: J=10 (P1), etc.
        var assignedSub = String(sheet.getRange(i + 1, subColumnIndex).getValue() || "").trim();

        if (assignedSub.toLowerCase() === userName.toLowerCase()) {
          // Get details BEFORE clearing the sub, just in case
          var coordinatorEmail = getCoordinatorEmail(ss);
          var details = getAbsenceDetails(absenceId, period);

          // Clear the sub from the sheet
          sheet.getRange(i + 1, subColumnIndex).setValue("");

          // Send email to sub coordinator, CCing the sub
          if (coordinatorEmail && details) {
            var subject = "SUB CANCELLATION: " + userName + " cancelled coverage";
            var body = userName + " has cancelled their assigned coverage.\n\n" +
                       "Date: " + details.date + "\n" +
                       "Period: " + details.period + "\n" +
                       "Teacher to Cover: " + details.teacherName + "\n" +
                       "Room: " + details.room + "\n" +
                       "Course: " + details.course + "\n\n" +
                       "This period is now UNFILLED. Please log into the Coverage Portal to reassign a sub.";

            GmailApp.sendEmail(coordinatorEmail, subject, body, { cc: userEmail });
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
  }
}

/**
 * Cancels an entire absence request.
 */
function cancelAbsence(absenceId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Absence Requests");
    if (!sheet) throw new Error("Absence Requests sheet not found.");

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        // Set Status to Canceled (Col 18, index 17)
        sheet.getRange(i + 1, 18).setValue("Canceled");

        // Notify all currently assigned subs
        for (var p = 1; p <= 8; p++) {
          var subIndex = 8 + p; // 9 for P1, etc.
          var subName = String(data[i][subIndex] || "").trim();
          if (subName) {
            var email = getSubEmail(subName);
            if (email) {
              var details = getAbsenceDetails(absenceId, p);
              if (details) sendSubNotification(email, "Canceled", details);
            }
          }
        }
        return { success: true };
      }
    }
    throw new Error("Absence ID not found.");
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Updates an absence request.
 */
function updateAbsence(absenceId, formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Absence Requests");
    if (!sheet) throw new Error("Absence Requests sheet not found.");

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {

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
             oldDateFormatted = String(oldDateRaw);
          }
        }
        var newDate = String(formData.date);

        var dateChanged = oldDateFormatted !== newDate;

        var urgencyFormatted = formData.urgency === 'Urgent' ? 'Urgent (Less than 24 hr notice)' : 'Standard (Advanced Notice)';
        var instructions = formData.specialInstructions;
        if (formData.hrConfirmed) instructions = "[HR Docs Provided] " + instructions;

        // Update basic info
        sheet.getRange(i + 1, 4).setValue(formData.date);
        sheet.getRange(i + 1, 5).setValue("'" + formData.periods);
        sheet.getRange(i + 1, 6).setValue(formData.reason);
        sheet.getRange(i + 1, 7).setValue(formData.duration);
        sheet.getRange(i + 1, 8).setValue(urgencyFormatted);
        sheet.getRange(i + 1, 9).setValue(instructions);

        // Notify subs
        for (var p = 1; p <= 8; p++) {
          var subIndex = 8 + p;
          var subName = String(data[i][subIndex] || "").trim();

          if (subName) {
            var email = getSubEmail(subName);
            var isPeriodStillNeeded = newPeriods.indexOf(String(p)) !== -1;

            if (email) {
               if (dateChanged || !isPeriodStillNeeded) {
                 // Cancel this sub for this period
                 var cancelDetails = getAbsenceDetails(absenceId, p);
                 // We pass old date since they are canceled for the old date
                 if (cancelDetails) {
                    cancelDetails.date = Utilities.formatDate(new Date(oldDateRaw), Session.getScriptTimeZone(), "MMM d, yyyy");
                    sendSubNotification(email, "Canceled", cancelDetails);
                 }
                 // Clear the sub from the sheet
                 sheet.getRange(i + 1, subIndex + 1).setValue("");
               } else {
                 // Sub still needed, but details modified
                 var oldReason = String(data[i][5]);
                 var oldDuration = String(data[i][6]);
                 var oldUrgency = String(data[i][7]);
                 var oldInstructions = String(data[i][8]);

                 var detailsChanged = (oldReason !== formData.reason) ||
                                      (oldDuration !== formData.duration) ||
                                      (oldUrgency !== urgencyFormatted) ||
                                      (oldInstructions !== instructions);

                 if (detailsChanged) {
                    var modDetails = getAbsenceDetails(absenceId, p);
                    if (modDetails) sendSubNotification(email, "Modified", modDetails);
                 }
               }
            } else if (dateChanged || !isPeriodStillNeeded) {
               // Clear sub from sheet even if email not found
               sheet.getRange(i + 1, subIndex + 1).setValue("");
            }
          }
        }

        return { success: true };
      }
    }
    throw new Error("Absence ID not found.");
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Helper to get a sub's email from the Staff Roster.
 */
function getSubEmail(subName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Staff Roster");
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === subName.trim()) {
      return String(data[i][1]).trim();
    }
  }
  return null;
}

/**
 * Helper to get teacher name from email.
 */
function getTeacherNameFromEmail(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Staff Roster");
  if (!sheet) return email;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === email.trim().toLowerCase()) {
      return String(data[i][0]).trim();
    }
  }
  return email;
}

/**
 * Helper to get full absence details.
 */
function getAbsenceDetails(absenceId, period) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Absence Requests");
  var masterScheduleSheet = ss.getSheetByName("Master Schedule");

  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();

  var scheduleLookup = {};
  if (masterScheduleSheet) {
    var scheduleData = masterScheduleSheet.getDataRange().getValues();
    if (scheduleData.length > 0) {
      var headers = scheduleData[0];
      var joinIdx = headers.indexOf("EMAIL_PERIOD_JOIN");
      var roomIdx = headers.indexOf("ROOM");
      var courseIdx = headers.indexOf("COURSE_NAMES");
      if (joinIdx > -1) {
        for (var s = 1; s < scheduleData.length; s++) {
          var joinKey = String(scheduleData[s][joinIdx]).toLowerCase();
          var room = roomIdx > -1 ? scheduleData[s][roomIdx] : "";
          var course = courseIdx > -1 ? scheduleData[s][courseIdx] : "";
          scheduleLookup[joinKey] = { room: room, course: course };
        }
      }
    }
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(absenceId)) {
      var teacherEmail = String(data[i][2]);
      var teacherName = getTeacherNameFromEmail(teacherEmail);
      var dateVal = data[i][3];
      var formattedDate = dateVal;
      if (dateVal instanceof Date) {
        formattedDate = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MMM d, yyyy");
      }
      var instructions = String(data[i][8]);

      var roomStr = "No Class Assigned";
      var courseStr = "No Class Assigned";

      if (period) {
        var joinKey = teacherEmail.toLowerCase() + "-" + period;
        var scheduleInfo = scheduleLookup[joinKey];
        if (scheduleInfo) {
          roomStr = scheduleInfo.room || roomStr;
          courseStr = scheduleInfo.course || courseStr;
        }
      }

      return {
        teacherName: teacherName,
        date: formattedDate,
        period: period,
        room: roomStr,
        course: courseStr,
        instructions: instructions,
        rowIndex: i + 1, // 1-based index for Apps Script Range
        row: data[i]
      };
    }
  }
  return null;
}

/**
 * Sends a notification email to a substitute.
 */
function sendSubNotification(subEmail, type, details) {
  if (!subEmail) return;

  var subject = "";
  var body = "";

  var detailsText = "Date: " + details.date + "\n";
  if (details.period) detailsText += "Period: " + details.period + "\n";
  detailsText += "Teacher: " + details.teacherName + "\n";
  if (details.period) {
    detailsText += "Room: " + details.room + "\n";
    detailsText += "Course: " + details.course + "\n";
  }
  detailsText += "\nSpecial Instructions:\n" + (details.instructions ? details.instructions : "None provided");

  if (type === 'Assigned') {
    subject = "Coverage Assignment: " + details.date + (details.period ? " Period " + details.period : "");
    body = "You have been assigned to cover a class.\n\n" + detailsText + "\n\nPlease check the Coverage Portal for more information.";
  } else if (type === 'Canceled') {
    subject = "CANCELED - Coverage Assignment: " + details.date + (details.period ? " Period " + details.period : "");
    body = "Your assigned coverage has been CANCELED. You are no longer needed for this assignment.\n\n" + detailsText;
  } else if (type === 'Modified') {
    subject = "UPDATED - Coverage Assignment: " + details.date + (details.period ? " Period " + details.period : "");
    body = "There has been an update to your assigned coverage.\n\nUpdated Details:\n" + detailsText + "\n\nPlease check the Coverage Portal for more information.";
  }

  try {
    GmailApp.sendEmail(subEmail, subject, body);
  } catch (e) {
    console.error("Failed to send email to " + subEmail + ": " + e.message);
  }
}

/**
 * Assigns a substitute to a specific period for an absence request.
 */
function assignSubToPeriod(absenceId, period, subName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Absence Requests");
    if (!sheet) throw new Error("Absence Requests sheet not found.");

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        var periodsRequested = String(data[i][4]).split(",").map(function(p) { return p.trim(); });

        if (periodsRequested.indexOf(String(period)) === -1) {
            throw new Error("Period " + period + " was not requested for this absence.");
        }

        var subColumnIndex = 10 + parseInt(period) - 1; // 1-based index for Apps Script Ranges: Col J is 10 (Period 1 Sub)

        var existingSub = String(sheet.getRange(i + 1, subColumnIndex).getValue() || "").trim();
        var newSub = String(subName || "").trim();

        if (existingSub === newSub) {
           return { success: true }; // No change
        }

        // Get details for email
        var details = getAbsenceDetails(absenceId, period);

        // Cancel existing sub if there is one
        if (existingSub && details) {
           var existingEmail = getSubEmail(existingSub);
           if (existingEmail) {
              sendSubNotification(existingEmail, 'Canceled', details);
           }
        }

        // Write the subname
        sheet.getRange(i + 1, subColumnIndex).setValue(newSub);

        // Notify new sub if there is one
        if (newSub && details) {
           var newEmail = getSubEmail(newSub);
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
  }
}

/**
 * Fetches data for the Admin Dashboard.
 * Returns an array of objects, one per period.
 */
function getAdminDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName("Absence Requests");
    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");

    if (!mainSheet) return [];

    var data = mainSheet.getDataRange().getValues();
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    // Master Schedule data mapping
    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];
    var scheduleLookup = {};
    if (scheduleData.length > 0) {
      var headers = scheduleData[0];
      var joinIdx = headers.indexOf("EMAIL_PERIOD_JOIN");
      var roomIdx = headers.indexOf("ROOM");
      var courseIdx = headers.indexOf("COURSE_NAMES");

      if (joinIdx > -1 && roomIdx > -1 && courseIdx > -1) {
        for (var s = 1; s < scheduleData.length; s++) {
          var key = String(scheduleData[s][joinIdx]).toLowerCase().trim();
          scheduleLookup[key] = {
            room: scheduleData[s][roomIdx] || "No Class Assigned",
            course: scheduleData[s][courseIdx] || "No Class Assigned"
          };
        }
      }
    }

    var adminData = [];

    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][17] || "").trim(); // 17 is Status
      if (status === "Canceled") continue;

      var id = data[i][0];
      var email = String(data[i][2]).toLowerCase().trim();
      var dateStr = data[i][3];
      var periodsStr = String(data[i][4]);
      var instructions = String(data[i][8] || "").trim();

      // Get teacher name
      var teacherName = getTeacherNameFromEmail(email);

      // Parse date to a comparable format, YYYY-MM-DD
      var dateObj = new Date(dateStr);
      var dateFormatted = "";
      if (!isNaN(dateObj.getTime())) {
          // Keep local timezone in mind, maybe just formatting using Utilities
          dateFormatted = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
          // Fallback if not a standard date
          dateFormatted = dateStr;
      }

      var periods = periodsStr.split(',').map(function(p) { return p.trim(); });

      for (var j = 0; j < periods.length; j++) {
        var p = parseInt(periods[j]);
        if (!isNaN(p)) {
            var subColumnIndex = 8 + p; // 9 for P1, 10 for P2, etc.
            var assignedSub = data[i][subColumnIndex] || "";

            var scheduleKey = email + "-" + p;
            var room = "";
            var course = "";

            if (scheduleLookup[scheduleKey]) {
              room = scheduleLookup[scheduleKey].room;
              course = scheduleLookup[scheduleKey].course;
            }

            adminData.push({
              id: id,
              originalDate: dateStr,
              date: dateFormatted,
              period: p,
              teacherName: teacherName,
              teacherEmail: email,
              course: course,
              room: room,
              assignedSub: String(assignedSub).trim(),
              instructions: instructions
            });
        }
      }
    }

    // Sort by Date, then by Period
    adminData.sort(function(a, b) {
      if (a.date < b.date) return -1;
      if (a.date > b.date) return 1;
      return a.period - b.period;
    });

    return adminData;

  } catch (e) {
    console.error("Error fetching Admin Dashboard Data: " + e.message);
    throw new Error("Failed to load admin dashboard data.");
  }
}
