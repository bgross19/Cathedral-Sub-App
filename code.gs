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
    "Period 4 Sub", "Period 5 Sub", "Period 6 Sub", "Period 7 Sub", "Period 8 Sub"
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
function getUserData() {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
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

        // Forced strings to ensure perfect serialization
        myAbsences.push({
          date: String(formattedDate),
          periods: String(data[i][4]), // Periods is index 4 now
          reason: String(data[i][5]),  // Reason is index 5 now
          urgency: urgencyStr.includes('Urgent') ? 'Urgent' : 'Standard' 
        });
      }
    }
    
    return myAbsences.reverse(); 
  } catch (err) {
    throw new Error("MyAbsences Error: " + err.message);
  }
}

/**
 * Fetches unfilled sub requests for the next 7 days for the Admin Dashboard.
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

    var nextWeek = new Date(today);
    nextWeek.setDate(today.getDate() + 7); 
    nextWeek.setHours(23, 59, 59, 999); 

    // Indices in new format:
    // 0:ID, 1:Timestamp, 2:Email, 3:Date, 4:Periods, 5:Reason, 6:Duration, 7:Urgency, 8:Instructions
    // 9:P1 Sub, 10:P2 Sub, ..., 16:P8 Sub
    for (var i = 1; i < data.length; i++) {
      var dateString = data[i][3];
      if (!dateString) continue; 
      
      var rowDate = new Date(dateString);
      if (isNaN(rowDate.getTime())) continue; 

      if (rowDate >= today && rowDate <= nextWeek) {
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

function getCoordinatorEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    // "Period 4 Sub", "Period 5 Sub", "Period 6 Sub", "Period 7 Sub", "Period 8 Sub"
    var newRow = [
      uniqueId, timestamp, email, formData.date, "'" + formData.periods,
      formData.reason, formData.duration, urgencyFormatted, instructions,
      "", "", "", "", "", "", "", ""
    ];
    mainSheet.appendRow(newRow);

    var teacherName = getUserData().name; 
    if (urgencyFormatted === 'Urgent (Less than 24 hr notice)') {
      var coordinatorEmail = getCoordinatorEmail();
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

        // Write the subname
        sheet.getRange(i + 1, subColumnIndex).setValue(subName);
        return { success: true };
      }
    }

    throw new Error("Absence Request ID not found.");
  } catch (err) {
    return { success: false, error: err.message };
  }
}
