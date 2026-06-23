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
      if (String(data[i][1]).toLowerCase() === String(email).toLowerCase()) {
        var dateVal = data[i][2]; 
        if (!dateVal) continue; 

        var formattedDate = "Unknown Date";
        if (dateVal instanceof Date) {
           formattedDate = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MMM d, yyyy");
        } else {
           formattedDate = new Date(dateVal).toLocaleDateString();
        }

        var urgencyStr = String(data[i][6] || '');

        // Forced strings to ensure perfect serialization
        myAbsences.push({
          date: String(formattedDate),
          periods: String(data[i][3]), 
          reason: String(data[i][4]),  
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
    var splitSheet = ss.getSheetByName("Split Responses");
    var rosterSheet = ss.getSheetByName("Staff Roster");

    if (!splitSheet) return [];

    var data = splitSheet.getDataRange().getValues();
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

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

    for (var i = 1; i < data.length; i++) {
      var dateString = data[i][3];
      if (!dateString) continue; 
      
      var rowDate = new Date(dateString);
      if (isNaN(rowDate.getTime())) continue; 

      var assignedSub = data[i][5]; 

      if (rowDate >= today && rowDate <= nextWeek && (!assignedSub || String(assignedSub).trim() === "")) {
        var teacherEmail = String(data[i][2]).toLowerCase();
        var teacherName = nameLookup[teacherEmail] || teacherEmail; 

        if (teacherName.includes(",")) {
          var parts = teacherName.split(",");
          teacherName = parts[1].trim() + " " + parts[0].trim();
        }

        // We force everything to String or Number to prevent silent serialization crashes!
        unfilled.push({
          id: String(data[i][0] || ""),
          teacherName: String(teacherName),
          date: String(Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MMM d, yyyy")),
          period: String(data[i][4] || ""),
          rawDate: Number(rowDate.getTime()) 
        });
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

    var newRow = [
      timestamp, email, formData.date, "'" + formData.periods, 
      formData.reason, formData.duration, urgencyFormatted, instructions
    ];
    mainSheet.appendRow(newRow);
    
    var splitSheet = ss.getSheetByName("Split Responses");
    if (splitSheet && formData.periods) {
      var parts = formData.periods.split(",").map(function(p) { return p.trim(); });
      parts.forEach(function(part) {
        var uniqueId = Utilities.getUuid(); 
        splitSheet.appendRow([
          uniqueId, timestamp, email, formData.date, part, "", instructions 
        ]);
      });
    }

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
