/**
 * Sets up the database headers in the Google Sheet.
 * Run this function once from the Apps Script editor.
 */
function setupDatabase() {
  var ss = getSS();
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

  // Setup Settings Sheet
  var settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    var settingsHeaders = ["Setting Name", "Setting Value"];
    var defaultSettings = [
      ["Email Mode", "Live"],
      ["Redirect Email", "Bgross@gocathedral.com"],
      ["App URL", "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec"],
      ["Urgency Cutoff Time", "15"],
      ["Absence Reasons", JSON.stringify([
        {reason: "Personal", hrRequired: false},
        {reason: "Professional Development", hrRequired: false},
        {reason: "Retreat", hrRequired: false},
        {reason: "Athletics", hrRequired: false},
        {reason: "Jury Duty", hrRequired: true},
        {reason: "Bereavement", hrRequired: true}
      ])]
    ];
    settingsSheet.getRange(1, 1, 1, 2).setValues([settingsHeaders]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
    settingsSheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  }
}

/**
 * Retrieves settings from the Settings sheet as an object.
 * Uses defaults in memory if the sheet does not exist or user lacks permission.
 */
function getSettings(ss) {
  var defaults = {
    "Email Mode": "Live",
    "Redirect Email": "Bgross@gocathedral.com",
    "App URL": "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec",
    "Urgency Cutoff Time": "15",
    "Absence Reasons": JSON.stringify([
      {reason: "Personal", hrRequired: false},
      {reason: "Professional Development", hrRequired: false},
      {reason: "Retreat", hrRequired: false},
      {reason: "Athletics", hrRequired: false},
      {reason: "Jury Duty", hrRequired: true},
      {reason: "Bereavement", hrRequired: true}
    ])
  };

  try {
    var sheetSS = ss || getSS();
    var settingsSheet = sheetSS.getSheetByName("Settings");

    var settings = {};

    if (settingsSheet) {
      var data = settingsSheet.getDataRange().getValues();
      // Skip header row
      for (var i = 1; i < data.length; i++) {
        var key = String(data[i][0]).trim();
        var value = String(data[i][1]).trim();
        if (key) {
          settings[key] = value;
        }
      }
    }

    // Merge with defaults if not present
    var settingsUpdated = false;
    for (var k in defaults) {
      if (!(k in settings)) {
        settings[k] = defaults[k];
        if (settingsSheet) {
          try {
            settingsSheet.appendRow([k, defaults[k]]);
          } catch (appendErr) {
            console.warn("Could not append default setting: " + appendErr.message);
          }
        }
      }
    }

    return settings;
  } catch (e) {
    console.warn("Could not read settings from spreadsheet, using defaults: " + e.message);
    return defaults;
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
 * Asserts that a user has one of the allowed roles.
 * @param {Object} user - The user object returned by getUserData().
 * @param {string|string[]} allowedRoles - A single role or an array of allowed roles.
 * @param {string} [customErrorMessage="Unauthorized"] - Optional custom error message.
 * @throws {Error} If the user's role is not in the allowed list.
 */
/**
 * Returns the active spreadsheet.
 */
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}


/**
 * Retrieves a sheet by name or throws an error if not found.
 * @param {SpreadsheetApp.Spreadsheet} ss - The spreadsheet object.
 * @param {string} sheetName - The exact name of the sheet.
 * @returns {SpreadsheetApp.Sheet} The requested sheet.
 * @throws {Error} If the sheet is not found.
 */
function getSheetOrThrow(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(sheetName + " sheet not found.");
  }
  return sheet;
}

/**
 * Retrieves a sheet case-insensitively or throws an error if not found.
 * @param {SpreadsheetApp.Spreadsheet} ss - The spreadsheet object.
 * @param {string} sheetName - The name of the sheet (case-insensitive).
 * @returns {SpreadsheetApp.Sheet} The requested sheet.
 * @throws {Error} If the sheet is not found.
 */
function getSheetCaseInsensitiveOrThrow(ss, sheetName) {
  var allSheets = ss.getSheets();
  var targetNameLower = sheetName.toLowerCase();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().toLowerCase() === targetNameLower) {
      return allSheets[i];
    }
  }
  throw new Error(sheetName + " sheet not found.");
}

function assertRole(user, allowedRoles, customErrorMessage) {
  if (!user || !user.role) {
    throw new Error(customErrorMessage || "Unauthorized");
  }
  var role = user.role.toLowerCase();
  var allowed = Array.isArray(allowedRoles) ? allowedRoles : [allowedRoles];
  var isAuthorized = allowed.some(function(r) { return r.toLowerCase() === role; });
  if (!isAuthorized) {
    throw new Error(customErrorMessage || "Unauthorized");
  }
}

/**
 * Grabs the user's name and role on startup.
 */
function getUserData(ss) {
  var email = Session.getActiveUser().getEmail();
  var ss = ss || getSS();
  
  var rosterSheet = ss.getSheetByName("Staff Roster");
  var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
  var name = "Teacher"; 
  var targetEmail = String(email).toLowerCase();
  
  for (var i = 1; i < rosterData.length; i++) {
    if (String(rosterData[i][1]).toLowerCase() === targetEmail) {
      name = rosterData[i][0]; 
      break;
    }
  }
  
  var roleSheet = ss.getSheetByName("User Roles");
  var roleData = roleSheet ? roleSheet.getDataRange().getValues() : [];
  var role = "User"; 
  
  for (var j = 1; j < roleData.length; j++) {
    if (String(roleData[j][0]).toLowerCase() === targetEmail) {
      role = roleData[j][1]; 
      break;
    }
  }
  
  var settings = getSettings(ss);
  var appUrl = settings["App URL"] || "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec";
  var urgencyCutoffTime = settings["Urgency Cutoff Time"] || "15";
  var defaultAbsenceReasons = JSON.stringify([
      {reason: "Personal", hrRequired: false},
      {reason: "Professional Development", hrRequired: false},
      {reason: "Retreat", hrRequired: false},
      {reason: "Athletics", hrRequired: false},
      {reason: "Jury Duty", hrRequired: true},
      {reason: "Bereavement", hrRequired: true}
  ]);
  var absenceReasons = settings["Absence Reasons"] || defaultAbsenceReasons;

  return {
    name: String(name),
    role: String(role),
    email: String(email),
    appUrl: String(appUrl),
    urgencyCutoffTime: String(urgencyCutoffTime),
    absenceReasons: String(absenceReasons)
  };
}

/**
 * Helper to send emails based on the current Email Mode setting.
 * Handles "Live", "Redirect", and "Off" modes.
 */
function sendEmailHelper(to, subject, body, options) {
  var settings = getSettings();
  var mode = settings["Email Mode"] || "Live";
  var redirectEmail = settings["Redirect Email"] || "";

  if (mode === "Off") {
    console.log("Email sending is turned Off. Suppressed email to: " + to);
    return;
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
}

/**
 * Helper to build a name lookup dictionary from Staff Roster data.
 */
function buildNameLookup(rosterData) {
  var nameLookup = {};
  for (var r = 1; r < rosterData.length; r++) {
    var rosterEmail = String(rosterData[r][1]).toLowerCase().trim();
    nameLookup[rosterEmail] = String(rosterData[r][0]).trim();
  }
  return nameLookup;
}

/**
 * Helper to build a schedule lookup dictionary from Master Schedule data.
 */
function buildScheduleLookup(scheduleData) {
  var scheduleLookup = {};
  if (scheduleData.length > 0) {
    var headers = scheduleData[0];
    var joinIdx = headers.indexOf("EMAIL_PERIOD_JOIN");
    var roomIdx = headers.indexOf("ROOM");
    var courseIdx = headers.indexOf("COURSE_NAMES");

    if (joinIdx > -1) {
      for (var s = 1; s < scheduleData.length; s++) {
        var joinKey = String(scheduleData[s][joinIdx]).toLowerCase().trim();
        var room = roomIdx > -1 ? scheduleData[s][roomIdx] : "No Class Assigned";
        var course = courseIdx > -1 ? scheduleData[s][courseIdx] : "No Class Assigned";
        // Do not overwrite with empty values if we don't have to, but be safe
        scheduleLookup[joinKey] = {
          room: room ? room : "No Class Assigned",
          course: course ? course : "No Class Assigned"
        };
      }
    }
  }
  return scheduleLookup;
}

/**
 * Fetches all user roles.
 */
function getUserRoles() {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, "admin");

    var roleSheet = ss.getSheetByName("User Roles");
    if (!roleSheet) return [];

    var data = roleSheet.getDataRange().getValues();
    var roles = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim()) {
        roles.push({
          email: String(data[i][0]).trim(),
          role: String(data[i][1]).trim()
        });
      }
    }
    return roles;
  } catch (err) {
    throw new Error("Failed to fetch roles: " + err.message);
  }
}

/**
 * Adds a new user role.
 */
function addUserRole(email, role) {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, "admin");

    var roleSheet = getSheetOrThrow(ss, "User Roles");

    roleSheet.appendRow([email.toLowerCase().trim(), role.trim()]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Edits an existing user role.
 */
function editUserRole(oldEmail, newEmail, role) {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, "admin");

    var roleSheet = getSheetOrThrow(ss, "User Roles");

    var data = roleSheet.getDataRange().getValues();
    var targetEmail = oldEmail.toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === targetEmail) {
        roleSheet.getRange(i + 1, 1, 1, 2).setValues([[newEmail.toLowerCase().trim(), role.trim()]]);
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Deletes a user role.
 */
function deleteUserRole(email) {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, "admin");

    var roleSheet = getSheetOrThrow(ss, "User Roles");

    var data = roleSheet.getDataRange().getValues();
    var targetEmail = email.toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === targetEmail) {
        roleSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Fetches settings for the frontend.
 */
function getSettingsForFrontend() {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, "admin");
    return getSettings();
  } catch (err) {
    throw new Error("Failed to fetch settings: " + err.message);
  }
}

/**
 * Updates settings in the Settings sheet.
 */
function updateSettings(newSettings) {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, "admin");

    var settingsSheet = getSheetOrThrow(ss, "Settings");

    var data = settingsSheet.getDataRange().getValues();
    var settingsMap = {};

    // Map existing rows
    for (var i = 1; i < data.length; i++) {
      settingsMap[String(data[i][0]).trim()] = i + 1;
    }

    for (var key in newSettings) {
      if (settingsMap[key]) {
        settingsSheet.getRange(settingsMap[key], 2).setValue(newSettings[key]);
      } else {
        settingsSheet.appendRow([key, newSettings[key]]);
      }
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Fetches the logged-in user's upcoming absences.
 */
function getMyAbsences() {
  try {
    var email = Session.getActiveUser().getEmail();
    var ss = getSS();
    var sheet = ss.getSheetByName("Absence Requests");
    
    if (!sheet) return []; 
    
    var data = sheet.getDataRange().getValues();
    var myAbsences = [];
    var targetEmail = String(email).toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][17] || 'Active');
      if (status === 'Canceled') continue;

      if (String(data[i][2]).toLowerCase() === targetEmail) { // Email is index 2 now
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
                console.error("Error formatting date: " + e.message);
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
    var ss = getSS();
    var userData = getUserData(ss);
    var userName = String(userData.name).trim().toLowerCase();

    var mainSheet = ss.getSheetByName("Absence Requests");
    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");

    if (!mainSheet) return [];

    var data = mainSheet.getDataRange().getValues();
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];
    var scheduleLookup = buildScheduleLookup(scheduleData);
    var nameLookup = buildNameLookup(rosterData);

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
        var formDateString = String(Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd"));
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
                formDateString: formDateString,
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
 * Fetches unfilled sub requests for today.
 */
function getTodaysOpenJobsData() {
  try {
    var ss = getSS();
    var mainSheet = ss.getSheetByName("Absence Requests");
    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");

    if (!mainSheet) return [];

    var data = mainSheet.getDataRange().getValues();
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];
    var scheduleLookup = buildScheduleLookup(scheduleData);
    var nameLookup = buildNameLookup(rosterData);

    var unfilled = [];
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var targetEnd = new Date(today);
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
        var formDateString = String(Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd"));
        var rawDate = Number(rowDate.getTime());

        var reason = String(data[i][5]);
        var duration = String(data[i][6]);
        var instructions = String(data[i][8]);

        for (var p = 1; p <= 8; p++) {
          if (periodsRequested.indexOf(String(p)) !== -1) {
            var subColumnIndex = 8 + p;
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
                formDateString: formDateString,
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
    throw new Error("Backend Error: " + err.message);
  }
}

/**
 * Fetches unfilled sub requests for the next 2 days (or through Monday if weekend) for the Admin Dashboard.
 */
function getQuickCoverData() {
  try {
    var ss = getSS();
    var mainSheet = ss.getSheetByName("Absence Requests");
    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");

    if (!mainSheet) return [];

    var data = mainSheet.getDataRange().getValues();
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];
    var scheduleLookup = buildScheduleLookup(scheduleData);
    var nameLookup = buildNameLookup(rosterData);

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
        var formDateString = String(Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd"));
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
                formDateString: formDateString,
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
    var ss = getSS();
    var sheet = ss.getSheetByName("Staff Roster");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var staffList = [];

    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][0]).trim();
      var duty = String(data[i][2] || "").trim(); // 3rd column

      if (name) {
        var display = name;
        if (duty) {
          display = name + " - " + duty;
        }

        staffList.push({
          name: name,
          display: display,
          duty: duty
        });
      }
    }

    // Sort alphabetically by name
    return staffList.sort(function(a, b) {
      var nameA = a.name.toLowerCase();
      var nameB = b.name.toLowerCase();
      if (nameA < nameB) return -1;
      if (nameA > nameB) return 1;
      return 0;
    });
  } catch (err) {
    return [];
  }
}

function getCoordinatorEmail(ss) {
  var ss = ss || getSS();
  var roleSheet = ss.getSheetByName("User Roles");
  if (!roleSheet) return null;
  
  var data = roleSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === "sub coordinator") return String(data[i][0]); 
  }
  return null;
}

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

/**
 * Helper to send an email for urgent coverage requests.
 */
function sendUrgentCoverageEmail(ss, teacherName, formData, instructions) {
  var coordinatorEmail = getCoordinatorEmail(ss);
  if (coordinatorEmail) {
    var subject = "URGENT COVERAGE NEEDED: " + teacherName;
    var settings = getSettings();
    var appUrl = settings["App URL"] || "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec";

    var body = "An urgent absence request has been submitted requiring immediate attention.\n\n" +
               "Teacher: " + teacherName + "\n" +
               "Date Needed: " + formData.date + "\n" +
               "Periods: " + formData.periods + "\n" +
               "Reason: " + formData.reason + "\n\n" +
               "Instructions: " + (instructions ? instructions : "None") + "\n\n" +
               "Please log into the Cathedral Sub App to assign a sub: " + appUrl;

    var htmlBody = "<p>An urgent absence request has been submitted requiring immediate attention.</p>" +
                   "<ul>" +
                   "<li><strong>Teacher:</strong> " + teacherName + "</li>" +
                   "<li><strong>Date Needed:</strong> " + formData.date + "</li>" +
                   "<li><strong>Periods:</strong> " + formData.periods + "</li>" +
                   "<li><strong>Reason:</strong> " + formData.reason + "</li>" +
                   "</ul>" +
                   "<p><strong>Instructions:</strong> " + (instructions ? instructions : "None") + "</p>" +
                   "<p>Please log into the <a href='" + appUrl + "'>Cathedral Sub App</a> to assign a sub.</p>";

    sendEmailHelper(coordinatorEmail, subject, body, { htmlBody: htmlBody });
  }
}

function submitAbsence(formData) {
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
    var confBody = "Your absence request has been successfully submitted.\n\n" +
                   "Details:\n" +
                   "Date: " + formData.date + "\n" +
                   "Periods: " + formData.periods + "\n" +
                   "Reason: " + formData.reason + "\n" +
                   "Duration: " + formData.duration + "\n" +
                   "Instructions: " + (instructions ? instructions : "None") + "\n\n" +
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
}

/**
 * Cancels a single assigned sub duty by the sub themselves.
 */
function cancelMySubDuty(absenceId, period) {
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

            var body = userName + " has cancelled their assigned coverage.\n\n" +
                       "Date: " + details.date + "\n" +
                       "Period: " + details.period + "\n" +
                       "Teacher to Cover: " + details.teacherName + "\n" +
                       "Room: " + details.room + "\n" +
                       "Course: " + details.course + "\n\n" +
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
  var instructions = String(row[8]);

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
    instructions: instructions
  };
}

function cancelAbsence(absenceId) {
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
           var teacherBody = "Your absence request for " + formattedDateForEmail + " has been canceled by an administrator.\n\n" +
                             "Reason: " + data[i][5] + "\n" +
                             "Periods: " + data[i][4] + "\n\n" +
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
}

/**
 * Updates an absence request.
 */
function updateAbsence(absenceId, formData) {
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
           var teacherBody = "Your absence request for " + formattedDateForEmail + " has been updated by an administrator.\n\n" +
                             "Updated Details:\n" +
                             "Date: " + formData.date + "\n" +
                             "Periods: " + formData.periods + "\n" +
                             "Reason: " + formData.reason + "\n" +
                             "Duration: " + formData.duration + "\n\n" +
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
}

/**
 * Helper to get a sub's email from the Staff Roster.
 */
function getSubEmail(subName) {
  var ss = getSS();
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
function getTeacherNameFromEmail(email, nameLookup) {
  var targetEmail = email.trim().toLowerCase();
  if (nameLookup !== undefined) {
    return nameLookup[targetEmail] || email;
  }

  // Fallback to slow lookup if nameLookup is not provided
  var ss = getSS();
  var sheet = ss.getSheetByName("Staff Roster");
  if (!sheet) return email;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === targetEmail) {
      return String(data[i][0]).trim();
    }
  }
  return email;
}

/**
 * Helper to get full absence details.
 */
function getAbsenceDetails(absenceId, period) {
  var ss = getSS();
  var sheet = ss.getSheetByName("Absence Requests");
  var masterScheduleSheet = ss.getSheetByName("Master Schedule");
  var rosterSheet = ss.getSheetByName("Staff Roster");

  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();

  var nameLookup = {};
  if (rosterSheet) {
    var rosterData = rosterSheet.getDataRange().getValues();
    nameLookup = buildNameLookup(rosterData);
  }

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
      var teacherName = getTeacherNameFromEmail(teacherEmail, nameLookup);
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
  var htmlBody = "";

  var settings = getSettings();
  var appUrl = settings["App URL"] || "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec";

  var detailsText = "Date: " + details.date + "\n";
  if (details.period) detailsText += "Period: " + details.period + "\n";
  detailsText += "Teacher: " + details.teacherName + "\n";
  if (details.period) {
    detailsText += "Room: " + details.room + "\n";
    detailsText += "Course: " + details.course + "\n";
  }
  detailsText += "\nSpecial Instructions:\n" + (details.instructions ? details.instructions : "None provided");

  var detailsHtml = "<ul>" +
                    "<li><strong>Date:</strong> " + details.date + "</li>" +
                    (details.period ? "<li><strong>Period:</strong> " + details.period + "</li>" : "") +
                    "<li><strong>Teacher:</strong> " + details.teacherName + "</li>" +
                    (details.period ? "<li><strong>Room:</strong> " + details.room + "</li>" : "") +
                    (details.period ? "<li><strong>Course:</strong> " + details.course + "</li>" : "") +
                    "</ul>" +
                    "<p><strong>Special Instructions:</strong><br>" + (details.instructions ? details.instructions : "None provided") + "</p>";

  if (type === 'Assigned') {
    subject = "Coverage Assignment: " + details.date + (details.period ? " Period " + details.period : "");
    body = "You have been assigned to cover a class.\n\n" + detailsText + "\n\nPlease check the Cathedral Sub App for more information: " + appUrl;
    htmlBody = "<p>You have been assigned to cover a class.</p>" + detailsHtml + "<p>Please check the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
  } else if (type === 'Canceled') {
    subject = "CANCELED - Coverage Assignment: " + details.date + (details.period ? " Period " + details.period : "");
    body = "Your assigned coverage has been CANCELED. You are no longer needed for this assignment.\n\n" + detailsText + "\n\nGo to the Cathedral Sub App for more information: " + appUrl;
    htmlBody = "<p>Your assigned coverage has been CANCELED. You are no longer needed for this assignment.</p>" + detailsHtml + "<p>Go to the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
  } else if (type === 'Modified') {
    subject = "UPDATED - Coverage Assignment: " + details.date + (details.period ? " Period " + details.period : "");
    body = "There has been an update to your assigned coverage.\n\nUpdated Details:\n" + detailsText + "\n\nPlease check the Cathedral Sub App for more information: " + appUrl;
    htmlBody = "<p>There has been an update to your assigned coverage.</p><h3>Updated Details:</h3>" + detailsHtml + "<p>Please check the <a href='" + appUrl + "'>Cathedral Sub App</a> for more information.</p>";
  }

  try {
    sendEmailHelper(subEmail, subject, body, { htmlBody: htmlBody });
  } catch (e) {
    console.error("Failed to send email to " + subEmail + ": " + e.message);
  }
}

/**
 * Assigns a substitute to a specific period for an absence request.
 */
function assignSubToPeriod(absenceId, period, subName) {
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
}

/**
 * Fetches data for the Admin Dashboard.
 * Returns an array of objects, one per period.
 */
function getAdminDashboardData() {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, ["admin", "sub coordinator"], "Unauthorized access. Admin or Sub Coordinator role required.");

    var mainSheet = ss.getSheetByName("Absence Requests");
    var rosterSheet = ss.getSheetByName("Staff Roster");
    var masterScheduleSheet = ss.getSheetByName("Master Schedule");

    if (!mainSheet) return [];

    var data = mainSheet.getDataRange().getValues();
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];
    var scheduleLookup = buildScheduleLookup(scheduleData);
    var nameLookup = buildNameLookup(rosterData);

    var adminData = [];

    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][17] || "").trim(); // 17 is Status
      if (status === "Canceled") continue;

      var id = data[i][0];
      var email = String(data[i][2]).toLowerCase().trim();
      var dateStr = data[i][3];
      var periodsStr = String(data[i][4]);
      var reason = String(data[i][5] || "").trim();
      var duration = String(data[i][6] || "").trim();
      var instructions = String(data[i][8] || "").trim();

      // Get teacher name
      var lookupEmail = email; // email is already lowercased and trimmed above
      var teacherName = nameLookup[lookupEmail] || email;

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
              id: String(id || ""),
              originalDate: String(dateStr || ""),
              date: String(dateFormatted || ""),
              formDateString: String(dateFormatted || ""), // Included for uniform date matching
              period: p,
              periodsString: String(periodsStr || ""), // Needed for edit modal
              urgency: String(data[i][7] || ""), // Urgency is index 7. Needed for edit modal
              teacherName: String(teacherName || ""),
              teacherEmail: String(email || ""),
              course: String(course || ""),
              room: String(room || ""),
              assignedSub: String(assignedSub || "").trim(),
              reason: String(reason || ""),
              duration: String(duration || ""),
              instructions: String(instructions || "")
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

/**
 * Fetches data for the HR Dashboard.
 * Returns a list of absence request summaries (ignoring canceled requests).
 */
function getHRDashboardData() {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, ["hr", "principal"], "Unauthorized access. HR or Principal role required.");

    var mainSheet = ss.getSheetByName("Absence Requests");
    var rosterSheet = ss.getSheetByName("Staff Roster");
    
    // Look for Payperiods sheet case-insensitively
    var allSheets = ss.getSheets();
    var payPeriodsSheet = null;
    for (var s = 0; s < allSheets.length; s++) {
      if (allSheets[s].getName().toLowerCase() === "payperiods") {
        payPeriodsSheet = allSheets[s];
        break;
      }
    }

    if (!mainSheet) return { requests: [], payPeriods: [] };

    var data = mainSheet.getDataRange().getValues();
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
    var payPeriodsData = payPeriodsSheet ? payPeriodsSheet.getDataRange().getValues() : [];

    var nameLookup = buildNameLookup(rosterData);

    var hrData = [];
    var payPeriods = [];

    // Process Payperiods - start from 0 in case there is no header
    for (var p = 0; p < payPeriodsData.length; p++) {
      var periodNum = String(payPeriodsData[p][0]).trim();
      var startDateRaw = payPeriodsData[p][1];
      var endDateRaw = payPeriodsData[p][2];
      
      // Attempt to validate dates to skip headers
      var isHeader = false;
      if (typeof startDateRaw === 'string' && startDateRaw.toLowerCase().includes('start')) isHeader = true;
      if (typeof endDateRaw === 'string' && endDateRaw.toLowerCase().includes('end')) isHeader = true;
      if (isHeader) continue;
      
      if (periodNum && startDateRaw && endDateRaw) {
        var startFormatted = "";
        if (startDateRaw instanceof Date) {
          startFormatted = Utilities.formatDate(startDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          try {
             startFormatted = Utilities.formatDate(new Date(startDateRaw), Session.getScriptTimeZone(), "yyyy-MM-dd");
          } catch(e) {
             console.error("Error formatting date: " + e.message);
             startFormatted = String(startDateRaw);
          }
        }
        
        var endFormatted = "";
        if (endDateRaw instanceof Date) {
          endFormatted = Utilities.formatDate(endDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          try {
             endFormatted = Utilities.formatDate(new Date(endDateRaw), Session.getScriptTimeZone(), "yyyy-MM-dd");
          } catch(e) {
             console.error("Error formatting date: " + e.message);
             endFormatted = String(endDateRaw);
          }
        }

        payPeriods.push({
          periodNumber: periodNum,
          startDate: startFormatted,
          endDate: endFormatted
        });
      }
    }

    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][17] || "").trim(); // 17 is Status
      if (status === "Canceled") continue;

      var id = data[i][0];
      var email = String(data[i][2]).toLowerCase().trim();
      var dateStr = data[i][3];
      var periodsStr = String(data[i][4]);
      var reason = String(data[i][5]).trim();
      var duration = String(data[i][6]).trim(); // "Full Day" or "Half Day"

      // Get teacher name
      var teacherName = nameLookup[email] || email;

      // Parse date to a comparable format, YYYY-MM-DD
      var dateObj = new Date(dateStr);
      var dateFormatted = "";
      if (!isNaN(dateObj.getTime())) {
          dateFormatted = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
          dateFormatted = dateStr;
      }

      var periods = periodsStr.split(',').map(function(p) { return p.trim(); });
      var assignedSubs = [];

      for (var j = 0; j < periods.length; j++) {
        var p = parseInt(periods[j]);
        if (!isNaN(p)) {
            var subColumnIndex = 8 + p; // 9 for P1, 10 for P2, etc.
            var assignedSub = data[i][subColumnIndex] || "";
            if (assignedSub && String(assignedSub).trim() !== "") {
              assignedSubs.push({
                  name: String(assignedSub).trim(),
                  period: String(p)
              });
            }
        }
      }

      hrData.push({
        id: String(id || ""),
        date: String(dateFormatted || ""),
        teacherName: String(teacherName || ""),
        reason: reason,
        duration: duration,
        assignedSubs: assignedSubs
      });
    }

    return {
      requests: hrData,
      payPeriods: payPeriods
    };

  } catch (e) {
    console.error("Error fetching HR Dashboard Data: " + e.message);
    throw new Error("Failed to load HR dashboard data.");
  }
}
/**
 * Fetches all necessary data for the initial application load in a single call.
 */
function getInitialPayload() {
  try {
    var ss = getSS();
    var email = Session.getActiveUser().getEmail();
    var targetEmail = String(email).toLowerCase();

    // 1. Fetch all required sheets
    var rosterSheet = ss.getSheetByName("Staff Roster");
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    var roleSheet = ss.getSheetByName("User Roles");
    var roleData = roleSheet ? roleSheet.getDataRange().getValues() : [];

    var settings = getSettings(ss); // Already passes ss to avoid fetching again

    var mainSheet = ss.getSheetByName("Absence Requests");
    var absenceData = mainSheet ? mainSheet.getDataRange().getValues() : [];

    var masterScheduleSheet = ss.getSheetByName("Master Schedule");
    var scheduleData = masterScheduleSheet ? masterScheduleSheet.getDataRange().getValues() : [];

    // Look for Payperiods sheet case-insensitively
    var allSheets = ss.getSheets();
    var payPeriodsSheet = null;
    for (var s = 0; s < allSheets.length; s++) {
      if (allSheets[s].getName().toLowerCase() === "payperiods") {
        payPeriodsSheet = allSheets[s];
        break;
      }
    }
    var payPeriodsData = payPeriodsSheet ? payPeriodsSheet.getDataRange().getValues() : [];


    // --- Build lookups ---
    var scheduleLookup = buildScheduleLookup(scheduleData);
    var nameLookup = buildNameLookup(rosterData);


    // --- 2. Extract User Data ---
    var name = "Teacher";
    for (var i = 1; i < rosterData.length; i++) {
      if (String(rosterData[i][1]).toLowerCase() === targetEmail) {
        name = rosterData[i][0];
        break;
      }
    }
    var userName = String(name).trim().toLowerCase();

    var role = "User";
    for (var j = 1; j < roleData.length; j++) {
      if (String(roleData[j][0]).toLowerCase() === targetEmail) {
        role = roleData[j][1];
        break;
      }
    }
    var lowerRole = String(role).toLowerCase();

    var appUrl = settings["App URL"] || "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec";
    var urgencyCutoffTime = settings["Urgency Cutoff Time"] || "15";
    var defaultAbsenceReasons = JSON.stringify([
        {reason: "Personal", hrRequired: false},
        {reason: "Professional Development", hrRequired: false},
        {reason: "Retreat", hrRequired: false},
        {reason: "Athletics", hrRequired: false},
        {reason: "Jury Duty", hrRequired: true},
        {reason: "Bereavement", hrRequired: true}
    ]);
    var absenceReasons = settings["Absence Reasons"] || defaultAbsenceReasons;

    var userData = {
      name: String(name),
      role: String(role),
      email: String(email),
      appUrl: String(appUrl),
      urgencyCutoffTime: String(urgencyCutoffTime),
      absenceReasons: String(absenceReasons)
    };


    // --- 3. Extract common data (My Absences, My Sub Duties, Open Jobs) ---
    var myAbsences = [];
    var mySubDuties = [];
    var todaysOpenJobs = [];

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var targetEndWeek = new Date(today);
    targetEndWeek.setDate(today.getDate() + 6); // Up to next week
    targetEndWeek.setHours(23, 59, 59, 999);

    var targetEndToday = new Date(today);
    targetEndToday.setHours(23, 59, 59, 999);

    for (var i = 1; i < absenceData.length; i++) {
      var row = absenceData[i];
      var status = String(row[17] || 'Active');
      if (status === 'Canceled') continue;

      var rowTeacherEmail = String(row[2]).toLowerCase();
      var dateVal = row[3];
      if (!dateVal) continue;

      var rowDate = new Date(dateVal);
      var isDateValid = !isNaN(rowDate.getTime());

      var formattedDate = "Unknown Date";
      var yyyymmdd = "";
      if (isDateValid) {
         formattedDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MMM d, yyyy");
         yyyymmdd = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
         try {
            formattedDate = new Date(dateVal).toLocaleDateString();
            yyyymmdd = Utilities.formatDate(new Date(dateVal), Session.getScriptTimeZone(), "yyyy-MM-dd");
         } catch(e) {
            console.error("Error formatting date: " + e.message);
            yyyymmdd = String(dateVal); // fallback
         }
      }

      // My Absences
      if (rowTeacherEmail === targetEmail) {
        var urgencyStr = String(row[7] || '');
        myAbsences.push({
          id: String(row[0]),
          date: String(formattedDate),
          rawDateString: String(dateVal),
          formDateString: String(yyyymmdd),
          periods: String(row[4]),
          reason: String(row[5]),
          urgency: urgencyStr.includes('Urgent') ? 'Urgent' : 'Standard',
          duration: String(row[6]),
          instructions: String(row[8])
        });
      }

      // My Sub Duties & Open Jobs
      if (isDateValid && rowDate >= today && rowDate <= targetEndWeek) {
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
        if (teacherName.includes(",")) {
          var parts = teacherName.split(",");
          teacherName = parts[1].trim() + " " + parts[0].trim();
        }

        var periodsRequested = String(row[4]).split(",").map(function(p) { return p.trim(); });
        var rowId = String(row[0]);
        var rawDate = Number(rowDate.getTime());
        var reason = String(row[5]);
        var duration = String(row[6]);
        var instructions = String(row[8]);

        for (var p = 1; p <= 8; p++) {
          if (periodsRequested.indexOf(String(p)) !== -1) {
            var subColumnIndex = 8 + p;
            var assignedSub = String(row[subColumnIndex] || "").trim().toLowerCase();

            var joinKey = rowTeacherEmail + "-" + p;
            var scheduleInfo = scheduleLookup[joinKey];
            var roomStr = scheduleInfo && scheduleInfo.room ? String(scheduleInfo.room) : "No Class Assigned";
            var courseStr = scheduleInfo && scheduleInfo.course ? String(scheduleInfo.course) : "No Class Assigned";

            var jobObj = {
                id: rowId,
                teacherName: String(teacherName),
                teacherEmail: String(rowTeacherEmail),
                date: formattedDate,
                formDateString: yyyymmdd,
                period: String(p),
                rawDate: rawDate,
                room: roomStr,
                course: courseStr,
                reason: reason,
                duration: duration,
                instructions: instructions
            };

            // My Sub Duties
            if (assignedSub === userName) {
              mySubDuties.push(jobObj);
            }

            // Open Jobs (Today only)
            if (assignedSub === "" && rowDate <= targetEndToday) {
              todaysOpenJobs.push(jobObj);
            }
          }
        }
      }
    }

    myAbsences.reverse();

    var sortJobs = function(a, b) {
      if (a.rawDate === b.rawDate) {
        return parseInt(a.period) - parseInt(b.period);
      }
      return a.rawDate - b.rawDate;
    };
    mySubDuties.sort(sortJobs);
    todaysOpenJobs.sort(sortJobs);


    var payload = {
      userData: userData,
      myAbsences: myAbsences,
      mySubDuties: mySubDuties,
      todaysOpenJobs: todaysOpenJobs
    };


    // --- 4. Extract Admin / Sub Coordinator data if applicable ---
    if (lowerRole === "admin" || lowerRole === "sub coordinator" || lowerRole === "hr" || lowerRole === "principal") {
      // Staff List
      var staffList = [];
      for (var i = 1; i < rosterData.length; i++) {
        var staffName = String(rosterData[i][0]).trim();
        var duty = String(rosterData[i][2] || "").trim();
        if (staffName) {
          var display = staffName;
          if (duty) display = staffName + " - " + duty;
          staffList.push({ name: staffName, display: display, duty: duty });
        }
      }
      staffList.sort(function(a, b) {
        var nA = a.name.toLowerCase();
        var nB = b.name.toLowerCase();
        if (nA < nB) return -1;
        if (nA > nB) return 1;
        return 0;
      });
      payload.staffList = staffList;
    }

    if (lowerRole === "admin" || lowerRole === "sub coordinator") {
      // Quick Cover Data
      var quickCover = [];
      var targetEndQC = new Date(today);
      var dayOfWeek = today.getDay();
      var daysToAdd = 1;
      if (dayOfWeek === 5) daysToAdd = 3;
      else if (dayOfWeek === 6) daysToAdd = 2;
      targetEndQC.setDate(today.getDate() + daysToAdd);
      targetEndQC.setHours(23, 59, 59, 999);

      for (var i = 1; i < absenceData.length; i++) {
        var row = absenceData[i];
        var status = String(row[17] || 'Active');
        if (status === 'Canceled') continue;

        var dateVal = row[3];
        if (!dateVal) continue;
        var rowDate = new Date(dateVal);
        if (isNaN(rowDate.getTime())) continue;

        if (rowDate >= today && rowDate <= targetEndQC) {
          var rowTeacherEmail = String(row[2]).toLowerCase();
          var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
          if (teacherName.includes(",")) {
            var parts = teacherName.split(",");
            teacherName = parts[1].trim() + " " + parts[0].trim();
          }

          var periodsRequested = String(row[4]).split(",").map(function(p) { return p.trim(); });
          var rowId = String(row[0]);
          var formattedDate = String(Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MMM d, yyyy"));
          var formDateString = String(Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd"));
          var rawDate = Number(rowDate.getTime());
          var reason = String(row[5]);
          var duration = String(row[6]);
          var instructions = String(row[8]);

          for (var p = 1; p <= 8; p++) {
            if (periodsRequested.indexOf(String(p)) !== -1) {
              var assignedSub = row[8 + p];
              if (!assignedSub || String(assignedSub).trim() === "") {
                var joinKey = rowTeacherEmail + "-" + p;
                var scheduleInfo = scheduleLookup[joinKey];
                quickCover.push({
                  id: rowId,
                  teacherName: String(teacherName),
                  teacherEmail: String(rowTeacherEmail),
                  date: formattedDate,
                  formDateString: formDateString,
                  period: String(p),
                  rawDate: rawDate,
                  room: scheduleInfo && scheduleInfo.room ? String(scheduleInfo.room) : "No Class Assigned",
                  course: scheduleInfo && scheduleInfo.course ? String(scheduleInfo.course) : "No Class Assigned",
                  reason: reason,
                  duration: duration,
                  instructions: instructions
                });
              }
            }
          }
        }
      }
      quickCover.sort(sortJobs);
      payload.quickCover = quickCover;

      // Admin Dashboard Data
      var adminData = [];
      for (var i = 1; i < absenceData.length; i++) {
        var row = absenceData[i];
        if (String(row[17] || "").trim() === "Canceled") continue;

        var rowTeacherEmail = String(row[2]).toLowerCase().trim();
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
        var dateStr = row[3];
        var dateObj = new Date(dateStr);
        var dateFormatted = !isNaN(dateObj.getTime()) ? Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dateStr;
        var periodsStr = String(row[4]);
        var periods = periodsStr.split(',').map(function(p) { return p.trim(); });

        for (var j = 0; j < periods.length; j++) {
          var p = parseInt(periods[j]);
          if (!isNaN(p)) {
            var scheduleKey = rowTeacherEmail + "-" + p;
            adminData.push({
              id: String(row[0] || ""),
              originalDate: String(dateStr || ""),
              date: String(dateFormatted || ""),
              formDateString: String(dateFormatted || ""),
              period: p,
              periodsString: String(periodsStr || ""),
              urgency: String(row[7] || ""),
              teacherName: String(teacherName || ""),
              teacherEmail: String(rowTeacherEmail || ""),
              course: scheduleLookup[scheduleKey] ? String(scheduleLookup[scheduleKey].course) : "",
              room: scheduleLookup[scheduleKey] ? String(scheduleLookup[scheduleKey].room) : "",
              assignedSub: String(row[8 + p] || "").trim(),
              reason: String(row[5] || "").trim(),
              duration: String(row[6] || "").trim(),
              instructions: String(row[8] || "").trim()
            });
          }
        }
      }
      adminData.sort(function(a, b) {
        if (a.date < b.date) return -1;
        if (a.date > b.date) return 1;
        return a.period - b.period;
      });
      payload.adminData = adminData;
    }


    // --- 5. Extract HR data if applicable ---
    if (lowerRole === "hr" || lowerRole === "principal") {
      var hrData = [];
      var payPeriods = [];

      for (var p = 0; p < payPeriodsData.length; p++) {
        var periodNum = String(payPeriodsData[p][0]).trim();
        var startDateRaw = payPeriodsData[p][1];
        var endDateRaw = payPeriodsData[p][2];

        var isHeader = false;
        if (typeof startDateRaw === 'string' && startDateRaw.toLowerCase().includes('start')) isHeader = true;
        if (typeof endDateRaw === 'string' && endDateRaw.toLowerCase().includes('end')) isHeader = true;
        if (isHeader) continue;

        if (periodNum && startDateRaw && endDateRaw) {
          var startFormatted = startDateRaw instanceof Date ? Utilities.formatDate(startDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") :
            (function(){ try { return Utilities.formatDate(new Date(startDateRaw), Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e) { console.error("Error formatting date: " + e.message); return String(startDateRaw); } })();
          var endFormatted = endDateRaw instanceof Date ? Utilities.formatDate(endDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") :
            (function(){ try { return Utilities.formatDate(new Date(endDateRaw), Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e) { console.error("Error formatting date: " + e.message); return String(endDateRaw); } })();

          payPeriods.push({
            periodNumber: periodNum,
            startDate: startFormatted,
            endDate: endFormatted
          });
        }
      }

      for (var i = 1; i < absenceData.length; i++) {
        var row = absenceData[i];
        if (String(row[17] || "").trim() === "Canceled") continue;

        var rowTeacherEmail = String(row[2]).toLowerCase().trim();
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
        var dateStr = row[3];
        var dateObj = new Date(dateStr);
        var dateFormatted = !isNaN(dateObj.getTime()) ? Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dateStr;

        var periodsStr = String(row[4]);
        var periods = periodsStr.split(',').map(function(p) { return p.trim(); });
        var assignedSubs = [];

        for (var j = 0; j < periods.length; j++) {
          var p = parseInt(periods[j]);
          if (!isNaN(p)) {
            var assignedSub = row[8 + p];
            if (assignedSub && String(assignedSub).trim() !== "") {
              assignedSubs.push({ name: String(assignedSub).trim(), period: String(p) });
            }
          }
        }

        hrData.push({
          id: String(row[0] || ""),
          date: String(dateFormatted || ""),
          teacherName: String(teacherName || ""),
          reason: String(row[5]).trim(),
          duration: String(row[6]).trim(),
          assignedSubs: assignedSubs
        });
      }

      payload.hrData = {
        requests: hrData,
        payPeriods: payPeriods
      };
    }

    return payload;

  } catch (err) {
    throw new Error("Failed to get initial payload: " + err.message);
  }
}
