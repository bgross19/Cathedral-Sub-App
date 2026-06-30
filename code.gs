
function notifyAdminOfError(funcName, e) {
  console.error("Global Error in " + funcName + ": " + e.message + "\nStack: " + e.stack);
  try {
    var settings = getSettings();
    var adminEmail = settings["Redirect Email"];
    if (adminEmail && adminEmail.trim() !== "") {
      var subject = "Critical App Error: " + funcName;
      var body = "An error occurred in the Cathedral Sub App.\n\n" +
                 "Function: " + funcName + "\n" +
                 "User: " + Session.getActiveUser().getEmail() + "\n" +
                 "Error Message: " + e.message + "\n\n" +
                 "Stack Trace:\n" + e.stack;

      MailApp.sendEmail({
        to: adminEmail,
        subject: subject,
        body: body
      });
    }
  } catch (mailError) {
    console.error("Failed to send admin error email: " + mailError.message);
  }
}

/**
 * Helper to parse settings from the sheet data.
 */
function _parseSettingsData(settingsSheet) {
  var settings = {};
  if (!settingsSheet) return settings;

  var data = settingsSheet.getDataRange().getValues();
  // Skip header row
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var value = String(data[i][1]).trim();
    if (key) {
      settings[key] = value;
    }
  }
  return settings;
}

/**
 * Helper to merge defaults and append missing ones to the sheet.
 */
function _mergeAndAppendMissingDefaults(settings, defaults, settingsSheet) {
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
}

/**
 * Retrieves settings from the Settings sheet as an object.
 * Uses defaults in memory if the sheet does not exist or user lacks permission.
 */
var _globalSettingsCache = null;

function getSettings(ss) {
  if (_globalSettingsCache) return _globalSettingsCache;

  var cache = CacheService.getScriptCache();
  var cachedSettings = cache.get("app_settings");
  if (cachedSettings) {
    try {
      _globalSettingsCache = JSON.parse(cachedSettings);
      return _globalSettingsCache;
    } catch (e) {
      console.warn("Failed to parse cached settings, reading from sheet.");
    }
  }

  var defaults = {
    "Email Mode": "Live",
    "Redirect Email": "Bgross@gocathedral.com",
    "App URL": "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec",
    "Urgency Cutoff Time": "15",
    "Term ID": "3503",
    "PS_CLIENT_ID": "",
    "PS_CLIENT_SECRET": "",
    "PS_URL": "",
    "Absence Reasons": JSON.stringify([
      {reason: "Personal", hrRequired: false},
      {reason: "Professional Development", hrRequired: false},
      {reason: "Retreat", hrRequired: false},
      {reason: "Athletics", hrRequired: false},
      {reason: "Jury Duty", hrRequired: true},
      {reason: "Bereavement", hrRequired: true}
    ]),
    "RolePermissions": JSON.stringify({
        "admin": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": true },
        "hr": { "Admin Dashboard": false, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": false, "My Past Absences": true, "Settings": true },
        "sub coordinator": { "Admin Dashboard": true, "HR Dashboard": false, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": false },
        "principal": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": false, "My Past Absences": true, "Settings": false },
        "user": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": false },
        "substitute": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": false }
      })
  };

  try {
    var sheetSS = ss || getSS();
    var settingsSheet = getSheetOrThrow(sheetSS, "Settings");

    var settings = _parseSettingsData(settingsSheet);
    settings = _mergeAndAppendMissingDefaults(settings, defaults, settingsSheet);

    _globalSettingsCache = settings;
    cache.put("app_settings", JSON.stringify(settings), 300); // Cache for 5 minutes
    return settings;
  } catch (e) {
    console.warn("Could not read settings from spreadsheet, using defaults: " + e.message);
    _globalSettingsCache = defaults;
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

/**
 * Asserts that a user has a specific permission view via Settings.
 */
function assertPermission(user, viewName, customErrorMessage) {
  if (!user || !user.role) {
    throw new Error(customErrorMessage || "Unauthorized");
  }

  var lowerRole = user.role.toLowerCase();
  if (lowerRole === "admin") return; // Admin bypass

  var settings = getSettings();
  var rolePermissionsStr = settings["RolePermissions"] || "{}";

  if (!hasPermission(lowerRole, viewName, rolePermissionsStr)) {
    throw new Error(customErrorMessage || "Unauthorized");
  }
}

function getUserData(ss) {
  var email = Session.getActiveUser().getEmail();
  var ss = ss || getSS();
  
  var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
  var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
  var name = "Teacher"; 
  var targetEmail = String(email).toLowerCase();
  
  for (var i = 1; i < rosterData.length; i++) {
    if (String(rosterData[i][1]).toLowerCase() === targetEmail) {
      name = rosterData[i][0]; 
      break;
    }
  }
  
  var roleSheet = getSheetOrThrow(ss, "User Roles");
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
      var rosterSheet = ss.getSheetByName("Staff Roster");
      if (rosterSheet) {
        var rosterData = rosterSheet.getDataRange().getValues();
        var targetEmail = String(to).toLowerCase().trim();
        for (var i = 1; i < rosterData.length; i++) {
          if (String(rosterData[i][1]).toLowerCase().trim() === targetEmail) {
            recipientName = String(rosterData[i][0]).trim();
            break;
          }
        }
      }
    } catch(e) {
      // Ignore roster lookup errors, fallback to email
    }

    // Reformat name if it contains a comma (e.g., "Last, First")
    if (recipientName.indexOf(",") > -1) {
      var parts = recipientName.split(",");
      recipientName = parts[1].trim() + " " + parts[0].trim();
    }

    // Split on space and get the first and last name
    var nameParts = recipientName.split(" ");
    var firstName = nameParts[0];
    var lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : "";
    var formattedName = lastName ? firstName + " " + lastName : firstName;

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

    for (var i = 1; i < data.length; i++) {
      if (data[i][5] === "Pending") {
        var to = data[i][1];
        var subject = data[i][2];
        var body = data[i][3];
        var optionsStr = data[i][4];
        var options = {};

        try {
          options = JSON.parse(optionsStr);
        } catch (e) {
          console.error("Failed to parse options for queued email: " + e.message);
        }

        try {
          var result = sendEmailHelper(to, subject, body, options);
          if (result === "SUPPRESSED") {
            sheet.getRange(i + 1, 6).setValue("Suppressed (Off)");
          } else {
            sheet.getRange(i + 1, 6).setValue("Sent");
          }
        } catch (e) {
          console.error("Failed to send queued email to " + to + ": " + e.message);
          sheet.getRange(i + 1, 6).setValue("Failed: " + e.message);
        }
      }
    }

    // Optional: Cleanup old sent/failed emails
    // We could delete rows that are marked "Sent" to keep the sheet small
    var rowsToDelete = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var status = String(sheet.getRange(i + 1, 6).getValue() || "");
      if (status === "Sent" || status.indexOf("Failed") > -1) {
         rowsToDelete.push(i + 1);
      }
    }

    // Delete from bottom up
    for (var j = 0; j < rowsToDelete.length; j++) {
       sheet.deleteRow(rowsToDelete[j]);
    }

  } catch (e) {
    console.error("Error in processEmailQueue: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sets up a time-driven trigger to process the email queue every 1 minute.
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
    assertPermission(user, "Settings");

    var roleSheet = getSheetOrThrow(ss, "User Roles");
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

/**
 * Fetches the staff roster for the Admin Settings dashboard.
 */
function getStaffRosterForAdmin() {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var data = rosterSheet.getDataRange().getValues();
    var roster = [];

    // Assuming row 0 is header: Name, Email, Duty
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][0] || "").trim();
      var email = String(data[i][1] || "").trim();
      var duty = String(data[i][2] || "").trim();

      if (name || email) {
        roster.push({ name: name, email: email, duty: duty });
      }
    }

    // Sort alphabetically by name
    roster.sort(function(a, b) {
        var nA = a.name.toLowerCase();
        var nB = b.name.toLowerCase();
        if (nA < nB) return -1;
        if (nA > nB) return 1;
        return 0;
    });

    return roster;
  } catch (err) {
    notifyAdminOfError("getStaffRosterForAdmin", err);
    throw new Error("Failed to load staff roster: " + err.message);
  }
}

var _globalRosterCache = null;

/**
 * Retrieves the staff roster data from CacheService, falling back to reading the sheet.
 */
function getRosterDataCached(ss) {
  if (_globalRosterCache) return _globalRosterCache;

  var cache = CacheService.getScriptCache();
  var cachedData = cache.get("staff_roster_data");
  if (cachedData) {
    try {
      _globalRosterCache = JSON.parse(cachedData);
      return _globalRosterCache;
    } catch (e) {
      console.warn("Failed to parse cached roster data.");
    }
  }

  var sheetSS = ss || getSS();
  var rosterSheet = getSheetOrThrow(sheetSS, "Staff Roster");
  var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

  // Try to cache if it's not huge
  try {
     var stringified = JSON.stringify(rosterData);
     if (stringified.length < 100000) { // Keep under 100KB cache limit
        cache.put("staff_roster_data", stringified, 1800); // 30 mins
     }
  } catch (e) {}

  _globalRosterCache = rosterData;
  return rosterData;
}

function clearRosterCache() {
  _globalRosterCache = null;
  CacheService.getScriptCache().remove("staff_roster_data");
}

/**
 * Saves a staff member (creates or updates) for the Admin Settings dashboard.
 */
function saveStaffMemberAdmin(staffData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("saveStaffMemberAdmin_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var data = rosterSheet.getDataRange().getValues();

    var originalEmail = staffData.originalEmail ? String(staffData.originalEmail).trim().toLowerCase() : "";
    var newEmail = String(staffData.email).trim();
    var newName = String(staffData.name).trim();
    var newDuty = String(staffData.duty || "").trim();

    if (!newEmail || !newName) {
       return { success: false, error: "Name and Email are required." };
    }

    var rowIndexToUpdate = -1;

    if (originalEmail) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][1]).trim().toLowerCase() === originalEmail) {
          rowIndexToUpdate = i + 1;
          break;
        }
      }
    }

    if (rowIndexToUpdate !== -1) {
       // Update existing
       rosterSheet.getRange(rowIndexToUpdate, 1, 1, 3).setValues([[newName, newEmail, newDuty]]);
       logAuditAction("STAFF_UPDATED", newEmail, "Updated staff member: " + newName + " (" + newDuty + ")");
    } else {
       // Check if new email already exists to prevent duplicates
       for (var i = 1; i < data.length; i++) {
         if (String(data[i][1]).trim().toLowerCase() === newEmail.toLowerCase()) {
            return { success: false, error: "A staff member with this email already exists." };
         }
       }
       // Append new
       rosterSheet.appendRow([newName, newEmail, newDuty]);
       logAuditAction("STAFF_ADDED", newEmail, "Added staff member: " + newName + " (" + newDuty + ")");
    }

    clearRosterCache();

        return { success: true };
  } catch (err) {
    notifyAdminOfError("saveStaffMemberAdmin", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a staff member from the roster.
 */
function deleteStaffMemberAdmin(email) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("deleteStaffMemberAdmin_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var data = rosterSheet.getDataRange().getValues();
    var targetEmail = String(email).trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === targetEmail) {
        logAuditAction("STAFF_DELETED", targetEmail, "Deleted staff member");
        rosterSheet.deleteRow(i + 1);
        clearRosterCache();
            return { success: true };
      }
    }

    return { success: false, error: "Staff member not found." };
  } catch (err) {
    notifyAdminOfError("deleteStaffMemberAdmin", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Processes a bulk upload/update of staff roster records.
 */
function bulkUpsertStaffRoster(updates) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("bulkUpsertStaffRoster_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var data = rosterSheet.getDataRange().getValues();

    var existingEmailsMap = {};
    // Map email to row index (1-based for getRange)
    for (var i = 1; i < data.length; i++) {
       var email = String(data[i][1] || "").trim().toLowerCase();
       if (email) {
          existingEmailsMap[email] = i + 1;
       }
    }

    var newRows = [];
    var processedCount = 0;

    for (var j = 0; j < updates.length; j++) {
       var update = updates[j];
       var email = String(update.email || "").trim();
       var name = String(update.name || "").trim();
       var duty = String(update.duty || "").trim();

       if (!email || !name) continue;

       var lowerEmail = email.toLowerCase();
       if (existingEmailsMap[lowerEmail] && existingEmailsMap[lowerEmail] > 0) {
          // Update existing row directly
          var rowIndex = existingEmailsMap[lowerEmail];
          rosterSheet.getRange(rowIndex, 1, 1, 3).setValues([[name, email, duty]]);
       } else {
          // Track for batch append
          newRows.push([name, email, duty]);
          // To handle duplicates within the upload batch itself
          existingEmailsMap[lowerEmail] = -1;
       }
       processedCount++;
    }

    if (newRows.length > 0) {
       // Append new rows at the end
       var startRow = rosterSheet.getLastRow() + 1;
       rosterSheet.getRange(startRow, 1, newRows.length, 3).setValues(newRows);
    }

    logAuditAction("STAFF_BULK_UPLOAD", "Multiple", "Processed " + processedCount + " staff records");
    clearRosterCache();
    return { success: true, updated: processedCount };
  } catch (err) {
    notifyAdminOfError("bulkUpsertStaffRoster", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function addUserRole(email, role) {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "Settings");

    var roleSheet = getSheetOrThrow(ss, "User Roles");

    logAuditAction("ROLE_ADDED", email, "Assigned role: " + role);
    roleSheet.appendRow([email.toLowerCase().trim(), role.trim()]);
        return { success: true };
  } catch (err) {
    notifyAdminOfError("addUserRole", err);
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
    assertPermission(user, "Settings");

    var roleSheet = getSheetOrThrow(ss, "User Roles");

    var data = roleSheet.getDataRange().getValues();
    var targetEmail = oldEmail.toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === targetEmail) {
        logAuditAction("ROLE_UPDATED", oldEmail, "Changed role to: " + role + " (Email: " + newEmail + ")");
        roleSheet.getRange(i + 1, 1, 1, 2).setValues([[newEmail.toLowerCase().trim(), role.trim()]]);
            return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch (err) {
    notifyAdminOfError("editUserRole", err);
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
    assertPermission(user, "Settings");

    var roleSheet = getSheetOrThrow(ss, "User Roles");

    var data = roleSheet.getDataRange().getValues();
    var targetEmail = email.toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === targetEmail) {
        logAuditAction("ROLE_DELETED", targetEmail, "Removed role");
        roleSheet.deleteRow(i + 1);
            return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch (err) {
    notifyAdminOfError("deleteUserRole", err);
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
    assertPermission(user, "Settings");
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
    assertPermission(user, "Settings");

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
    logAuditAction("SETTINGS_UPDATED", "Global", "Updated application settings");

    // Clear the cache
    _globalSettingsCache = null;
    var cache = CacheService.getScriptCache();
    cache.remove("app_settings");

        return { success: true };
  } catch (err) {
    notifyAdminOfError("updateSettings", err);
    return { success: false, error: err.message };
  }
}

/**
 * Fetches the logged-in user's upcoming absences.
 */

/**
 * Fetches the sub duties assigned to the logged-in user over the next calendar week.
 */

/**
 * Fetches unfilled sub requests for today.
 */

/**
 * Fetches unfilled sub requests for the next 2 days (or through Monday if weekend) for the Admin Dashboard.
 */

/**
 * Fetches the list of staff names from the Staff Roster.
 */

function getCoordinatorEmail(ss) {
  var ss = ss || getSS();
  var roleSheet = getSheetOrThrow(ss, "User Roles");
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

    enqueueEmail(coordinatorEmail, subject, body, { htmlBody: htmlBody });
  }
}

function submitAbsence(formData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("submitAbsence", e);
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var mainSheet = getSheetOrThrow(ss, "Absence Requests");
    
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
    
    logAuditAction("ABSENCE_SUBMITTED", uniqueId, "Requested coverage for " + formData.date + " (Periods: " + formData.periods + ")");
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

    enqueueEmail(email, confSubject, confBody, { htmlBody: confHtmlBody });

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
    notifyAdminOfError("cancelMySubDuty", e);
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
          logAuditAction("SUB_DUTY_CANCELLED", absenceId, "Cancelled coverage for period " + period);

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

            enqueueEmail(coordinatorEmail, subject, body, { cc: userEmail, htmlBody: htmlBody });
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
    notifyAdminOfError("cancelAbsence", e);
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var rosterData = getRosterDataCached(ss);
    var scheduleData = getMasterScheduleData();

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
          assertPermission(user, "Admin Dashboard", "Unauthorized to cancel this absence.");
        }

        sheet.getRange(i + 1, 18).setValue("Canceled");
        logAuditAction("ABSENCE_CANCELLED", absenceId, "Cancelled entire absence request");

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

           enqueueEmail(teacherEmail, teacherSubject, teacherBody, {htmlBody: teacherHtml});
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
    notifyAdminOfError("updateAbsence", e);
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var rosterData = getRosterDataCached(ss);
    var scheduleData = getMasterScheduleData();

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
          assertPermission(user, "Admin Dashboard", "Unauthorized to modify this absence.");
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
        logAuditAction("ABSENCE_UPDATED", absenceId, "Updated absence details (Date: " + formData.date + ", Periods: " + formData.periods + ")");

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

           enqueueEmail(teacherEmail, teacherSubject, teacherBody, {htmlBody: teacherHtml});
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
 * Helper to get teacher name from email.
 */
/**
 * Helper to get full absence details.
 */
function getAbsenceDetails(absenceId, period) {
  var ss = getSS();
  var sheet = getSheetOrThrow(ss, "Absence Requests");
  var rosterSheet = getSheetOrThrow(ss, "Staff Roster");

  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();

  var nameLookup = {};
  if (rosterSheet) {
    var rosterData = rosterSheet.getDataRange().getValues();
    nameLookup = buildNameLookup(rosterData);
  }

  var scheduleLookup = {};
  var scheduleData = getMasterScheduleData();
  if (scheduleData.length > 0) {
    scheduleLookup = buildScheduleLookup(scheduleData);
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(absenceId)) {
      var details = getAbsenceDetailsLocal(data[i], period, scheduleLookup, nameLookup);
      details.rowIndex = i + 1;
      return details;
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
    enqueueEmail(subEmail, subject, body, { htmlBody: htmlBody });
  } catch (e) {
    console.error("Failed to enqueue email to " + subEmail + ": " + e.message);
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
    notifyAdminOfError("assignSubToPeriod", e);
    return { success: false, error: "The server is currently busy. Please try again in a few moments." };
  }

  try {
    var ss = getSS();
    var sheet = getSheetOrThrow(ss, "Absence Requests");

    var rosterData = getRosterDataCached(ss);

    var subEmailLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

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

        // Check if the new sub is absent for a full day on the same date
        if (newSub !== "") {
          var targetDateRaw = data[i][3]; // Date object or string
          var targetDateStr = (targetDateRaw instanceof Date) ? Utilities.formatDate(targetDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(targetDateRaw).trim();
          var newSubEmail = (subEmailLookup[newSub] || "").toLowerCase();

          if (newSubEmail !== "") {
            for (var j = 1; j < data.length; j++) {
              var rowEmail = String(data[j][2] || "").toLowerCase();
              if (rowEmail === "") continue;

              var rowStatus = String(data[j][17] || "Active");
              var rowDuration = String(data[j][6] || "").trim();

              if (rowEmail === newSubEmail && rowStatus !== "Canceled" && rowDuration === "Full Day") {
                var rowDateRaw = data[j][3];
                var rowDateStr = (rowDateRaw instanceof Date) ? Utilities.formatDate(rowDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(rowDateRaw).trim();

                if (rowDateStr === targetDateStr) {
                  var currentUserEmail = (Session.getActiveUser().getEmail() || "").toLowerCase();
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
              var scheduleData = getMasterScheduleData();
              var scheduleLookup = buildScheduleLookup(scheduleData);
              var details = getAbsenceDetailsLocal(data[i], period, scheduleLookup, nameLookup);
              sendSubNotification(existingEmail, 'Canceled', details);
           }
        }

        // Write the subname
        sheet.getRange(i + 1, subColumnIndex).setValue(newSub);
        logAuditAction("SUB_ASSIGNED", absenceId, "Assigned " + (newSub || "NO ONE") + " to period " + period);

        // Notify new sub if there is one
        if (newSub) {
           var newEmail = subEmailLookup[newSub];
           if (newEmail) {
              var scheduleData = getMasterScheduleData();
              var scheduleLookup = buildScheduleLookup(scheduleData);
              var details = getAbsenceDetailsLocal(data[i], period, scheduleLookup, nameLookup);
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

/**
 * Fetches data for the HR Dashboard.
 * Returns a list of absence request summaries (ignoring canceled requests).
 */
/**
 * Fetches all necessary data for the initial application load in a single call.
 */

/**
 * Helper to check if a role has a specific permission.
 */
function hasPermission(role, view, rolePermissionsStr) {
  try {
    var permissions = JSON.parse(rolePermissionsStr || '{}');
    var lowerRole = String(role).toLowerCase();
    if (permissions[lowerRole] && permissions[lowerRole][view] === true) {
      return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

function getInitialPayload() {
  try {
    var ss = getSS();
    var email = Session.getActiveUser().getEmail();
    var targetEmail = String(email).toLowerCase();

    // 1. Fetch all required sheets
    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];

    var roleSheet = getSheetOrThrow(ss, "User Roles");
    var roleData = roleSheet ? roleSheet.getDataRange().getValues() : [];

    var settings = getSettings(ss); // Already passes ss to avoid fetching again

    var mainSheet = getSheetOrThrow(ss, "Absence Requests");
    var absenceData = mainSheet ? mainSheet.getDataRange().getValues() : [];

    var scheduleData = getMasterScheduleData();

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
    var rolePermissions = settings["RolePermissions"] || "{}";

    // Set permissions to return in payload
    var permissions = {
      "Admin Dashboard": hasPermission(lowerRole, "Admin Dashboard", rolePermissions),
      "HR Dashboard": hasPermission(lowerRole, "HR Dashboard", rolePermissions),
      "Today at a Glance": hasPermission(lowerRole, "Today at a Glance", rolePermissions),
      "My Upcoming Sub Duties": hasPermission(lowerRole, "My Upcoming Sub Duties", rolePermissions),
      "Today's Open Jobs": hasPermission(lowerRole, "Today's Open Jobs", rolePermissions),
      "My Past Absences": hasPermission(lowerRole, "My Past Absences", rolePermissions),
      "Settings": hasPermission(lowerRole, "Settings", rolePermissions) || lowerRole === "admin"
    };

    var userData = {
      name: String(name),
      role: String(role),
      email: String(email),
      appUrl: String(appUrl),
      urgencyCutoffTime: String(urgencyCutoffTime),
      absenceReasons: String(absenceReasons)
    };


    // --- 3. Extract common data (My Absences, My Sub Duties, Open Jobs) ---
    var myUpcomingAbsences = [];
    var myPastAbsences = [];
    var mySubDuties = [];
    var todaysOpenJobs = [];

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var targetEndWeek = new Date(today);
    targetEndWeek.setDate(today.getDate() + 6); // Up to next week
    targetEndWeek.setHours(23, 59, 59, 999);

    var targetEndToday = new Date(today);
    targetEndToday.setHours(23, 59, 59, 999);

    var settings = getSettings();
    var fetchWindowDays = parseInt(settings["Data Fetch Window (Days)"]);
    if (isNaN(fetchWindowDays)) fetchWindowDays = 30; // default to 30 days
    var cutoffDate = new Date(today);
    cutoffDate.setDate(today.getDate() - fetchWindowDays);
    cutoffDate.setHours(0, 0, 0, 0);

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
        var absenceObj = {
          id: String(row[0]),
          date: String(formattedDate),
          rawDate: isDateValid ? Number(rowDate.getTime()) : 0,
          rawDateString: String(dateVal),
          formDateString: String(yyyymmdd),
          periods: String(row[4]),
          reason: String(row[5]),
          urgency: urgencyStr.includes('Urgent') ? 'Urgent' : 'Standard',
          duration: String(row[6]),
          instructions: String(row[8])
        };
        if (isDateValid && rowDate < today) {
          myPastAbsences.push(absenceObj);
        } else {
          myUpcomingAbsences.push(absenceObj);
        }
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

    var sortAbsencesAsc = function(a, b) { return a.rawDate - b.rawDate; };
    var sortAbsencesDesc = function(a, b) { return b.rawDate - a.rawDate; };
    myUpcomingAbsences.sort(sortAbsencesAsc);
    myPastAbsences.sort(sortAbsencesDesc);

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
      myAbsences: myUpcomingAbsences,
      myPastAbsences: myPastAbsences,
      mySubDuties: mySubDuties,
      todaysOpenJobs: todaysOpenJobs,
      permissions: permissions
    };


    // --- 4. Extract Admin / Sub Coordinator data if applicable ---
    if (permissions["Admin Dashboard"] || permissions["HR Dashboard"] || permissions["Today at a Glance"]) {
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

    if (permissions["Admin Dashboard"] || permissions["Today at a Glance"] || permissions["Today's Open Jobs"]) {
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
        if (rowDate < cutoffDate) continue;

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

        var dateStr = row[3];
        var dateObj = new Date(dateStr);
        if (!isNaN(dateObj.getTime()) && dateObj < cutoffDate) continue;

        var rowTeacherEmail = String(row[2]).toLowerCase().trim();
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
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
    if (permissions["HR Dashboard"]) {
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

        var dateStr = row[3];
        var dateObj = new Date(dateStr);
        if (!isNaN(dateObj.getTime()) && dateObj < cutoffDate) continue;

        var rowTeacherEmail = String(row[2]).toLowerCase().trim();
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
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
    notifyAdminOfError("getInitialPayload", err);
    throw new Error("Failed to get initial payload: " + err.message);
  }
}



/**
 * Executes a function and catches any unhandled exceptions to notify the admin.
 * @param {Function} func - The function to execute.
 * @param {string} funcName - The name of the function for logging.
 * @returns {any} The result of the function.
 */
function withGlobalExceptionHandler(funcName, func) {
  return function() {
    try {
      return func.apply(this, arguments);
    } catch (e) {
      console.error("Global Error in " + funcName + ": " + e.message + "\nStack: " + e.stack);

      try {
        var settings = getSettings();
        var adminEmail = settings["Redirect Email"];
        if (adminEmail && adminEmail.trim() !== "") {
          var subject = "Critical App Error: " + funcName;
          var body = "An error occurred in the Cathedral Sub App.\n\n" +
                     "Function: " + funcName + "\n" +
                     "User: " + Session.getActiveUser().getEmail() + "\n" +
                     "Error Message: " + e.message + "\n\n" +
                     "Stack Trace:\n" + e.stack;

          MailApp.sendEmail({
            to: adminEmail,
            subject: subject,
            body: body
          });
        }
      } catch (mailError) {
        console.error("Failed to send admin error email: " + mailError.message);
      }

      throw e; // Re-throw so frontend still sees an error
    }
  };
}



/**
 * Clears the Master Schedule cache from the Script Cache.
 * Returns a success object.
 */
function clearMasterScheduleCache() {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "Settings");

    var sheet = ss.getSheetByName("Master Schedule Cache");
    if (sheet) {
      sheet.clearContents();
    }

    // Attempt to warm it immediately
    warmMasterScheduleCache();

    return { success: true };
  } catch (err) {
    notifyAdminOfError("clearMasterScheduleCache", err);
    return { success: false, error: err.message };
  }
}


/**
 * Refreshes requested data components.
 * @param {Array<string>} components - The components to fetch (e.g. ['myAbsences', 'quickCover'])
 * @returns {Object} The requested data.
 */
function refreshData(components) {
  try {
    var payload = getInitialPayload();
    var response = {};
    for (var i = 0; i < components.length; i++) {
        var comp = components[i];
        if (payload[comp] !== undefined) {
            response[comp] = payload[comp];
        }
    }
    return response;
  } catch (e) {
    notifyAdminOfError("refreshData", e);
    throw new Error("Failed to refresh data: " + e.message);
  }
}


/**
 * Logs an action to the hidden Audit Log sheet.
 * @param {string} actionType - The type of action (e.g., "ASSIGN_SUB", "CANCEL_ABSENCE").
 * @param {string} targetId - The ID of the affected record.
 * @param {string} details - A description of the action.
 */
function logAuditAction(actionType, targetId, details) {
  try {
    var ss = getSS();
    var auditSheet = ss.getSheetByName("Audit Log");
    if (!auditSheet) return; // Fail silently if not set up

    var timestamp = new Date();
    var actor = Session.getActiveUser().getEmail() || "Unknown";

    // Sanitize inputs for appendRow
    var safeActionType = String(actionType != null ? actionType : "");
    var safeTargetId = typeof targetId === 'object' ? JSON.stringify(targetId) : String(targetId != null ? targetId : "");
    var safeDetails = typeof details === 'object' ? JSON.stringify(details) : String(details != null ? details : "");

    auditSheet.appendRow([timestamp, actor, safeActionType, safeTargetId, safeDetails]);
  } catch (e) {
    console.error("Failed to log audit action: " + e.message);
  }
}

/**
 * Fetches audit logs within a specific date range for the Admin dashboard.
 */

/**
 * Exports all absence requests as a JSON string to be converted to CSV on the frontend.
 */
function exportAllAbsenceRequests() {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, "admin");

    var sheet = getSheetOrThrow(ss, "Absence Requests");
    var data = sheet.getDataRange().getValues();

    return JSON.stringify(data);
  } catch (err) {
    notifyAdminOfError("exportAllAbsenceRequests", err);
    throw new Error("Failed to export absence requests: " + err.message);
  }
}

/**
 * Archives absence requests before the given date cutoff.
 */
function archiveAbsenceRequests(cutoffDateStr) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("archiveAbsenceRequests_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, "admin");

    var mainSheet = getSheetOrThrow(ss, "Absence Requests");
    var archiveSheet = ss.getSheetByName("Archived Data");
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet("Archived Data");
      var headersRow = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
      archiveSheet.getRange(1, 1, 1, headersRow.length).setValues([headersRow]);
    }

    var data = mainSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, count: 0 };
    }

    var headers = data[0];
    var cutoffDate = new Date(cutoffDateStr);
    cutoffDate.setHours(0, 0, 0, 0);

    var rowsToKeep = [headers];
    var rowsToArchive = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowDate = new Date(row[3]);

      if (!isNaN(rowDate.getTime()) && rowDate < cutoffDate) {
        rowsToArchive.push(row);
      } else {
        rowsToKeep.push(row);
      }
    }

    if (rowsToArchive.length > 0) {
      var startRow = archiveSheet.getLastRow() + 1;
      archiveSheet.getRange(startRow, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);

      mainSheet.clearContents();
      mainSheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);

      logAuditAction("ARCHIVE_DATA", "N/A", "Archived " + rowsToArchive.length + " absence requests older than " + cutoffDateStr);
    }

    return { success: true, count: rowsToArchive.length };
  } catch (err) {
    notifyAdminOfError("archiveAbsenceRequests", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function getAuditLogs(startDateStr, endDateStr) {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "Settings");

    var auditSheet = ss.getSheetByName("Audit Log");
    if (!auditSheet) return [];

    var data = auditSheet.getDataRange().getValues();
    var logs = [];

    var startDate = new Date(startDateStr);
    startDate.setHours(0,0,0,0);
    var endDate = new Date(endDateStr);
    endDate.setHours(23,59,59,999);

    // Skip header
    for (var i = 1; i < data.length; i++) {
      var rowDate = new Date(data[i][0]);
      if (rowDate >= startDate && rowDate <= endDate) {
        logs.push({
          timestamp: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
          actor: String(data[i][1]),
          actionType: String(data[i][2]),
          targetId: String(data[i][3]),
          details: String(data[i][4])
        });
      }
    }

    // Reverse to show newest first
    return logs.reverse();
  } catch (err) {
    notifyAdminOfError("getAuditLogs", err);
    throw new Error("Failed to load audit logs: " + err.message);
  }
}


function bulkUpsertPayPeriods(updates) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("bulkUpsertPayPeriods_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "HR Dashboard");

    var sheet = getSheetOrThrow(ss, "PayPeriods");

    // updates is an array of objects: {period: "", start: "", end: ""}
    var newRows = [];
    for (var i = 0; i < updates.length; i++) {
       var u = updates[i];
       newRows.push([u.period, u.start, u.end]);
    }

    if (newRows.length > 0) {
       var startRow = sheet.getLastRow() + 1;
       sheet.getRange(startRow, 1, newRows.length, 3).setValues(newRows);
    }

    logAuditAction("PAY_PERIODS_BULK_UPLOAD", "Multiple", "Added " + newRows.length + " pay periods");

    // Clear cache because PayPeriods data changed
    // In hrData, pay periods are retrieved directly from the sheet, so it's live
    // but we can clear cache just in case.

    return { success: true, updated: newRows.length };
  } catch (err) {
    notifyAdminOfError("bulkUpsertPayPeriods", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteAllPayPeriods() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("deleteAllPayPeriods_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "HR Dashboard");

    var sheet = getSheetOrThrow(ss, "PayPeriods");
    var lastRow = sheet.getLastRow();

    if (lastRow > 1) {
       sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    }

    logAuditAction("PAY_PERIODS_DELETE_ALL", "All", "Deleted all pay periods");

    return { success: true };
  } catch (err) {
    notifyAdminOfError("deleteAllPayPeriods", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function loadPayPeriodsSettings() {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertPermission(user, "HR Dashboard");

    var sheet = getSheetOrThrow(ss, "PayPeriods");
    var data = sheet.getDataRange().getValues();
    var payPeriods = [];

    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (row[0] && row[1] && row[2]) {
            var periodNum = String(row[0]).trim();
            var startFormatted = row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[1]);
            var endFormatted = row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[2]);

            payPeriods.push({
                period: periodNum,
                start: startFormatted,
                end: endFormatted
            });
        }
    }
    return payPeriods;
  } catch(err) {
    throw new Error("Failed to load pay periods: " + err.message);
  }
}
