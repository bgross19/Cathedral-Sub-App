
function getActiveUserEmail(clientEmail) {
  if (clientEmail && typeof clientEmail === 'string' && clientEmail.indexOf('@') > 0) {
    return clientEmail.trim().toLowerCase();
  }
  try {
    return Session.getActiveUser().getEmail().toLowerCase();
  } catch(e) {
    return "";
  }
}


/**
 * Helper to get the 1-based column index for a given period.
 * 0-based array index is one less.
 */
function getSubColumnIndex(period) {
  var p = String(period).toLowerCase().trim();
  if (p === '0') return 18; // Col R
  if (p === 'a' || p === 'advisory') return 19; // Col S
  var pNum = parseInt(p);
  if (!isNaN(pNum) && pNum >= 1 && pNum <= 8) {
    return 10 + pNum - 1; // Col J is 10 for Period 1
  }
  return -1;
}

/**
 * Helper to get the master schedule join period.
 */
function getScheduleJoinPeriod(period) {
  var p = String(period).toLowerCase().trim();
  if (p === '0') return '9';
  if (p === 'a' || p === 'advisory') return '10';
  return String(period);
}

const APP_VERSION = "1.0.0";
const DEFAULT_APP_URL = "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec";


function notifyAdminOfError(funcName, e) {
  console.error("Global Error in " + funcName + ": " + e.message + "\nStack: " + e.stack);
  try {
    var settings = getSettings();
    var adminEmail = settings["Redirect Email"];
    var senderName = settings["Email Sender Name"] || "Cathedral Sub App";
    var replyToEmail = settings["Reply To Email"] || "";
    if (adminEmail && adminEmail.trim() !== "") {
      var subject = "Critical App Error: " + funcName;
      var body = "An error occurred in the Cathedral Sub App.\n\n" +
                 "Function: " + funcName + "\n" +
                 "User: " + getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined) + "\n" +
                 "Error Message: " + e.message;

      var options = { name: senderName };
      if (replyToEmail && replyToEmail.trim() !== "") options.replyTo = replyToEmail.trim();
      GmailApp.sendEmail(adminEmail, subject, body, options);
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
    "Reply To Email": "",
    "App URL": DEFAULT_APP_URL,
    "Max Multi-Select Days": "5",
    "Urgency Cutoff Time": "15",
    "Term ID": "3503",
    "Green Day Pay Rate": "10",
    "Blue/Gold Day Pay Rate": "20",
    "Absence Reasons": JSON.stringify([
      {reason: "Personal", hrRequired: false, principalRequired: false},
      {reason: "Professional Development", hrRequired: false, principalRequired: false},
      {reason: "Retreat", hrRequired: false, principalRequired: false},
      {reason: "Athletics", hrRequired: false, principalRequired: true},
      {reason: "Jury Duty", hrRequired: true, principalRequired: false},
      {reason: "Bereavement", hrRequired: true, principalRequired: false}
    ]),
    "RolePermissions": JSON.stringify({
        "admin": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "hr": { "Admin Dashboard": false, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": false, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "sub coordinator": { "Admin Dashboard": true, "HR Dashboard": false, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "principal": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": false, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "teacher": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": false, "Add Request on Behalf": false },
        "substitute": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": false, "My Past Absences": false, "Settings": false, "Add Request on Behalf": false }
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
  var email = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName("Staff Roster");

  var isAuthorized = false;
  if (rosterSheet) {
    var rosterData = rosterSheet.getDataRange().getValues();
    var targetEmail = String(email).toLowerCase();
    for (var i = 1; i < rosterData.length; i++) {
      if (String(rosterData[i][1]).toLowerCase() === targetEmail) {
        isAuthorized = true;
        break;
      }
    }
  }

  // Access Denied check moved to frontend to allow custom login

  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = email;
  
  var htmlOutput = template.evaluate();
  htmlOutput.setTitle('Cathedral Sub Coverage');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  return htmlOutput;
}

/**
 * Asserts that a user has one of the allowed roles.
 * @param {Object} user - The user object returned by getUserData(undefined, typeof clientEmail !== 'undefined' ? clientEmail : undefined).
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

function getUserData(ss, clientEmail) {
  var email = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
  ss = ss || getSS();
  
  var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
  var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
  var name = null;
  var role = null;
  var targetEmail = String(email).toLowerCase();
  
  for (var i = 1; i < rosterData.length; i++) {
    if (String(rosterData[i][1]).toLowerCase() === targetEmail) {
      name = rosterData[i][0]; 
      role = rosterData[i][2] ? String(rosterData[i][2]).trim() : "Teacher";
      break;
    }
  }

  if (!name) {
    throw new Error("Access Denied: Please Contact Technology@gocathedral.com");
  }
  
  var settings = getSettings(ss);
  var appUrl = settings["App URL"] || DEFAULT_APP_URL;
  var urgencyCutoffTime = settings["Urgency Cutoff Time"] || "15";
  var defaultAbsenceReasons = JSON.stringify([
      {reason: "Personal", hrRequired: false, principalRequired: false},
      {reason: "Professional Development", hrRequired: false, principalRequired: false},
      {reason: "Retreat", hrRequired: false, principalRequired: false},
      {reason: "Athletics", hrRequired: false, principalRequired: true},
      {reason: "Jury Duty", hrRequired: true, principalRequired: false},
      {reason: "Bereavement", hrRequired: true, principalRequired: false}
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

        if (String(rosterData[i][1]).toLowerCase().trim() === targetEmail) {
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
      formattedName = parts[1].trim() + " " + parts[0].trim();
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

    for (var i = 0; i < data.length; i++) {
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

    for (var j = 1; j < data.length; j++) {
      if (data[j][5] === "Pending") {
        var to = data[j][1];
        var subject = data[j][2];
        var body = data[j][3];
        var optionsStr = data[j][4];
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

    // Optional: Cleanup old sent/failed emails
    // We could delete rows that are marked "Sent" to keep the sheet small
    var rowsToDelete = [];
    for (var i = statuses.length - 1; i >= 1; i--) {
      var status = String(statuses[i][0] || "");
      if (status === "Sent" || status.indexOf("Failed") > -1) {
         rowsToDelete.push(i + 1);
      }
    }

    // Delete from bottom up in contiguous batches
    if (rowsToDelete.length > 0) {
      var startRow = rowsToDelete[0];
      var numRows = 1;

      for (var j = 1; j < rowsToDelete.length; j++) {
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
 * Adds a new user role.
 */

/**
 * Fetches the staff roster for the Admin Settings dashboard.
 */
function getStaffRosterForAdmin(clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var data = rosterSheet.getDataRange().getValues();
    var roster = [];

    // Assuming row 0 is header: Name, Email, Role, Duty
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][0] || "").trim();
      var email = String(data[i][1] || "").trim();
      var role = String(data[i][2] || "").trim();
      var duty = String(data[i][3] || "").trim();

      if (name || email) {
        roster.push({ name: name, email: email, role: role, duty: duty });
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
 * Updates a staff member's role inline from the admin settings dashboard.
 */
function updateStaffRoleInlineAdmin(email, newRole, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("updateStaffRoleInlineAdmin_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var data = rosterSheet.getDataRange().getValues();
    var targetEmail = String(email).trim().toLowerCase();

    var rowIndexToUpdate = -1;
    var currentName = "";
    var currentDuty = "";

    for (var i = 1; i < data.length; i++) {
        if (String(data[i][1]).trim().toLowerCase() === targetEmail) {
            rowIndexToUpdate = i + 1;
            currentName = String(data[i][0]).trim();
            currentDuty = String(data[i][3]).trim();
            break;
        }
    }

    if (rowIndexToUpdate !== -1) {
       rosterSheet.getRange(rowIndexToUpdate, 1, 1, 4).setValues([[currentName, targetEmail, newRole, currentDuty]]);
       logAuditAction("STAFF_UPDATED", targetEmail, "Updated staff role inline: " + currentName + " to " + newRole);
       clearRosterCache();
       return { success: true };
    }
    return { success: false, error: "Staff member not found." };
  } catch (err) {
    notifyAdminOfError("updateStaffRoleInlineAdmin", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates a staff member's duty inline from the admin settings dashboard.
 */
function updateStaffDutyInlineAdmin(email, newDuty, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("updateStaffDutyInlineAdmin_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var data = rosterSheet.getDataRange().getValues();
    var targetEmail = String(email).trim().toLowerCase();

    var rowIndexToUpdate = -1;
    var currentName = "";
    var currentRole = "";

    for (var i = 1; i < data.length; i++) {
        if (String(data[i][1]).trim().toLowerCase() === targetEmail) {
            rowIndexToUpdate = i + 1;
            currentName = String(data[i][0]).trim();
            currentRole = String(data[i][2]).trim();
            break;
        }
    }

    if (rowIndexToUpdate !== -1) {
       rosterSheet.getRange(rowIndexToUpdate, 1, 1, 4).setValues([[currentName, targetEmail, currentRole, newDuty]]);
       logAuditAction("STAFF_UPDATED", targetEmail, "Updated staff duty inline: " + currentName + " to " + newDuty);
       clearRosterCache();
       return { success: true };
    }
    return { success: false, error: "Staff member not found." };
  } catch (err) {
    notifyAdminOfError("updateStaffDutyInlineAdmin", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Saves a staff member (creates or updates) for the Admin Settings dashboard.
 */
function saveStaffMemberAdmin(staffData, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("saveStaffMemberAdmin_lock", e);
    return {
      success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var data = rosterSheet.getDataRange().getValues();

    var originalEmail = staffData.originalEmail ? String(staffData.originalEmail).trim().toLowerCase() : "";
    var newEmail = String(staffData.email).trim();
    var newName = String(staffData.name).trim();
    var newRole = String(staffData.role || "Teacher").trim();
    var newDuty = String(staffData.duty || "").trim();

    if (!newEmail || !newName) {
       return {
      success: false, error: "Name and Email are required." };
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
       rosterSheet.getRange(rowIndexToUpdate, 1, 1, 4).setValues([[newName, newEmail, newRole, newDuty]]);
       logAuditAction("STAFF_UPDATED", newEmail, "Updated staff member: " + newName + " (" + newRole + ", " + newDuty + ")");
    } else {
       // Check if new email already exists to prevent duplicates
       for (var i = 1; i < data.length; i++) {
         if (String(data[i][1]).trim().toLowerCase() === newEmail.toLowerCase()) {
            return {
      success: false, error: "A staff member with this email already exists." };
         }
       }
       // Append new
       rosterSheet.appendRow([newName, newEmail, newRole, newDuty]);
       logAuditAction("STAFF_ADDED", newEmail, "Added staff member: " + newName + " (" + newRole + ", " + newDuty + ")");
    }

    clearRosterCache();

        return {
      success: true };
  } catch (err) {
    notifyAdminOfError("saveStaffMemberAdmin", err);
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a staff member from the roster.
 */
function deleteStaffMemberAdmin(email, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("deleteStaffMemberAdmin_lock", e);
    return {
      success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var data = rosterSheet.getDataRange().getValues();
    var targetEmail = String(email).trim().toLowerCase();
    var targetIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === targetEmail) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex !== -1) {
      logAuditAction("STAFF_DELETED", targetEmail, "Deleted staff member");
      rosterSheet.deleteRow(targetIndex + 1);
      clearRosterCache();
      return {
      success: true };
    }

    return {
      success: false, error: "Staff member not found." };
  } catch (err) {
    notifyAdminOfError("deleteStaffMemberAdmin", err);
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Clears all duties from the Staff Roster.
 */
function clearAllStaffDuties(clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("clearAllStaffDuties_lock", e);
    return {
      success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var lastRow = rosterSheet.getLastRow();

    if (lastRow > 1) {
      // Clear only the 4th column (Duty) from row 2 down to the last row
      rosterSheet.getRange(2, 4, lastRow - 1, 1).clearContent();
    }

    logAuditAction("STAFF_DUTIES_CLEARED", "All", "Admin cleared all staff duties");
    clearRosterCache();

    return {
      success: true };
  } catch (err) {
    notifyAdminOfError("clearAllStaffDuties", err);
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Processes a bulk upload/update of staff roster records.
 */
function bulkUpsertStaffRoster(updates, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("bulkUpsertStaffRoster_lock", e);
    return {
      success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
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
    var dataChanged = false;

    for (var j = 0; j < updates.length; j++) {
       var update = updates[j];
       var email = String(update.email || "").trim();
       var name = String(update.name || "").trim();
       var role = String(update.role || "Teacher").trim();
       var duty = String(update.duty || "").trim();

       if (!email || !name) continue;

       var lowerEmail = email.toLowerCase();
       if (existingEmailsMap[lowerEmail] && existingEmailsMap[lowerEmail] > 0) {
          // Update in memory
          var rowIndex = existingEmailsMap[lowerEmail];
          // getRange is 1-based, array is 0-based
          var dataIndex = rowIndex - 1;

          if (data[dataIndex][0] !== name || data[dataIndex][1] !== email || data[dataIndex][2] !== role || data[dataIndex][3] !== duty) {
            data[dataIndex][0] = name;
            data[dataIndex][1] = email;
            data[dataIndex][2] = role;
            data[dataIndex][3] = duty;
            dataChanged = true;
          }
       } else {
          // Track for batch append
          newRows.push([name, email, role, duty]);
          // To handle duplicates within the upload batch itself
          existingEmailsMap[lowerEmail] = -1;
       }

       processedCount++;
    }

    // Perform a single batch update for existing rows if any were modified
    if (dataChanged) {
       // Using getRange(1, 1, data.length, 4) because we are only updating the first 4 columns,
       // but data might have more columns, so we map the first 4 columns of the data array.
       var updateData = data.map(function(row) {
         return [row[0], row[1], row[2], row[3]];
       });
       rosterSheet.getRange(1, 1, updateData.length, 4).setValues(updateData);
    }

    if (newRows.length > 0) {
       // Append new rows at the end
       var startRow = rosterSheet.getLastRow() + 1;
       rosterSheet.getRange(startRow, 1, newRows.length, 4).setValues(newRows);
    }

    logAuditAction("STAFF_BULK_UPLOAD", "Multiple", "Processed " + processedCount + " staff records");
    clearRosterCache();
    return {
      success: true, updated: processedCount };
  } catch (err) {
    notifyAdminOfError("bulkUpsertStaffRoster", err);
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function refreshMasterScheduleCache(clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    // Warm it immediately
    warmMasterScheduleCache();

    return {
      success: true };
  } catch (err) {
    notifyAdminOfError("refreshMasterScheduleCache", err);
    return {
      success: false, error: err.message };
  }
}

/**
 * Edits an existing user role.
 */

/**
 * Deletes a user role.
 */



/**
 * Fetches settings for the frontend.
 */
function getSettingsForFrontend(clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");
    return getSettings();
  } catch (err) {
    throw new Error("Failed to fetch settings: " + err.message);
  }
}

/**
 * Updates settings in the Settings sheet.
 */
function updateSettings(newSettings, clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var settingsSheet = getSheetOrThrow(ss, "Settings");

    var data = settingsSheet.getDataRange().getValues();
    var settingsMap = {};

    // Map existing rows (0-based index)
    for (var i = 1; i < data.length; i++) {
      settingsMap[String(data[i][0]).trim()] = i;
    }

    var rowsToAppend = [];
    var dataChanged = false;

    for (var key in newSettings) {
      if (settingsMap[key] !== undefined) {
        if (data[settingsMap[key]][1] !== newSettings[key]) {
          data[settingsMap[key]][1] = newSettings[key];
          dataChanged = true;
        }
      } else {
        rowsToAppend.push([key, newSettings[key]]);
      }
    }

    if (dataChanged) {
      settingsSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    if (rowsToAppend.length > 0) {
      settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, rowsToAppend.length, 2).setValues(rowsToAppend);
    }

    logAuditAction("SETTINGS_UPDATED", "Global", "Updated application settings");

    // Clear the cache
    _globalSettingsCache = null;
    var cache = CacheService.getScriptCache();
    cache.remove("app_settings");

        return {
      success: true };
  } catch (err) {
    notifyAdminOfError("updateSettings", err);
    return {
      success: false, error: err.message };
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
  ss = ss || getSS();
  var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
  if (!rosterSheet) return null;

  var data = rosterSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    // Column 2 is Role, Column 1 is Email
    if (String(data[i][2]).toLowerCase().trim() === "sub coordinator") return String(data[i][1]).trim();
  }
  return null;
}

/**
 * Helper to get a list of emails by role.
 */
function getEmailsByRole(sheetSS, roleStr) {
  var ss = sheetSS || getSS();
  var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
  if (!rosterSheet) return [];

  var data = rosterSheet.getDataRange().getValues();
  var emails = [];
  var targetRole = roleStr.toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    var roles = String(data[i][2]).toLowerCase().split(",").map(function(r) { return r.trim(); });
    if (roles.indexOf(targetRole) !== -1) {
       emails.push(String(data[i][1]).trim());
    }
  }
  return emails;
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
    var appUrl = settings["App URL"] || DEFAULT_APP_URL;

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

function submitMultipleAbsenceRequests(requestsToSubmit, clientEmail) {
  var results = [];
  var failedCount = 0;
  var lastError = "";
  for (var i = 0; i < requestsToSubmit.length; i++) {
    var req = requestsToSubmit[i];
    var res = submitAbsence(req, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    if (!res.success) {
      failedCount++;
      lastError = res.error + " (on date " + req.date + ")";
    } else {
      results.push(res);
    }
  }

  if (failedCount > 0) {
      if (results.length > 0) {
          return { success: false, error: "Partial success. " + results.length + " saved, but " + failedCount + " failed. Last error: " + lastError };
      }
      return { success: false, error: lastError };
  }
  return { success: true, count: results.length };
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
    if (formData.onBehalfTeacher) {
      // Find the teacher's email based on the provided name
      var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
      var rosterData = rosterSheet.getDataRange().getValues();

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

    var newRow = [
      uniqueId, timestamp, targetEmail, formData.date, "'" + formData.periods,
      formData.reason, formData.duration, urgencyFormatted, instructions,
      "", "", "", "", "", "", "", "", "", "", "Active"
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
    } catch(e) {}

    var hrRequired = false;
    var principalRequired = false;
    for (var i = 0; i < reasons.length; i++) {
        if (reasons[i].reason === formData.reason) {
            hrRequired = reasons[i].hrRequired === true;
            principalRequired = reasons[i].principalRequired === true;
            break;
        }
    }

    if (hrRequired) {
        var hrEmails = getEmailsByRole(ss, "hr");
        for (var h = 0; h < hrEmails.length; h++) {
            var hrEmail = hrEmails[h];
            var hrSubject = teacherName + " " + formData.reason + " Absence Request";
            var hrBody = "An absence request has been submitted requiring attention.\n\n" +
                         "Teacher: " + teacherName + "\n" +
                         "Date Needed: " + formData.date + "\n" +
                         "Periods: " + formData.periods + "\n" +
                         "Reason: " + formData.reason + "\n\n" +
                         "Instructions: " + (instructions ? instructions : "None") + "\n\n" +
                         "Please log into the Cathedral Sub App for more information: " + appUrl;

            var hrHtmlBody = "<p>An absence request has been submitted requiring attention.</p>" +
                             "<ul>" +
                             "<li><strong>Teacher:</strong> " + teacherName + "</li>" +
                             "<li><strong>Date Needed:</strong> " + formData.date + "</li>" +
                             "<li><strong>Periods:</strong> " + formData.periods + "</li>" +
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
                           "Periods: " + formData.periods + "\n" +
                           "Reason: " + formData.reason + "\n\n" +
                           "Instructions: " + (instructions ? instructions : "None") + "\n\n" +
                           "Please log into the Cathedral Sub App for more information: " + appUrl;

            var prinHtmlBody = "<p>An absence request has been submitted requiring attention.</p>" +
                               "<ul>" +
                               "<li><strong>Teacher:</strong> " + teacherName + "</li>" +
                               "<li><strong>Date Needed:</strong> " + formData.date + "</li>" +
                               "<li><strong>Periods:</strong> " + formData.periods + "</li>" +
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
          logAuditAction("SUB_DUTY_CANCELLED", absenceId, "Cancelled coverage for period " + period);

          if (coordinatorEmail && details) {
            var subject = "SUB CANCELLATION: " + userName + " cancelled coverage";
            var settings = getSettings();
            var appUrl = settings["App URL"] || DEFAULT_APP_URL;

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
    var joinKey = teacherEmail.toLowerCase() + "-" + getScheduleJoinPeriod(period);
    if (scheduleLookup) {
      var scheduleInfo = scheduleLookup[joinKey];
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

    teacherName: teacherName,
    date: formattedDate,
    period: period,
    room: roomStr,
    course: courseStr,
    instructions: instructions
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
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

    var scheduleLookup = null; // Deferred to email queue to save time
    var nameLookup = buildNameLookup(rosterData);

    var data = sheet.getDataRange().getValues();
    var targetIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex !== -1) {
      var i = targetIndex;
      var currentUserEmail = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined).toLowerCase();
      var teacherEmail = String(data[i][2]).toLowerCase();

      if (currentUserEmail !== teacherEmail) {
        var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
        assertPermission(user, "Admin Dashboard", "Unauthorized to cancel this absence.");
      }

      sheet.getRange(i + 1, 20).setValue("Canceled");
      logAuditAction("ABSENCE_CANCELLED", absenceId, "Cancelled entire absence request");

      var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', 'Advisory'];
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

      var teacherSubject = "Absence Request Canceled";
      var teacherBody = "Your absence request for " + formattedDateForEmail + " has been canceled.\n\n" +
                        "Reason: " + data[i][5] + "\n" +
                        "Periods: " + data[i][4] + "\n\n" +
                        "If you have questions, please contact the sub coordinator.";

      var teacherHtml = "<p>Your absence request for <strong>" + formattedDateForEmail + "</strong> has been canceled.</p>" +
                        "<ul><li><strong>Reason:</strong> " + data[i][5] + "</li>" +
                        "<li><strong>Periods:</strong> " + data[i][4] + "</li></ul>" +
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
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

    var scheduleLookup = null; // Deferred to email queue to save time
    var nameLookup = buildNameLookup(rosterData);

    var data = sheet.getDataRange().getValues();
    var targetIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex !== -1) {
      var i = targetIndex;
      var currentUserEmail = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined).toLowerCase();
      var teacherEmail = String(data[i][2]).toLowerCase();

      if (currentUserEmail !== teacherEmail) {
        var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
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

      // Identify sub clears
      var subClearUpdates = [];
      var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', 'Advisory'];
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
      logAuditAction("ABSENCE_UPDATED", absenceId, "Updated absence details (Date: " + formData.date + ", Periods: " + formData.periods + ")");


      var rawDate = data[i][3];
      var formattedDateForEmail = rawDate;
      if (rawDate instanceof Date) {
          formattedDateForEmail = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MMM d, yyyy");
      } else {
          try {
              formattedDateForEmail = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "MMM d, yyyy");
          } catch(e) { console.error("Error formatting date: " + e.message); }
      }

      var teacherSubject = "Absence Request Updated";
      var teacherBody = "Your absence request for " + formattedDateForEmail + " has been updated.\n\n" +
                        "Updated Details:\n" +
                        "Date: " + formData.date + "\n" +
                        "Periods: " + formData.periods + "\n" +
                        "Reason: " + formData.reason + "\n" +
                        "Duration: " + formData.duration + "\n\n" +
                        "If you have questions, please contact the sub coordinator.";

      var teacherHtml = "<p>Your absence request for <strong>" + formattedDateForEmail + "</strong> has been updated.</p>" +
                        "<ul><li><strong>Date:</strong> " + formData.date + "</li>" +
                        "<li><strong>Periods:</strong> " + formData.periods + "</li>" +
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
    if (String(data[i][0]) === String(absenceId)) {
      var details = getAbsenceDetailsLocal(data[i], period, _cachedScheduleLookup, _cachedNameLookup);
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
  var appUrl = settings["App URL"] || DEFAULT_APP_URL;

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

    var rosterData = getRosterDataCached(ss);

    var subEmailLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }

    var nameLookup = buildNameLookup(rosterData);

    // Re-read data under lock
    var data = sheet.getDataRange().getValues();

    var targetIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex !== -1) {
      var i = targetIndex;
      var periodsRequested = String(data[i][4]).split(",").map(function(p) { return p.trim(); });

      if (periodsRequested.indexOf(String(period)) === -1) {
          throw new Error("Period " + period + " was not requested for this absence.");
      }

      var subColumnIndex = getSubColumnIndex(period); // 1-based index for Apps Script Ranges: Col J is 10 (Period 1 Sub)

      var existingSub = String(data[i][subColumnIndex - 1] || "").trim(); // array is 0-indexed, so subColumnIndex - 1
      var newSub = String(subName || "").trim();

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
        for (var r = 1; r < rosterData.length; r++) {
          if (String(rosterData[r][0]).trim() === newSub) {
            newSubRole = String(rosterData[r][2]).trim(); // Role is at index 2
            break;
          }
        }
      }

      // Normalize "No Sub Needed"
      if (newSub.toLowerCase() === "no sub needed") {
          newSub = "No Sub Needed";
      }

      // Check if the new sub is available based on Dates and SubstituteAvailability or Master Schedule
      if (newSub !== "" && newSub !== "No Sub Needed" && !forceOverride) {
        var newSubEmail = (subEmailLookup[newSub] || "").toLowerCase();
        if (newSubEmail !== "") {
           var isSubstitute = newSubRole.indexOf("Substitute") !== -1;
           var targetDateRaw = data[i][3];
           var targetDateStr = (targetDateRaw instanceof Date) ? Utilities.formatDate(targetDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(targetDateRaw).trim();

           if (isSubstitute) {
             // For substitutes, check SubstituteAvailability sheet
             var allSubAvail = getAllSubstituteAvailability();
             var subAvail = allSubAvail[newSubEmail] || {};
             var availStatus = subAvail[targetDateStr] || "Not Available";

             if (availStatus !== "Available" && availStatus !== "AM Only" && availStatus !== "PM Only") {
                 throw new Error(JSON.stringify({type: "AVAILABILITY_ERROR", message: "Sub not listed as available, proceed?"}));
             }

             // Fetch day color using cached payload optimization
             var dayColor = "Green"; // default
             try {
               var payload = getInitialPayload(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
               if (payload.dateColors && payload.dateColors[targetDateStr]) {
                  dayColor = payload.dateColors[targetDateStr];
               }
             } catch(e) {
               console.error("Failed to fetch day color from payload cache: " + e.message);
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
             // For teachers/others, check master schedule
             var scheduleData = getMasterScheduleData();
             var teacherSchedule = [];
             if (scheduleData && scheduleData.length > 0) {
               var headers = scheduleData[0];
               var emailIdx = headers.indexOf("EMAIL_ADDR");
               var periodIdx = headers.indexOf("PERIOD");
               if (emailIdx > -1 && periodIdx > -1) {
                 for (var s = 1; s < scheduleData.length; s++) {
                   if (String(scheduleData[s][emailIdx]).toLowerCase().trim() === newSubEmail) {
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
                 // The teacher has a class during this period
                 throw new Error(JSON.stringify({type: "AVAILABILITY_ERROR", message: "Sub not listed as available, proceed?"}));
             }
           }
        }
      }

      // Check if the new sub is absent for a full day on the same date
      if (newSub !== "" && newSub !== "No Sub Needed") {
        var targetDateRaw = data[i][3]; // Date object or string
        var targetDateStr = (targetDateRaw instanceof Date) ? Utilities.formatDate(targetDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(targetDateRaw).trim();
        var newSubEmail = (subEmailLookup[newSub] || "").toLowerCase();

        if (newSubEmail !== "") {
          for (var j = 1; j < data.length; j++) {
            var rowEmail = String(data[j][2] || "").toLowerCase();
            if (rowEmail === "") continue;

            var rowStatus = String(data[j][19] || "Active");
            var rowDuration = String(data[j][6] || "").trim();

            if (rowEmail === newSubEmail && rowStatus !== "Canceled" && rowDuration === "Full Day") {
              var rowDateRaw = data[j][3];
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
            var scheduleLookup = null; // Deferred to email queue to save time
            var details = getAbsenceDetailsLocal(data[i], period, scheduleLookup, nameLookup);
            sendSubNotification(existingEmail, 'Canceled', details);
         }
      }

      // Write the subname
      sheet.getRange(i + 1, subColumnIndex).setValue(newSub);
      logAuditAction("SUB_ASSIGNED", absenceId, "Assigned " + (newSub || "NO ONE") + " to period " + period);

      // Notify new sub if there is one
      if (newSub && newSub !== "No Sub Needed") {
         var newEmail = subEmailLookup[newSub];
         if (newEmail) {
            var scheduleLookup = null; // Deferred to email queue to save time
            var details = getAbsenceDetailsLocal(data[i], period, scheduleLookup, nameLookup);
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

    // Check if the permission is explicitly set
    if (permissions[lowerRole] && typeof permissions[lowerRole][view] !== 'undefined') {
      return permissions[lowerRole][view] === true;
    }

    // Fallback to default if not explicitly set in the saved JSON
    var defaultSettings = {
        "admin": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "hr": { "Admin Dashboard": false, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": false, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "sub coordinator": { "Admin Dashboard": true, "HR Dashboard": false, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "principal": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": false, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "teacher": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": true, "My Past Absences": true, "Settings": false, "Add Request on Behalf": false },
        "substitute": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Open Jobs": false, "My Past Absences": false, "Settings": false, "Add Request on Behalf": false }
    };

    if (defaultSettings[lowerRole] && typeof defaultSettings[lowerRole][view] !== 'undefined') {
        return defaultSettings[lowerRole][view] === true;
    }

    return false;
  } catch (e) {
    return false;
  }
}

function getInitialPayload(clientEmail) {
  try {
    var ss = getSS();
    var email = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    var targetEmail = String(email).toLowerCase();

    // 1. Fetch all required sheets
    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];


    var settings = getSettings(ss); // Already passes ss to avoid fetching again

    var mainSheet = getSheetOrThrow(ss, "Absence Requests");
    var absenceData = mainSheet ? mainSheet.getDataRange().getValues() : [];

    var scheduleData = getMasterScheduleData();

    var payPeriodsSheet = ss.getSheetByName("PayPeriods");
    var payPeriodsData = payPeriodsSheet ? payPeriodsSheet.getDataRange().getValues() : [];

    var datesSheet = ss.getSheetByName("Dates");
    var datesData = datesSheet ? datesSheet.getDataRange().getValues() : [];

    var dateColors = {};
    if (datesData && datesData.length > 0) {
      for (var d = 1; d < datesData.length; d++) {
        var dateRaw = datesData[d][0];
        var colorRaw = datesData[d][1];
        if (dateRaw && colorRaw) {
          var dateFormatted = dateRaw instanceof Date ? Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd") :
            (function(){ try { return Utilities.formatDate(new Date(dateRaw), Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e) { return String(dateRaw); } })();
          dateColors[dateFormatted] = String(colorRaw).trim();
        }
      }
    }



    // --- Build lookups ---
    var scheduleLookup = buildScheduleLookup(scheduleData);
    var nameLookup = buildNameLookup(rosterData);


    // --- 2. Extract User Data ---
    var name = null;
    var role = null;
    for (var i = 1; i < rosterData.length; i++) {
      if (String(rosterData[i][1]).toLowerCase() === targetEmail) {
        name = rosterData[i][0];
        role = rosterData[i][2] ? String(rosterData[i][2]).trim() : "Teacher";
        break;
      }
    }

    if (!name) {
      throw new Error("Access Denied: Please Contact Technology@gocathedral.com");
    }

    var userName = String(name).trim().toLowerCase();
    var lowerRole = String(role).toLowerCase();

    // Compute teacherSchedule
    var teacherSchedule = [];
    if (scheduleData && scheduleData.length > 0) {
      var headers = scheduleData[0];
      var emailIdx = headers.indexOf("EMAIL_ADDR");
      var periodIdx = headers.indexOf("PERIOD");
      if (emailIdx > -1 && periodIdx > -1) {
        for (var s = 1; s < scheduleData.length; s++) {
          if (String(scheduleData[s][emailIdx]).toLowerCase().trim() === targetEmail) {
            var pVal = String(scheduleData[s][periodIdx]).trim();
            var joinP = getScheduleJoinPeriod(pVal);
            if (joinP && teacherSchedule.indexOf(joinP) === -1) {
              teacherSchedule.push(joinP);
            }
          }
        }
      }
    }
    teacherSchedule.sort(function(a, b) {
      var numA = parseInt(a, 10);
      var numB = parseInt(b, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      if (a < b) return -1;
      if (a > b) return 1;
      return 0;
    });

    var appUrl = settings["App URL"] || DEFAULT_APP_URL;
    var urgencyCutoffTime = settings["Urgency Cutoff Time"] || "15";
    var defaultAbsenceReasons = JSON.stringify([
        {reason: "Personal", hrRequired: false, principalRequired: false},
        {reason: "Professional Development", hrRequired: false, principalRequired: false},
        {reason: "Retreat", hrRequired: false, principalRequired: false},
        {reason: "Athletics", hrRequired: false, principalRequired: true},
        {reason: "Jury Duty", hrRequired: true, principalRequired: false},
        {reason: "Bereavement", hrRequired: true, principalRequired: false}
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
      "Settings": hasPermission(lowerRole, "Settings", rolePermissions) || lowerRole === "admin",
      "Add Request on Behalf": hasPermission(lowerRole, "Add Request on Behalf", rolePermissions)
    };

    var userData = {
      appVersion: APP_VERSION,
      name: String(name),
      role: String(role),
      email: String(email),
      appUrl: String(appUrl),
      urgencyCutoffTime: String(urgencyCutoffTime),
      absenceReasons: String(absenceReasons),
      teacherSchedule: teacherSchedule
    };


    // --- 3. Extract common data (My Absences, My Sub Duties, Open Jobs) ---
    var myUpcomingAbsences = [];
    var myPastAbsences = [];
    var mySubDuties = [];
    var myPastSubDuties = [];
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
      var status = String(row[19] || 'Active');
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
        var subFeedbackRaw = String(row[20] || "[]");
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
          instructions: String(row[8]),
          subFeedback: subFeedbackRaw
        };
        if (isDateValid && rowDate < today) {
          myPastAbsences.push(absenceObj);
        } else {
          myUpcomingAbsences.push(absenceObj);
        }
      }

      // My Sub Duties & Open Jobs
      if (isDateValid && rowDate <= targetEndWeek) {
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

        var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', 'Advisory'];
      for (var pIdx = 0; pIdx < allPeriods.length; pIdx++) {
        var p = allPeriods[pIdx];
          if (periodsRequested.indexOf(String(p)) !== -1) {
            var subColumnIndex = getSubColumnIndex(p) - 1;
            var assignedSub = String(row[subColumnIndex] || "").trim().toLowerCase();

            var joinKey = rowTeacherEmail + "-" + getScheduleJoinPeriod(p);
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
              if (rowDate < today && rowDate >= cutoffDate) {
                 var subFeedbackRaw = String(row[20] || "[]");
                 var subFeedbackParsed = [];
                 try {
                     subFeedbackParsed = JSON.parse(subFeedbackRaw);
                 } catch(e) {}
                 var matchingFeedback = subFeedbackParsed.filter(function(fb) { return fb.period === String(p) && String(fb.subName).toLowerCase() === userName; });
                 jobObj.rating = matchingFeedback.length > 0 ? matchingFeedback[0].rating : 0;
                 jobObj.note = matchingFeedback.length > 0 ? matchingFeedback[0].note : "";
                 myPastSubDuties.push(jobObj);
              } else if (rowDate >= today) {
                 mySubDuties.push(jobObj);
              }
            }

            // Open Jobs (Today only)
            if (assignedSub === "" && rowDate <= targetEndToday && rowDate >= today) {
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
    var sortJobsDesc = function(a, b) {
      if (a.rawDate === b.rawDate) {
        return parseInt(a.period) - parseInt(b.period);
      }
      return b.rawDate - a.rawDate;
    };
    mySubDuties.sort(sortJobs);
    myPastSubDuties.sort(sortJobsDesc);
    todaysOpenJobs.sort(sortJobs);


    var payload = {
      userData: userData,
      myAbsences: myUpcomingAbsences,
      myPastAbsences: myPastAbsences,
      mySubDuties: mySubDuties,
      myPastSubDuties: myPastSubDuties,
      todaysOpenJobs: todaysOpenJobs,
      permissions: permissions,
      dateColors: dateColors
    };


    // --- 4. Extract Admin / Sub Coordinator data if applicable ---
    if (permissions["Admin Dashboard"] || permissions["HR Dashboard"] || permissions["Today at a Glance"]) {
      // Staff List
      var staffList = [];
      var allSubAvail = getAllSubstituteAvailability();

      // Pre-calculate all staff schedules
      var allSchedules = {};
      if (scheduleData && scheduleData.length > 0) {
        var headers = scheduleData[0];
        var emailIdx = headers.indexOf("EMAIL_ADDR");
        var periodIdx = headers.indexOf("PERIOD");
        if (emailIdx > -1 && periodIdx > -1) {
          for (var s = 1; s < scheduleData.length; s++) {
            var sEmail = String(scheduleData[s][emailIdx]).toLowerCase().trim();
            var pVal = String(scheduleData[s][periodIdx]).trim();
            var joinP = getScheduleJoinPeriod(pVal);
            if (joinP) {
              if (!allSchedules[sEmail]) allSchedules[sEmail] = [];
              if (allSchedules[sEmail].indexOf(joinP) === -1) {
                allSchedules[sEmail].push(joinP);
              }
            }
          }
          for (var e in allSchedules) {
            allSchedules[e].sort(function(a, b) {
              var numA = parseInt(a, 10);
              var numB = parseInt(b, 10);
              if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
              if (a < b) return -1;
              if (a > b) return 1;
              return 0;
            });
          }
        }
      }



      for (var i = 1; i < rosterData.length; i++) {
        var staffName = String(rosterData[i][0]).trim();
        var staffEmail = String(rosterData[i][1]).toLowerCase().trim();
        var staffRole = String(rosterData[i][2]).toLowerCase().trim();
        var duty = String(rosterData[i][3] || "").trim();
        if (staffName) {
          var display = staffName;
          if (duty) display = staffName + " - " + duty;

          staffList.push({
             name: staffName,
             display: display,
             duty: duty,
             role: staffRole,
             email: staffEmail,
             availability: allSubAvail[staffEmail] || {},
             teacherSchedule: allSchedules[staffEmail] || []
          });
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
        var status = String(row[19] || 'Active');
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

          var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', 'Advisory'];
      for (var pIdx = 0; pIdx < allPeriods.length; pIdx++) {
        var p = allPeriods[pIdx];
            if (periodsRequested.indexOf(String(p)) !== -1) {
              var assignedSub = row[getSubColumnIndex(p) - 1];
              if (!assignedSub || String(assignedSub).trim() === "") {
                var joinKey = rowTeacherEmail + "-" + getScheduleJoinPeriod(p);
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
        if (String(row[19] || "").trim() === "Canceled") continue;

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
            var scheduleKey = rowTeacherEmail + "-" + getScheduleJoinPeriod(p);
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
              assignedSub: String(row[getSubColumnIndex(p) - 1] || "").trim(),
              reason: String(row[5] || "").trim(),
              duration: String(row[6] || "").trim(),
              instructions: String(row[8] || "").trim(),
              subPlanUrl: String(row[20] || "").trim()
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
    if (lowerRole === 'substitute') {
      payload.subAvailability = getSubstituteAvailability(email);
    }

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
          var approved = String(payPeriodsData[p][3] || "").toLowerCase().trim() === "true";

          payPeriods.push({
            periodNumber: periodNum,
            startDate: startFormatted,
            endDate: endFormatted,
            approved: approved
          });
        }
      }

      for (var i = 1; i < absenceData.length; i++) {
        var row = absenceData[i];
        if (String(row[19] || "").trim() === "Canceled") continue;

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
          var pStr = periods[j];
          var p = pStr; // Keep as string for getSubColumnIndex
          if (p) {
            var assignedSub = row[getSubColumnIndex(p) - 1];
            if (assignedSub && String(assignedSub).trim() !== "") {
              var scheduleKey = rowTeacherEmail + "-" + getScheduleJoinPeriod(p);
              var courseStr = scheduleLookup[scheduleKey] ? String(scheduleLookup[scheduleKey].course) : "No Class Assigned";
              assignedSubs.push({ name: String(assignedSub).trim(), period: String(p), course: courseStr });
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
        payPeriods: payPeriods,
        dateColors: dateColors,
        rates: {
          green: settings["Green Day Pay Rate"] || "10",
          blueGold: settings["Blue/Gold Day Pay Rate"] || "20"
        }
      };
    }

    return payload;

  } catch (err) {
    notifyAdminOfError("getInitialPayload", err);
    throw new Error("Failed to get initial payload: " + err.message);
  }
}



/**
 * Clears the Master Schedule cache from the Script Cache.
 * Returns a success object.
 */
function clearMasterScheduleCache(clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var sheet = ss.getSheetByName("Master Schedule Cache");
    if (sheet) {
      sheet.clearContents();
    }

    // Attempt to warm it immediately
    warmMasterScheduleCache();

    return {
      success: true };
  } catch (err) {
    notifyAdminOfError("clearMasterScheduleCache", err);
    return {
      success: false, error: err.message };
  }
}


/**
 * Refreshes requested data components.
 * @param {Array<string>} components - The components to fetch (e.g. ['myAbsences', 'quickCover'])
 * @returns {Object} The requested data.
 */
function refreshData(components, clientEmail) {
  try {
    var payload = getInitialPayload(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
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
    var actor = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined) || "Unknown";

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
function exportAllAbsenceRequests(clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
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
function archiveAbsenceRequests(cutoffDateStr, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("archiveAbsenceRequests_lock", e);
    return {
      success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
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
      return {
      success: true, count: 0 };
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

    return {
      success: true, count: rowsToArchive.length };
  } catch (err) {
    notifyAdminOfError("archiveAbsenceRequests", err);
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function getAuditLogs(startDateStr, endDateStr, clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
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


function approvePayPeriod(periodNumber, rangeString, csvString, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("approvePayPeriod_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "HR Dashboard");

    var payPeriodsSheet = getSheetOrThrow(ss, "PayPeriods");
    var payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < payPeriodsData.length; i++) {
      if (String(payPeriodsData[i][0]).trim() === String(periodNumber).trim()) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: "Pay period not found." };
    }

    // Set approved to true in the 4th column
    payPeriodsSheet.getRange(rowIndex, 4).setValue(true);

    // Get emails
    var hrEmails = getEmailsByRole(ss, "hr");
    var subCoordinatorEmails = getEmailsByRole(ss, "sub coordinator");

    if (hrEmails.length === 0 && subCoordinatorEmails.length === 0) {
       return { success: true, message: "Pay period approved, but no HR or Sub Coordinator emails found to notify." };
    }

    var subject = "Pay Period " + periodNumber + " with the date range approved";
    var body = user.name + " has approved the payroll for the subs for Pay Period " + periodNumber + " (" + rangeString + "). Please reach out with any questions.";

    var blob = Utilities.newBlob(csvString, MimeType.CSV, "Pay_Period_" + periodNumber + ".csv");

    var options = {
       attachments: [blob]
    };

    if (subCoordinatorEmails.length > 0) {
       options.cc = subCoordinatorEmails.join(",");
    }

    // Ensure there is at least one "To" recipient. If HR is empty, use sub coordinator as fallback if possible.
    var toEmail = hrEmails.length > 0 ? hrEmails.join(",") : subCoordinatorEmails.join(",");

    sendEmailHelper(toEmail, subject, body, options);

    logAuditAction("PAY_PERIOD_APPROVED", "Period " + periodNumber, user.name + " approved pay period " + periodNumber);

    return { success: true };
  } catch (err) {
    notifyAdminOfError("approvePayPeriod", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function bulkUpsertPayPeriods(updates, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("bulkUpsertPayPeriods_lock", e);
    return {
      success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "HR Dashboard");

    var sheet = getSheetOrThrow(ss, "PayPeriods");

    // updates is an array of objects: {period: "", start: "", end: ""}
    var newRows = [];
    for (var i = 0; i < updates.length; i++) {
       var u = updates[i];
       newRows.push([u.period, u.start, u.end, "FALSE"]); // Default approved to FALSE
    }

    if (newRows.length > 0) {
       var startRow = sheet.getLastRow() + 1;
       sheet.getRange(startRow, 1, newRows.length, 4).setValues(newRows);
    }

    logAuditAction("PAY_PERIODS_BULK_UPLOAD", "Multiple", "Added " + newRows.length + " pay periods");

    // Clear cache because PayPeriods data changed
    // In hrData, pay periods are retrieved directly from the sheet, so it's live
    // but we can clear cache just in case.

    return {
      success: true, updated: newRows.length };
  } catch (err) {
    notifyAdminOfError("bulkUpsertPayPeriods", err);
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteAllPayPeriods(clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("deleteAllPayPeriods_lock", e);
    return {
      success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "HR Dashboard");

    var sheet = getSheetOrThrow(ss, "PayPeriods");
    var lastRow = sheet.getLastRow();

    if (lastRow > 1) {
       sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
    }

    logAuditAction("PAY_PERIODS_DELETE_ALL", "All", "Deleted all pay periods");

    return {
      success: true };
  } catch (err) {
    notifyAdminOfError("deleteAllPayPeriods", err);
    return {
      success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function loadPayPeriodsSettings(clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
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
            var approved = String(row[3] || "").toLowerCase().trim() === "true";

            payPeriods.push({
                period: periodNum,
                start: startFormatted,
                end: endFormatted,
                approved: approved
            });
        }
    }
    return payPeriods;
  } catch(err) {
    throw new Error("Failed to load pay periods: " + err.message);
  }
}

/**
 * Returns the current backend app version for frontend polling.
 * @returns {string} The APP_VERSION.
 */
function getAppVersion(clientEmail) {
  return APP_VERSION;
}

function getSubstituteAvailability(email) {
  var ss = getSS();
  var subAvailSheet = ss.getSheetByName("SubstituteAvailability");
  if (!subAvailSheet) return {};

  var data = subAvailSheet.getDataRange().getValues();
  var availability = {};
  var targetEmail = String(email).toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === targetEmail) {
      var cellVal = data[i][1];
      var dateStr = (cellVal instanceof Date) ? _formatDateToYYYYMMDD(cellVal) : String(cellVal).trim();
      availability[dateStr] = data[i][2];
    }
  }
  return availability;
}

function saveSubstituteAvailability(dateStr, status, clientEmail) {
  var email = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
  var targetEmail = String(email).toLowerCase();
  var ss = getSS();
  var subAvailSheet = ss.getSheetByName("SubstituteAvailability");

  if (!subAvailSheet) {
    throw new Error("SubstituteAvailability sheet not found.");
  }

  var data = subAvailSheet.getDataRange().getValues();
  var matchingRows = [];

  for (var i = 1; i < data.length; i++) {
    var cellVal = data[i][1];
    var rowDateStr = (cellVal instanceof Date) ? _formatDateToYYYYMMDD(cellVal) : String(cellVal).trim();
    if (String(data[i][0]).toLowerCase() === targetEmail && rowDateStr === dateStr) {
      matchingRows.push(i + 1); // 1-based index for sheets
    }
  }

  if (matchingRows.length > 0) {
    // Update the first matching row with the new status (even if 'Not Available')
    subAvailSheet.getRange(matchingRows[0], 3).setValue(status);

    // Delete any subsequent duplicate rows (iterate backwards to avoid shifting issues)
    for (var j = matchingRows.length - 1; j > 0; j--) {
      subAvailSheet.deleteRow(matchingRows[j]);
    }
  } else {
    subAvailSheet.appendRow([targetEmail, dateStr, status]);
  }
}

function getAllSubstituteAvailability() {
  var ss = getSS();
  var subAvailSheet = ss.getSheetByName("SubstituteAvailability");
  if (!subAvailSheet) return {};

  var data = subAvailSheet.getDataRange().getValues();
  var availabilityMap = {}; // Format: { "email": { "YYYY-MM-DD": "status" } }

  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][0]).toLowerCase().trim();
    var cellVal = data[i][1];
    var dateStr = (cellVal instanceof Date) ? _formatDateToYYYYMMDD(cellVal) : String(cellVal).trim();
    var status = String(data[i][2]).trim();

    if (!availabilityMap[email]) {
      availabilityMap[email] = {};
    }
    availabilityMap[email][dateStr] = status;
  }

  return availabilityMap;
}

function _formatDateToYYYYMMDD(dateObj) {
  var yyyy = dateObj.getFullYear();
  var mm = String(dateObj.getMonth() + 1);
  if (mm.length < 2) mm = '0' + mm;
  var dd = String(dateObj.getDate());
  if (dd.length < 2) dd = '0' + dd;
  return yyyy + '-' + mm + '-' + dd;
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
  var hrRates = {
    green: parseFloat(settings["Green Day Pay Rate"] || "10"),
    blueGold: parseFloat(settings["Blue/Gold Day Pay Rate"] || "20")
  };

  // Get data
  var dateColors = {};
  var datesSheet = ss.getSheetByName("Dates");
  if (datesSheet) {
    var datesRaw = datesSheet.getDataRange().getValues();
    for (var i = 1; i < datesRaw.length; i++) {
        var dStr = datesRaw[i][0];
        if (dStr) {
           var formatted = dStr instanceof Date ? Utilities.formatDate(dStr, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dStr).trim();
           dateColors[formatted] = String(datesRaw[i][1]).trim();
        }
    }
  }

  var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
  var rosterData = rosterSheet.getDataRange().getValues();
  var nameLookup = {};
  var dutyLookup = {};
  for (var i = 1; i < rosterData.length; i++) {
    var e = String(rosterData[i][1]).toLowerCase().trim();
    var name = String(rosterData[i][0]).trim();
    if (e) nameLookup[e] = name;
    dutyLookup[name] = String(rosterData[i][3] || "").trim();
  }

  var absenceSheet = getSheetOrThrow(ss, "Absence Requests");
  var absenceData = absenceSheet.getDataRange().getValues();

  var currentWeekAbsences = {};
  var currentWeekCoverage = {};
  var nextWeekAbsences = {};

  for (var i = 1; i < absenceData.length; i++) {
    var row = absenceData[i];
    var status = String(row[19] || "").trim();
    if (status.toLowerCase() !== "active") continue; // Only active requests

    var dateVal = row[3];
    var dateObjRow;
    if (dateVal instanceof Date) {
      dateObjRow = new Date(dateVal);
    } else {
      var dStr = String(dateVal).trim();
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
                    subEmailLookup = String(rosterData[r][1]).toLowerCase().trim();
                    break;
                }
            }
            
            var isSubstituteRole = false;
            var teacherDuty = dutyLookup[assignedSub];
            var teacherDutyArr = teacherDuty ? String(teacherDuty).split(',').map(function(d){return d.trim();}) : [];
            var isDuty = teacherDutyArr.includes(String(p));

            if (subEmailLookup) {
                for (var r = 1; r < rosterData.length; r++) {
                    if (String(rosterData[r][1]).toLowerCase().trim() === subEmailLookup) {
                         var roles = String(rosterData[r][2]).toLowerCase();
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
                   if (dayColor.toLowerCase() === "blue" || dayColor.toLowerCase() === "gold") {
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
  for (var name in currentWeekAbsences) {
    if (currentWeekAbsences[name].days > 0 || currentWeekAbsences[name].periods > 0) {
      sec1 = true;
      var text = [];
      if (currentWeekAbsences[name].days > 0) text.push(currentWeekAbsences[name].days + " days");
      if (currentWeekAbsences[name].periods > 0) text.push(currentWeekAbsences[name].periods + " periods");
      html += "<li><strong>" + name + "</strong>: " + text.join(" and ") + "</li>";
    }
  }
  if (!sec1) html += "<li>No absences this week.</li>";
  html += "</ul>";

  // SECTION 2
  html += "<h3 style='color: #002147; border-bottom: 1px solid #ccc; padding-bottom: 5px;'>Coverage This Week</h3>";
  var sec2 = false;
  html += "<ul>";
  for (var name in currentWeekCoverage) {
    if (currentWeekCoverage[name].periods > 0) {
      sec2 = true;
      var pText = currentWeekCoverage[name].periods + (currentWeekCoverage[name].periods === 1 ? " class" : " classes");
      var payText = currentWeekCoverage[name].pay > 0 ? " ($" + currentWeekCoverage[name].pay.toFixed(2) + " extra pay)" : "";
      html += "<li><strong>" + name + "</strong>: covered " + pText + payText + "</li>";
    }
  }
  if (!sec2) html += "<li>No classes covered this week.</li>";
  html += "</ul>";

  // SECTION 3
  html += "<h3 style='color: #002147; border-bottom: 1px solid #ccc; padding-bottom: 5px;'>Upcoming Absences Next Week (" + Utilities.formatDate(nextWeekMonday, Session.getScriptTimeZone(), "MMM d") + " - " + Utilities.formatDate(nextWeekFriday, Session.getScriptTimeZone(), "MMM d") + ")</h3>";
  var sec3 = false;
  html += "<ul>";
  for (var name in nextWeekAbsences) {
    if (nextWeekAbsences[name].days > 0 || nextWeekAbsences[name].periods > 0) {
      sec3 = true;
      var text = [];
      if (nextWeekAbsences[name].days > 0) text.push(nextWeekAbsences[name].days + " days");
      if (nextWeekAbsences[name].periods > 0) text.push(nextWeekAbsences[name].periods + " periods");
      html += "<li><strong>" + name + "</strong>: " + text.join(" and ") + "</li>";
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
  var dateObj = dateStr ? new Date(dateStr) : new Date();
  return generatePrincipalsDigestHTML(dateObj);
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
       if (String(rosterData[r][1]).toLowerCase().trim() === targetEmail) {
           subName = String(rosterData[r][0]).trim();
           break;
       }
    }

    var data = sheet.getDataRange().getValues();
    var targetIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(absenceId)) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex === -1) {
       throw new Error("Could not find absence request with ID " + absenceId);
    }

    var row = data[targetIndex];
    var subFeedbackRaw = String(row[20] || "[]");
    var subFeedbackParsed = [];
    try {
      subFeedbackParsed = JSON.parse(subFeedbackRaw);
    } catch(e) {}

    var existingIndex = -1;
    for (var j = 0; j < subFeedbackParsed.length; j++) {
       if (subFeedbackParsed[j].period === String(period) && String(subFeedbackParsed[j].subName).toLowerCase() === subName.toLowerCase()) {
           existingIndex = j;
           break;
       }
    }

    var feedbackObj = {
        subName: subName,
        period: String(period),
        rating: parseInt(rating) || 0,
        note: String(note || "").trim()
    };

    if (existingIndex !== -1) {
       subFeedbackParsed[existingIndex] = feedbackObj;
    } else {
       subFeedbackParsed.push(feedbackObj);
    }

    sheet.getRange(targetIndex + 1, 21).setValue(JSON.stringify(subFeedbackParsed));

    var teacherEmail = String(row[2]).toLowerCase().trim();
    var absenceDateVal = row[3];
    var absenceDateStr = (absenceDateVal instanceof Date) ? Utilities.formatDate(absenceDateVal, Session.getScriptTimeZone(), "M/d/yyyy") : String(absenceDateVal);

    var nameLookup = buildNameLookup(rosterData);
    var teacherName = nameLookup[teacherEmail] || teacherEmail;
    if (teacherName.includes(",")) {
        var parts = teacherName.split(",");
        teacherName = parts[1].trim() + " " + parts[0].trim();
    }

    var subject = "Substitute Feedback Received - Period " + period;
    var parsedRating = parseInt(rating) || 0;
    var ratingStars = "★".repeat(parsedRating) + "☆".repeat(Math.max(0, 5 - parsedRating));

    var bodyText = "Hello " + teacherName + ",\n\n";
    bodyText += subName + " just submitted feedback for covering your class on " + absenceDateStr + ":\n\n";
    bodyText += "- Period: " + period + "\n";
    bodyText += "- Rating: " + rating + "/5\n";
    bodyText += "- Notes: " + (note || "No notes") + "\n\n";
    bodyText += "Thank you!";

    var bodyHtml = "<p>Hello " + teacherName + ",</p>";
    bodyHtml += "<p>" + subName + " just submitted feedback for covering your class on " + absenceDateStr + ":</p>";
    bodyHtml += "<ul>";
    bodyHtml += "<li><strong>Period:</strong> " + period + "</li>";
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
      subEmailLookup[String(rosterData[r][0]).trim()] = String(rosterData[r][1]).trim();
    }
    var nameLookup = buildNameLookup(rosterData);

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var requestsToSend = {}; // Keyed by substitute email

    for (var i = 1; i < data.length; i++) {
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
          teacherName = parts[1].trim() + " " + parts[0].trim();
        }

        var periodsRequested = String(row[4]).split(",").map(function(p) { return p.trim(); });
        var rowId = String(row[0]);
        var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', 'Advisory'];

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
            var jobLink = appUrl + "?feedbackAbsenceId=" + encodeURIComponent(job.id) + "&feedbackPeriod=" + encodeURIComponent(job.period);
            bodyText += "- Period " + job.period + " (" + job.teacherName + "): " + jobLink + "\n";
            bodyHtml += "<li>Period " + job.period + " (" + job.teacherName + "): <a href='" + jobLink + "'>Leave Feedback</a></li>";
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

