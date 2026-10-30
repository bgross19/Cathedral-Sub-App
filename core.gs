// ============================================================================
// CORE.GS — Web App Entrypoint, Auth & Permissions, Period Helpers & Audit Log
// ============================================================================

function getActiveUserEmail(clientEmail) {
  try {
    var sessionEmail = Session.getActiveUser().getEmail();
    if (sessionEmail && typeof sessionEmail === 'string' && sessionEmail.indexOf('@') > 0) {
      return sessionEmail.trim().toLowerCase();
    }
  } catch(e) {
    // Session unavailable or restricted; fall back to clientEmail below
  }
  if (clientEmail && typeof clientEmail === 'string' && clientEmail.indexOf('@') > 0) {
    return clientEmail.trim().toLowerCase();
  }
  return "";
}


/**
 * Helper to normalize a period string.
 */
function parsePeriod(period) {
  var p = String(period).toLowerCase().trim();
  if (p === 'a' || p === 'advisory') return '10';
  return p;
}

/**
 * Helper to format a single period value ('10' -> 'Advisory') for user-facing display/emails.
 */
function formatPeriodDisplay(period) {
  var p = parsePeriod(period);
  if (p === '10') return 'Advisory';
  return String(period).trim();
}

/**
 * Helper to format a single period with "Period " prefix (or "Advisory" for 10) for email subjects/lists.
 */
function formatPeriodLabelDisplay(period) {
  var p = parsePeriod(period);
  if (p === '10') return 'Advisory';
  return 'Period ' + String(period).trim();
}

/**
 * Helper to format a comma-separated periods string (e.g. "1, 10, 2" -> "1, Advisory, 2") for user-facing display/emails.
 */
function formatPeriodsListDisplay(periodsStr) {
  if (!periodsStr) return "";
  return String(periodsStr)
    .split(",")
    .map(function(p) { return p.trim(); })
    .filter(function(p) { return p !== ""; })
    .map(function(p) { return formatPeriodDisplay(p); })
    .join(", ");
}

/**
 * Helper to get the 1-based column index for a given period.
 * 0-based array index is one less.
 */
function getSubColumnIndex(period) {
  var p = parsePeriod(period);
  if (p === '0') return 18; // Col R
  if (p === '10') return 19; // Col S
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
  var p = parsePeriod(period);
  if (p === '0') return '9';
  if (p === '10') return '10';
  return String(period);
}

/**
 * Helper to extract Sub Plan URL from an Absence Requests row (Col 21 / index 20),
 * ignoring legacy Sub Feedback JSON arrays if they were previously written into Col 21.
 */
function extractSubPlanUrl(row) {
  if (!row || row.length <= 20) return "";
  var val = String(row[20] || "").trim();
  if (val.indexOf("[") === 0) return "";
  return val;
}

/**
 * Helper to extract Sub Feedback JSON string from an Absence Requests row (Col 22 / index 21),
 * falling back to Col 21 (index 20) if legacy feedback JSON was stored there.
 */
function extractSubFeedbackRaw(row) {
  if (!row) return "[]";
  if (row.length > 21 && row[21] !== undefined && row[21] !== null && String(row[21]).trim() !== "") {
    return String(row[21]).trim();
  }
  if (row.length > 20) {
    var legacyVal = String(row[20] || "").trim();
    if (legacyVal.indexOf("[") === 0) return legacyVal;
  }
  return "[]";
}

const APP_VERSION = "1.0.0";
const DEFAULT_APP_URL = "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec";
const MASKED_SECRET_VALUE = "••••••••";

/**
 * Returns a copy of settings safe to send to client payloads.
 * - When includeAdminFields is false (general initial payload), strips PowerSchool API credentials.
 * - When includeAdminFields is true (Settings page), masks PS_CLIENT_SECRET if present.
 */
function sanitizeSettingsForClient(settings, includeAdminFields) {
  var safe = {};
  if (!settings) return safe;
  var sensitiveKeys = { "PS_CLIENT_SECRET": true, "PS_CLIENT_ID": true, "PS_URL": true };
  for (var k in settings) {
    if (!Object.prototype.hasOwnProperty.call(settings, k)) continue;
    if (!includeAdminFields && sensitiveKeys[k]) continue;
    if (includeAdminFields && k === "PS_CLIENT_SECRET") {
      safe[k] = settings[k] ? MASKED_SECRET_VALUE : "";
    } else {
      safe[k] = settings[k];
    }
  }
  return safe;
}


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
      if (!rosterData[i] || rosterData[i].length < 3) continue;
      if (rosterData[i] && String(rosterData[i][1]).toLowerCase() === targetEmail) {
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
 * Includes an HTML partial file into a template.
 * @param {string} filename - Name of the HTML file to include.
 * @returns {string} Raw HTML/JS content of the partial.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
  
  var rosterData = getRosterDataCached(ss);
  var name = null;
  var role = null;
  var targetEmail = String(email).toLowerCase();
  
  for (var i = 1; i < rosterData.length; i++) {
      if (!rosterData[i] || rosterData[i].length < 3) continue;
      if (rosterData[i] && String(rosterData[i][1]).toLowerCase() === targetEmail) {
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
        "admin": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Absences": true, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "hr": { "Admin Dashboard": false, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Absences": false, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "sub coordinator": { "Admin Dashboard": true, "HR Dashboard": false, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Absences": true, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "principal": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Absences": false, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "teacher": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Absences": true, "My Past Absences": true, "Settings": false, "Add Request on Behalf": false },
        "substitute": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Absences": false, "My Past Absences": false, "Settings": false, "Add Request on Behalf": false }
    };

    if (defaultSettings[lowerRole] && typeof defaultSettings[lowerRole][view] !== 'undefined') {
        return defaultSettings[lowerRole][view] === true;
    }

    return false;
  } catch (e) {
    return false;
  }
}


function getPeriodSortValue(p) {
  if (String(p).toLowerCase() === '10') return 1.5;
  var num = parseInt(p, 10);
  return isNaN(num) ? 99 : num;
}

/**
 * Logs an action to the hidden Audit Log sheet.
 * @param {string} actionType - The type of action (e.g., "ASSIGN_SUB", "CANCEL_ABSENCE").
 * @param {string} targetId - The ID of the affected record.
 * @param {string} details - A description of the action.
 * @param {string} [clientEmail] - Optional fallback email from caller.
 */
function logAuditAction(actionType, targetId, details, clientEmail) {
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
    if (!data[i]) continue;
      var rowDate = new Date(data[i][0]);
      if (rowDate >= startDate && rowDate <= endDate) {
        logs.push({
          timestamp: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
          actor: (data[i] ? String(data[i][1]) : ""),
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

/**
 * Returns the current backend app version for frontend polling.
 * @returns {string} The APP_VERSION.
 */
function getAppVersion(clientEmail) {
  return APP_VERSION;
}

/**
 * Default period start/end clock times (24-hour HH:MM) by day type / bell schedule.
 * Synced automatically from PowerSchool `com.cathedral.subapp.bellscheduletimes` when available,
 * or customizable via Settings["Period Bell Times"].
 */
var DEFAULT_PERIOD_BELL_TIMES = {
  "Green": {
    "0":  { start: "07:00", end: "07:45" },
    "9":  { start: "07:00", end: "07:45" },
    "1":  { start: "07:50", end: "08:35" },
    "10": { start: "08:40", end: "09:05" },
    "2":  { start: "09:10", end: "09:55" },
    "3":  { start: "10:00", end: "10:45" },
    "4":  { start: "10:50", end: "11:35" },
    "5":  { start: "11:40", end: "12:55" },
    "6":  { start: "13:00", end: "13:45" },
    "7":  { start: "13:50", end: "14:35" },
    "8":  { start: "14:40", end: "15:25" }
  },
  "Blue": {
    "0":  { start: "07:00", end: "07:45" },
    "9":  { start: "07:00", end: "07:45" },
    "1":  { start: "07:50", end: "09:15" },
    "10": { start: "09:20", end: "09:50" },
    "2":  { start: "09:55", end: "11:20" },
    "3":  { start: "11:25", end: "13:25" },
    "4":  { start: "13:30", end: "14:55" }
  },
  "Gold": {
    "0":  { start: "07:00", end: "07:45" },
    "9":  { start: "07:00", end: "07:45" },
    "5":  { start: "07:50", end: "09:15" },
    "10": { start: "09:20", end: "09:50" },
    "6":  { start: "09:55", end: "11:20" },
    "7":  { start: "11:25", end: "13:25" },
    "8":  { start: "13:30", end: "14:55" }
  }
};

/**
 * Resolves the PowerSchool Semester Term ID (e.g. "3601" for Semester 1, "3603" for Semester 2)
 * for a given date string.
 */
function resolveTermIdForDate(dateStr, optionalSettings) {
  var settings = optionalSettings || (typeof getSettings === "function" ? getSettings() : {});
  var configuredTerm = String(settings["Term ID"] || "3601").trim();
  var termNum = parseInt(configuredTerm, 10);
  var yearPrefix = !isNaN(termNum) && termNum >= 100 ? Math.floor(termNum / 100) : 36;
  var sem1Term = String(yearPrefix * 100 + 1);
  var sem2Term = String(yearPrefix * 100 + 3);

  if (!dateStr) return configuredTerm;
  var isoDate = typeof getStandardizedDateString === "function" ? getStandardizedDateString(dateStr) : String(dateStr).trim();
  if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) {
    return configuredTerm;
  }

  var sem2Start = String(settings["Semester 2 Start Date"] || "").trim();
  if (sem2Start && /^\d{4}-\d{2}-\d{2}$/.test(sem2Start)) {
    return isoDate >= sem2Start ? sem2Term : sem1Term;
  }

  var month = parseInt(isoDate.split("-")[1], 10);
  if (month >= 1 && month <= 7) {
    return sem2Term;
  }
  return sem1Term;
}

/**
 * Resolves the exact start and end clock times for a specific date and period
 * using the day's Bell Schedule (from PowerSchool or Settings) and Day Color.
 *
 * @return {Object|null} { startTime: "HH:MM", endTime: "HH:MM", startLocalStr: "YYYYMMDDTHHMM00", endLocalStr: "YYYYMMDDTHHMM00", startDateObj: Date, endDateObj: Date, scheduleName: string }
 */
function getPeriodTimeRangeForDate(dateStr, period, optionalDateColors, optionalSettings) {
  if (!dateStr || period === undefined || period === null || String(period).trim() === "") return null;

  var isoDate = typeof getStandardizedDateString === "function" ? getStandardizedDateString(dateStr) : String(dateStr).trim();
  if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return null;

  var settings = optionalSettings || (typeof getSettings === "function" ? getSettings() : {});
  var dateColors = optionalDateColors || (typeof getDateColorsCached === "function" ? getDateColorsCached() : {});
  var dayColor = (dateColors && dateColors[isoDate]) ? String(dateColors[isoDate]).trim() : "Green";
  if (dayColor === "Red") return null;

  var bellTimesMap = DEFAULT_PERIOD_BELL_TIMES;
  if (settings && settings["Period Bell Times"]) {
    try {
      var parsed = JSON.parse(settings["Period Bell Times"]);
      if (parsed && typeof parsed === "object") {
        bellTimesMap = parsed;
      }
    } catch (e) {
      // fallback to DEFAULT_PERIOD_BELL_TIMES
    }
  }

  var scheduleName = dayColor;
  if (settings && settings["Date Bell Schedule Names"]) {
    try {
      var dateSchedMap = JSON.parse(settings["Date Bell Schedule Names"]);
      if (dateSchedMap && dateSchedMap[isoDate] && bellTimesMap[dateSchedMap[isoDate]]) {
        scheduleName = dateSchedMap[isoDate];
      }
    } catch (e) {
      // ignore
    }
  }

  var schedTimes = bellTimesMap[scheduleName] || bellTimesMap[dayColor] || DEFAULT_PERIOD_BELL_TIMES[dayColor] || DEFAULT_PERIOD_BELL_TIMES["Green"];
  if (!schedTimes) return null;

  var normP = parsePeriod(period);
  var timeEntry = schedTimes[normP] || (normP === "0" ? schedTimes["9"] : (normP === "9" ? schedTimes["0"] : null));
  if (!timeEntry && DEFAULT_PERIOD_BELL_TIMES[dayColor]) {
    timeEntry = DEFAULT_PERIOD_BELL_TIMES[dayColor][normP];
  }
  if (!timeEntry && DEFAULT_PERIOD_BELL_TIMES["Green"]) {
    timeEntry = DEFAULT_PERIOD_BELL_TIMES["Green"][normP];
  }
  if (!timeEntry || !timeEntry.start || !timeEntry.end) return null;

  var startMatch = String(timeEntry.start).trim().match(/^(\d{1,2}):(\d{2})$/);
  var endMatch = String(timeEntry.end).trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!startMatch || !endMatch) return null;

  var dParts = isoDate.split("-");
  var yr = parseInt(dParts[0], 10);
  var mo = parseInt(dParts[1], 10);
  var dy = parseInt(dParts[2], 10);

  var sh = parseInt(startMatch[1], 10);
  var sm = parseInt(startMatch[2], 10);
  var eh = parseInt(endMatch[1], 10);
  var em = parseInt(endMatch[2], 10);

  var pad2 = function(n) { return n < 10 ? "0" + n : String(n); };
  var ymd = String(yr) + pad2(mo) + pad2(dy);
  var startLocalStr = ymd + "T" + pad2(sh) + pad2(sm) + "00";
  var endLocalStr = ymd + "T" + pad2(eh) + pad2(em) + "00";

  var startDateObj = new Date(yr, mo - 1, dy, sh, sm, 0);
  var endDateObj = new Date(yr, mo - 1, dy, eh, em, 0);

  var fmt12 = function(h, m) {
    var h12 = h === 0 ? 12 : (h > 12 ? h - 12 : h);
    return String(h12) + ":" + pad2(m);
  };

  return {
    startTime: pad2(sh) + ":" + pad2(sm),
    endTime: pad2(eh) + ":" + pad2(em),
    displayRange: fmt12(sh, sm) + "-" + fmt12(eh, em),
    startLocalStr: startLocalStr,
    endLocalStr: endLocalStr,
    startDateObj: startDateObj,
    endDateObj: endDateObj,
    scheduleName: scheduleName,
    dayColor: dayColor
  };
}


