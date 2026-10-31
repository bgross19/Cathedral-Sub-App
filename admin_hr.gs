// ============================================================================
// ADMIN_HR.GS — Settings, Staff Roster CRUD, Pay Periods & Data Archiving
// ============================================================================

/**
 * Helper to parse settings from the sheet data.
 */
function _parseSettingsData(settingsSheet) {
  var settings = {};
  if (!settingsSheet) return settings;

  var data = settingsSheet.getDataRange().getValues();
  // Skip header row
  for (var i = 1; i < data.length; i++) {
  if (!data[i]) continue;
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
    "Sub Availability Window (Days)": "5",
    "Urgency Cutoff Time": "15",
    "Term ID": "3601",
    "Auto Sync Google Calendar": "true",
    "Period Bell Times": JSON.stringify(typeof DEFAULT_PERIOD_BELL_TIMES !== "undefined" ? DEFAULT_PERIOD_BELL_TIMES : {}),
    "Green Day Pay Rate": "10",
    "Blue/Gold Day Pay Rate": "20",
    "Special Pay Rates": "[]",
    "Absence Reasons": JSON.stringify([
      {reason: "Personal", hrRequired: false, principalRequired: false},
      {reason: "Professional Development", hrRequired: false, principalRequired: false},
      {reason: "Retreat", hrRequired: false, principalRequired: false},
      {reason: "Athletics", hrRequired: false, principalRequired: true},
      {reason: "Jury Duty", hrRequired: true, principalRequired: false},
      {reason: "Bereavement", hrRequired: true, principalRequired: false}
    ]),
    "RolePermissions": JSON.stringify({
        "admin": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Absences": true, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "hr": { "Admin Dashboard": false, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Absences": false, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "sub coordinator": { "Admin Dashboard": true, "HR Dashboard": false, "Today at a Glance": true, "My Upcoming Sub Duties": true, "Today's Absences": true, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "principal": { "Admin Dashboard": true, "HR Dashboard": true, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Absences": false, "My Past Absences": true, "Settings": true, "Add Request on Behalf": true },
        "teacher": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Absences": true, "My Past Absences": true, "Settings": false, "Add Request on Behalf": false },
        "substitute": { "Admin Dashboard": false, "HR Dashboard": false, "Today at a Glance": false, "My Upcoming Sub Duties": true, "Today's Absences": false, "My Past Absences": false, "Settings": false, "Add Request on Behalf": false }
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
    if (!data[i]) continue;
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
  var rawRosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
  var rosterData = [];
  for (var rIdx = 0; rIdx < rawRosterData.length; rIdx++) {
    var rRow = rawRosterData[rIdx];
    if (!rRow) continue;
    if (rIdx === 0 || (rRow[0] && String(rRow[0]).trim() !== "") || (rRow[1] && String(rRow[1]).trim() !== "")) {
      rosterData.push(rRow.slice(0, 12));
    }
  }

  // Try to cache if it's not huge
  try {
     var stringified = JSON.stringify(rosterData);
     if (stringified.length < 100000) { // Keep under 100KB cache limit
        cache.put("staff_roster_data", stringified, 1800); // 30 mins
     }
  } catch (e) {
    // Ignore cache put error
  }

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
    if (!data[i]) continue;
        if (data[i] && String(data[i][1]).trim().toLowerCase() === targetEmail) {
            rowIndexToUpdate = i + 1;
            currentName = String(data[i][0]).trim();
            currentDuty = String(data[i][3]).trim();
            break;
        }
    }

    if (rowIndexToUpdate !== -1) {
       rosterSheet.getRange(rowIndexToUpdate, 1, 1, 4).setValues([[currentName, targetEmail, newRole, currentDuty]]);
       logAuditAction("STAFF_UPDATED", targetEmail, "Updated staff role inline: " + currentName + " to " + newRole, clientEmail);
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
    if (!data[i]) continue;
        if (data[i] && String(data[i][1]).trim().toLowerCase() === targetEmail) {
            rowIndexToUpdate = i + 1;
            currentName = String(data[i][0]).trim();
            currentRole = String(data[i][2]).trim();
            break;
        }
    }

    if (rowIndexToUpdate !== -1) {
       rosterSheet.getRange(rowIndexToUpdate, 1, 1, 4).setValues([[currentName, targetEmail, currentRole, newDuty]]);
       logAuditAction("STAFF_UPDATED", targetEmail, "Updated staff duty inline: " + currentName + " to " + newDuty, clientEmail);
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
      if (!data[i]) continue;
        if (data[i] && String(data[i][1]).trim().toLowerCase() === originalEmail) {
          rowIndexToUpdate = i + 1;
          break;
        }
      }
    }

    if (rowIndexToUpdate !== -1) {
       // Update existing
       rosterSheet.getRange(rowIndexToUpdate, 1, 1, 4).setValues([[newName, newEmail, newRole, newDuty]]);
       logAuditAction("STAFF_UPDATED", newEmail, "Updated staff member: " + newName + " (" + newRole + ", " + newDuty + ")", clientEmail);
    } else {
       // Check if new email already exists to prevent duplicates
       for (let i = 1; i < data.length; i++) {
       if (!data[i]) continue;
         if (data[i] && String(data[i][1]).trim().toLowerCase() === newEmail.toLowerCase()) {
            return {
      success: false, error: "A staff member with this email already exists." };
         }
       }
       // Append new
       rosterSheet.appendRow([newName, newEmail, newRole, newDuty]);
       logAuditAction("STAFF_ADDED", newEmail, "Added staff member: " + newName + " (" + newRole + ", " + newDuty + ")", clientEmail);
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
    if (!data[i]) continue;
      if (data[i] && String(data[i][1]).trim().toLowerCase() === targetEmail) {
        targetIndex = i;
        break;
      }
    }

    if (targetIndex !== -1) {
      logAuditAction("STAFF_DELETED", targetEmail, "Deleted staff member", clientEmail);
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

    logAuditAction("STAFF_DUTIES_CLEARED", "All", "Admin cleared all staff duties", clientEmail);
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
    if (!data[i]) continue;
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
       let email = String(update.email || "").trim();
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

    logAuditAction("STAFF_BULK_UPLOAD", "Multiple", "Processed " + processedCount + " staff records", clientEmail);
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
    return sanitizeSettingsForClient(getSettings(ss), true);
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
    if (!data[i]) continue;
      settingsMap[String(data[i][0]).trim()] = i;
    }

    var rowsToAppend = [];
    var dataChanged = false;

    for (var key in newSettings) {
      if (key === "PS_CLIENT_SECRET" && newSettings[key] === MASKED_SECRET_VALUE) {
        // Preserve existing secret when the masked placeholder is submitted unchanged
        continue;
      }
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

    logAuditAction("SETTINGS_UPDATED", "Global", "Updated application settings", clientEmail);

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
    if (!data[i]) continue;
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
      if (typeof hasV2AbsenceSheets === "function" && hasV2AbsenceSheets(ss) && typeof rebuildAllV2AbsenceSheets === "function") {
        rebuildAllV2AbsenceSheets(ss);
      }

      logAuditAction("ARCHIVE_DATA", "N/A", "Archived " + rowsToArchive.length + " absence requests older than " + cutoffDateStr, clientEmail);
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

/**
 * Parses a single CSV line into an array of unquoted cell strings.
 */
function _parsePayrollCsvLine(line) {
  var result = [];
  var current = "";
  var inQuotes = false;
  var s = String(line || "");
  for (var i = 0; i < s.length; i++) {
    var ch = s.charAt(i);
    if (ch === '"') {
      if (inQuotes && i + 1 < s.length && s.charAt(i + 1) === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      result.push(current.trim());
      current = "";
    } else {
      current += ch;
    }
  }
  result.push(current.trim());
  return result;
}

/**
 * Appends a pay period approval (or re-approval) snapshot to the "Payroll" sheet.
 * Creates the "Payroll" sheet with headers if it does not exist yet.
 */
function _appendPayrollSheetSnapshot(ss, periodNumber, rangeString, csvString, approvedBy) {
  var payrollHeaders = [
    "Pay Period",
    "Date Range",
    "Teacher Name",
    "Duties",
    "Extras",
    "Extra (Pay)",
    "Totals",
    "Approved By",
    "Approved At"
  ];

  var payrollSheet = ss.getSheetByName("Payroll");
  if (!payrollSheet) {
    payrollSheet = ss.insertSheet("Payroll");
    payrollSheet.getRange(1, 1, 1, payrollHeaders.length).setValues([payrollHeaders]);
    payrollSheet.getRange(1, 1, 1, payrollHeaders.length).setFontWeight("bold");
  } else if (payrollSheet.getLastRow() === 0) {
    payrollSheet.getRange(1, 1, 1, payrollHeaders.length).setValues([payrollHeaders]);
    payrollSheet.getRange(1, 1, 1, payrollHeaders.length).setFontWeight("bold");
  }

  var tz = (typeof Session !== "undefined" && Session.getScriptTimeZone) ? Session.getScriptTimeZone() : "America/Indiana/Indianapolis";
  var approvedAtStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
  var actorLabel = String(approvedBy || "Admin").trim();
  var periodStr = String(periodNumber !== undefined && periodNumber !== null ? periodNumber : "").trim();
  var rangeStr = String(rangeString || "").trim();

  var lines = String(csvString || "").split(/\r?\n/).filter(function(l) {
    return String(l || "").trim() !== "";
  });

  var rowsToAppend = [];
  for (var i = 0; i < lines.length; i++) {
    var cols = _parsePayrollCsvLine(lines[i]);
    if (cols.length === 0) continue;
    // Skip header row
    if (i === 0 && (cols[0].toLowerCase() === "teacher name" || cols[0].toLowerCase() === "name")) {
      continue;
    }
    var teacherName = cols[0] || "";
    if (!teacherName) continue;
    var duties = cols.length > 1 ? cols[1] : "0";
    var extras = cols.length > 2 ? cols[2] : "0";
    var extraPay = cols.length > 3 ? cols[3] : "$0.00";
    var totals = cols.length > 4 ? cols[4] : "0";

    rowsToAppend.push([
      periodStr,
      rangeStr,
      teacherName,
      duties,
      extras,
      extraPay,
      totals,
      actorLabel,
      approvedAtStr
    ]);
  }

  if (rowsToAppend.length === 0) {
    rowsToAppend.push([
      periodStr,
      rangeStr,
      "No Paid Sub Coverages",
      "0",
      "0",
      "$0.00",
      "0",
      actorLabel,
      approvedAtStr
    ]);
  }

  var startRow = payrollSheet.getLastRow() + 1;
  payrollSheet.getRange(startRow, 1, rowsToAppend.length, payrollHeaders.length).setValues(rowsToAppend);
  return rowsToAppend.length;
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
    var wasAlreadyApproved = false;

    for (var i = 1; i < payPeriodsData.length; i++) {
      if (String(payPeriodsData[i][0]).trim() === String(periodNumber).trim()) {
        rowIndex = i + 1; // 1-based index
        wasAlreadyApproved = String(payPeriodsData[i][3] || "").toLowerCase().trim() === "true";
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: "Pay period not found." };
    }

    // Set approved to true in the 4th column
    payPeriodsSheet.getRange(rowIndex, 4).setValue(true);

    // Write snapshot rows to the Payroll sheet on every approval or re-approval
    var actorName = (user && user.name) ? user.name : (clientEmail || "Admin");
    var writtenRows = _appendPayrollSheetSnapshot(ss, periodNumber, rangeString, csvString, actorName);

    var actionWord = wasAlreadyApproved ? "re-approved" : "approved";
    var auditAction = wasAlreadyApproved ? "PAY_PERIOD_REAPPROVED" : "PAY_PERIOD_APPROVED";
    logAuditAction(auditAction, "Period " + periodNumber, actorName + " " + actionWord + " pay period " + periodNumber + " (" + writtenRows + " payroll row(s) recorded)", clientEmail);

    // Get emails
    var hrEmails = getEmailsByRole(ss, "hr");
    var subCoordinatorEmails = getEmailsByRole(ss, "sub coordinator");

    if (hrEmails.length === 0 && subCoordinatorEmails.length === 0) {
       return { success: true, reapproved: wasAlreadyApproved, rowsWritten: writtenRows, message: "Pay period " + actionWord + " and written to Payroll sheet, but no HR or Sub Coordinator emails found to notify." };
    }

    var subject = "Pay Period " + periodNumber + " with the date range " + actionWord;
    var body = actorName + " has " + actionWord + " the payroll for the subs for Pay Period " + periodNumber + " (" + rangeString + "). The summary has been recorded to the Payroll sheet. Please reach out with any questions.";

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

    return { success: true, reapproved: wasAlreadyApproved, rowsWritten: writtenRows };
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

    logAuditAction("PAY_PERIODS_BULK_UPLOAD", "Multiple", "Added " + newRows.length + " pay periods", clientEmail);

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

    logAuditAction("PAY_PERIODS_DELETE_ALL", "All", "Deleted all pay periods", clientEmail);

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
    if (!data[i]) continue;
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

// ============================================================================
// BELL SCHEDULE (DATES SHEET) MANAGEMENT & CACHING
// ============================================================================
var _globalDatesCache = null;

function _normalizeScheduleDateStr(dateRaw) {
  if (!dateRaw) return "";
  if (dateRaw instanceof Date) {
    if (isNaN(dateRaw.getTime())) return "";
    return Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(dateRaw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // Support M/D/YYYY or M-D-YYYY
  var parts = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (parts) {
    var mm = parts[1].length < 2 ? "0" + parts[1] : parts[1];
    var dd = parts[2].length < 2 ? "0" + parts[2] : parts[2];
    return parts[3] + "-" + mm + "-" + dd;
  }

  try {
    var parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
  } catch (e) {}
  return "";
}

function _normalizeBellColor(colorRaw) {
  if (!colorRaw) return "";
  var c = String(colorRaw).trim().toLowerCase();
  if (c === "none" || c === "clear" || c === "unset" || c === "") return "";
  if (c === "green" || c === "green day" || c === "g") return "Green";
  if (c === "blue" || c === "blue day" || c === "b") return "Blue";
  if (c === "gold" || c === "gold day" || c === "y") return "Gold";
  if (c === "red" || c === "red day" || c === "no school" || c === "holiday" || c === "r") return "Red";
  // Preserve capitalized custom day types if ever used
  var rawTrim = String(colorRaw).trim();
  return rawTrim.charAt(0).toUpperCase() + rawTrim.slice(1);
}

/**
 * Retrieves the Date -> Color map from CacheService or the Dates sheet.
 */
function getDateColorsCached(ss) {
  if (_globalDatesCache) return _globalDatesCache;

  var cache = CacheService.getScriptCache();
  var cached = cache ? cache.get("bell_schedule_dates") : null;
  if (cached) {
    try {
      _globalDatesCache = JSON.parse(cached);
      return _globalDatesCache;
    } catch (e) {
      console.warn("Failed to parse cached bell schedule dates.");
    }
  }

  var sheetSS = ss || getSS();
  var datesSheet = sheetSS.getSheetByName("Dates");
  var dateColors = {};

  if (datesSheet) {
    var datesData = datesSheet.getDataRange().getValues();
    for (var d = 1; d < datesData.length; d++) {
      if (!datesData[d] || datesData[d].length < 2) continue;
      var dateFormatted = _normalizeScheduleDateStr(datesData[d][0]);
      var colorNormalized = _normalizeBellColor(datesData[d][1]);
      if (dateFormatted && colorNormalized) {
        dateColors[dateFormatted] = colorNormalized;
      }
    }
  }

  try {
    if (cache) {
      var str = JSON.stringify(dateColors);
      if (str.length < 100000) {
        cache.put("bell_schedule_dates", str, 1800); // 30 minutes
      }
    }
  } catch (e) {}

  _globalDatesCache = dateColors;
  return dateColors;
}

function clearDatesCache() {
  _globalDatesCache = null;
  try {
    var cache = CacheService.getScriptCache();
    if (cache) cache.remove("bell_schedule_dates");
  } catch (e) {}
}

function _getOrCreateDatesSheet(ss) {
  var sheet = ss.getSheetByName("Dates");
  if (!sheet) {
    sheet = ss.insertSheet("Dates");
    sheet.getRange(1, 1, 1, 2).setValues([["Date", "Color"]]);
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold");
    sheet.getRange("A:A").setNumberFormat("@");
  }
  return sheet;
}

/**
 * Returns the full Date -> Color map for the Settings Bell Schedule editor.
 */
function getBellScheduleDatesAdmin(clientEmail) {
  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");
    return getDateColorsCached(ss);
  } catch (err) {
    notifyAdminOfError("getBellScheduleDatesAdmin", err);
    throw new Error("Failed to load bell schedule dates: " + err.message);
  }
}

/**
 * Saves or clears a single date's bell schedule color from the visual calendar.
 */
function saveBellScheduleDateAdmin(dateStr, color, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("saveBellScheduleDateAdmin_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var normalizedDate = _normalizeScheduleDateStr(dateStr);
    if (!normalizedDate) {
      throw new Error("Invalid date format: " + dateStr);
    }
    var normalizedColor = _normalizeBellColor(color);

    var sheet = _getOrCreateDatesSheet(ss);
    var data = sheet.getDataRange().getValues();
    var matchingRows = [];

    for (var i = 1; i < data.length; i++) {
      if (!data[i]) continue;
      var rowDate = _normalizeScheduleDateStr(data[i][0]);
      if (rowDate === normalizedDate) {
        matchingRows.push(i + 1); // 1-based sheet row
      }
    }

    if (!normalizedColor) {
      // Delete matching rows bottom-up if clearing the date
      for (var d = matchingRows.length - 1; d >= 0; d--) {
        sheet.deleteRow(matchingRows[d]);
      }
    } else if (matchingRows.length > 0) {
      sheet.getRange(matchingRows[0], 1, 1, 2).setValues([[normalizedDate, normalizedColor]]);
      for (var k = matchingRows.length - 1; k > 0; k--) {
        sheet.deleteRow(matchingRows[k]);
      }
    } else {
      sheet.appendRow([normalizedDate, normalizedColor]);
    }

    clearDatesCache();
    logAuditAction("BELL_SCHEDULE_UPDATED", normalizedDate, "Set bell schedule color to " + (normalizedColor || "Cleared"), clientEmail);

    return { success: true, date: normalizedDate, color: normalizedColor };
  } catch (err) {
    notifyAdminOfError("saveBellScheduleDateAdmin", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Bulk upserts (or replaces) Bell Schedule dates from CSV/TSV upload.
 */
function bulkUpsertBellScheduleDates(updates, replaceAll, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("bulkUpsertBellScheduleDates_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    if (!updates || !Array.isArray(updates) || updates.length === 0) {
      return { success: false, error: "No valid schedule rows provided." };
    }

    var sheet = _getOrCreateDatesSheet(ss);
    var dateMap = {};

    if (!replaceAll) {
      var existingData = sheet.getDataRange().getValues();
      for (var i = 1; i < existingData.length; i++) {
        if (!existingData[i] || existingData[i].length < 2) continue;
        var exDate = _normalizeScheduleDateStr(existingData[i][0]);
        var exColor = _normalizeBellColor(existingData[i][1]);
        if (exDate && exColor) {
          dateMap[exDate] = exColor;
        }
      }
    }

    var validUpdatesCount = 0;
    for (var u = 0; u < updates.length; u++) {
      var item = updates[u];
      if (!item) continue;
      var normDate = _normalizeScheduleDateStr(item.date);
      var normColor = _normalizeBellColor(item.color);
      if (!normDate) continue;
      if (!normColor) {
        delete dateMap[normDate];
      } else {
        dateMap[normDate] = normColor;
      }
      validUpdatesCount++;
    }

    var sortedDates = Object.keys(dateMap).sort();
    var allRows = [["Date", "Color"]];
    for (var s = 0; s < sortedDates.length; s++) {
      var dt = sortedDates[s];
      allRows.push([dt, dateMap[dt]]);
    }

    sheet.clearContents();
    sheet.getRange(1, 1, allRows.length, 2).setValues(allRows);
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold");

    clearDatesCache();
    logAuditAction("BELL_SCHEDULE_BULK_UPLOAD", "Multiple (" + validUpdatesCount + ")", (replaceAll ? "Replaced" : "Merged") + " bell schedule with " + validUpdatesCount + " entries", clientEmail);

    return { success: true, updated: validUpdatesCount, total: sortedDates.length, dateColors: dateMap };
  } catch (err) {
    notifyAdminOfError("bulkUpsertBellScheduleDates", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Clears all rows in the Dates sheet.
 */
function clearAllBellScheduleDates(clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    notifyAdminOfError("clearAllBellScheduleDates_lock", e);
    return { success: false, error: "The server is currently busy. Please try again." };
  }

  try {
    var ss = getSS();
    var user = getUserData(ss, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    assertPermission(user, "Settings");

    var sheet = _getOrCreateDatesSheet(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
    }

    clearDatesCache();
    logAuditAction("BELL_SCHEDULE_CLEARED", "All", "Cleared all bell schedule dates", clientEmail);

    return { success: true };
  } catch (err) {
    notifyAdminOfError("clearAllBellScheduleDates", err);
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


