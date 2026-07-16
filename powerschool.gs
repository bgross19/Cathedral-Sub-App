/**
 * Authenticates with PowerSchool and retrieves a valid OAuth Bearer Token.
 * Caches the token to optimize app performance.
 */
function getPowerSchoolToken() {
  const cache = CacheService.getScriptCache();
  const cachedToken = cache.get("ps_access_token");

  if (cachedToken) {
    return cachedToken;
  }

  const settings = typeof getSettings === "function" ? getSettings() : {};
  const rawClientId = settings['PS_CLIENT_ID'];
  const rawClientSecret = settings['PS_CLIENT_SECRET'];
  const rawUrl = settings['PS_URL'];

  const CLIENT_ID = rawClientId ? rawClientId.trim() : null;
  const CLIENT_SECRET = rawClientSecret ? rawClientSecret.trim() : null;
  const POWERSCHOOL_URL = rawUrl ? rawUrl.trim().replace(/\/$/, '') : null;

  if (!CLIENT_ID || !CLIENT_SECRET || !POWERSCHOOL_URL) {
    throw new Error("Missing PowerSchool API configuration in Settings.");
  }

  // PowerSchool OAuth requires Base64 encoding the ClientID:ClientSecret
  const credentialString = CLIENT_ID + ":" + CLIENT_SECRET;
  const base64Credentials = Utilities.base64Encode(credentialString);

  const options = {
    method: "POST",
    headers: {
      "Authorization": "Basic " + base64Credentials,
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"
    },
    payload: "grant_type=client_credentials",
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(POWERSCHOOL_URL + "/oauth/access_token", options);
    const json = JSON.parse(response.getContentText());

    if (json.access_token) {
      // Cache the token for 55 minutes (it typically expires in 60)
      cache.put("ps_access_token", json.access_token, 3300);
      return json.access_token;
    } else {
      throw new Error("Failed to retrieve token: " + response.getContentText());
    }
  } catch (error) {
    Logger.log("Authentication Error: " + error.toString());
    return null;
  }
}

let global_master_schedule_cache = null;

/**
 * Fetches the Master Schedule from PowerSchool API and formats it into a 2D array.
 */
function fetchMasterScheduleFromAPI() {
  const token = getPowerSchoolToken();
  if (!token) {
    Logger.log("Failed to get PowerSchool token in fetchMasterScheduleFromAPI.");
    return [];
  }

  const settings = typeof getSettings === "function" ? getSettings() : {};
  const rawUrl = settings['PS_URL'];
  const POWERSCHOOL_URL = rawUrl ? rawUrl.trim().replace(/\/$/, '') : null;
  if (!POWERSCHOOL_URL) {
    Logger.log("Missing PS_URL property.");
    return [];
  }

  // Need to get dynamic Term ID from Settings
  const termId = settings["Term ID"] || "3503";

  const endpoint = "/ws/schema/query/com.cathedral.subapp.masterschedule?pagesize=0";

  const payload = {
    "target_term": String(termId)
  };

  const options = {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let responseText;
  let statusCode;
  try {
    let url = POWERSCHOOL_URL + endpoint;
    const response = UrlFetchApp.fetch(url, options);
    statusCode = response.getResponseCode();
    responseText = response.getContentText();
  } catch (error) {
    Logger.log("API Fetch Error in fetchMasterScheduleFromAPI: " + error.toString());
    return [];
  }

  if (statusCode !== 200) {
    Logger.log("PowerSchool API returned status: " + statusCode + ", " + responseText);
    return [];
  }

  try {
    const json = JSON.parse(responseText);
    const records = json.record || [];

    // Headers matching expected 2D array structure
    const scheduleData = [
      ["LASTFIRST", "EMAIL_ADDR", "PERIOD", "ROOM", "COURSE_NAMES", "TERM", "EMAIL_PERIOD_JOIN"]
    ];

    records.forEach(r => {
      const email = String(r.email_addr || r.EMAIL_ADDR || "").trim();

      let periodRaw = String(r.period || r.PERIOD || "");
      let periodMatch = periodRaw.match(/\d+/);
      let periodClean = periodMatch ? periodMatch[0] : periodRaw.trim();

      const joinKey = (email && periodClean) ? (email + "-" + periodClean) : "";

      scheduleData.push([
        r.lastfirst || r.LASTFIRST || "",
        email,
        periodClean,
        r.room || r.ROOM || r.room_number || r.ROOM_NUMBER || "",
        r.course_names || r.COURSE_NAMES || r.course_name || r.COURSE_NAME || "",
        termId,
        joinKey
      ]);
    });

    return scheduleData;

  } catch (parseError) {
    Logger.log("Failed to parse Master Schedule JSON: " + parseError.toString());
    return [];
  }
}

/**
 * Gets the Master Schedule data by reading directly from the "Master Schedule Cache" sheet.
 * This completely avoids API calls during normal app usage.
 * Uses a global variable to avoid repeated sheet reads in the same execution context.
 */
function getMasterScheduleData() {
  if (global_master_schedule_cache) {
    return global_master_schedule_cache;
  }

  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("master_schedule_data");
  if (cachedData) {
    try {
      global_master_schedule_cache = JSON.parse(cachedData);
      return global_master_schedule_cache;
    } catch (e) {
      Logger.log("Failed to parse cached Master Schedule data.");
    }
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master Schedule Cache");
    if (!sheet) {
      Logger.log("Master Schedule Cache sheet not found. Attempting to fetch from API as fallback.");
      const data = fetchMasterScheduleFromAPI();
      global_master_schedule_cache = data;
      return data;
    }

    const data = sheet.getDataRange().getValues();
    if (data && data.length > 0) {
      global_master_schedule_cache = data;

      try {
        const stringified = JSON.stringify(data);
        if (stringified.length < 100000) { // Max 100KB cache
           cache.put("master_schedule_data", stringified, 1800); // 30 mins
        }
      } catch (e) {}

      return data;
    }
    return [];
  } catch (e) {
    Logger.log("Error reading Master Schedule Cache sheet: " + e.toString());
    return [];
  }
}

/**
 * Warms the Master Schedule cache by retrieving fresh data from PowerSchool
 * and saving it to the "Master Schedule Cache" sheet.
 * Intended to be run periodically via a time-driven trigger.
 */
function warmMasterScheduleCache() {
  // Clear the global in-memory variable
  global_master_schedule_cache = null;
  CacheService.getScriptCache().remove("master_schedule_data");

  // Fetch from API
  const scheduleData = fetchMasterScheduleFromAPI();

  if (scheduleData && scheduleData.length > 0) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Master Schedule Cache");
    if (!sheet) {
      sheet = ss.insertSheet("Master Schedule Cache");
    }

    // Clear old data and write new data
    sheet.clearContents();
    sheet.getRange(1, 1, scheduleData.length, scheduleData[0].length).setValues(scheduleData);

    Logger.log("Master Schedule Cache sheet updated successfully with " + scheduleData.length + " rows.");

    // Sync to Staff Roster
    try {
      syncMasterScheduleToStaffRoster(scheduleData, ss);
    } catch (e) {
      Logger.log("Error syncing master schedule to staff roster: " + e.toString());
    }
  } else {
    Logger.log("Master Schedule fetch returned empty data. Cache sheet not updated.");
  }
}

/**
 * Synchronizes the Master Schedule teachers to the Staff Roster.
 * Adds any missing teachers with default role "Teacher" and blank duty.
 */
function syncMasterScheduleToStaffRoster(scheduleData, ss) {
  if (!scheduleData || scheduleData.length <= 1) return;

  const rosterSheet = ss.getSheetByName("Staff Roster");
  if (!rosterSheet) return;

  const rosterData = rosterSheet.getDataRange().getValues();
  const existingEmails = {};

  // Roster format: Name, Email, Role, Duty
  for (let i = 1; i < rosterData.length; i++) {
    const email = String(rosterData[i][1]).toLowerCase().trim();
    if (email) {
      existingEmails[email] = true;
    }
  }

  const newTeachers = {};
  const headers = scheduleData[0];
  const nameIdx = headers.indexOf("LASTFIRST");
  const emailIdx = headers.indexOf("EMAIL_ADDR");

  if (nameIdx === -1 || emailIdx === -1) return;

  for (let s = 1; s < scheduleData.length; s++) {
    const row = scheduleData[s];
    const teacherName = String(row[nameIdx]).trim();
    // Maintain the exact casing for the email in the sheet, but lower-case for comparison
    const rawEmail = String(row[emailIdx]).trim();
    const teacherEmail = rawEmail.toLowerCase();

    if (teacherEmail && teacherName && !existingEmails[teacherEmail] && !newTeachers[teacherEmail]) {
      newTeachers[teacherEmail] = {
        name: teacherName,
        email: rawEmail
      };
    }
  }

  const newRows = [];
  let addedCount = 0;
  for (const key in newTeachers) {
    const teacher = newTeachers[key];
    newRows.push([teacher.name, teacher.email, "Teacher", ""]);

    // Also add to User Roles
    try {
      if (typeof upsertUserRoleInternal === 'function') {
        upsertUserRoleInternal(ss, teacher.email, "Teacher");
      }
    } catch (e) {
      Logger.log("Failed to add user role for " + teacher.email + ": " + e.message);
    }

    // Log the action
    try {
      if (typeof logAuditAction === 'function') {
        logAuditAction("STAFF_ADDED_SYNC", teacher.email, "Added staff member from Master Schedule: " + teacher.name + " (Teacher)");
      }
    } catch (e) {
      Logger.log("Failed to log audit action for " + teacher.email + ": " + e.message);
    }
    addedCount++;
  }

  if (newRows.length > 0) {
    const startRow = rosterSheet.getLastRow() + 1;
    rosterSheet.getRange(startRow, 1, newRows.length, 4).setValues(newRows);
    Logger.log("Successfully synced " + addedCount + " new teachers to Staff Roster.");

    // Invalidate roster cache
    try {
      if (typeof clearRosterCache === 'function') {
        clearRosterCache();
      }
    } catch (e) {}
  }
}

/**
 * Sets up a time-driven trigger to warm the Master Schedule cache daily between 1 AM and 2 AM.
 * Should be run once during initial setup or if triggers are lost.
 */
function setupCacheWarmingTrigger() {
  // First, remove any existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'warmMasterScheduleCache') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create a new trigger to run daily between 1 AM and 2 AM
  ScriptApp.newTrigger('warmMasterScheduleCache')
    .timeBased()
    .everyDays(1)
    .atHour(1) // Runs between 1 AM and 2 AM
    .create();

  Logger.log("Cache warming trigger created successfully to run daily between 1 AM and 2 AM.");
}
