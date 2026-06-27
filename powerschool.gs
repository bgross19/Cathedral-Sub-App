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

/**
 * Test function to fetch the master schedule from PowerSchool and dump it into a temporary sheet.
 * 
 * NOTE FOR POWERSCHOOL ADMIN: 
 * Because PowerSchool's default API doesn't expose a clean, single endpoint for this, 
 * you will need to create a "PowerQuery" plugin in PowerSchool with the endpoint path:
 * /ws/schema/query/com.cathedral.subapp.masterschedule
 * 
 * You can base the PowerQuery on this provided SQL:
 * WITH DistinctClasses AS (
 *     -- Step 1: Get a clean list with only ONE row per teacher, per period, per course, per room
 *     SELECT DISTINCT
 *         t.LASTFIRST, 
 *         t.EMAIL_Addr, 
 *         REPLACE(cc.expression, '(A)', '') AS period,
 *         t.EMAIL_Addr || '-' || REPLACE(cc.expression, '(A)', '') AS email_period_join,
 *         c.course_name,
 *         s.room,  -- Added Room Number here
 *         t.id AS teacher_id
 *     FROM cc cc
 *     JOIN courses c
 *         ON c.course_number = cc.course_number
 *     JOIN teachers t
 *         ON cc.TEACHERID = t.id
 *     JOIN sections s                     -- New JOIN for the Sections table
 *         ON cc.sectionid = s.id          -- Linking the enrollment to the specific section
 *     WHERE cc.termid = :target_term      -- Use a parameterized term ID
 * )
 * -- Step 2: Combine the co-seated courses from that clean list
 * SELECT 
 *     LASTFIRST,
 *     EMAIL_Addr,
 *     period,
 *     email_period_join,
 *     room,  -- Pulling the Room Number through to the final result
 *     LISTAGG(course_name, ' / ') WITHIN GROUP (ORDER BY course_name) AS course_names,
 *     MAX(teacher_id) AS teacher_id
 * FROM DistinctClasses
 * GROUP BY 
 *     LASTFIRST,
 *     EMAIL_Addr,
 *     period,
 *     email_period_join,
 *     room   -- Grouping by Room Number as well
 * ORDER BY LASTFIRST
 */
function testPowerSchoolMasterScheduleFetch() {
  const token = getPowerSchoolToken();
  if (!token) {
    Logger.log("Failed to get PowerSchool token.");
    return;
  }
  
  const settings = typeof getSettings === "function" ? getSettings() : {};
  const rawUrl = settings['PS_URL'];
  const POWERSCHOOL_URL = rawUrl ? rawUrl.trim().replace(/\/$/, '') : null;

  if (!POWERSCHOOL_URL) {
    Logger.log("Missing PS_URL property.");
    return;
  }
  
  // The placeholder PowerQuery endpoint. 
  // Update this if you name your PowerQuery differently.
  // By default, PowerSchool PowerQueries limit results to 100 records.
  // Appending ?pagesize=0 instructs it to return all records at once.
  const endpoint = "/ws/schema/query/com.cathedral.subapp.masterschedule?pagesize=0";
  
  const options = {
    method: "POST", // PowerQueries require POST, even for retrieving data
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    // PowerQueries require a JSON payload, even if empty, when using POST.
    payload: JSON.stringify({}),
    muteHttpExceptions: true
  };
  
  let responseText;
  let statusCode;
  try {
    let url = POWERSCHOOL_URL + endpoint;
    Logger.log('Fetching URL: ' + url);
    const response = UrlFetchApp.fetch(url, options);
    statusCode = response.getResponseCode();
    responseText = response.getContentText();
  } catch (error) {
    Logger.log("API Fetch Error: " + error.toString());
    responseText = error.toString();
    statusCode = "ERROR";
  }
  
  // Now write to Google Sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "PS Master Schedule Test";
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  // Set Headers
  sheet.appendRow(["LASTFIRST", "EMAIL_ADDR", "PERIOD", "ROOM", "COURSE_NAMES", "TERM"]);
  
  // If we have an error code (like 404 because the query doesn't exist yet)
  if (statusCode !== 200) {
    sheet.appendRow(["API ERROR", "Status Code: " + statusCode, responseText, "", "", ""]);
    Logger.log("API returned status: " + statusCode);
    return;
  }
  
  try {
    const json = JSON.parse(responseText);
    // PowerQueries usually return data in a `record` array.
    const records = json.record || [];
    
    if (records.length === 0) {
      sheet.appendRow(["NO DATA RETURNED", JSON.stringify(json), "", "", "", ""]);
    } else {
      const rows = records.map(r => {
        let periodRaw = String(r.period || r.PERIOD || "");
        let periodMatch = periodRaw.match(/\d+/);
        let periodClean = periodMatch ? periodMatch[0] : periodRaw.trim();

        return [
          r.lastfirst || r.LASTFIRST || "",
          r.email_addr || r.EMAIL_ADDR || "",
          periodClean,
          r.room || r.ROOM || r.room_number || r.ROOM_NUMBER || "",
          r.course_names || r.COURSE_NAMES || r.course_name || r.COURSE_NAME || "",
          "Target Term" // The actual term ID should be mapped if returned by the API
        ];
      });
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }
  } catch (parseError) {
    sheet.appendRow(["JSON PARSE ERROR", parseError.toString(), responseText, "", "", ""]);
    Logger.log("Failed to parse JSON: " + parseError.toString());
  }
}

let global_master_schedule_cache = null;

/**
 * Fetches the Master Schedule from PowerSchool API and formats it into a 2D array.
 * Caches the result in Script Cache to avoid excessive API calls.
 * Uses a global variable to avoid repeated JSON parsing in the same execution context.
 */
function getMasterScheduleData() {
  if (global_master_schedule_cache) {
    return global_master_schedule_cache;
  }

  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("ps_master_schedule");

  if (cachedData) {
    try {
      global_master_schedule_cache = JSON.parse(cachedData);
      return global_master_schedule_cache;
    } catch (e) {
      Logger.log("Error parsing cached master schedule: " + e.toString());
    }
  }

  const token = getPowerSchoolToken();
  if (!token) {
    Logger.log("Failed to get PowerSchool token in getMasterScheduleData.");
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
    Logger.log("API Fetch Error in getMasterScheduleData: " + error.toString());
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

    // Cache limits to 6 hours (21600 seconds)
    // 100kb string limit for Script Cache, if it exceeds, we may need to handle it or compress
    const stringifiedData = JSON.stringify(scheduleData);
    if (stringifiedData.length <= 100000) {
      cache.put("ps_master_schedule", stringifiedData, 21600);
    } else {
      Logger.log("Master Schedule data exceeds cache size limit. Not caching.");
      // In a robust implementation, you might chunk the cache.
      // E.g., caching multiple chunks. Let's start simple.
    }

    global_master_schedule_cache = scheduleData;
    return scheduleData;

  } catch (parseError) {
    Logger.log("Failed to parse Master Schedule JSON: " + parseError.toString());
    return [];
  }
}


/**
 * Warms the Master Schedule cache by retrieving fresh data from PowerSchool.
 * Intended to be run periodically via a time-driven trigger.
 */
function warmMasterScheduleCache() {
  const cache = CacheService.getScriptCache();
  // Clear existing cache
  cache.remove("ps_master_schedule");

  // Clear the global in-memory variable to force a fresh fetch
  global_master_schedule_cache = null;

  // Call the function to fetch and re-cache the data
  getMasterScheduleData();

  Logger.log("Master Schedule cache warmed successfully.");
}

/**
 * Sets up a time-driven trigger to warm the Master Schedule cache every 2 hours.
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

  // Create a new trigger to run every 2 hours
  ScriptApp.newTrigger('warmMasterScheduleCache')
    .timeBased()
    .everyHours(2)
    .create();

  Logger.log("Cache warming trigger created successfully to run every 2 hours.");
}
