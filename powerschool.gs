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

  const properties = PropertiesService.getScriptProperties();
  const CLIENT_ID = properties.getProperty('PS_CLIENT_ID');
  const CLIENT_SECRET = properties.getProperty('PS_CLIENT_SECRET');
  const POWERSCHOOL_URL = properties.getProperty('PS_URL');

  if (!CLIENT_ID || !CLIENT_SECRET || !POWERSCHOOL_URL) {
    throw new Error("Missing PowerSchool API configuration in Script Properties.");
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
 *     WHERE cc.termid = 3503
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

  const properties = PropertiesService.getScriptProperties();
  const POWERSCHOOL_URL = properties.getProperty('PS_URL');

  // The placeholder PowerQuery endpoint.
  // Update this if you name your PowerQuery differently.
  const endpoint = "/ws/schema/query/com.cathedral.subapp.masterschedule";

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
    const response = UrlFetchApp.fetch(POWERSCHOOL_URL + endpoint, options);
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
      const rows = records.map(r => [
        r.lastfirst || r.LASTFIRST || "",
        r.email_addr || r.EMAIL_ADDR || "",
        r.period || r.PERIOD || "",
        r.room || r.ROOM || "",
        r.course_names || r.COURSE_NAMES || "",
        "3503" // Term is hardcoded in the SQL for now (3503), but you can map it if added to output
      ]);
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }
  } catch (parseError) {
    sheet.appendRow(["JSON PARSE ERROR", parseError.toString(), responseText, "", "", ""]);
    Logger.log("Failed to parse JSON: " + parseError.toString());
  }
}
