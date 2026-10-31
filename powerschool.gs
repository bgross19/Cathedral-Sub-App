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
/**
 * Internal helper to fetch raw records from `com.cathedral.subapp.masterschedule` for a target term.
 */
function _fetchMasterScheduleRecordsForTerm(POWERSCHOOL_URL, token, targetTerm) {
  const endpoint = "/ws/schema/query/com.cathedral.subapp.masterschedule?pagesize=0";
  const options = {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    payload: JSON.stringify({ "target_term": String(targetTerm) }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(POWERSCHOOL_URL + endpoint, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    if (statusCode !== 200) {
      Logger.log("PowerSchool API returned status: " + statusCode + ", " + responseText);
      return null;
    }
    const json = JSON.parse(responseText);
    return json.record || [];
  } catch (error) {
    Logger.log("API Fetch Error in _fetchMasterScheduleRecordsForTerm (" + targetTerm + "): " + error.toString());
    return null;
  }
}

/**
 * Fetches the Master Schedule from PowerSchool API across Semester 1 & Semester 2
 * and formats it into a 2D array with each row's TERM preserved.
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

  // Primary configured Term ID from Settings (e.g. "3601")
  const termId = String(settings["Term ID"] || "3601").trim();
  const primaryRecords = _fetchMasterScheduleRecordsForTerm(POWERSCHOOL_URL, token, termId);
  if (!primaryRecords) {
    return [];
  }

  // Headers matching expected 2D array structure
  const scheduleData = [
    ["LASTFIRST", "EMAIL_ADDR", "PERIOD", "ROOM", "COURSE_NAMES", "TERM", "EMAIL_PERIOD_JOIN"]
  ];

  const seenRowKeys = {};
  const appendRecords = function(records, fallbackTerm) {
    if (!records || !Array.isArray(records)) return;
    records.forEach(r => {
      const email = String(r.email_addr || r.EMAIL_ADDR || "").trim();
      let periodRaw = String(r.period || r.PERIOD || "");
      let periodMatch = periodRaw.match(/\d+/);
      let periodClean = periodMatch ? periodMatch[0] : periodRaw.trim();
      const rowTerm = String(r.term_id || r.TERM_ID || fallbackTerm || termId).trim();
      const joinKey = (email && periodClean) ? (email + "-" + periodClean) : "";
      const dedupeKey = (joinKey + "|" + rowTerm + "|" + (r.room || r.ROOM || "") + "|" + (r.course_names || r.COURSE_NAMES || "")).toLowerCase();
      if (seenRowKeys[dedupeKey]) return;
      seenRowKeys[dedupeKey] = true;

      scheduleData.push([
        r.lastfirst || r.LASTFIRST || "",
        email,
        periodClean,
        r.room || r.ROOM || r.room_number || r.ROOM_NUMBER || "",
        r.course_names || r.COURSE_NAMES || r.course_name || r.COURSE_NAME || "",
        rowTerm,
        joinKey
      ]);
    });
  };

  appendRecords(primaryRecords, termId);

  // If the installed PowerSchool plugin is v2.0.0 (which filters `cc.termid = :target_term`
  // and does not return `term_id` per row), automatically fetch the sibling semester (e.g. 3603 / 3601)
  // so both Semester 1 and Semester 2 schedules are cached for date-aware lookups.
  const hasMultiTermColumn = primaryRecords.some(r => r && (r.term_id !== undefined || r.TERM_ID !== undefined));
  if (!hasMultiTermColumn && primaryRecords.length > 0) {
    const termNum = parseInt(termId, 10);
    if (!isNaN(termNum) && termNum >= 100) {
      const yearPrefix = Math.floor(termNum / 100);
      const sem1 = String(yearPrefix * 100 + 1);
      const sem2 = String(yearPrefix * 100 + 3);
      const siblingTerms = [sem1, sem2].filter(t => t !== termId);
      siblingTerms.forEach(sibTerm => {
        const sibRecords = _fetchMasterScheduleRecordsForTerm(POWERSCHOOL_URL, token, sibTerm);
        if (sibRecords && sibRecords.length > 0) {
          appendRecords(sibRecords, sibTerm);
        }
      });
    }
  }

  return scheduleData;
}

const MASTER_SCHEDULE_CACHE_CHUNK_SIZE = 90000;
const MASTER_SCHEDULE_CACHE_MAX_CHUNKS = 10;

function _readMasterScheduleFromScriptCache(cache) {
  if (!cache || typeof cache.get !== "function") return null;
  try {
    const chunkCountStr = cache.get("master_schedule_chunks");
    if (chunkCountStr) {
      const count = parseInt(chunkCountStr, 10);
      if (!isNaN(count) && count > 0 && count <= MASTER_SCHEDULE_CACHE_MAX_CHUNKS) {
        let combined = "";
        for (let i = 0; i < count; i++) {
          const part = cache.get("master_schedule_chunk_" + i);
          if (!part) return null;
          combined += part;
        }
        return JSON.parse(combined);
      }
    }
    const legacySingle = cache.get("master_schedule_data");
    if (legacySingle) {
      return JSON.parse(legacySingle);
    }
  } catch (e) {
    Logger.log("Failed to parse cached Master Schedule data.");
  }
  return null;
}

function _writeMasterScheduleToScriptCache(cache, data) {
  if (!cache || typeof cache.put !== "function" || !data) return;
  try {
    const compact = [];
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!r) continue;
      const sliced = r.slice(0, 7);
      const hasAny = sliced.some(c => c !== undefined && c !== null && String(c).trim() !== "");
      if (i === 0 || hasAny) {
        compact.push(sliced);
      }
    }
    const stringified = JSON.stringify(compact);
    if (stringified.length < MASTER_SCHEDULE_CACHE_CHUNK_SIZE) {
      cache.put("master_schedule_data", stringified, 1800);
      if (typeof cache.remove === "function") cache.remove("master_schedule_chunks");
      return;
    }
    const numChunks = Math.ceil(stringified.length / MASTER_SCHEDULE_CACHE_CHUNK_SIZE);
    if (numChunks <= MASTER_SCHEDULE_CACHE_MAX_CHUNKS) {
      for (let c = 0; c < numChunks; c++) {
        const chunkStr = stringified.substring(c * MASTER_SCHEDULE_CACHE_CHUNK_SIZE, (c + 1) * MASTER_SCHEDULE_CACHE_CHUNK_SIZE);
        cache.put("master_schedule_chunk_" + c, chunkStr, 1800);
      }
      cache.put("master_schedule_chunks", String(numChunks), 1800);
    }
  } catch (e) {
    /* ignored */
  }
}

function _clearMasterScheduleFromScriptCache(cache) {
  if (!cache || typeof cache.remove !== "function") return;
  try {
    cache.remove("master_schedule_data");
    cache.remove("master_schedule_chunks");
    for (let c = 0; c < MASTER_SCHEDULE_CACHE_MAX_CHUNKS; c++) {
      cache.remove("master_schedule_chunk_" + c);
    }
  } catch (e) {
    /* ignored */
  }
}

/**
 * Gets the Master Schedule data by reading from CacheService (with multi-chunk support)
 * or falling back to the "Master Schedule Cache" sheet.
 * Uses a global variable to avoid repeated sheet reads in the same execution context.
 */
function getMasterScheduleData() {
  if (global_master_schedule_cache) {
    return global_master_schedule_cache;
  }

  const cache = CacheService.getScriptCache();
  const cachedData = _readMasterScheduleFromScriptCache(cache);
  if (cachedData) {
    global_master_schedule_cache = cachedData;
    return global_master_schedule_cache;
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
    if (data && data.length > 1) {
      global_master_schedule_cache = data;
      _writeMasterScheduleToScriptCache(cache, data);
      return data;
    }

    Logger.log("Master Schedule Cache sheet is empty. Attempting to fetch from API as fallback.");
    const fallbackData = fetchMasterScheduleFromAPI();
    if (fallbackData && fallbackData.length > 1) {
      global_master_schedule_cache = fallbackData;
    }
    return fallbackData || [];
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
  _clearMasterScheduleFromScriptCache(CacheService.getScriptCache());

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
    global_master_schedule_cache = scheduleData;
    _writeMasterScheduleToScriptCache(CacheService.getScriptCache(), scheduleData);

    Logger.log("Master Schedule Cache sheet updated successfully with " + scheduleData.length + " rows.");

    // Sync to Staff Roster
    try {
      syncMasterScheduleToStaffRoster(scheduleData, ss);
    } catch (e) {
      Logger.log("Error syncing master schedule to staff roster: " + e.toString());
    }

    // Also attempt a non-destructive merge sync of the Bell Schedule from PowerSchool
    try {
      syncBellScheduleFromPowerSchool(false);
    } catch (e) {
      Logger.log("Error syncing bell schedule during cache warm: " + e.toString());
    }

    // Prune past/Not Available rows from SubstituteAvailability
    try {
      if (typeof pruneSubstituteAvailabilitySheet === "function") {
        pruneSubstituteAvailabilitySheet(ss);
      }
    } catch (e) {
      Logger.log("Error pruning SubstituteAvailability during cache warm: " + e.toString());
    }
  } else {
    Logger.log("Master Schedule fetch returned empty data. Cache sheet not updated.");
  }
}

/**
 * Maps a single PowerSchool `com.cathedral.subapp.bellschedule` record to a normalized
 * Sub App day color ("Green", "Blue", "Gold", "Red", or "").
 */
function resolveBellScheduleColorFromPSRecord(r) {
  if (!r || typeof r !== 'object') return "";

  const insessionRaw = r.insession !== undefined ? r.insession : r.INSESSION;
  if (insessionRaw !== undefined && insessionRaw !== null && String(insessionRaw).trim() !== "") {
    if (Number(insessionRaw) === 0) {
      return "Red";
    }
  }

  const candidates = [
    r.cycle_day_name || r.CYCLE_DAY_NAME || "",
    r.bell_schedule_name || r.BELL_SCHEDULE_NAME || "",
    r.day_note || r.DAY_NOTE || ""
  ];

  for (let i = 0; i < candidates.length; i++) {
    const text = String(candidates[i] || "").trim();
    if (!text) continue;

    if (typeof normalizeBellScheduleColor === "function") {
      const exact = normalizeBellScheduleColor(text);
      if (exact === "Green" || exact === "Blue" || exact === "Gold" || exact === "Red") {
        return exact;
      }
    }

    const lower = text.toLowerCase();
    if (lower.indexOf("green") !== -1) return "Green";
    if (lower.indexOf("blue") !== -1) return "Blue";
    if (lower.indexOf("gold") !== -1) return "Gold";
    if (lower.indexOf("no school") !== -1 || lower.indexOf("holiday") !== -1 || lower.indexOf("break") !== -1 || lower.indexOf("red") !== -1) {
      return "Red";
    }
  }

  const letter = String(r.cycle_day_letter || r.CYCLE_DAY_LETTER || "").trim().toUpperCase();
  if (letter === "G") return "Green";
  if (letter === "B") return "Blue";
  if (letter === "AU" || letter === "Y") return "Gold";
  if (letter === "R" || letter === "N") return "Red";

  return "";
}

/**
 * Fetches the Bell Schedule / Calendar days from PowerSchool using the
 * `com.cathedral.subapp.bellschedule` PowerQuery endpoint.
 *
 * @return {{success: boolean, updates: Array<{date: string, color: string}>, rawCount?: number, termId?: string, error?: string}}
 */
function fetchBellScheduleFromAPI() {
  let token;
  try {
    token = getPowerSchoolToken();
  } catch (err) {
    return { success: false, error: err.message, updates: [] };
  }
  if (!token) {
    return { success: false, error: "Failed to authenticate with PowerSchool API.", updates: [] };
  }

  const settings = typeof getSettings === "function" ? getSettings() : {};
  const rawUrl = settings['PS_URL'];
  const POWERSCHOOL_URL = rawUrl ? rawUrl.trim().replace(/\/$/, '') : null;
  if (!POWERSCHOOL_URL) {
    return { success: false, error: "Missing PS_URL in Settings.", updates: [] };
  }

  const termId = String(settings["Term ID"] || "3601").trim();
  const endpoint = "/ws/schema/query/com.cathedral.subapp.bellschedule?pagesize=0";

  const options = {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    payload: JSON.stringify({ "target_term": termId }),
    muteHttpExceptions: true
  };

  let responseText;
  let statusCode;
  try {
    const response = UrlFetchApp.fetch(POWERSCHOOL_URL + endpoint, options);
    statusCode = response.getResponseCode();
    responseText = response.getContentText();
  } catch (error) {
    Logger.log("API Fetch Error in fetchBellScheduleFromAPI: " + error.toString());
    return { success: false, error: "PowerSchool request failed: " + error.toString(), updates: [] };
  }

  if (statusCode !== 200) {
    Logger.log("PowerSchool Bell Schedule API returned status: " + statusCode + ", " + responseText);
    return {
      success: false,
      error: "PowerSchool API returned HTTP " + statusCode + ". Ensure the Cathedral Sub App plugin v2.0.0 (com.cathedral.subapp.bellschedule) is installed and enabled.",
      updates: []
    };
  }

  try {
    const json = JSON.parse(responseText);
    const records = json.record || [];
    const updates = [];
    const dateBellScheduleNames = {};

    records.forEach(r => {
      const rawDate = String(r.calendar_date || r.CALENDAR_DATE || r.date_value || r.DATE_VALUE || "").trim();
      const stdDate = typeof getStandardizedDateString === "function" ? getStandardizedDateString(rawDate) : rawDate;
      const color = resolveBellScheduleColorFromPSRecord(r);
      const bsName = String(r.bell_schedule_name || r.BELL_SCHEDULE_NAME || "").trim();
      if (stdDate && color) {
        updates.push({ date: stdDate, color: color, scheduleName: bsName });
        if (bsName) {
          dateBellScheduleNames[stdDate] = bsName;
        }
      }
    });

    return {
      success: true,
      updates: updates,
      dateBellScheduleNames: dateBellScheduleNames,
      rawCount: records.length,
      termId: termId
    };
  } catch (parseError) {
    Logger.log("Failed to parse Bell Schedule JSON: " + parseError.toString());
    return { success: false, error: "Failed to parse PowerSchool Bell Schedule JSON: " + parseError.toString(), updates: [] };
  }
}

/**
 * Converts PowerSchool BELL_SCHEDULE_ITEMS start/end time (seconds or minutes from midnight)
 * into a 24-hour "HH:MM" string.
 */
function formatPSMinutesOrSecondsToHHMM(rawVal) {
  if (rawVal === undefined || rawVal === null || String(rawVal).trim() === "") return "";
  const num = Number(rawVal);
  if (isNaN(num) || num < 0) return "";
  // PowerSchool stores BELL_SCHEDULE_ITEMS.START_TIME / END_TIME in seconds from midnight (0..86400)
  // or minutes (0..1440) depending on view; > 1440 is unambiguously seconds.
  const totalMinutes = num > 1440 ? Math.round(num / 60) : Math.round(num);
  if (totalMinutes < 0 || totalMinutes >= 1440) return "";
  const hh = Math.floor(totalMinutes / 60);
  const mm = totalMinutes % 60;
  return (hh < 10 ? "0" + hh : String(hh)) + ":" + (mm < 10 ? "0" + mm : String(mm));
}

/**
 * Fetches exact period start/end times per Bell Schedule from PowerSchool
 * (`com.cathedral.subapp.bellscheduletimes`).
 */
function fetchBellScheduleTimesFromAPI() {
  let token;
  try {
    token = getPowerSchoolToken();
  } catch (err) {
    return { success: false, error: err.message };
  }
  if (!token) return { success: false, error: "Missing PowerSchool token." };

  const settings = typeof getSettings === "function" ? getSettings() : {};
  const rawUrl = settings['PS_URL'];
  const POWERSCHOOL_URL = rawUrl ? rawUrl.trim().replace(/\/$/, '') : null;
  if (!POWERSCHOOL_URL) return { success: false, error: "Missing PS_URL." };

  const termId = String(settings["Term ID"] || "3601").trim();
  const endpoint = "/ws/schema/query/com.cathedral.subapp.bellscheduletimes?pagesize=0";

  const options = {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    payload: JSON.stringify({ "target_term": termId }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(POWERSCHOOL_URL + endpoint, options);
    if (response.getResponseCode() !== 200) {
      return { success: false, error: "HTTP " + response.getResponseCode() };
    }
    const json = JSON.parse(response.getContentText());
    const records = json.record || [];
    if (records.length === 0) {
      return { success: false, error: "No bell schedule time records returned." };
    }

    const baseMap = JSON.parse(JSON.stringify(typeof DEFAULT_PERIOD_BELL_TIMES !== "undefined" ? DEFAULT_PERIOD_BELL_TIMES : {}));
    records.forEach(r => {
      const bsName = String(r.bell_schedule_name || r.BELL_SCHEDULE_NAME || "").trim();
      const abbr = String(r.period_abbr || r.PERIOD_ABBR || r.period_number || r.PERIOD_NUMBER || "").trim();
      if (!bsName || !abbr) return;

      const normP = typeof parsePeriod === "function" ? parsePeriod(abbr) : abbr;
      const startHHMM = formatPSMinutesOrSecondsToHHMM(r.start_time !== undefined ? r.start_time : r.START_TIME);
      const endHHMM = formatPSMinutesOrSecondsToHHMM(r.end_time !== undefined ? r.end_time : r.END_TIME);
      if (!normP || !startHHMM || !endHHMM) return;

      if (!baseMap[bsName]) baseMap[bsName] = {};
      baseMap[bsName][normP] = { start: startHHMM, end: endHHMM };
      if (normP === "0") baseMap[bsName]["9"] = { start: startHHMM, end: endHHMM };
      if (normP === "9") baseMap[bsName]["0"] = { start: startHHMM, end: endHHMM };

      // Also update canonical Green / Blue / Gold schedule defaults if this schedule is a regular Green/Blue/Gold schedule
      const lowerBs = bsName.toLowerCase();
      ["Green", "Blue", "Gold"].forEach(colorKey => {
        if (lowerBs === colorKey.toLowerCase() || lowerBs === colorKey.toLowerCase() + " day" || lowerBs === colorKey.toLowerCase() + " regular") {
          if (!baseMap[colorKey]) baseMap[colorKey] = {};
          baseMap[colorKey][normP] = { start: startHHMM, end: endHHMM };
        }
      });
    });

    return { success: true, periodBellTimes: baseMap, count: records.length };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Pulls the Bell Schedule from PowerSchool (`com.cathedral.subapp.bellschedule`)
 * and upserts the mapped dates into the `Dates` sheet.
 *
 * @param {boolean} replaceAll - Whether to replace all existing dates in the `Dates` sheet.
 * @param {string} [clientEmail] - Active user email when invoked from the Settings UI.
 */
function syncBellScheduleFromPowerSchool(replaceAll, clientEmail) {
  try {
    if (clientEmail !== undefined) {
      const ss = getSS();
      const user = getUserData(ss, clientEmail);
      assertPermission(user, "Settings");
    }

    const apiRes = fetchBellScheduleFromAPI();
    if (!apiRes.success) {
      return { success: false, error: apiRes.error || "Failed to fetch bell schedule from PowerSchool." };
    }

    if (!apiRes.updates || apiRes.updates.length === 0) {
      return {
        success: false,
        error: "PowerSchool returned " + (apiRes.rawCount || 0) + " calendar records, but 0 dates mapped to Green/Blue/Gold/Red for Term ID " + (apiRes.termId || "") + "."
      };
    }

    const upsertRes = bulkUpsertBellScheduleDates(apiRes.updates, Boolean(replaceAll), clientEmail);
    if (!upsertRes || !upsertRes.success) {
      return upsertRes || { success: false, error: "Failed to save synced bell schedule dates." };
    }

    // Opportunistically sync exact period bell times if `com.cathedral.subapp.bellscheduletimes` is available
    try {
      const timesRes = fetchBellScheduleTimesFromAPI();
      if (timesRes && timesRes.success && timesRes.periodBellTimes && typeof updateSettings === "function") {
        const settingsPatch = {
          "Period Bell Times": JSON.stringify(timesRes.periodBellTimes)
        };
        if (apiRes.dateBellScheduleNames && Object.keys(apiRes.dateBellScheduleNames).length > 0) {
          settingsPatch["Date Bell Schedule Names"] = JSON.stringify(apiRes.dateBellScheduleNames);
        }
        updateSettings(settingsPatch, clientEmail);
      }
    } catch (timesErr) {
      Logger.log("Optional bell schedule period times sync skipped: " + timesErr.toString());
    }

    if (typeof logAuditAction === "function") {
      logAuditAction(
        "BELL_SCHEDULE_PS_SYNC",
        "Dates",
        (replaceAll ? "Replaced" : "Merged") + " " + upsertRes.updated + " dates from PowerSchool (Term " + apiRes.termId + ").",
        clientEmail
      );
    }

    return {
      success: true,
      updated: upsertRes.updated,
      total: upsertRes.total,
      rawCount: apiRes.rawCount,
      termId: apiRes.termId,
      dateColors: upsertRes.dateColors
    };
  } catch (e) {
    Logger.log("Error in syncBellScheduleFromPowerSchool: " + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Synchronizes the Master Schedule teachers and their teaching periods (`TEACHING_PERIODS` in Col K
 * and `SCHEDULE_BY_TERM` JSON in Col L) to the Staff Roster.
 * Automatically updates teaching periods whenever the Master Schedule refreshes (e.g. overnight or across semesters)
 * and adds any missing teachers with default role "Teacher" and blank duty.
 */
function syncMasterScheduleToStaffRoster(scheduleData, ss) {
  if (!scheduleData || scheduleData.length <= 1) return;

  const rosterSheet = ss.getSheetByName("Staff Roster");
  if (!rosterSheet) return;

  const rosterData = rosterSheet.getDataRange().getValues();
  const existingEmails = {};

  const settings = (typeof getSettings === "function") ? getSettings(ss) : {};
  const activeTermId = String(settings["Term ID"] || "3601").trim();

  const sortPeriods = function(arr) {
    arr.sort(function(a, b) {
      const numA = parseInt(a, 10);
      const numB = parseInt(b, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      if (a < b) return -1;
      if (a > b) return 1;
      return 0;
    });
    return arr;
  };

  // Roster format: Name (0), Email (1), Role (2), Duty (3), Dept (4), Homebase (5), Building (6), ID (7), TeacherNum (8), Phone (9), TEACHING_PERIODS (10), SCHEDULE_BY_TERM (11)
  for (let i = 1; i < rosterData.length; i++) {
    if (!rosterData[i]) continue;
    const email = String(rosterData[i][1] || "").toLowerCase().trim();
    if (email) {
      existingEmails[email] = i;
    }
  }

  const newTeachers = {};
  const allSchedules = {};
  const allSchedulesByTerm = {};

  const headers = scheduleData[0];
  const nameIdx = headers.indexOf("LASTFIRST");
  const emailIdx = headers.indexOf("EMAIL_ADDR");
  const periodIdx = headers.indexOf("PERIOD");
  const termIdx = headers.indexOf("TERM");

  if (nameIdx === -1 || emailIdx === -1) return;

  for (let s = 1; s < scheduleData.length; s++) {
    const row = scheduleData[s];
    if (!row) continue;
    const teacherName = String(row[nameIdx] || "").trim();
    const rawEmail = String(row[emailIdx] || "").trim();
    const teacherEmail = rawEmail.toLowerCase();
    if (!teacherEmail) continue;

    if (teacherName && existingEmails[teacherEmail] === undefined && !newTeachers[teacherEmail]) {
      newTeachers[teacherEmail] = {
        name: teacherName,
        email: rawEmail
      };
    }

    if (periodIdx !== -1 && row[periodIdx] !== undefined) {
      const pVal = String(row[periodIdx] || "").trim();
      const joinP = (typeof getScheduleJoinPeriod === "function") ? getScheduleJoinPeriod(pVal) : pVal;
      const rowTerm = (termIdx !== -1 && row[termIdx] !== undefined && String(row[termIdx]).trim() !== "")
        ? String(row[termIdx]).trim()
        : activeTermId;
      if (joinP) {
        if (!allSchedules[teacherEmail]) allSchedules[teacherEmail] = [];
        if ((!rowTerm || rowTerm === activeTermId || Math.floor(Number(rowTerm) / 100) * 100 === Number(rowTerm)) && allSchedules[teacherEmail].indexOf(joinP) === -1) {
          allSchedules[teacherEmail].push(joinP);
        }
        if (rowTerm) {
          if (!allSchedulesByTerm[teacherEmail]) allSchedulesByTerm[teacherEmail] = {};
          if (!allSchedulesByTerm[teacherEmail][rowTerm]) allSchedulesByTerm[teacherEmail][rowTerm] = [];
          if (allSchedulesByTerm[teacherEmail][rowTerm].indexOf(joinP) === -1) {
            allSchedulesByTerm[teacherEmail][rowTerm].push(joinP);
          }
        }
      }
    }
  }

  for (const emailKey in allSchedulesByTerm) {
    if (!allSchedules[emailKey] || allSchedules[emailKey].length === 0) {
      allSchedules[emailKey] = [];
      for (const tk in allSchedulesByTerm[emailKey]) {
        for (let ti = 0; ti < allSchedulesByTerm[emailKey][tk].length; ti++) {
          const pItem = allSchedulesByTerm[emailKey][tk][ti];
          if (allSchedules[emailKey].indexOf(pItem) === -1) {
            allSchedules[emailKey].push(pItem);
          }
        }
      }
    }
    for (const termKey in allSchedulesByTerm[emailKey]) {
      sortPeriods(allSchedulesByTerm[emailKey][termKey]);
    }
  }
  for (const emailKey in allSchedules) {
    sortPeriods(allSchedules[emailKey]);
  }

  // Update Col K (11: TEACHING_PERIODS) and Col L (12: SCHEDULE_BY_TERM) for all existing roster rows
  if (rosterData.length > 0) {
    const klBlock = [["TEACHING_PERIODS", "SCHEDULE_BY_TERM"]];
    for (let r = 1; r < rosterData.length; r++) {
      const rEmail = rosterData[r] ? String(rosterData[r][1] || "").toLowerCase().trim() : "";
      const pList = (rEmail && allSchedules[rEmail]) ? allSchedules[rEmail] : [];
      const byTerm = (rEmail && allSchedulesByTerm[rEmail]) ? allSchedulesByTerm[rEmail] : {};
      // Prefix TEACHING_PERIODS with "'" so Google Sheets never parses 3 periods (e.g. "1, 2, 3") as a Date ("1, 2, 2003")
      klBlock.push([
        pList.length > 0 ? ("'" + pList.join(", ")) : "",
        Object.keys(byTerm).length > 0 ? JSON.stringify(byTerm) : ""
      ]);
    }
    try {
      const klRange = rosterSheet.getRange(1, 11, klBlock.length, 2);
      if (typeof klRange.setNumberFormat === "function") {
        klRange.setNumberFormat("@");
      }
      klRange.setValues(klBlock);
    } catch (klErr) {
      Logger.log("Could not update Staff Roster K:L schedule columns: " + klErr.message);
    }
  }

  const newRows = [];
  const addedNames = [];
  let addedCount = 0;
  for (const key in newTeachers) {
    const teacher = newTeachers[key];
    const pList = allSchedules[key] || [];
    const byTerm = allSchedulesByTerm[key] || {};
    newRows.push([
      teacher.name,
      teacher.email,
      "Teacher",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      pList.length > 0 ? ("'" + pList.join(", ")) : "",
      Object.keys(byTerm).length > 0 ? JSON.stringify(byTerm) : ""
    ]);
    addedNames.push(teacher.name + " <" + teacher.email + ">");
    addedCount++;
  }

  if (newRows.length > 0) {
    const startRow = rosterSheet.getLastRow() + 1;
    rosterSheet.getRange(startRow, 1, newRows.length, 12).setValues(newRows);
    Logger.log("Successfully synced " + addedCount + " new teachers to Staff Roster.");

    try {
      if (typeof logAuditAction === 'function') {
        logAuditAction("STAFF_ADDED_SYNC", "Multiple (" + addedCount + ")", "Synced new teachers from Master Schedule: " + addedNames.join(", "));
      }
    } catch (e) {
      Logger.log("Failed to log audit action for staff sync: " + e.message);
    }
  }

  // Always invalidate roster cache after syncing schedules or new teachers
  try {
    if (typeof clearRosterCache === 'function') {
      clearRosterCache();
    }
  } catch (e) {
    /* ignored */
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
