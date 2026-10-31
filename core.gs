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
const SPREADSHEET_ID = "1PaHptRCCSweEeRpYsuY1nDBTGD9DcGtGte84B96wc3A";
const DEFAULT_APP_URL = "https://script.google.com/a/macros/gocathedral.com/s/AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0/exec";
const MASKED_SECRET_VALUE = "••••••••";

const V2_ABSENCES_HEADERS = [
  "Absence ID",
  "Created At",
  "Date",
  "Teacher Email",
  "Teacher Name",
  "Duration",
  "Reason",
  "Urgency",
  "Status",
  "HR Docs Provided",
  "Submitted By",
  "Sub Plan URL",
  "Instructions"
];

const V2_ASSIGNMENTS_HEADERS = [
  "Assignment ID",
  "Absence ID",
  "Date",
  "Period",
  "Teacher Email",
  "Teacher Name",
  "Course Name",
  "Room",
  "Status",
  "Sub Email",
  "Sub Name",
  "Pay Type",
  "Pay Amount",
  "Feedback Rating",
  "Feedback Note"
];

/**
 * Checks whether the normalized v2 absence tables (`Absences` + `Coverage_Assignments`) exist
 * and have been initialized on the given spreadsheet.
 */
function hasV2AbsenceSheets(ss) {
  if (!ss || typeof ss.getSheetByName !== "function") return false;
  var absSheet = ss.getSheetByName("Absences");
  var asgSheet = ss.getSheetByName("Coverage_Assignments");
  return Boolean(absSheet && asgSheet);
}

/**
 * Ensures the `Absences` and `Coverage_Assignments` sheets exist with headers.
 */
function ensureV2AbsenceSheets(ss) {
  if (!ss) return null;
  var absSheet = ss.getSheetByName("Absences");
  if (!absSheet && typeof ss.insertSheet === "function") {
    absSheet = ss.insertSheet("Absences");
    absSheet.getRange(1, 1, 1, V2_ABSENCES_HEADERS.length).setValues([V2_ABSENCES_HEADERS]);
    absSheet.getRange(1, 1, 1, V2_ABSENCES_HEADERS.length).setFontWeight("bold");
  }
  var asgSheet = ss.getSheetByName("Coverage_Assignments");
  if (!asgSheet && typeof ss.insertSheet === "function") {
    asgSheet = ss.insertSheet("Coverage_Assignments");
    asgSheet.getRange(1, 1, 1, V2_ASSIGNMENTS_HEADERS.length).setValues([V2_ASSIGNMENTS_HEADERS]);
    asgSheet.getRange(1, 1, 1, V2_ASSIGNMENTS_HEADERS.length).setFontWeight("bold");
  }
  return { absSheet: absSheet, asgSheet: asgSheet };
}

/**
 * Builds lookup maps from Staff Roster data for v2 normalization.
 */
function buildV2RosterLookups(rosterData) {
  var emailToName = {};
  var nameToEmail = {};
  var nameToDuty = {};
  var nameToRole = {};
  if (rosterData && Array.isArray(rosterData)) {
    for (var i = 1; i < rosterData.length; i++) {
      var r = rosterData[i];
      if (!r || r.length < 2) continue;
      var name = String(r[0] || "").trim();
      var email = String(r[1] || "").trim().toLowerCase();
      var role = String(r[2] || "").trim();
      var duty = String(r[3] || "").trim();
      if (email && name) {
        emailToName[email] = name;
        nameToEmail[name.toLowerCase()] = email;
      }
      if (name) {
        nameToDuty[name.toLowerCase()] = duty;
        nameToRole[name.toLowerCase()] = role;
      }
    }
  }
  return {
    emailToName: emailToName,
    nameToEmail: nameToEmail,
    nameToDuty: nameToDuty,
    nameToRole: nameToRole
  };
}

/**
 * Converts a single 22-column legacy `Absence Requests` row into one 13-column `Absences` row
 * and N 15-column `Coverage_Assignments` rows.
 */
function convertLegacyAbsenceRowToV2(row, rosterLookups, scheduleLookup, dateColors, settings) {
  if (!row || !row[0]) return null;
  var rLookups = rosterLookups || { emailToName: {}, nameToEmail: {}, nameToDuty: {}, nameToRole: {} };
  var dColors = dateColors || {};
  var cfg = settings || {};

  var absenceId = String(row[0]).trim();
  var createdAt = row[1] || "";
  var teacherEmail = String(row[2] || "").trim().toLowerCase();
  var teacherName = rLookups.emailToName[teacherEmail] || String(row[2] || "").trim();
  var rawDate = row[3];
  var isoDate = getStandardizedDateString(rawDate) || String(rawDate || "").trim();
  if (isoDate.indexOf(",") !== -1) {
    isoDate = isoDate.split(",")[0].trim();
  }

  var rawPeriods = String(row[4] || "").trim();
  var reason = String(row[5] || "").trim();
  var duration = String(row[6] || "Full Day").trim();
  var urgency = String(row[7] || "Standard (Advanced Notice)").trim();
  var rawInstructions = String(row[8] || "").trim();
  var status = String(row[19] || "Active").trim() || "Active";
  var subPlanUrl = extractSubPlanUrl(row);
  var feedbackRaw = extractSubFeedbackRaw(row);

  var hrDocsProvided = "FALSE";
  var cleanInstructions = rawInstructions;
  if (cleanInstructions.indexOf("[HR Docs Provided]") !== -1) {
    hrDocsProvided = "TRUE";
    cleanInstructions = cleanInstructions.replace(/(\[HR Docs Provided\]\s*)+/g, "").trim();
  }

  var submittedBy = teacherEmail;
  var subMatch = cleanInstructions.match(/\(Submitted by ([^)]+)\)/i);
  if (subMatch && subMatch[1]) {
    submittedBy = subMatch[1].trim();
  }

  var v2AbsenceRow = [
    absenceId,
    createdAt,
    isoDate,
    teacherEmail,
    teacherName,
    duration,
    reason,
    urgency,
    status,
    hrDocsProvided,
    submittedBy,
    subPlanUrl,
    cleanInstructions
  ];

  var feedbackByPeriod = {};
  if (feedbackRaw && feedbackRaw.indexOf("[") === 0) {
    try {
      var fbArr = JSON.parse(feedbackRaw);
      if (Array.isArray(fbArr)) {
        for (var f = 0; f < fbArr.length; f++) {
          var fbItem = fbArr[f];
          if (fbItem && fbItem.period !== undefined) {
            var fbPKey = parsePeriod(fbItem.period);
            feedbackByPeriod[fbPKey] = {
              rating: fbItem.rating !== undefined && fbItem.rating !== null ? String(fbItem.rating) : "",
              note: String(fbItem.note || "")
            };
          }
        }
      }
    } catch (e) {}
  }

  var specialRatesMap = {};
  try {
    var specialArr = JSON.parse(cfg["Special Pay Rates"] || "[]");
    if (Array.isArray(specialArr)) {
      for (var sp = 0; sp < specialArr.length; sp++) {
        if (specialArr[sp] && specialArr[sp].date) {
          specialRatesMap[String(specialArr[sp].date).trim()] = parseFloat(specialArr[sp].rate || 0);
        }
      }
    }
  } catch (e) {}

  var greenRate = parseFloat(cfg["Green Day Pay Rate"] || "10") || 10;
  var blueGoldRate = parseFloat(cfg["Blue/Gold Day Pay Rate"] || "20") || 20;

  var periodTokens = rawPeriods
    .split(",")
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s !== "" && s.toLowerCase() !== "none"; });

  var seenPeriods = {};
  var v2AssignmentRows = [];

  for (var pIdx = 0; pIdx < periodTokens.length; pIdx++) {
    var pKey = parsePeriod(periodTokens[pIdx]);
    if (!pKey || seenPeriods[pKey]) continue;
    seenPeriods[pKey] = true;

    var subColOneBased = getSubColumnIndex(pKey);
    var assignedSubRaw = (subColOneBased > 0 && subColOneBased - 1 < row.length)
      ? String(row[subColOneBased - 1] || "").trim()
      : "";

    var schedInfo = (typeof lookupTeacherScheduleInfo === "function" && scheduleLookup)
      ? lookupTeacherScheduleInfo(scheduleLookup, teacherEmail, pKey, isoDate, cfg)
      : null;
    var courseStr = (schedInfo && schedInfo.course) ? String(schedInfo.course) : "No Class Assigned";
    var roomStr = (schedInfo && schedInfo.room) ? String(schedInfo.room) : "No Class Assigned";

    var assignStatus = "Unfilled";
    var subEmail = "";
    var subName = "";
    var payType = "";
    var payAmount = "";

    if (status.toLowerCase() === "canceled" || status.toLowerCase() === "cancelled") {
      assignStatus = "Canceled";
      subName = assignedSubRaw;
      if (subName && subName.toLowerCase() !== "no sub needed") {
        subEmail = rLookups.nameToEmail[subName.toLowerCase()] || "";
      }
    } else if (assignedSubRaw.toLowerCase() === "no sub needed") {
      assignStatus = "No Sub Needed";
      subName = "No Sub Needed";
    } else if (assignedSubRaw !== "") {
      assignStatus = "Assigned";
      subName = assignedSubRaw;
      subEmail = rLookups.nameToEmail[subName.toLowerCase()] || "";

      var subRole = String(rLookups.nameToRole[subName.toLowerCase()] || "").toLowerCase();
      var dutyStr = rLookups.nameToDuty[subName.toLowerCase()] || "";
      var dutyArr = dutyStr ? dutyStr.split(",").map(function(d) { return parsePeriod(d); }) : [];
      var isDuty = dutyArr.indexOf(pKey) !== -1;

      if (subRole.indexOf("substitute") !== -1) {
        payType = "External Sub";
        payAmount = "0";
      } else if (isDuty) {
        payType = "Duty";
        payAmount = "0";
      } else if (specialRatesMap[isoDate] !== undefined) {
        payType = "Special";
        payAmount = String(specialRatesMap[isoDate]);
      } else {
        var dayColor = dColors[isoDate] || "Green";
        payType = dayColor;
        payAmount = (String(dayColor).toLowerCase() === "blue" || String(dayColor).toLowerCase() === "gold")
          ? String(blueGoldRate)
          : String(greenRate);
      }
    }

    var fb = feedbackByPeriod[pKey] || { rating: "", note: "" };

    v2AssignmentRows.push([
      absenceId + "_P" + pKey,
      absenceId,
      isoDate,
      "'" + pKey,
      teacherEmail,
      teacherName,
      courseStr,
      roomStr,
      assignStatus,
      subEmail,
      subName,
      payType,
      payAmount,
      fb.rating,
      fb.note
    ]);
  }

  return {
    absenceRow: v2AbsenceRow,
    assignmentRows: v2AssignmentRows
  };
}

/**
 * Synchronizes one or more legacy `Absence Requests` rows into the v2 `Absences` and `Coverage_Assignments` sheets.
 * Safe no-op if the v2 sheets are not present on `ss`.
 */
function syncV2SheetsForAbsenceRows(ss, legacyRows, optionalScheduleLookup, optionalRosterData, optionalDateColors, optionalSettings) {
  try {
    if (!hasV2AbsenceSheets(ss) || !legacyRows || legacyRows.length === 0) return false;
    var absSheet = ss.getSheetByName("Absences");
    var asgSheet = ss.getSheetByName("Coverage_Assignments");
    if (!absSheet || !asgSheet) return false;

    var rosterData = optionalRosterData || (typeof getRosterDataCached === "function" ? getRosterDataCached(ss) : []);
    var settings = optionalSettings || (typeof getSettings === "function" ? getSettings(ss) : {});
    var dateColors = optionalDateColors || (typeof getDateColorsCached === "function" ? getDateColorsCached(ss) : {});
    var scheduleLookup = optionalScheduleLookup;
    if (!scheduleLookup && typeof getMasterScheduleData === "function" && typeof buildScheduleLookup === "function") {
      var schedData = getMasterScheduleData();
      var activeTerm = String(settings["Term ID"] || "3601").trim();
      scheduleLookup = buildScheduleLookup(schedData, activeTerm);
    }

    var rosterLookups = buildV2RosterLookups(rosterData);

    var targetIds = {};
    var convertedById = {};
    for (var i = 0; i < legacyRows.length; i++) {
      var conv = convertLegacyAbsenceRowToV2(legacyRows[i], rosterLookups, scheduleLookup, dateColors, settings);
      if (conv && conv.absenceRow && conv.absenceRow[0]) {
        var idKey = String(conv.absenceRow[0]);
        targetIds[idKey] = true;
        convertedById[idKey] = conv;
      }
    }

    var existingAbs = absSheet.getDataRange().getValues();
    if (!existingAbs || existingAbs.length === 0) {
      existingAbs = [V2_ABSENCES_HEADERS];
    }
    var updatedAbs = [existingAbs[0]];
    var seenAbsIds = {};
    for (var a = 1; a < existingAbs.length; a++) {
      if (!existingAbs[a] || !existingAbs[a][0]) continue;
      var rowAbsId = String(existingAbs[a][0]);
      if (targetIds[rowAbsId]) {
        if (!seenAbsIds[rowAbsId]) {
          updatedAbs.push(convertedById[rowAbsId].absenceRow);
          seenAbsIds[rowAbsId] = true;
        }
      } else {
        updatedAbs.push(existingAbs[a]);
      }
    }
    for (var newId in convertedById) {
      if (Object.prototype.hasOwnProperty.call(convertedById, newId) && !seenAbsIds[newId]) {
        updatedAbs.push(convertedById[newId].absenceRow);
      }
    }

    var existingAsg = asgSheet.getDataRange().getValues();
    if (!existingAsg || existingAsg.length === 0) {
      existingAsg = [V2_ASSIGNMENTS_HEADERS];
    }
    var updatedAsg = [existingAsg[0]];
    for (var g = 1; g < existingAsg.length; g++) {
      if (!existingAsg[g] || !existingAsg[g][0]) continue;
      var parentAbsId = String(existingAsg[g][1] || "");
      if (!targetIds[parentAbsId]) {
        updatedAsg.push(existingAsg[g]);
      }
    }
    for (var cId in convertedById) {
      if (Object.prototype.hasOwnProperty.call(convertedById, cId)) {
        var asgRows = convertedById[cId].assignmentRows;
        for (var k = 0; k < asgRows.length; k++) {
          updatedAsg.push(asgRows[k]);
        }
      }
    }

    if (typeof absSheet.clearContents === "function") absSheet.clearContents();
    absSheet.getRange(1, 1, updatedAbs.length, V2_ABSENCES_HEADERS.length).setValues(updatedAbs);

    if (typeof asgSheet.clearContents === "function") asgSheet.clearContents();
    asgSheet.getRange(1, 1, updatedAsg.length, V2_ASSIGNMENTS_HEADERS.length).setValues(updatedAsg);
    return true;
  } catch (e) {
    console.warn("syncV2SheetsForAbsenceRows warning: " + e.message);
    return false;
  }
}

/**
 * Rebuilds the entire `Absences` and `Coverage_Assignments` sheets from the full `Absence Requests` table.
 */
function rebuildAllV2AbsenceSheets(ss) {
  ss = ss || getSS();
  var mainSheet = ss.getSheetByName("Absence Requests");
  if (!mainSheet) return { success: false, error: "Absence Requests sheet not found." };

  var sheets = ensureV2AbsenceSheets(ss);
  var absSheet = sheets.absSheet;
  var asgSheet = sheets.asgSheet;

  var legacyData = mainSheet.getDataRange().getValues();
  var rosterData = typeof getRosterDataCached === "function" ? getRosterDataCached(ss) : [];
  var settings = typeof getSettings === "function" ? getSettings(ss) : {};
  var dateColors = typeof getDateColorsCached === "function" ? getDateColorsCached(ss) : {};
  var schedData = typeof getMasterScheduleData === "function" ? getMasterScheduleData() : [];
  var activeTerm = String(settings["Term ID"] || "3601").trim();
  var scheduleLookup = typeof buildScheduleLookup === "function" ? buildScheduleLookup(schedData, activeTerm) : {};
  var rosterLookups = buildV2RosterLookups(rosterData);

  var v2AbsRows = [V2_ABSENCES_HEADERS];
  var v2AsgRows = [V2_ASSIGNMENTS_HEADERS];

  for (var i = 1; i < legacyData.length; i++) {
    if (!legacyData[i] || !legacyData[i][0]) continue;
    var conv = convertLegacyAbsenceRowToV2(legacyData[i], rosterLookups, scheduleLookup, dateColors, settings);
    if (conv) {
      v2AbsRows.push(conv.absenceRow);
      for (var j = 0; j < conv.assignmentRows.length; j++) {
        v2AsgRows.push(conv.assignmentRows[j]);
      }
    }
  }

  if (typeof absSheet.clearContents === "function") absSheet.clearContents();
  absSheet.getRange(1, 1, v2AbsRows.length, V2_ABSENCES_HEADERS.length).setValues(v2AbsRows);

  if (typeof asgSheet.clearContents === "function") asgSheet.clearContents();
  asgSheet.getRange(1, 1, v2AsgRows.length, V2_ASSIGNMENTS_HEADERS.length).setValues(v2AsgRows);

  return {
    success: true,
    absencesCount: v2AbsRows.length - 1,
    assignmentsCount: v2AsgRows.length - 1
  };
}

/**
 * Reads the normalized v2 `Absences` and `Coverage_Assignments` sheets and returns:
 * - `absenceData`: reconstructed 22-column virtual rows compatible with all existing payload/digest/export logic
 * - `snapshottedScheduleLookup`: pre-joined `{ room, course }` map built directly from `Coverage_Assignments`
 *   so `getInitialPayload()` does not need to read `Master Schedule Cache` for every absence row.
 */
function readV2AbsenceDataCompat(ss) {
  if (!hasV2AbsenceSheets(ss)) return null;
  var absSheet = ss.getSheetByName("Absences");
  var asgSheet = ss.getSheetByName("Coverage_Assignments");
  if (!absSheet || !asgSheet) return null;

  var absValues = absSheet.getDataRange().getValues();
  var asgValues = asgSheet.getDataRange().getValues();
  if (!absValues || absValues.length <= 1) return null;

  var assignmentsByAbsenceId = {};
  var snapshottedScheduleLookup = {};

  for (var g = 1; g < asgValues.length; g++) {
    var gRow = asgValues[g];
    if (!gRow || !gRow[1]) continue;
    var parentId = String(gRow[1]).trim();
    if (!assignmentsByAbsenceId[parentId]) {
      assignmentsByAbsenceId[parentId] = [];
    }
    assignmentsByAbsenceId[parentId].push(gRow);

    var gEmail = String(gRow[4] || "").toLowerCase().trim();
    var gPeriod = parsePeriod(String(gRow[3] || "").replace(/^'/, ""));
    var gCourse = String(gRow[6] || "").trim();
    var gRoom = String(gRow[7] || "").trim();
    var gDate = getStandardizedDateString(gRow[2]) || String(gRow[2] || "").trim();
    if (gEmail && gPeriod && (gCourse || gRoom)) {
      var joinKey = gEmail + "-" + getScheduleJoinPeriod(gPeriod);
      var entry = {
        room: gRoom || "No Class Assigned",
        course: gCourse || "No Class Assigned"
      };
      snapshottedScheduleLookup[joinKey] = entry;
      if (gDate && typeof resolveTermIdForDate === "function") {
        var termForDate = String(resolveTermIdForDate(gDate) || "").toLowerCase().trim();
        if (termForDate) {
          snapshottedScheduleLookup[joinKey + "-" + termForDate] = entry;
        }
      }
    }
  }

  var compatRows = [[
    "ID", "Timestamp", "Email", "Date", "Periods", "Reason", "Duration",
    "Urgency", "Instructions", "Period 1 Sub", "Period 2 Sub", "Period 3 Sub",
    "Period 4 Sub", "Period 5 Sub", "Period 6 Sub", "Period 7 Sub", "Period 8 Sub",
    "Period 0 Sub", "Advisory Sub", "Status", "Sub Plan URL", "Sub Feedback"
  ]];

  for (var a = 1; a < absValues.length; a++) {
    var aRow = absValues[a];
    if (!aRow || !aRow[0]) continue;

    var absId = String(aRow[0]).trim();
    var createdAt = aRow[1];
    var dateVal = aRow[2];
    var teacherEmail = String(aRow[3] || "").trim();
    var duration = String(aRow[5] || "").trim();
    var reason = String(aRow[6] || "").trim();
    var urgency = String(aRow[7] || "").trim();
    var status = String(aRow[8] || "Active").trim();
    var hrDocs = String(aRow[9] || "").toLowerCase().trim() === "true";
    var subPlanUrl = String(aRow[11] || "").trim();
    var instructions = String(aRow[12] || "").trim();
    if (hrDocs && instructions.indexOf("[HR Docs Provided]") === -1) {
      instructions = ("[HR Docs Provided] " + instructions).trim();
    }

    var childAsgs = assignmentsByAbsenceId[absId] || [];
    var periodsList = [];
    var subCols = ["", "", "", "", "", "", "", "", "", ""]; // indices 0..9 -> cols 10..19
    var feedbackArr = [];

    for (var c = 0; c < childAsgs.length; c++) {
      var cRow = childAsgs[c];
      var pClean = parsePeriod(String(cRow[3] || "").replace(/^'/, ""));
      if (!pClean) continue;
      periodsList.push(pClean);

      var subColOneBased = getSubColumnIndex(pClean);
      var subNameVal = String(cRow[10] || "").trim();
      var asgStatus = String(cRow[8] || "").trim();
      if (asgStatus === "No Sub Needed" && !subNameVal) {
        subNameVal = "No Sub Needed";
      }
      if (subColOneBased >= 10 && subColOneBased <= 19) {
        subCols[subColOneBased - 10] = subNameVal;
      }

      var fbRating = cRow[13];
      var fbNote = String(cRow[14] || "").trim();
      if ((fbRating !== undefined && fbRating !== null && String(fbRating).trim() !== "") || fbNote !== "") {
        feedbackArr.push({
          subName: subNameVal,
          period: pClean,
          rating: parseInt(fbRating, 10) || 0,
          note: fbNote
        });
      }
    }

    compatRows.push([
      absId,
      createdAt,
      teacherEmail,
      dateVal,
      periodsList.join(", "),
      reason,
      duration,
      urgency,
      instructions,
      subCols[0],
      subCols[1],
      subCols[2],
      subCols[3],
      subCols[4],
      subCols[5],
      subCols[6],
      subCols[7],
      subCols[8],
      subCols[9],
      status,
      subPlanUrl,
      JSON.stringify(feedbackArr)
    ]);
  }

  return {
    absenceData: compatRows,
    snapshottedScheduleLookup: snapshottedScheduleLookup,
    rawAbsences: absValues,
    rawAssignments: asgValues
  };
}

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
 * Normalizes any Date object or date string ("YYYY-MM-DD", "M/D/YYYY", "MMM d, yyyy", etc.)
 * into a standardized "YYYY-MM-DD" string.
 */
function getStandardizedDateString(dateRaw) {
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

/**
 * Normalizes a bell schedule color / day type string to "Green", "Blue", "Gold", "Red", or "".
 */
function normalizeBellScheduleColor(colorRaw) {
  if (!colorRaw) return "";
  var c = String(colorRaw).trim().toLowerCase();
  if (c === "none" || c === "clear" || c === "unset" || c === "") return "";
  if (c === "green" || c === "green day" || c === "g") return "Green";
  if (c === "blue" || c === "blue day" || c === "b") return "Blue";
  if (c === "gold" || c === "gold day" || c === "y") return "Gold";
  if (c === "red" || c === "red day" || c === "no school" || c === "holiday" || c === "r") return "Red";
  var rawTrim = String(colorRaw).trim();
  return rawTrim.charAt(0).toUpperCase() + rawTrim.slice(1);
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


