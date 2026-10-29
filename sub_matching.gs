// ============================================================================
// SUB_MATCHING.GS — Substitute Availability & Auto-Suggestion Engine
// ============================================================================

function getSubstituteAvailability(email) {
  var ss = getSS();
  var subAvailSheet = ss.getSheetByName("SubstituteAvailability");
  if (!subAvailSheet) return {};

  var data = subAvailSheet.getDataRange().getValues();
  var availability = {};
  var targetEmail = String(email).toLowerCase();

  for (var i = 1; i < data.length; i++) {
  if (!data[i]) continue;
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
  if (!data[i]) continue;
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
  if (!data[i]) continue;
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
function getSubstituteEmailByName(subName) {
  var ss = getSS();
  var rosterData = getRosterDataCached(ss);
  if (!rosterData) return null;

  for (var i = 1; i < rosterData.length; i++) {
    if (!rosterData[i]) continue;
    if (String(rosterData[i][0]).trim().toLowerCase() === subName.trim().toLowerCase()) {
      return String(rosterData[i][1]).trim().toLowerCase();
    }
  }
  return null;
}

function getSubstituteAvailabilityForAdmin(subName, clientEmail) {
  var user = getUserData(null, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
  assertPermission(user, "Add Request on Behalf");

  var targetEmail = getSubstituteEmailByName(subName);
  if (!targetEmail) {
    throw new Error("Could not find substitute email for name: " + subName);
  }

  return getSubstituteAvailability(targetEmail);
}

function saveSubstituteAvailabilityAdmin(subName, dateStr, status, clientEmail) {
  var user = getUserData(null, typeof clientEmail !== 'undefined' ? clientEmail : undefined);
  assertPermission(user, "Add Request on Behalf");

  var targetEmail = getSubstituteEmailByName(subName);
  if (!targetEmail) {
    throw new Error("Could not find substitute email for name: " + subName);
  }

  var ss = getSS();
  var subAvailSheet = ss.getSheetByName("SubstituteAvailability");

  if (!subAvailSheet) {
    throw new Error("SubstituteAvailability sheet not found.");
  }

  var data = subAvailSheet.getDataRange().getValues();
  var matchingRows = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i]) continue;
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

function getSuggestedSubs(dateStr, clientEmail) {
  var ss = getSS();
  var email = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
  var user = getUserData(ss, email);
  var roleLower = user && user.role ? String(user.role).toLowerCase().trim() : "";
  var settings = getSettings(ss);
  var rolePermissionsStr = settings["RolePermissions"] || "{}";
  var isAuthorized = user && (
    roleLower === 'admin' ||
    roleLower === 'sub coordinator' ||
    (user.permissions && (user.permissions["Admin Dashboard"] === true || user.permissions["Today at a Glance"] === true)) ||
    hasPermission(roleLower, "Admin Dashboard", rolePermissionsStr) ||
    hasPermission(roleLower, "Today at a Glance", rolePermissionsStr)
  );
  if (!isAuthorized) {
    return { success: false, error: "Insufficient permissions to view substitute suggestions." };
  }

  // 1. Get all requests for the date
  var requests = fetchAbsenceRequestsForDate(ss, dateStr);

  // 2. Get Roster Data
  var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
  if (!rosterSheet) return { success: false, error: "Staff Roster sheet not found." };
  var rosterData = rosterSheet.getDataRange().getValues();

  // 3. Build a map of teachers by email for quick lookup
  var staffMap = {};
  var staffList = []; // Array to preserve some order if needed

  for (var r = 1; r < rosterData.length; r++) {
    var row = rosterData[r];
    if (!row || !row[1]) continue;

    var email = String(row[1]).toLowerCase().trim();
    var staff = {
      name: String(row[0] || "").trim(),
      email: email,
      role: String(row[2] || "").trim().toLowerCase(),
      duty: String(row[3] || "").trim(),
      department: String(row[4] || "").trim(),
      homebase: String(row[5] || "").trim(),
      building: String(row[6] || "").trim(),
      absentPeriods: {} // Track which periods they are absent themselves
    };
    staffMap[email] = staff;
    staffList.push(staff);
  }

  // 4. Determine which teachers are absent on this day (can't sub)
  // And which requests need subs
  // Track who is already assigned for a specific period to avoid double booking
  var openRequests = [];
  var assignedByPeriod = {};

  for (var i = 0; i < requests.length; i++) {
    var req = requests[i];

    // Mark the absent teacher as unavailable for this period
    var absentEmail = String(req.email).toLowerCase().trim();
    if (staffMap[absentEmail]) {
       staffMap[absentEmail].absentPeriods[req.period] = true;
    }

    if (req.status === 'Active' && (!req.assignedSub || req.assignedSub === "")) {
       // Needs a sub
       openRequests.push(req);
    } else if (req.status === 'Active' && req.assignedSub && req.assignedSub !== "") {
       // Someone is already assigned to this request
       var cleanSubName = String(req.assignedSub).replace(/\s+-\s+.*$/, '').trim().toLowerCase();
       if (!assignedByPeriod[req.period]) {
           assignedByPeriod[req.period] = {};
       }
       assignedByPeriod[req.period][cleanSubName] = true;
    }
  }

  // Track who we've already suggested for a specific period to avoid double booking
  var suggestedByPeriod = {};

  // 5. Suggest subs
  for (i = 0; i < openRequests.length; i++) {
    req = openRequests[i];
    var period = String(req.period);
    absentEmail = String(req.email).toLowerCase().trim();
    var absentTeacher = staffMap[absentEmail] || null;

    if (!suggestedByPeriod[period]) {
        suggestedByPeriod[period] = {};
    }

    var bestMatch = null;
    var bestScore = -1;

    for (var j = 0; j < staffList.length; j++) {
      var candidate = staffList[j];

      // Skip if it's the absent teacher
      if (candidate.email === absentEmail) continue;

      // Skip if candidate is absent this period
      if (candidate.absentPeriods[period]) continue;

      // Skip if candidate is already suggested for this period
      if (suggestedByPeriod[period][candidate.email]) continue;

      // Skip if candidate is already assigned for this period
      var candidateCleanName = String(candidate.name).trim().toLowerCase();
      if (assignedByPeriod[period] && assignedByPeriod[period][candidateCleanName]) continue;

      // Check if candidate has duty this period
      var duties = candidate.duty.split(',').map(function(s) { return parsePeriod(s); });
      if (duties.indexOf(parsePeriod(period)) === -1) continue;

      // Calculate score based on priority:
      // Priority 1: Same Department
      // Priority 2: Building Proximity (Cunningham=1, Loretto=2, Kelly Hall=3, default Kelly)
      // Priority 3: Homebase Proximity (Same floor, closest wing)
      var score = 0;

      if (absentTeacher) {
          // 1. Same Department (+1000)
          if (candidate.department && candidate.department === absentTeacher.department) {
             score += 1000;
          }

          // 2. Building Proximity
          var getBuildingScore = function(b) {
             b = b.toLowerCase();
             if (b === 'cunningham') return 1;
             if (b === 'loretto') return 2;
             return 3; // default Kelly Hall
          };
          var b1 = getBuildingScore(candidate.building);
          var b2 = getBuildingScore(absentTeacher.building);

          // Max difference is 2 (e.g. 1 and 3).
          // Smaller difference means closer.
          var bDiff = Math.abs(b1 - b2);
          // Add score based on closeness (+200 for same building, +100 for 1 apart, +0 for 2 apart)
          if (bDiff === 0) score += 200;
          else if (bDiff === 1) score += 100;

          // 3. Homebase Proximity
          if (candidate.homebase && absentTeacher.homebase) {
             var h1 = String(candidate.homebase);
             var h2 = String(absentTeacher.homebase);
             if (h1.length === 4 && h2.length === 4) {
                 b1 = h1.charAt(0);
                 b2 = h2.charAt(0);
                 var f1 = h1.charAt(1);
                 var f2 = h2.charAt(1);
                 var w1 = parseInt(h1.substring(2));
                 var w2 = parseInt(h2.substring(2));

                 // if same building and same floor
                 if (b1 === b2 && f1 === f2 && !isNaN(w1) && !isNaN(w2)) {
                    var wDiff = Math.abs(w1 - w2);
                    // Add up to 99 points based on how close the wings are
                    // wDiff is typically small (e.g. 0 to 20).
                    // Use a formula that gives higher points for smaller wDiff
                    var wingScore = 99 - Math.min(wDiff, 99);
                    score += wingScore;
                 }
             }
          }
      } else {
          // If absent teacher not found in roster, just give a baseline score
          score = 10;
      }

      if (score > bestScore) {
          bestScore = score;
          bestMatch = candidate;
      }
    }

    if (bestMatch) {
       req.suggestedSub = bestMatch.name + (bestMatch.duty ? " - " + bestMatch.duty : "");
       suggestedByPeriod[period][bestMatch.email] = true;
    }
  }

  // Only return open requests that need a sub or where we have a suggestion,
  // or return all so UI can render them properly?
  // Returning all requests for this date (including assigned) is probably best
  // so the UI can show the full picture.

  return { success: true, suggestions: requests };
}

function fetchAbsenceRequestsForDate(ss, dateStr) {
  var reqSheet = getSheetOrThrow(ss, "Absence Requests");
  if (!reqSheet) return [];

  var data = reqSheet.getDataRange().getValues();
  var requests = [];
  var headers = data[0] || [];
  var colMap = {};
  for (var i = 0; i < headers.length; i++) {
     colMap[String(headers[i]).toLowerCase().trim()] = i;
  }

  var rosterData = getRosterDataCached(ss);
  var nameLookup = buildNameLookup(rosterData || []);

  var idxId = colMap['id'] !== undefined ? colMap['id'] : 0;
  var idxTimestamp = colMap['timestamp'] !== undefined ? colMap['timestamp'] : 1;
  var idxEmail = colMap['email'] !== undefined ? colMap['email'] : (colMap['teacher email'] !== undefined ? colMap['teacher email'] : 2);
  var idxDate = colMap['date'] !== undefined ? colMap['date'] : 3;
  var idxPeriod = colMap['periods'] !== undefined ? colMap['periods'] : (colMap['period'] !== undefined ? colMap['period'] : 4);
  var idxReason = colMap['reason'] !== undefined ? colMap['reason'] : 5;
  var idxDuration = colMap['duration'] !== undefined ? colMap['duration'] : (colMap['absence type'] !== undefined ? colMap['absence type'] : 6);
  var idxUrgency = colMap['urgency'] !== undefined ? colMap['urgency'] : 7;
  var idxNotes = colMap['instructions'] !== undefined ? colMap['instructions'] : (colMap['notes'] !== undefined ? colMap['notes'] : 8);
  var idxStatus = colMap['status'] !== undefined ? colMap['status'] : 19;

  for (i = 1; i < data.length; i++) {
     var row = data[i];
     if (!row || row.length === 0 || !row[idxId]) continue;

     // Note: Date parsing in GS can be tricky. We assume row[idxDate] is a Date object or string
     var reqDateObj;
     if (row[idxDate] instanceof Date) {
        reqDateObj = row[idxDate];
     } else {
        reqDateObj = new Date(row[idxDate]);
     }

     if (isNaN(reqDateObj.getTime())) continue;

     // Format req date to YYYY-MM-DD
     var reqDateStr = Utilities.formatDate(reqDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");

     if (reqDateStr === dateStr) {
        var rowEmail = String(row[idxEmail] || "").trim();
        var teacherName = colMap['teacher name'] !== undefined
          ? String(row[colMap['teacher name']] || "").trim()
          : (nameLookup[rowEmail.toLowerCase()] || rowEmail);
        var periodsRequested = String(row[idxPeriod] || "").split(",").map(function(p) { return parsePeriod(p); });
        var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', '10'];
        var rowStatus = String(row[idxStatus] || "Active").trim();

        for (var pIdx = 0; pIdx < allPeriods.length; pIdx++) {
           var p = allPeriods[pIdx];
           if (periodsRequested.indexOf(String(p)) !== -1) {
              var subColumnIndex = getSubColumnIndex(p) - 1;
              var assignedSub = String(row[subColumnIndex] || "").trim();

              requests.push({
                 id: String(row[idxId]),
                 timestamp: row[idxTimestamp] instanceof Date ? Utilities.formatDate(row[idxTimestamp], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : String(row[idxTimestamp]),
                 teacherName: String(teacherName),
                 email: rowEmail,
                 date: reqDateStr,
                 period: String(p),
                 room: colMap['room'] !== undefined ? String(row[colMap['room']] || "") : "",
                 urgency: String(row[idxUrgency] || ""),
                 type: String(row[idxDuration] || ""),
                 reason: String(row[idxReason] || ""),
                 notes: String(row[idxNotes] || ""),
                 status: rowStatus,
                 assignedSub: assignedSub
              });
           }
        }
     }
  }
  return requests;
}

function batchAssignSubs(assignments, clientEmail) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = getSS();
    var email = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    var user = getUserData(ss, email);

    // Ensure permission using role and configured RolePermissions
    var roleLower = user && user.role ? String(user.role).toLowerCase().trim() : "";
    var settings = getSettings(ss);
    var rolePermissionsStr = settings["RolePermissions"] || "{}";
    var isAuthorized = user && (
      roleLower === 'admin' ||
      roleLower === 'sub coordinator' ||
      (user.permissions && user.permissions["Admin Dashboard"] === true) ||
      hasPermission(roleLower, "Admin Dashboard", rolePermissionsStr)
    );
    if (!isAuthorized) {
        return { success: false, error: "Insufficient permissions to assign substitutes." };
    }

    var reqSheet = getSheetOrThrow(ss, "Absence Requests");
    if (!reqSheet) return { success: false, error: "Absence Requests sheet not found." };
    var data = reqSheet.getDataRange().getValues();

    var rosterSheet = getSheetOrThrow(ss, "Staff Roster");
    if (!rosterSheet) return { success: false, error: "Staff Roster sheet not found." };
    var rosterData = rosterSheet.getDataRange().getValues();

    var nameLookup = buildNameLookup(rosterData);
    var canonicalSubLookup = {};
    for (var r = 1; r < rosterData.length; r++) {
       if (!rosterData[r]) continue;
       var rName = String(rosterData[r][0] || "").trim();
       var rEmail = String(rosterData[r][1] || "").trim().toLowerCase();
       if (rName) {
          canonicalSubLookup[rName.toLowerCase()] = { name: rName, email: rEmail };
       }
    }

    // Extract column indices
    var headers = data[0] || [];
    var colMap = {};
    for (var i = 0; i < headers.length; i++) {
       colMap[String(headers[i]).toLowerCase().trim()] = i;
    }

    var idxId = colMap['id'] !== undefined ? colMap['id'] : 0;
    var idxPeriod = colMap['periods'] !== undefined ? colMap['periods'] : (colMap['period'] !== undefined ? colMap['period'] : 4);
    var idxStatus = colMap['status'] !== undefined ? colMap['status'] : 19;

    // Convert assignments to a map keyed by request ID and normalized period
    var assignmentMap = {};
    for (i = 0; i < assignments.length; i++) {
        var a = assignments[i];
        if (!a || !a.id || a.period === undefined) continue;
        var reqKey = String(a.id);
        var normPeriod = parsePeriod(a.period);
        if (!assignmentMap[reqKey]) {
            assignmentMap[reqKey] = {};
        }
        assignmentMap[reqKey][normPeriod] = a.subName;
    }

    var successCount = 0;
    var errors = [];
    var notifications = [];
    var hasChanges = false;

    // Scan sheet for matching requests
    for (i = 1; i < data.length; i++) {
       var row = data[i];
       if (!row || !row[idxId]) continue;

       var reqId = String(row[idxId]);
       if (!assignmentMap[reqId]) continue;

       var rowStatus = String(row[idxStatus] || "Active").trim();
       if (rowStatus === "Canceled") continue;

       var periodsRequested = String(row[idxPeriod] || "").split(",").map(function(p) { return parsePeriod(p); });
       var periodsToAssign = assignmentMap[reqId];

       for (var targetPeriod in periodsToAssign) {
          if (!Object.prototype.hasOwnProperty.call(periodsToAssign, targetPeriod)) continue;

          var subColIdx = getSubColumnIndex(targetPeriod); // 1-based column (10..19)
          if (subColIdx === -1 || periodsRequested.indexOf(targetPeriod) === -1) {
             errors.push("Period " + targetPeriod + " was not requested for absence " + reqId);
             continue;
          }

          var subNameRaw = String(periodsToAssign[targetPeriod] || "").trim();
          var isNoSubNeeded = (subNameRaw.toLowerCase() === "no sub needed");
          var finalSubName = subNameRaw;
          var subEmail = null;

          if (isNoSubNeeded) {
             finalSubName = "No Sub Needed";
          } else if (subNameRaw !== "") {
             var cleanSubName = subNameRaw.replace(/\s+-\s+.*$/, '').trim();
             var matchedSub = canonicalSubLookup[cleanSubName.toLowerCase()];
             if (!matchedSub) {
                errors.push("Invalid sub name: " + subNameRaw);
                continue;
             }
             finalSubName = matchedSub.name;
             subEmail = matchedSub.email;
          }

          // Check if an existing sub is being cleared or replaced
          var prevSubRaw = String(data[i][subColIdx - 1] || "").trim();
          var prevSubClean = prevSubRaw.replace(/\s+-\s+.*$/, '').trim();
          if (prevSubClean && prevSubClean.toLowerCase() !== "no sub needed" && prevSubClean.toLowerCase() !== finalSubName.toLowerCase()) {
             var prevMatchedSub = canonicalSubLookup[prevSubClean.toLowerCase()];
             if (prevMatchedSub && prevMatchedSub.email) {
                var cancelDetails = getAbsenceDetailsLocal(row, targetPeriod, null, nameLookup);
                if (cancelDetails) {
                   notifications.push({
                      subEmail: prevMatchedSub.email,
                      type: "Canceled",
                      details: cancelDetails
                   });
                }
             }
          }

          // Update only the specific period's substitute column (0-based index is subColIdx - 1)
          data[i][subColIdx - 1] = finalSubName;
          hasChanges = true;

          if (!isNoSubNeeded && finalSubName && subEmail && prevSubClean.toLowerCase() !== finalSubName.toLowerCase()) {
             var details = getAbsenceDetailsLocal(row, targetPeriod, null, nameLookup);
             if (details) {
                notifications.push({
                   subEmail: subEmail,
                   type: "Assigned",
                   details: details
                });
             }
          }

          successCount++;
       }
    }

    // Apply all updates at once with a single setValues call
    if (hasChanges && data.length > 0) {
        reqSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    // Log audit
    logAuditAction("Batch Assign", "Multiple", "Assigned " + successCount + " subs", clientEmail);

    // Send notifications
    for (i = 0; i < notifications.length; i++) {
        var n = notifications[i];
        try {
           sendSubNotification(n.subEmail, n.type || "Assigned", n.details);
        } catch (e) {
           console.error("Failed to send notification to " + n.subEmail, e);
        }
    }

    return { success: true, count: successCount, errors: errors.length > 0 ? errors : null };

  } catch (e) {
    console.error("batchAssignSubs Error: " + e.stack);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
