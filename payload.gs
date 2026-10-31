// ============================================================================
// PAYLOAD.GS — Initial Application Payload & Selective Component Refresh
// ============================================================================

/**
 * Helper to build a name lookup dictionary from Staff Roster data.
 */
function buildNameLookup(rosterData) {
  var nameLookup = {};
  for (var r = 1; r < rosterData.length; r++) {
    if (!rosterData[r]) continue;
    var rosterEmail = String(rosterData[r][1]).toLowerCase().trim();
    nameLookup[rosterEmail] = String(rosterData[r][0]).trim();
  }
  return nameLookup;
}

/**
 * Helper to build a schedule lookup dictionary from Master Schedule data.
 * Stores both term-naive keys (`email-period`, preferring `preferredTermId` when multiple terms exist)
 * and term-specific keys (`email-period-term`) for date-aware Semester 1 / Semester 2 lookups.
 */
function buildScheduleLookup(scheduleData, preferredTermId) {
  var scheduleLookup = {};
  if (scheduleData && scheduleData.length > 0) {
    var headers = scheduleData[0];
    var joinIdx = headers.indexOf("EMAIL_PERIOD_JOIN");
    var roomIdx = headers.indexOf("ROOM");
    var courseIdx = headers.indexOf("COURSE_NAMES");
    var termIdx = headers.indexOf("TERM");
    var prefTerm = preferredTermId ? String(preferredTermId).trim() : "";

    if (joinIdx > -1) {
      for (var s = 1; s < scheduleData.length; s++) {
        if (!scheduleData[s]) continue;
        var joinKey = String(scheduleData[s][joinIdx]).toLowerCase().trim();
        var room = roomIdx > -1 ? scheduleData[s][roomIdx] : "No Class Assigned";
        var course = courseIdx > -1 ? scheduleData[s][courseIdx] : "No Class Assigned";
        var rowTerm = (termIdx > -1 && scheduleData[s][termIdx] !== undefined) ? String(scheduleData[s][termIdx]).trim() : "";
        var entry = {
          room: room ? room : "No Class Assigned",
          course: course ? course : "No Class Assigned"
        };

        if (rowTerm) {
          scheduleLookup[joinKey + "-" + rowTerm.toLowerCase()] = entry;
        }
        // Populate default key if not yet set, or if this row matches preferredTermId
        if (!scheduleLookup[joinKey] || (prefTerm && rowTerm === prefTerm)) {
          scheduleLookup[joinKey] = entry;
        }
      }
    }
  }
  return scheduleLookup;
}

/**
 * Looks up a teacher's room and course for a given period and date, automatically selecting
 * Semester 1 (e.g. 3601) vs. Semester 2 (e.g. 3603) when multi-term data is available.
 */
function lookupTeacherScheduleInfo(scheduleLookup, email, period, dateStr, optionalSettings) {
  if (!scheduleLookup || !email || period === undefined || period === null) return null;
  var joinKey = String(email).toLowerCase().trim() + "-" + getScheduleJoinPeriod(period);

  if (dateStr && typeof resolveTermIdForDate === "function") {
    var targetTerm = String(resolveTermIdForDate(dateStr, optionalSettings) || "").toLowerCase().trim();
    if (targetTerm && scheduleLookup[joinKey + "-" + targetTerm]) {
      return scheduleLookup[joinKey + "-" + targetTerm];
    }
    var termNum = parseInt(targetTerm, 10);
    if (!isNaN(termNum) && termNum >= 100) {
      var fullYearTerm = String(Math.floor(termNum / 100) * 100);
      if (scheduleLookup[joinKey + "-" + fullYearTerm]) {
        return scheduleLookup[joinKey + "-" + fullYearTerm];
      }
    }
  }

  return scheduleLookup[joinKey] || null;
}

function getInitialPayload(clientEmail, requestedComponents) {
  try {
    var ss = getSS();
    var email = getActiveUserEmail(typeof clientEmail !== 'undefined' ? clientEmail : undefined);
    var targetEmail = String(email).toLowerCase();

    var wantAll = !requestedComponents || !Array.isArray(requestedComponents) || requestedComponents.length === 0;
    var wantMap = {};
    if (!wantAll) {
      for (var rcIdx = 0; rcIdx < requestedComponents.length; rcIdx++) {
        wantMap[String(requestedComponents[rcIdx])] = true;
      }
    }
    var want = function(comp) { return wantAll || wantMap[comp] === true; };

    // 1. Fetch required sheets (using cached roster data)
    var rosterData = getRosterDataCached(ss);
    var settings = getSettings(ss);
    var activeTermId = String(settings["Term ID"] || "3601").trim();

    var sortPeriodList = function(arr) {
      arr.sort(function(a, b) {
        var numA = parseInt(a, 10);
        var numB = parseInt(b, 10);
        if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
        if (a < b) return -1;
        if (a > b) return 1;
        return 0;
      });
      return arr;
    };

    // Check if Staff Roster has pre-computed TEACHING_PERIODS (Col K / index 10) or SCHEDULE_BY_TERM (Col L / index 11)
    var rosterHasPrecomputedSchedules = false;
    var rosterSchedulesByEmail = {};
    var rosterSchedulesByTermByEmail = {};
    if (rosterData && rosterData.length > 0) {
      var rHeaders = rosterData[0] || [];
      if (String(rHeaders[10] || "").trim().toUpperCase() === "TEACHING_PERIODS" || String(rHeaders[11] || "").trim().toUpperCase() === "SCHEDULE_BY_TERM") {
        rosterHasPrecomputedSchedules = true;
      }
      for (var rIdx = 1; rIdx < rosterData.length; rIdx++) {
        var rRow = rosterData[rIdx];
        if (!rRow || rRow.length < 2) continue;
        var rEmail = String(rRow[1] || "").toLowerCase().trim();
        if (!rEmail) continue;
        var rawTeachingPeriods = rRow.length > 10 ? String(rRow[10] || "").trim() : "";
        var rawScheduleByTerm = rRow.length > 11 ? String(rRow[11] || "").trim() : "";
        if (rawTeachingPeriods || rawScheduleByTerm) {
          rosterHasPrecomputedSchedules = true;
          var byTermMap = {};
          if (rawScheduleByTerm && rawScheduleByTerm.indexOf("{") === 0) {
            try {
              var parsedMap = JSON.parse(rawScheduleByTerm);
              if (parsedMap && typeof parsedMap === "object") {
                byTermMap = parsedMap;
              }
            } catch (e) {}
          }
          rosterSchedulesByTermByEmail[rEmail] = byTermMap;

          // Resolve active-term schedule first from byTermMap[activeTermId], falling back to TEACHING_PERIODS column
          var activeList = [];
          if (byTermMap[activeTermId] && Array.isArray(byTermMap[activeTermId])) {
            activeList = byTermMap[activeTermId].slice();
          } else if (rawTeachingPeriods) {
            activeList = String(rawTeachingPeriods).replace(/^'/, "").split(",").map(function(p) { return getScheduleJoinPeriod(p.trim()); }).filter(function(p) { return p !== ""; });
          }
          rosterSchedulesByEmail[rEmail] = sortPeriodList(activeList);
        }
      }
    }

    var needAbsenceData = wantAll || want("myAbsences") || want("myPastAbsences") || want("mySubDuties") || want("myPastSubDuties") || want("todaysAbsences") || want("quickCover") || want("adminData") || want("hrData");
    var v2Compat = (needAbsenceData && typeof readV2AbsenceDataCompat === "function") ? readV2AbsenceDataCompat(ss) : null;
    var mainSheet = (needAbsenceData && !v2Compat) ? getSheetOrThrow(ss, "Absence Requests") : null;
    var absenceData = v2Compat ? v2Compat.absenceData : (mainSheet ? mainSheet.getDataRange().getValues() : []);

    var needScheduleData = (!rosterHasPrecomputedSchedules && (wantAll || want("userData") || want("staffList"))) || (!v2Compat && (want("mySubDuties") || want("myPastSubDuties") || want("todaysAbsences") || want("quickCover") || want("adminData") || want("hrData")));
    var scheduleData = needScheduleData ? getMasterScheduleData() : [];

    var needPayPeriods = want("hrData");
    var payPeriodsSheet = needPayPeriods ? (ss.getSheetByName("PayPeriods") || ss.getSheetByName("Payperiods")) : null;
    var payPeriodsData = payPeriodsSheet ? payPeriodsSheet.getDataRange().getValues() : [];

    var needDates = want("dateColors") || want("hrData");
    var dateColors = needDates ? getDateColorsCached(ss) : {};
    var dateBellScheduleNames = (needDates && typeof getDateBellScheduleNamesCached === "function") ? getDateBellScheduleNamesCached(ss) : {};



    // --- Build lookups ---
    var scheduleLookup = buildScheduleLookup(scheduleData, activeTermId);
    if (v2Compat && v2Compat.snapshottedScheduleLookup) {
      for (var snapKey in v2Compat.snapshottedScheduleLookup) {
        if (Object.prototype.hasOwnProperty.call(v2Compat.snapshottedScheduleLookup, snapKey) && !scheduleLookup[snapKey]) {
          scheduleLookup[snapKey] = v2Compat.snapshottedScheduleLookup[snapKey];
        }
      }
    }
    var nameLookup = buildNameLookup(rosterData);


    // --- 2. Extract User Data ---
    var name = null;
    var role = null;
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

    var userName = String(name).trim().toLowerCase();
    var lowerRole = String(role).toLowerCase();

    // Compute teacherSchedule (default active term) and teacherScheduleByTerm (all terms)
    var teacherSchedule = rosterSchedulesByEmail[targetEmail] ? rosterSchedulesByEmail[targetEmail].slice() : [];
    var teacherScheduleByTerm = rosterSchedulesByTermByEmail[targetEmail] ? rosterSchedulesByTermByEmail[targetEmail] : {};
    if (!rosterHasPrecomputedSchedules && scheduleData && scheduleData.length > 0) {
      var headers = scheduleData[0];
      var emailIdx = headers.indexOf("EMAIL_ADDR");
      var periodIdx = headers.indexOf("PERIOD");
      var termColIdx = headers.indexOf("TERM");
      if (emailIdx > -1 && periodIdx > -1) {
        for (var s = 1; s < scheduleData.length; s++) {
          if (!scheduleData[s] || scheduleData[s].length <= Math.max(emailIdx, periodIdx)) continue;
          if (String(scheduleData[s][emailIdx]).toLowerCase().trim() === targetEmail) {
            var pVal = String(scheduleData[s][periodIdx]).trim();
            var joinP = getScheduleJoinPeriod(pVal);
            var rowTermVal = (termColIdx > -1 && scheduleData[s][termColIdx] !== undefined) ? String(scheduleData[s][termColIdx]).trim() : activeTermId;
            if (joinP) {
              if (rowTermVal) {
                if (!teacherScheduleByTerm[rowTermVal]) teacherScheduleByTerm[rowTermVal] = [];
                if (teacherScheduleByTerm[rowTermVal].indexOf(joinP) === -1) {
                  teacherScheduleByTerm[rowTermVal].push(joinP);
                }
              }
              if ((!rowTermVal || rowTermVal === activeTermId || Math.floor(Number(rowTermVal) / 100) * 100 === Number(rowTermVal)) && teacherSchedule.indexOf(joinP) === -1) {
                teacherSchedule.push(joinP);
              }
            }
          }
        }
        // Fallback if no rows matched activeTermId directly
        if (teacherSchedule.length === 0) {
          for (var tk in teacherScheduleByTerm) {
            for (var ti = 0; ti < teacherScheduleByTerm[tk].length; ti++) {
              if (teacherSchedule.indexOf(teacherScheduleByTerm[tk][ti]) === -1) {
                teacherSchedule.push(teacherScheduleByTerm[tk][ti]);
              }
            }
          }
        }
        for (var termKey in teacherScheduleByTerm) {
          sortPeriodList(teacherScheduleByTerm[termKey]);
        }
      }
    }
    sortPeriodList(teacherSchedule);

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
      "Today's Absences": hasPermission(lowerRole, "Today's Absences", rolePermissions),
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
      teacherSchedule: teacherSchedule,
      teacherScheduleByTerm: teacherScheduleByTerm
    };


    // --- 3. Extract common data (My Absences, My Sub Duties, Open Jobs) ---
    var myUpcomingAbsences = [];
    var myPastAbsences = [];
    var mySubDuties = [];
    var myPastSubDuties = [];
    var todaysAbsences = [];

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var targetEndWeek = new Date(today);
    targetEndWeek.setDate(today.getDate() + 6); // Up to next week
    targetEndWeek.setHours(23, 59, 59, 999);

    var targetEndToday = new Date(today);
    targetEndToday.setHours(23, 59, 59, 999);

    var fetchWindowDays = parseInt(settings["Data Fetch Window (Days)"]);
    if (isNaN(fetchWindowDays)) fetchWindowDays = 30; // default to 30 days
    var cutoffDate = new Date(today);
    cutoffDate.setDate(today.getDate() - fetchWindowDays);
    cutoffDate.setHours(0, 0, 0, 0);

    var tz = Session.getScriptTimeZone();
    var formattedDateCache = {};
    var yyyymmddCache = {};
    var getCachedDateFormats = function(rawVal, dateObj, isValid) {
      var cacheKey = (rawVal instanceof Date) ? String(dateObj.getTime()) : String(rawVal);
      if (formattedDateCache[cacheKey] !== undefined) {
        return { formattedDate: formattedDateCache[cacheKey], yyyymmdd: yyyymmddCache[cacheKey] };
      }
      var fDate = "Unknown Date";
      var ymd = "";
      if (isValid) {
        fDate = Utilities.formatDate(dateObj, tz, "MMM d, yyyy");
        ymd = Utilities.formatDate(dateObj, tz, "yyyy-MM-dd");
      } else {
        try {
          var fallbackObj = new Date(rawVal);
          fDate = fallbackObj.toLocaleDateString();
          ymd = Utilities.formatDate(fallbackObj, tz, "yyyy-MM-dd");
        } catch (e) {
          ymd = String(rawVal);
        }
      }
      formattedDateCache[cacheKey] = fDate;
      yyyymmddCache[cacheKey] = ymd;
      return { formattedDate: fDate, yyyymmdd: ymd };
    };

    var formatFirstLastTeacherName = function(emailKey) {
      var tName = nameLookup[emailKey] || emailKey;
      if (tName.includes(",")) {
        var parts = tName.split(",");
        if (parts.length > 1) tName = parts[1].trim() + " " + parts[0].trim();
      }
      return tName;
    };

    var wantQuickCover = want("quickCover") && (permissions["Admin Dashboard"] || permissions["Today at a Glance"] || permissions["Today's Absences"]);
    var wantAdminData = want("adminData") && (permissions["Admin Dashboard"] || permissions["Today at a Glance"] || permissions["Today's Absences"]);
    var wantHrData = want("hrData") && permissions["HR Dashboard"];

    var quickCover = [];
    var targetEndQC = new Date(today);
    if (wantQuickCover) {
      var dayOfWeek = today.getDay();
      var daysToAdd = 1;
      if (dayOfWeek === 5) daysToAdd = 3;
      else if (dayOfWeek === 6) daysToAdd = 2;
      targetEndQC.setDate(today.getDate() + daysToAdd);
      targetEndQC.setHours(23, 59, 59, 999);
    }

    var adminData = [];
    var hrData = [];

    for (let i = 1; i < absenceData.length; i++) {
      var row = absenceData[i];
      if (!row || row.length < 4) continue;
      var status = String(row[19] || 'Active').trim();
      if (status === 'Canceled') continue;

      var rowTeacherEmail = String(row[2] || "").toLowerCase().trim();
      var dateVal = row[3];
      if (!dateVal) continue;

      var rowDate = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
      var isDateValid = !isNaN(rowDate.getTime());
      var isWithinWindow = !isDateValid || rowDate >= cutoffDate;

      var dateFmts = getCachedDateFormats(dateVal, rowDate, isDateValid);
      var formattedDate = dateFmts.formattedDate;
      var yyyymmdd = dateFmts.yyyymmdd;
      var rawDate = isDateValid ? Number(rowDate.getTime()) : 0;

      var rowId = String(row[0] || "");
      var periodsStrRaw = String(row[4] || "");
      var reason = String(row[5] || "");
      var duration = String(row[6] || "");
      var urgencyStr = String(row[7] || "");
      var instructions = String(row[8] || "");
      var rowSubPlanUrl = extractSubPlanUrl(row);

      // 3a. My Absences
      if (rowTeacherEmail === targetEmail) {
        let subFeedbackRaw = extractSubFeedbackRaw(row);
        var periodsRequestedForMe = periodsStrRaw.split(",").map(function(p) { return p.trim(); }).filter(function(p) { return p !== ""; });
        var periodsWithSubs = periodsRequestedForMe.map(function(p) {
          var subColIdx = getSubColumnIndex(p);
          var subName = (subColIdx > 0 && subColIdx <= row.length) ? row[subColIdx - 1] : "";
          var displayP = formatPeriodDisplay(p);
          if (subName && String(subName).trim() !== "") {
            return displayP + " (" + String(subName).trim() + ")";
          } else {
            return displayP + " (Not Yet Assigned)";
          }
        }).join(", ");

        var absenceObj = {
          id: rowId,
          date: String(formattedDate),
          rawDate: rawDate,
          rawDateString: String(dateVal),
          formDateString: String(yyyymmdd),
          periods: periodsStrRaw,
          periodsWithSubs: periodsWithSubs,
          reason: reason,
          urgency: urgencyStr.includes('Urgent') ? 'Urgent' : 'Standard',
          duration: duration,
          instructions: instructions,
          subPlanUrl: rowSubPlanUrl,
          subFeedback: subFeedbackRaw
        };
        if (isDateValid && rowDate < today) {
          myPastAbsences.push(absenceObj);
        } else {
          myUpcomingAbsences.push(absenceObj);
        }
      }

      // 3b. My Sub Duties & Today's Absences
      var isWithinSubWeek = isDateValid && rowDate <= targetEndWeek;
      var isWithinQC = wantQuickCover && isDateValid && isWithinWindow && rowDate >= today && rowDate <= targetEndQC;

      if (isWithinSubWeek || isWithinQC) {
        var displayTeacherName = formatFirstLastTeacherName(rowTeacherEmail);
        var normPeriodsRequested = periodsStrRaw.split(",").map(function(p) { return parsePeriod(p); });
        var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', '10'];

        for (var pIdx = 0; pIdx < allPeriods.length; pIdx++) {
          var p = allPeriods[pIdx];
          if (normPeriodsRequested.indexOf(String(p)) === -1) continue;

          var subColumnIndex = getSubColumnIndex(p) - 1;
          var rawAssignedSub = String(row[subColumnIndex] || "").trim();
          var assignedSubLower = rawAssignedSub.toLowerCase();

          var scheduleInfo = lookupTeacherScheduleInfo(scheduleLookup, rowTeacherEmail, p, yyyymmdd, settings);
          var roomStr = scheduleInfo && scheduleInfo.room ? String(scheduleInfo.room) : "No Class Assigned";
          var courseStr = scheduleInfo && scheduleInfo.course ? String(scheduleInfo.course) : "No Class Assigned";

          if (isWithinSubWeek) {
            var jobObj = {
              id: rowId,
              teacherName: String(displayTeacherName),
              teacherEmail: String(rowTeacherEmail),
              date: formattedDate,
              formDateString: yyyymmdd,
              period: String(p),
              rawDate: rawDate,
              room: roomStr,
              course: courseStr,
              reason: reason,
              duration: duration,
              instructions: instructions,
              subPlanUrl: rowSubPlanUrl,
              assignedSub: assignedSubLower
            };

            if (assignedSubLower === userName) {
              if (rowDate < today && rowDate >= cutoffDate) {
                let subFeedbackRaw = extractSubFeedbackRaw(row);
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

            if (rowDate <= targetEndToday && rowDate >= today) {
              todaysAbsences.push(jobObj);
            }
          }

          if (isWithinQC && rawAssignedSub === "") {
            quickCover.push({
              id: rowId,
              teacherName: String(displayTeacherName),
              teacherEmail: String(rowTeacherEmail),
              date: String(formattedDate),
              formDateString: String(yyyymmdd),
              period: String(p),
              rawDate: rawDate,
              room: roomStr,
              course: courseStr,
              reason: reason,
              duration: duration,
              instructions: instructions,
              subPlanUrl: rowSubPlanUrl
            });
          }
        }
      }

      // 3c. Admin Dashboard Data & HR Dashboard Requests (single pass)
      if (isWithinWindow && (wantAdminData || wantHrData)) {
        var rawRosterTeacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
        var pdsTokens = periodsStrRaw.split(',').map(function(p) { return p.trim(); });
        var hrAssignedSubs = wantHrData ? [] : null;

        for (var pdsIdx = 0; pdsIdx < pdsTokens.length; pdsIdx++) {
          var rawPd = pdsTokens[pdsIdx];
          if (!rawPd) continue;
          var normPd = parsePeriod(rawPd);
          var pdNum = parseInt(normPd, 10);
          var subColOneBased = getSubColumnIndex(rawPd);
          var assignedSubVal = (subColOneBased > 0 && subColOneBased <= row.length) ? String(row[subColOneBased - 1] || "").trim() : "";
          var adSchedInfo = lookupTeacherScheduleInfo(scheduleLookup, rowTeacherEmail, !isNaN(pdNum) ? pdNum : rawPd, yyyymmdd, settings);

          if (wantAdminData && !isNaN(pdNum)) {
            adminData.push({
              id: rowId,
              originalDate: String(dateVal || ""),
              date: String(yyyymmdd || ""),
              formDateString: String(yyyymmdd || ""),
              period: pdNum,
              periodsString: periodsStrRaw,
              urgency: urgencyStr,
              teacherName: String(rawRosterTeacherName || ""),
              teacherEmail: String(rowTeacherEmail || ""),
              course: adSchedInfo ? String(adSchedInfo.course) : "",
              room: adSchedInfo ? String(adSchedInfo.room) : "",
              assignedSub: assignedSubVal,
              reason: reason.trim(),
              duration: duration.trim(),
              instructions: instructions.trim(),
              subPlanUrl: rowSubPlanUrl
            });
          }

          if (wantHrData && assignedSubVal !== "") {
            var hrCourseStr = adSchedInfo ? String(adSchedInfo.course) : "No Class Assigned";
            hrAssignedSubs.push({ name: assignedSubVal, period: String(rawPd), course: hrCourseStr });
          }
        }

        if (wantHrData) {
          hrData.push({
            id: rowId,
            date: String(yyyymmdd || ""),
            teacherName: String(rawRosterTeacherName || ""),
            reason: reason.trim(),
            duration: duration.trim(),
            assignedSubs: hrAssignedSubs
          });
        }
      }
    }

    var sortAbsencesAsc = function(a, b) { return a.rawDate - b.rawDate; };
    var sortAbsencesDesc = function(a, b) { return b.rawDate - a.rawDate; };
    myUpcomingAbsences.sort(sortAbsencesAsc);
    myPastAbsences.sort(sortAbsencesDesc);

    var sortJobs = function(a, b) {
      if (a.rawDate === b.rawDate) {
        return getPeriodSortValue(a.period) - getPeriodSortValue(b.period);
      }
      return a.rawDate - b.rawDate;
    };
    var sortJobsDesc = function(a, b) {
      if (a.rawDate === b.rawDate) {
        return getPeriodSortValue(a.period) - getPeriodSortValue(b.period);
      }
      return b.rawDate - a.rawDate;
    };
    mySubDuties.sort(sortJobs);
    myPastSubDuties.sort(sortJobsDesc);
    todaysAbsences.sort(sortJobs);


    var safeSettings = sanitizeSettingsForClient(settings, false);

    var payload = {
      userData: userData,
      myAbsences: myUpcomingAbsences,
      myPastAbsences: myPastAbsences,
      mySubDuties: mySubDuties,
      myPastSubDuties: myPastSubDuties,
      todaysAbsences: todaysAbsences,
      permissions: permissions,
      dateColors: dateColors,
      dateBellScheduleNames: dateBellScheduleNames,
      settings: safeSettings
    };


    // --- 4. Extract Admin / Sub Coordinator data if applicable ---
    if (want("staffList") && (permissions["Admin Dashboard"] || permissions["HR Dashboard"] || permissions["Today at a Glance"])) {
      // Staff List
      var staffList = [];
      var allSubAvail = getAllSubstituteAvailability();

      // Pre-calculate all staff schedules (using Staff Roster pre-computed schedules when available, or Master Schedule fallback)
      var allSchedules = rosterHasPrecomputedSchedules ? rosterSchedulesByEmail : {};
      var allSchedulesByTerm = rosterHasPrecomputedSchedules ? rosterSchedulesByTermByEmail : {};
      if (!rosterHasPrecomputedSchedules && scheduleData && scheduleData.length > 0) {
        var hdrs = scheduleData[0];
        var emlIdx = hdrs.indexOf("EMAIL_ADDR");
        var prIdx = hdrs.indexOf("PERIOD");
        var tmIdx = hdrs.indexOf("TERM");
        if (emlIdx > -1 && prIdx > -1) {
          for (var sIdx = 1; sIdx < scheduleData.length; sIdx++) {
            if (!scheduleData[sIdx]) continue;
            var sEmail = String(scheduleData[sIdx][emlIdx]).toLowerCase().trim();
            let pVal = String(scheduleData[sIdx][prIdx]).trim();
            var sTerm = (tmIdx > -1 && scheduleData[sIdx][tmIdx] !== undefined) ? String(scheduleData[sIdx][tmIdx]).trim() : activeTermId;
            var jP = getScheduleJoinPeriod(pVal);
            if (jP) {
              if (!allSchedules[sEmail]) allSchedules[sEmail] = [];
              if ((!sTerm || sTerm === activeTermId || Math.floor(Number(sTerm) / 100) * 100 === Number(sTerm)) && allSchedules[sEmail].indexOf(jP) === -1) {
                allSchedules[sEmail].push(jP);
              }
              if (sTerm) {
                if (!allSchedulesByTerm[sEmail]) allSchedulesByTerm[sEmail] = {};
                if (!allSchedulesByTerm[sEmail][sTerm]) allSchedulesByTerm[sEmail][sTerm] = [];
                if (allSchedulesByTerm[sEmail][sTerm].indexOf(jP) === -1) {
                  allSchedulesByTerm[sEmail][sTerm].push(jP);
                }
              }
            }
          }
          for (var e in allSchedules) {
            sortPeriodList(allSchedules[e]);
          }
          for (var em in allSchedulesByTerm) {
            for (var tm in allSchedulesByTerm[em]) {
              sortPeriodList(allSchedulesByTerm[em][tm]);
            }
          }
        }
      }



      for (var rsIdx = 1; rsIdx < rosterData.length; rsIdx++) {
        var staffName = String(rosterData[rsIdx][0]).trim();
        if (!rosterData[rsIdx]) continue;
        var staffEmail = String(rosterData[rsIdx][1]).toLowerCase().trim();
        var staffRole = String(rosterData[rsIdx][2]).toLowerCase().trim();
        var duty = String(rosterData[rsIdx][3] || "").trim();
        if (staffName) {
          var display = staffName;
          if (duty) display = staffName + " - " + duty;

          var department = rosterData[rsIdx].length > 4 ? String(rosterData[rsIdx][4] || "").trim() : "";
          var homebase = rosterData[rsIdx].length > 5 ? String(rosterData[rsIdx][5] || "").trim() : "";
          var building = rosterData[rsIdx].length > 6 ? String(rosterData[rsIdx][6] || "").trim() : "";
          staffList.push({
             name: staffName,
             display: display,
             duty: duty,
             role: staffRole,
             email: staffEmail,
             department: department,
             homebase: homebase,
             building: building,
             availability: allSubAvail[staffEmail] || {},
             teacherSchedule: allSchedules[staffEmail] || [],
             teacherScheduleByTerm: allSchedulesByTerm[staffEmail] || {}
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

    if (wantQuickCover) {
      quickCover.sort(sortJobs);
      payload.quickCover = quickCover;
    }
    if (wantAdminData) {
      adminData.sort(function(a, b) {
        if (a.date < b.date) return -1;
        if (a.date > b.date) return 1;
        return getPeriodSortValue(a.period) - getPeriodSortValue(b.period);
      });
      payload.adminData = adminData;
    }

    // --- 5. Extract HR data if applicable ---
    if (want("subAvailability") && lowerRole === 'substitute') {
      payload.subAvailability = getSubstituteAvailability(email);
    }

    if (wantHrData) {
      var payPeriods = [];

      for (let p = 0; p < payPeriodsData.length; p++) {
        if (!payPeriodsData[p] || payPeriodsData[p].length < 3) continue;
        var periodNum = String(payPeriodsData[p][0]).trim();
        var startDateRaw = payPeriodsData[p][1];
        var endDateRaw = payPeriodsData[p][2];

        var isHeader = false;
        if (typeof startDateRaw === 'string' && startDateRaw.toLowerCase().includes('start')) isHeader = true;
        if (typeof endDateRaw === 'string' && endDateRaw.toLowerCase().includes('end')) isHeader = true;
        if (isHeader) continue;

        if (periodNum && startDateRaw && endDateRaw) {
          var startFormatted = startDateRaw instanceof Date ? Utilities.formatDate(startDateRaw, tz, "yyyy-MM-dd") :
            (function(){ try { return Utilities.formatDate(new Date(startDateRaw), tz, "yyyy-MM-dd"); } catch(e) { return String(startDateRaw); } })();
          var endFormatted = endDateRaw instanceof Date ? Utilities.formatDate(endDateRaw, tz, "yyyy-MM-dd") :
            (function(){ try { return Utilities.formatDate(new Date(endDateRaw), tz, "yyyy-MM-dd"); } catch(e) { return String(endDateRaw); } })();
          var approved = String(payPeriodsData[p][3] || "").toLowerCase().trim() === "true";

          payPeriods.push({
            periodNumber: periodNum,
            startDate: startFormatted,
            endDate: endFormatted,
            approved: approved
          });
        }
      }

      payload.hrData = {
        requests: hrData,
        payPeriods: payPeriods,
        dateColors: dateColors,
        settings: safeSettings,
        rates: {
          green: settings["Green Day Pay Rate"] || "10",
          blueGold: settings["Blue/Gold Day Pay Rate"] || "20",
          special: settings["Special Pay Rates"] || "[]"
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
 * Refreshes requested data components.
 * @param {Array<string>} components - The components to fetch (e.g. ['myAbsences', 'quickCover'])
 * @returns {Object} The requested data.
 */
function refreshData(components, clientEmail) {
  try {
    var payload = getInitialPayload(typeof clientEmail !== 'undefined' ? clientEmail : undefined, components);
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

