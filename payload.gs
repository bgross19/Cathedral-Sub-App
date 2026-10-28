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

    var needAbsenceData = wantAll || want("myAbsences") || want("myPastAbsences") || want("mySubDuties") || want("myPastSubDuties") || want("todaysAbsences") || want("quickCover") || want("adminData") || want("hrData");
    var mainSheet = needAbsenceData ? getSheetOrThrow(ss, "Absence Requests") : null;
    var absenceData = mainSheet ? mainSheet.getDataRange().getValues() : [];

    var needScheduleData = wantAll || want("userData") || want("mySubDuties") || want("myPastSubDuties") || want("todaysAbsences") || want("staffList") || want("quickCover") || want("adminData") || want("hrData");
    var scheduleData = needScheduleData ? getMasterScheduleData() : [];

    var needPayPeriods = want("hrData");
    var payPeriodsSheet = needPayPeriods ? ss.getSheetByName("PayPeriods") : null;
    var payPeriodsData = payPeriodsSheet ? payPeriodsSheet.getDataRange().getValues() : [];

    var needDates = want("dateColors") || want("hrData");
    var dateColors = needDates ? getDateColorsCached(ss) : {};



    // --- Build lookups ---
    var scheduleLookup = buildScheduleLookup(scheduleData);
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

    // Compute teacherSchedule
    var teacherSchedule = [];
    if (scheduleData && scheduleData.length > 0) {
      var headers = scheduleData[0];
      var emailIdx = headers.indexOf("EMAIL_ADDR");
      var periodIdx = headers.indexOf("PERIOD");
      if (emailIdx > -1 && periodIdx > -1) {
        for (var s = 1; s < scheduleData.length; s++) {
          if (!scheduleData[s] || scheduleData[s].length <= Math.max(emailIdx, periodIdx)) continue;
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
      teacherSchedule: teacherSchedule
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

    for (let i = 1; i < absenceData.length; i++) {
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
        let subFeedbackRaw = extractSubFeedbackRaw(row);
        var subPlanUrl = extractSubPlanUrl(row);

        var periodsRequested = String(row[4] || "").split(",").map(function(p) { return p.trim(); }).filter(function(p) { return p !== ""; });
        var periodsWithSubs = periodsRequested.map(function(p) {
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
          id: String(row[0]),
          date: String(formattedDate),
          rawDate: isDateValid ? Number(rowDate.getTime()) : 0,
          rawDateString: String(dateVal),
          formDateString: String(yyyymmdd),
          periods: String(row[4]),
          periodsWithSubs: periodsWithSubs,
          reason: String(row[5]),
          urgency: urgencyStr.includes('Urgent') ? 'Urgent' : 'Standard',
          duration: String(row[6]),
          instructions: String(row[8]),
          subPlanUrl: subPlanUrl,
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
          if (parts.length > 1) teacherName = parts[1].trim() + " " + parts[0].trim();
        }

        periodsRequested = String(row[4]).split(",").map(function(p) { return parsePeriod(p); });
        var rowId = String(row[0]);
        var rawDate = Number(rowDate.getTime());
        var reason = String(row[5]);
        var duration = String(row[6]);
        var instructions = String(row[8]);
        var jobSubPlanUrl = extractSubPlanUrl(row);

        var allPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', '10'];
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
                instructions: instructions,
                subPlanUrl: jobSubPlanUrl,
                assignedSub: assignedSub
            };

            // My Sub Duties
            if (assignedSub === userName) {
              if (rowDate < today && rowDate >= cutoffDate) {
                 let subFeedbackRaw = extractSubFeedbackRaw(row);
                 var subFeedbackParsed = [];
                 try {
                     subFeedbackParsed = JSON.parse(subFeedbackRaw);
                 } catch(e) {
                     // Ignore parse error
                 }
                 var matchingFeedback = subFeedbackParsed.filter(function(fb) { return fb.period === String(p) && String(fb.subName).toLowerCase() === userName; });
                 jobObj.rating = matchingFeedback.length > 0 ? matchingFeedback[0].rating : 0;
                 jobObj.note = matchingFeedback.length > 0 ? matchingFeedback[0].note : "";
                 myPastSubDuties.push(jobObj);
              } else if (rowDate >= today) {
                 mySubDuties.push(jobObj);
              }
            }

            // Open Jobs (Today only)
            if (rowDate <= targetEndToday && rowDate >= today) {
              todaysAbsences.push(jobObj);
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
      settings: safeSettings
    };


    // --- 4. Extract Admin / Sub Coordinator data if applicable ---
    if (want("staffList") && (permissions["Admin Dashboard"] || permissions["HR Dashboard"] || permissions["Today at a Glance"])) {
      // Staff List
      var staffList = [];
      var allSubAvail = getAllSubstituteAvailability();

      // Pre-calculate all staff schedules
      var allSchedules = {};
      if (scheduleData && scheduleData.length > 0) {
        var hdrs = scheduleData[0];
        var emlIdx = hdrs.indexOf("EMAIL_ADDR");
        var prIdx = hdrs.indexOf("PERIOD");
        if (emlIdx > -1 && prIdx > -1) {
          for (var sIdx = 1; sIdx < scheduleData.length; sIdx++) {
            var sEmail = String(scheduleData[sIdx][emlIdx]).toLowerCase().trim();
            let pVal = String(scheduleData[sIdx][prIdx]).trim();
            var jP = getScheduleJoinPeriod(pVal);
            if (jP) {
              if (!allSchedules[sEmail]) allSchedules[sEmail] = [];
              if (allSchedules[sEmail].indexOf(jP) === -1) {
                allSchedules[sEmail].push(jP);
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

    if ((want("quickCover") || want("adminData")) && (permissions["Admin Dashboard"] || permissions["Today at a Glance"] || permissions["Today's Absences"])) {
      if (want("quickCover")) {
        // Quick Cover Data
        var quickCover = [];
        var targetEndQC = new Date(today);
        var dayOfWeek = today.getDay();
        var daysToAdd = 1;
        if (dayOfWeek === 5) daysToAdd = 3;
        else if (dayOfWeek === 6) daysToAdd = 2;
        targetEndQC.setDate(today.getDate() + daysToAdd);
        targetEndQC.setHours(23, 59, 59, 999);

        for (var abIdx = 1; abIdx < absenceData.length; abIdx++) {
          let row = absenceData[abIdx];
          if (!row || row.length < 4) continue;
          var stat = String(row[19] || 'Active');
          if (stat === 'Canceled') continue;

          var dVal = row[3];
          if (!dVal) continue;
          var rDate = new Date(dVal);
          if (isNaN(rDate.getTime())) continue;
          if (rDate < cutoffDate) continue;

          if (rDate >= today && rDate <= targetEndQC) {
            var rTeacherEmail = String(row[2]).toLowerCase();
            var tName = nameLookup[rTeacherEmail] || rTeacherEmail;
            if (tName.includes(",")) {
              var prts = tName.split(",");
              if (prts.length > 1) tName = prts[1].trim() + " " + prts[0].trim();
            }

            var prdsRequested = String(row[4]).split(",").map(function(p) { return parsePeriod(p); });
            var rId = String(row[0]);
            var fmtDate = String(Utilities.formatDate(rDate, Session.getScriptTimeZone(), "MMM d, yyyy"));
            var fDateString = String(Utilities.formatDate(rDate, Session.getScriptTimeZone(), "yyyy-MM-dd"));
            var rwDate = Number(rDate.getTime());
            var rsn = String(row[5]);
            var dur = String(row[6]);
            var inst = String(row[8]);
            var qcSubPlanUrl = extractSubPlanUrl(row);

            var aPeriods = ['1', '2', '3', '4', '5', '6', '7', '8', '0', '10'];
            for (let prIdx = 0; prIdx < aPeriods.length; prIdx++) {
              var prd = aPeriods[prIdx];
              if (prdsRequested.indexOf(String(prd)) !== -1) {
                var asgSub = row[getSubColumnIndex(prd) - 1];
                if (!asgSub || String(asgSub).trim() === "") {
                  var jKey = rTeacherEmail + "-" + getScheduleJoinPeriod(prd);
                  var schedInfo = scheduleLookup[jKey];
                  quickCover.push({
                    id: rId,
                    teacherName: String(tName),
                    teacherEmail: String(rTeacherEmail),
                    date: fmtDate,
                    formDateString: fDateString,
                    period: String(prd),
                    rawDate: rwDate,
                    room: schedInfo && schedInfo.room ? String(schedInfo.room) : "No Class Assigned",
                    course: schedInfo && schedInfo.course ? String(schedInfo.course) : "No Class Assigned",
                    reason: rsn,
                    duration: dur,
                    instructions: inst,
                    subPlanUrl: qcSubPlanUrl
                  });
                }
              }
            }
          }
        }
        quickCover.sort(sortJobs);
        payload.quickCover = quickCover;
      }

      if (want("adminData")) {
        // Admin Dashboard Data
        var adminData = [];
        for (var adIdx = 1; adIdx < absenceData.length; adIdx++) {
          var adRow = absenceData[adIdx];
          if (String(adRow[19] || "").trim() === "Canceled") continue;

          var dtStr = adRow[3];
          var dtObj = new Date(dtStr);
          if (!isNaN(dtObj.getTime()) && dtObj < cutoffDate) continue;

          var adTeacherEmail = String(adRow[2]).toLowerCase().trim();
          var adTeacherName = nameLookup[adTeacherEmail] || adTeacherEmail;
          var dtFormatted = !isNaN(dtObj.getTime()) ? Utilities.formatDate(dtObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : dtStr;
          var pdsStr = String(adRow[4]);
          var pds = pdsStr.split(',').map(function(p) { return p.trim(); });

          for (var pdsIdx = 0; pdsIdx < pds.length; pdsIdx++) {
            var rawPd = pds[pdsIdx];
            var pdNum = parseInt(parsePeriod(rawPd), 10);

            if (!isNaN(pdNum)) {
              var pd = pdNum;
              var schedKey = adTeacherEmail + "-" + getScheduleJoinPeriod(pd);
              adminData.push({
                id: String(adRow[0] || ""),
                originalDate: String(dtStr || ""),
                date: String(dtFormatted || ""),
                formDateString: String(dtFormatted || ""),
                period: pd,
                periodsString: String(pdsStr || ""),
                urgency: String(adRow[7] || ""),
                teacherName: String(adTeacherName || ""),
                teacherEmail: String(adTeacherEmail || ""),
                course: scheduleLookup[schedKey] ? String(scheduleLookup[schedKey].course) : "",
                room: scheduleLookup[schedKey] ? String(scheduleLookup[schedKey].room) : "",
                assignedSub: String(adRow[getSubColumnIndex(pd) - 1] || "").trim(),
                reason: String(adRow[5] || "").trim(),
                duration: String(adRow[6] || "").trim(),
                instructions: String(adRow[8] || "").trim(),
                subPlanUrl: extractSubPlanUrl(adRow)
              });
            }
          }
        }
        adminData.sort(function(a, b) {
          if (a.date < b.date) return -1;
          if (a.date > b.date) return 1;
          return getPeriodSortValue(a.period) - getPeriodSortValue(b.period);
        });
        payload.adminData = adminData;
      }
    }


    // --- 5. Extract HR data if applicable ---
    if (want("subAvailability") && lowerRole === 'substitute') {
      payload.subAvailability = getSubstituteAvailability(email);
    }

    if (want("hrData") && permissions["HR Dashboard"]) {
      var hrData = [];
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

      for (var hrIdx = 1; hrIdx < absenceData.length; hrIdx++) {
        var hrRow = absenceData[hrIdx];
        if (!hrRow || hrRow.length < 4) continue;
        if (String(hrRow[19] || "").trim() === "Canceled") continue;

        var hrDateStr = hrRow[3];
        var hrDateObj = new Date(hrDateStr);
        if (!isNaN(hrDateObj.getTime()) && hrDateObj < cutoffDate) continue;

        var hrTeacherEmail = String(hrRow[2]).toLowerCase().trim();
        var hrTeacherName = nameLookup[hrTeacherEmail] || hrTeacherEmail;
        var hrDateFormatted = !isNaN(hrDateObj.getTime()) ? Utilities.formatDate(hrDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : hrDateStr;

        var hrPeriodsStr = String(hrRow[4]);
        var hrPeriods = hrPeriodsStr.split(',').map(function(p) { return p.trim(); });
        var assignedSubs = [];

        for (var hrPIdx = 0; hrPIdx < hrPeriods.length; hrPIdx++) {
          var pStr = hrPeriods[hrPIdx];
          let p = pStr; // Keep as string for getSubColumnIndex
          if (p) {
            var hrAssignedSub = hrRow[getSubColumnIndex(p) - 1];
            if (hrAssignedSub && String(hrAssignedSub).trim() !== "") {
              var hrSchedKey = hrTeacherEmail + "-" + getScheduleJoinPeriod(p);
              var hrCourseStr = scheduleLookup[hrSchedKey] ? String(scheduleLookup[hrSchedKey].course) : "No Class Assigned";
              assignedSubs.push({ name: String(hrAssignedSub).trim(), period: String(p), course: hrCourseStr });
            }
          }
        }

        hrData.push({
          id: String(hrRow[0] || ""),
          date: String(hrDateFormatted || ""),
          teacherName: String(hrTeacherName || ""),
          reason: String(hrRow[5]).trim(),
          duration: String(hrRow[6]).trim(),
          assignedSubs: assignedSubs
        });
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

