/**
 * Sets up the database headers in the Google Sheet.
 * Run this function once from the Apps Script editor.
 */
function setupDatabase() {
  var ss = getSS();

  // 1. Absence Requests
  var absenceSheet = ss.getSheetByName("Absence Requests");
  if (!absenceSheet) {
    absenceSheet = ss.insertSheet("Absence Requests");
  }
  var absenceHeaders = [
    "ID", "Timestamp", "Email", "Date", "Periods", "Reason", "Duration",
    "Urgency", "Instructions", "Period 1 Sub", "Period 2 Sub", "Period 3 Sub",
    "Period 4 Sub", "Period 5 Sub", "Period 6 Sub", "Period 7 Sub", "Period 8 Sub", "Period 0 Sub", "Advisory Sub", "Status", "Sub Plan URL", "Sub Feedback"
  ];
  absenceSheet.getRange(1, 1, 1, absenceHeaders.length).setValues([absenceHeaders]);
  absenceSheet.getRange(1, 1, 1, absenceHeaders.length).setFontWeight("bold");

  // 2. Audit Log Sheet
  var auditSheet = ss.getSheetByName("Audit Log");
  if (!auditSheet) {
    auditSheet = ss.insertSheet("Audit Log");
  }
  var auditHeaders = ["Timestamp", "Actor", "Action Type", "Target ID", "Details"];
  auditSheet.getRange(1, 1, 1, auditHeaders.length).setValues([auditHeaders]);
  auditSheet.getRange(1, 1, 1, auditHeaders.length).setFontWeight("bold");
  auditSheet.hideSheet(); // Keep it hidden from normal view

  // 3. Settings Sheet
  var settingsSheet = ss.getSheetByName("Settings");

  // 7. SubstituteAvailability Sheet
  var subAvailSheet = ss.getSheetByName("SubstituteAvailability");
  if (!subAvailSheet) {
    subAvailSheet = ss.insertSheet("SubstituteAvailability");
    var subAvailHeaders = ["Email", "Date", "Availability"];
    subAvailSheet.getRange(1, 1, 1, subAvailHeaders.length).setValues([subAvailHeaders]);
    subAvailSheet.getRange(1, 1, 1, subAvailHeaders.length).setFontWeight("bold");
    subAvailSheet.getRange("B:B").setNumberFormat("@"); // Set Date column to plain text
  }
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    var settingsHeaders = ["Setting Name", "Setting Value"];
    var defaultSettings = [
      ["Email Mode", "Live"],
      ["Redirect Email", "Bgross@gocathedral.com"],
      ["Email Sender Name", "Cathedral Sub App"],
      ["App URL", DEFAULT_APP_URL],
      ["Max Multi-Select Days", "5"],
      ["Urgency Cutoff Time", "15"],
      ["Term ID", "3601"],
      ["Green Day Pay Rate", "10"],
      ["Blue/Gold Day Pay Rate", "20"],
      ["Special Pay Rates", "[]"],
      ["Absence Reasons", JSON.stringify([
        {reason: "Personal", hrRequired: false},
        {reason: "Professional Development", hrRequired: false},
        {reason: "Retreat", hrRequired: false},
        {reason: "Athletics", hrRequired: false},
        {reason: "Jury Duty", hrRequired: true},
        {reason: "Bereavement", hrRequired: true}
      ])]
    ];
    settingsSheet.getRange(1, 1, 1, 2).setValues([settingsHeaders]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
    settingsSheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  }

  // 4. Staff Roster Sheet
  var rosterSheet = ss.getSheetByName("Staff Roster");
  if (!rosterSheet) {
    rosterSheet = ss.insertSheet("Staff Roster");
  }
  var rosterHeaders = ["Name", "Email", "Role", "Duty"];
  rosterSheet.getRange(1, 1, 1, rosterHeaders.length).setValues([rosterHeaders]);
  rosterSheet.getRange(1, 1, 1, rosterHeaders.length).setFontWeight("bold");

  // 6. PayPeriods Sheet
  var payPeriodsSheet = ss.getSheetByName("PayPeriods");
  if (!payPeriodsSheet) {
    payPeriodsSheet = ss.insertSheet("PayPeriods");
  }
  var payPeriodsHeaders = ["Period Number", "Start Date", "End Date", "Approved"];
  payPeriodsSheet.getRange(1, 1, 1, payPeriodsHeaders.length).setValues([payPeriodsHeaders]);
  payPeriodsSheet.getRange(1, 1, 1, payPeriodsHeaders.length).setFontWeight("bold");

  // 8. Dates (Bell Schedule) Sheet
  var datesSheet = ss.getSheetByName("Dates");
  if (!datesSheet) {
    datesSheet = ss.insertSheet("Dates");
  }
  var datesHeaders = ["Date", "Color"];
  datesSheet.getRange(1, 1, 1, datesHeaders.length).setValues([datesHeaders]);
  datesSheet.getRange(1, 1, 1, datesHeaders.length).setFontWeight("bold");
  datesSheet.getRange("A:A").setNumberFormat("@");

  // 9. Email Queue Sheet
  var emailQueueSheet = ss.getSheetByName("Email Queue");
  if (!emailQueueSheet) {
    emailQueueSheet = ss.insertSheet("Email Queue");
  }
  var emailQueueHeaders = ["Timestamp", "To", "Subject", "Body", "Options", "Status"];
  emailQueueSheet.getRange(1, 1, 1, emailQueueHeaders.length).setValues([emailQueueHeaders]);
  emailQueueSheet.getRange(1, 1, 1, emailQueueHeaders.length).setFontWeight("bold");

  // 10. Payroll Sheet (Snapshots written when Pay Periods are approved or re-approved)
  var payrollSheet = ss.getSheetByName("Payroll");
  if (!payrollSheet) {
    payrollSheet = ss.insertSheet("Payroll");
  }
  var payrollHeaders = [
    "Pay Period", "Date Range", "Teacher Name", "Duties", "Extras", "Extra (Pay)", "Totals", "Approved By", "Approved At"
  ];
  payrollSheet.getRange(1, 1, 1, payrollHeaders.length).setValues([payrollHeaders]);
  payrollSheet.getRange(1, 1, 1, payrollHeaders.length).setFontWeight("bold");

  // 11. Absences Sheet (v2 Normalized Parent Table)
  var v2AbsencesSheet = ss.getSheetByName("Absences");
  if (!v2AbsencesSheet) {
    v2AbsencesSheet = ss.insertSheet("Absences");
  }
  v2AbsencesSheet.getRange(1, 1, 1, V2_ABSENCES_HEADERS.length).setValues([V2_ABSENCES_HEADERS]);
  v2AbsencesSheet.getRange(1, 1, 1, V2_ABSENCES_HEADERS.length).setFontWeight("bold");

  // 12. Coverage_Assignments Sheet (v2 Normalized Child Table)
  var v2AssignmentsSheet = ss.getSheetByName("Coverage_Assignments");
  if (!v2AssignmentsSheet) {
    v2AssignmentsSheet = ss.insertSheet("Coverage_Assignments");
  }
  v2AssignmentsSheet.getRange(1, 1, 1, V2_ASSIGNMENTS_HEADERS.length).setValues([V2_ASSIGNMENTS_HEADERS]);
  v2AssignmentsSheet.getRange(1, 1, 1, V2_ASSIGNMENTS_HEADERS.length).setFontWeight("bold");
}

/**
 * Admin migration helper that backfills or re-syncs the normalized v2 `Absences`
 * and `Coverage_Assignments` tables from the legacy `Absence Requests` sheet.
 */
function migrateAbsenceRequestsToV2(clientEmail) {
  try {
    var ss = getSS();
    if (clientEmail !== undefined) {
      var user = getUserData(ss, clientEmail);
      assertRole(user, "admin");
    }
    var result = rebuildAllV2AbsenceSheets(ss);
    if (result && result.success) {
      logAuditAction(
        "V2_SCHEMA_MIGRATED",
        "Absences,Coverage_Assignments",
        "Synced " + result.absencesCount + " absences and " + result.assignmentsCount + " coverage assignments",
        clientEmail
      );
    }
    return result;
  } catch (err) {
    notifyAdminOfError("migrateAbsenceRequestsToV2", err);
    return { success: false, error: err.message };
  }
}
