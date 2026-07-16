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
    "Period 4 Sub", "Period 5 Sub", "Period 6 Sub", "Period 7 Sub", "Period 8 Sub", "Status"
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
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    var settingsHeaders = ["Setting Name", "Setting Value"];
    var defaultSettings = [
      ["Email Mode", "Live"],
      ["Redirect Email", "Bgross@gocathedral.com"],
      ["App URL", DEFAULT_APP_URL],
      ["Urgency Cutoff Time", "15"],
      ["Term ID", "3503"],
      ["PS_CLIENT_ID", ""],
      ["PS_CLIENT_SECRET", ""],
      ["PS_URL", ""],
      ["Green Day Pay Rate", "10"],
      ["Blue/Gold Day Pay Rate", "20"],
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

  // 5. User Roles Sheet
  var rolesSheet = ss.getSheetByName("User Roles");
  if (!rolesSheet) {
    rolesSheet = ss.insertSheet("User Roles");
  }
  var rolesHeaders = ["Email", "Role"];
  rolesSheet.getRange(1, 1, 1, rolesHeaders.length).setValues([rolesHeaders]);
  rolesSheet.getRange(1, 1, 1, rolesHeaders.length).setFontWeight("bold");

  // 6. PayPeriods Sheet
  var payPeriodsSheet = ss.getSheetByName("PayPeriods");
  if (!payPeriodsSheet) {
    payPeriodsSheet = ss.insertSheet("PayPeriods");
  }
  var payPeriodsHeaders = ["Period Number", "Start Date", "End Date"];
  payPeriodsSheet.getRange(1, 1, 1, payPeriodsHeaders.length).setValues([payPeriodsHeaders]);
  payPeriodsSheet.getRange(1, 1, 1, payPeriodsHeaders.length).setFontWeight("bold");
}
