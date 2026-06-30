const fs = require('fs');
const content = fs.readFileSync('code.gs', 'utf8');
if (!content.includes('var rosterSheet = getSheetOrThrow(ss, "Staff Roster");\n    var rosterData = rosterSheet ? rosterSheet.getDataRange().getValues() : [];\n    var scheduleData = getMasterScheduleData();') && content.includes('var rosterData = getRosterDataCached(ss);')) {
  console.log("Success: Used getRosterDataCached where applicable.");
} else {
  console.log("Error: Optimization not fully implemented.");
}
