const fs = require('fs');
const content = fs.readFileSync('code.gs', 'utf8');
if (content.includes('auditSheet.appendRow([timestamp, actor, safeActionType, safeTargetId, safeDetails]);') && !content.includes('SpreadsheetApp.flush();')) {
  console.log("Success: logAuditAction optimized");
} else {
  console.log("Error: logAuditAction not optimized properly");
}
