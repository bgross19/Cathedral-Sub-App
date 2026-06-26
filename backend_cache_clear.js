const fs = require('fs');

let content = fs.readFileSync('code.gs', 'utf8');

// Add clearMasterScheduleCache to code.gs
const clearCacheCode = `
/**
 * Clears the Master Schedule cache from the Script Cache.
 * Returns a success object.
 */
function clearMasterScheduleCache() {
  try {
    var ss = getSS();
    var user = getUserData(ss);
    assertRole(user, "admin");

    var cache = CacheService.getScriptCache();
    cache.remove("ps_master_schedule");

    return { success: true };
  } catch (err) {
    notifyAdminOfError("clearMasterScheduleCache", err);
    return { success: false, error: err.message };
  }
}
`;

content += '\n' + clearCacheCode;
fs.writeFileSync('code.gs', content, 'utf8');
