const fs = require('fs');
const content = fs.readFileSync('powerschool.gs', 'utf8');
if (content.includes('CacheService.getScriptCache()') && content.includes('cache.put("master_schedule_data", stringified, 1800);')) {
  console.log("Success: getMasterScheduleData optimized");
} else {
  console.log("Error: getMasterScheduleData not optimized properly");
}
