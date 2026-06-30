const fs = require('fs');
const content = fs.readFileSync('code.gs', 'utf8');
if (content.includes('_globalSettingsCache = null;') && content.includes('cache.put("app_settings", JSON.stringify(settings), 300);')) {
  console.log("Success: getSettings optimized");
} else {
  console.log("Error: getSettings not optimized properly");
}
