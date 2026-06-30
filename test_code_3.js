const fs = require('fs');
const content = fs.readFileSync('code.gs', 'utf8');
if (content.includes('function getRosterDataCached(ss)') && content.includes('clearRosterCache()') && content.split('clearRosterCache();').length > 3) {
  console.log("Success: getRosterDataCached implemented");
} else {
  console.log("Error: getRosterDataCached not optimized properly");
}
