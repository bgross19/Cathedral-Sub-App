const fs = require('fs');
const html = fs.readFileSync('Index.html', 'utf8');

if (html.includes('id="settingsUrgencyCutoff"') && html.includes('Absence Reasons')) {
  console.log("UI updates look correct");
} else {
  console.log("UI updates missing");
}
