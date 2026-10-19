const fs = require('fs');

const html = fs.readFileSync('Index.html', 'utf8');

const regex = /function\s+saveDayPayRates\(\)\s*{[\s\S]*?}/;
const match = html.match(regex);
console.log(match ? match[0] : "Not found");
