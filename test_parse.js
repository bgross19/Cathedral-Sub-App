const fs = require('fs');
const content = fs.readFileSync('Index.html', 'utf8');

// Check script tags balancing
const scriptOpen = (content.match(/<script\b[^>]*>/g) || []).length;
const scriptClose = (content.match(/<\/script>/g) || []).length;
console.log(`Open script tags: ${scriptOpen}`);
console.log(`Close script tags: ${scriptClose}`);
