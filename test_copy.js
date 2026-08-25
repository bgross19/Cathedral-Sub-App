const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

// The classes to remove
const classesToRemove = ['sticky', 'left-0', 'z-10', 'z-20', 'shadow-[1px_0_0_0_#d1d5db]'];

// The classes to keep
// Let's look at how the table is rendered
