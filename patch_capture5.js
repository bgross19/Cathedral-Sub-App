const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

// Hook inside openAddRoleModal
html = html.replace('body.classList.add(\'modal-active\');\n      }\n\n      function closeAddRoleModal()', 'body.classList.add(\'modal-active\');\n        // Delay slightly to let inputs initialize if needed\n        setTimeout(captureModalState, 50);\n      }\n\n      function closeAddRoleModal()');

fs.writeFileSync('Index.html', html);
