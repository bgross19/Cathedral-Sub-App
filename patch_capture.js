const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

// Hook inside openAddRoleModal
html = html.replace('body.classList.add(\'modal-active\');\n      }\n\n      function closeAddRoleModal()', 'body.classList.add(\'modal-active\');\n        // Delay slightly to let inputs initialize if needed\n        setTimeout(captureModalState, 50);\n      }\n\n      function closeAddRoleModal()');

// Hook inside toggleModal (opening)
html = html.replace('body.classList.add(\'modal-active\');\n\n          if (!isEdit) {', 'body.classList.add(\'modal-active\');\n\n          if (!isEdit) {'); // Need to find where to hook inside toggleModal.
// For toggleModal, we have !isEdit branch and edit branch. We can hook it at the end of the 'if' block.
