const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

// For editAbsence
html = html.replace('document.getElementById(\'agreement\').checked = true;\n\n        toggleModal(true);', 'document.getElementById(\'agreement\').checked = true;\n\n        toggleModal(true);\n        setTimeout(captureModalState, 50);');

// For openAdminEditModal
html = html.replace('document.getElementById(\'agreement\').checked = true;\n\n         toggleModal(true);', 'document.getElementById(\'agreement\').checked = true;\n\n         toggleModal(true);\n         setTimeout(captureModalState, 50);');

fs.writeFileSync('Index.html', html);
