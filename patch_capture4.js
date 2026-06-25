const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

// For editAbsence
html = html.replace('document.getElementById(\'adminDeleteBtn\').classList.add(\'hidden\'); // Regular users don\'t see delete in the modal, they use the cancel button on the list\n        toggleModal(true);', 'document.getElementById(\'adminDeleteBtn\').classList.add(\'hidden\'); // Regular users don\'t see delete in the modal, they use the cancel button on the list\n        toggleModal(true);\n        setTimeout(captureModalState, 50);');

fs.writeFileSync('Index.html', html);
