const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

// The end of the "opening" block inside toggleModal is at line 1081.
html = html.replace('document.querySelectorAll(\'.period-cb\').forEach(cb => cb.checked = false);\n          }', 'document.querySelectorAll(\'.period-cb\').forEach(cb => cb.checked = false);\n          }\n          // Delay slightly to let inputs initialize if needed\n          setTimeout(captureModalState, 50);');

fs.writeFileSync('Index.html', html);
