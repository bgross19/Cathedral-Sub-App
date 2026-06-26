const fs = require('fs');
const codeGs = fs.readFileSync('code.gs', 'utf8');
const testsGs = fs.readFileSync('tests.gs', 'utf8');

// Simple mocks required for global scope
global.getSettings = function() {};
global.GmailApp = { sendEmail: function() {} };
global.console.log = console.log;

// Evaluate the scripts
eval(codeGs);
eval(testsGs);

// Run the tests
const result = runTests();
if (result.failed > 0) {
  process.exit(1);
}
