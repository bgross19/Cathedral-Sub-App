const fs = require('fs');

global.Logger = { log: console.log };
global.console = { log: console.log, error: console.error, warn: console.warn };
global.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({}) }) };
global.Session = { getActiveUser: () => ({ getEmail: () => "admin@example.com" }) };
global.PropertiesService = { getScriptProperties: () => ({ getProperty: () => null }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };
global.LockService = { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) };
global.HtmlService = { createHtmlOutputFromFile: () => ({ setTitle: () => {}, setXFrameOptionsMode: () => {} }), XFrameOptionsMode: { ALLOWALL: 1 } };
global.Utilities = { computeDigest: () => [] };
global.ScriptApp = { getOAuthToken: () => "" };
global.GmailApp = { sendEmail: () => {} };

let staffRosterData = [
  ['Name', 'Email', 'Duty'],
  ['Alice', 'alice@example.com', '1'],
  ['Bob', 'bob@example.com', '2']
];

let mockSheet = {
  getDataRange: () => ({ getValues: () => [...staffRosterData] }),
  appendRow: (row) => { staffRosterData.push(row); },
  getRange: (r, c, numRows, numCols) => ({
    setValues: (values) => {
        if (numRows === 1) {
            staffRosterData[r-1] = values[0];
        } else {
            for (let i = 0; i < numRows; i++) {
               staffRosterData[r-1+i] = values[i];
            }
        }
    }
  }),
  deleteRow: (index) => { staffRosterData.splice(index - 1, 1); },
  getLastRow: () => staffRosterData.length
};

global.getSS = () => ({
  getSheetByName: (name) => {
    if (name === "Staff Roster") return mockSheet;
    if (name === "Settings") return { getDataRange: () => ({ getValues: () => [] }) };
    if (name === "User Roles") return { getDataRange: () => ({ getValues: () => [['Email', 'Role'], ['admin@example.com', 'Admin']] }) };
    return null;
  }
});

global.SpreadsheetApp = {
  getActiveSpreadsheet: global.getSS
};

// Evaluate the modified code
const code = fs.readFileSync('code.gs', 'utf8');
const script = new (require('vm')).Script(code);
const sandbox = { ...global, console: global.console };
script.runInNewContext(sandbox);

console.log("--- Testing getStaffRosterForAdmin ---");
let roster = sandbox.getStaffRosterForAdmin();
if (roster.length === 2 && roster[0].name === 'Alice') {
    console.log("Passed fetch test.");
} else {
    console.log("Failed fetch test.");
    process.exit(1);
}

console.log("--- Testing saveStaffMemberAdmin (Add) ---");
sandbox.saveStaffMemberAdmin({name: 'Charlie', email: 'charlie@example.com', duty: '3'});
roster = sandbox.getStaffRosterForAdmin();
if (roster.length === 3 && roster[2].name === 'Charlie') {
    console.log("Passed add staff test.");
} else {
    console.log("Failed add staff test.");
    process.exit(1);
}

console.log("--- Testing saveStaffMemberAdmin (Edit) ---");
sandbox.saveStaffMemberAdmin({originalEmail: 'alice@example.com', name: 'Alice M', email: 'alice.m@example.com', duty: '1B'});
roster = sandbox.getStaffRosterForAdmin();
if (roster[0].name === 'Alice M' && roster[0].email === 'alice.m@example.com') {
    console.log("Passed edit staff test.");
} else {
    console.log("Failed edit staff test.");
    process.exit(1);
}

console.log("--- Testing deleteStaffMemberAdmin ---");
sandbox.deleteStaffMemberAdmin('bob@example.com');
roster = sandbox.getStaffRosterForAdmin();
if (roster.length === 2 && roster.findIndex(r => r.name === 'Bob') === -1) {
    console.log("Passed delete staff test.");
} else {
    console.log("Failed delete staff test.");
    process.exit(1);
}

console.log("--- Testing bulkUpsertStaffRoster ---");
sandbox.bulkUpsertStaffRoster([
    {name: 'Alice M Updated', email: 'alice.m@example.com', duty: '10'},
    {name: 'Dave', email: 'dave@example.com', duty: '5'}
]);
roster = sandbox.getStaffRosterForAdmin();
if (roster.length === 3 && roster.find(r => r.name === 'Alice M Updated') && roster.find(r => r.name === 'Dave')) {
    console.log("Passed bulk upsert test.");
} else {
    console.log("Failed bulk upsert test.");
    process.exit(1);
}

console.log("All custom tests passed.");
process.exit(0);
