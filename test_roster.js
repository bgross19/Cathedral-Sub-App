const fs = require('fs');
global.Logger = { log: console.log };
global.console = { log: console.log, error: console.error, warn: console.warn };

let lockAcquired = false;
global.LockService = {
  getScriptLock: () => ({
    waitLock: (timeout) => { lockAcquired = true; },
    releaseLock: () => { lockAcquired = false; }
  })
};

global.PropertiesService = { getScriptProperties: () => ({ getProperty: () => null }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };

let activeUserEmail = 'admin@example.com';
global.Session = {
  getActiveUser: () => ({ getEmail: () => activeUserEmail })
};

global.getUserData = () => ({ role: 'admin' });
global.assertRole = () => {};

let staffRosterData = [
  ['Name', 'Email', 'Duty'],
  ['Alice', 'alice@example.com', '1'],
  ['Bob', 'bob@example.com', '2']
];

let lastRowAppended = [];
let lastRangeSetValues = [];
let deletedRowIndex = -1;

const mockSheet = {
  getDataRange: () => ({ getValues: () => [...staffRosterData] }),
  appendRow: (row) => { lastRowAppended = row; },
  getRange: (r, c, numRows, numCols) => ({
    setValues: (values) => { lastRangeSetValues = {r, c, numRows, numCols, values}; }
  }),
  deleteRow: (index) => { deletedRowIndex = index; },
  getLastRow: () => staffRosterData.length
};

global.getSS = () => ({
  getSheetByName: (name) => name === 'Staff Roster' ? mockSheet : null
});

global.getSheetOrThrow = (ss, name) => {
  if (name === 'Staff Roster') return mockSheet;
  throw new Error("Sheet not found");
};

// Evaluate the modified code
eval(fs.readFileSync('code.gs', 'utf8'));

// Test fetching staff roster
console.log("--- Testing getStaffRosterForAdmin ---");
const roster = getStaffRosterForAdmin();
if (roster.length === 2 && roster[0].name === 'Alice' && roster[1].name === 'Bob') {
  console.log("Passed fetch test");
} else {
  console.error("Failed fetch test", roster);
  process.exit(1);
}

// Test adding a staff member
console.log("--- Testing saveStaffMemberAdmin (Add) ---");
lastRowAppended = [];
saveStaffMemberAdmin({ name: 'Charlie', email: 'charlie@example.com', duty: '3' });
if (lastRowAppended.length === 3 && lastRowAppended[0] === 'Charlie') {
  console.log("Passed add staff test");
} else {
  console.error("Failed add staff test", lastRowAppended);
  process.exit(1);
}

// Test updating a staff member
console.log("--- Testing saveStaffMemberAdmin (Edit) ---");
lastRangeSetValues = [];
saveStaffMemberAdmin({ originalEmail: 'alice@example.com', name: 'Alice M', email: 'alice.m@example.com', duty: '1B' });
if (lastRangeSetValues.r === 2 && lastRangeSetValues.values[0][1] === 'alice.m@example.com') {
  console.log("Passed edit staff test");
} else {
  console.error("Failed edit staff test", lastRangeSetValues);
  process.exit(1);
}

// Test deleting a staff member
console.log("--- Testing deleteStaffMemberAdmin ---");
deletedRowIndex = -1;
deleteStaffMemberAdmin('bob@example.com');
if (deletedRowIndex === 3) {
  console.log("Passed delete staff test");
} else {
  console.error("Failed delete staff test", deletedRowIndex);
  process.exit(1);
}

// Test Bulk Upsert
console.log("--- Testing bulkUpsertStaffRoster ---");
lastRangeSetValues = null;
const upsertResult = bulkUpsertStaffRoster([
  { name: 'Alice New', email: 'alice@example.com', duty: '10' },
  { name: 'Dave', email: 'dave@example.com', duty: '5' }
]);
if (upsertResult.success && upsertResult.updated === 2) {
  console.log("Passed bulk upsert test");
} else {
  console.error("Failed bulk upsert test", upsertResult);
  process.exit(1);
}

console.log("All tests passed!");
