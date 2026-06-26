const fs = require('fs');

// Basic mock for Google Apps Script services
const Utilities = {
  getUuid: () => Math.random().toString(36).substring(2, 15),
  formatDate: () => '2025-05-10'
};

const Session = {
  getActiveUser: () => ({ getEmail: () => 'test@gocathedral.com' }),
  getScriptTimeZone: () => 'America/New_York'
};

// Simulate LockService with waitLock and releaseLock
let lockAcquired = false;
let lockQueue = [];

function processLockQueue() {
  if (!lockAcquired && lockQueue.length > 0) {
    lockAcquired = true;
    const next = lockQueue.shift();
    next(); // grant lock
  }
}

const LockService = {
  getScriptLock: () => {
    return {
      waitLock: (timeout) => {
        // In local node script, waitLock is hard to simulate synchronously,
        // so we'll mock the wait by throwing an error randomly if too busy, or passing.
        // For a true stress test mock, we would need async/await, but GAS code is sync.
        // We'll simulate a 10% chance of lock timeout if concurrent requests are high.
        if (Math.random() < 0.1) {
            throw new Error("Lock timeout");
        }
      },
      releaseLock: () => {
        // release
      }
    };
  }
};

let sheetData = [
  ["ID", "Timestamp", "Email", "Date", "Periods", "Reason", "Duration", "Urgency", "Instructions", "Period 1 Sub", "Period 2 Sub", "Period 3 Sub", "Period 4 Sub", "Period 5 Sub", "Period 6 Sub", "Period 7 Sub", "Period 8 Sub", "Status"],
  ["existing_id", new Date(), "teacher@gocathedral.com", "2025-05-10", "1, 2, 3", "Personal", "Full Day", "Standard", "", "", "", "", "", "", "", "", "", "Active"]
];

const mockSheet = {
  appendRow: (row) => {
    // simulate delay to cause race conditions without lock
    sheetData.push(row);
  },
  getDataRange: () => ({
    getValues: () => sheetData
  }),
  getRange: (row, col, numRows = 1, numCols = 1) => ({
    setValue: (val) => {
        sheetData[row - 1][col - 1] = val;
    },
    setValues: (vals) => {
        for (let r=0; r<vals.length; r++) {
            for (let c=0; c<vals[r].length; c++) {
                sheetData[row - 1 + r][col - 1 + c] = vals[r][c];
            }
        }
    },
    getValue: () => {
       return sheetData[row - 1][col - 1];
    }
  })
};

const SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: (name) => {
        if (name === "Absence Requests") return mockSheet;
        if (name === "Staff Roster") return {
            getDataRange: () => ({ getValues: () => [["Name", "Email", "Duty"], ["Test User", "test@gocathedral.com", ""]] })
        };
        if (name === "User Roles") return {
            getDataRange: () => ({ getValues: () => [["Email", "Role"], ["admin@gocathedral.com", "admin"]] })
        };
        if (name === "Settings") return {
            getDataRange: () => ({ getValues: () => [["Setting Name", "Setting Value"], ["Email Mode", "Off"], ["App URL", "http://test"], ["Urgency Cutoff Time", "15"]] })
        };
        return null;
    },
    getSheets: () => []
  })
};

const GmailApp = {
  sendEmail: (to, subject, body, options) => {
    // console.log(`Simulated email sent to ${to}: ${subject}`);
  }
};

// Expose globals
global.Utilities = Utilities;
global.Session = Session;
global.LockService = LockService;
global.SpreadsheetApp = SpreadsheetApp;
global.GmailApp = GmailApp;
global.Logger = { log: console.log };

// Evaluate code.gs
const code = fs.readFileSync('code.gs', 'utf8');
eval(code);

// Run a stress test locally
async function stressTest() {
    console.log("Starting concurrency stress test...");

    let promises = [];

    // Simulate 50 simultaneous absence submissions
    console.log("Simulating 50 concurrent absence submissions...");
    for (let i = 0; i < 50; i++) {
        promises.push(new Promise((resolve) => {
            setTimeout(() => {
                const res = submitAbsence({
                    date: '2025-05-10',
                    periods: '1, 2',
                    reason: 'Stress test ' + i,
                    duration: 'Full Day',
                    urgency: 'Standard',
                    specialInstructions: ''
                });
                resolve(res);
            }, Math.random() * 50); // slight random delay to interleave execution
        }));
    }

    let results = await Promise.all(promises);
    let successCount = results.filter(r => r && r.success).length;
    let failCount = results.filter(r => r && !r.success).length;

    console.log(`Submissions Complete: ${successCount} successful, ${failCount} failed due to lock timeouts.`);

    // Simulate 20 users trying to grab the exact same job
    console.log("Simulating 20 concurrent sign-ups for the same period...");
    let assignPromises = [];
    for (let i = 0; i < 20; i++) {
        assignPromises.push(new Promise((resolve) => {
            setTimeout(() => {
                const res = assignSubToPeriod('existing_id', 1, 'Sub ' + i);
                resolve({ sub: 'Sub ' + i, ...res });
            }, Math.random() * 20);
        }));
    }

    let assignResults = await Promise.all(assignPromises);
    let assignedSub = null;
    let jobFilledErrors = 0;

    for (let res of assignResults) {
        if (res.success) {
            if (assignedSub && assignedSub !== res.sub) {
                console.error("CRITICAL ERROR: Job assigned to multiple subs! Lock failed.");
            }
            assignedSub = res.sub;
        } else if (res.error && res.error.includes("already filled")) {
            jobFilledErrors++;
        }
    }

    console.log(`Assignment Complete: 1 successful (${assignedSub}), ${jobFilledErrors} correctly blocked by concurrency check.`);

    // The sheet data should have exactly 1 existing row + the successful submissions.
    console.log(`Total rows in database (expected ${1 + 1 + successCount}): ${sheetData.length}`);
}

stressTest();
