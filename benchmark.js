const fs = require('fs');

// Mock Google Apps Script environment
let apiCallCount = 0;
let apiWaitTime = 0; // ms

function sleep(ms) {
    // In a real script we would sleep, but here we just add to a counter
    apiWaitTime += ms;
}

class MockRange {
    constructor(row, col, numRows, numCols, sheet) {
        this.row = row;
        this.col = col;
        this.numRows = numRows || 1;
        this.numCols = numCols || 1;
        this.sheet = sheet;
    }
    getValue() {
        apiCallCount++;
        apiWaitTime += 100;
        return this.sheet.data[this.row - 1][this.col - 1];
    }
    setValue(val) {
        apiCallCount++;
        apiWaitTime += 100;
        this.sheet.data[this.row - 1][this.col - 1] = val;
        return this;
    }
    getValues() {
        apiCallCount++;
        apiWaitTime += 100;
        let res = [];
        for (let i = 0; i < this.numRows; i++) {
            let r = [];
            for (let j = 0; j < this.numCols; j++) {
                r.push(this.sheet.data[this.row - 1 + i][this.col - 1 + j]);
            }
            res.push(r);
        }
        return res;
    }
    setValues(vals) {
        apiCallCount++;
        apiWaitTime += 100;
        for (let i = 0; i < this.numRows; i++) {
            for (let j = 0; j < this.numCols; j++) {
                this.sheet.data[this.row - 1 + i][this.col - 1 + j] = vals[i][j];
            }
        }
        return this;
    }
    clearContent() {
        apiCallCount++;
        apiWaitTime += 100;
        return this;
    }
}

class MockSheet {
    constructor(name, data) {
        this.name = name;
        this.data = JSON.parse(JSON.stringify(data)); // deep copy
    }
    getDataRange() {
        return new MockRange(1, 1, this.data.length, this.data[0].length, this);
    }
    getRange(row, col, numRows, numCols) {
        return new MockRange(row, col, numRows, numCols, this);
    }
    appendRow(row) {
        apiCallCount++;
        apiWaitTime += 100;
        this.data.push(row);
        return this;
    }
    deleteRow(rowPos) {
        apiCallCount++;
        apiWaitTime += 100;
        this.data.splice(rowPos - 1, 1);
        return this;
    }
    getLastRow() {
        return this.data.length;
    }
    getLastColumn() {
        return this.data[0].length;
    }
}

function runBenchmark(name, fn) {
    apiCallCount = 0;
    apiWaitTime = 0;
    const start = Date.now();
    fn();
    const duration = Date.now() - start;
    console.log(`--- ${name} ---`);
    console.log(`Wall time: ${duration}ms`);
    console.log(`API Call Count: ${apiCallCount}`);
    console.log(`Simulated API Wait Time: ${apiWaitTime}ms\n`);
    return { name, apiCallCount, apiWaitTime };
}

// -------------------------------------------------------------
// Benchmarks for processEmailQueue
// -------------------------------------------------------------
function processEmailQueueOld(sheet) {
    let data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    for (let i = 1; i < data.length; i++) {
        if (data[i][5] === "Pending") {
            // simulate sending email
            sheet.getRange(i + 1, 6).setValue("Sent");
        }
    }

    let rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) {
        let status = String(sheet.getRange(i + 1, 6).getValue() || "");
        if (status === "Sent" || status.indexOf("Failed") > -1) {
            rowsToDelete.push(i + 1);
        }
    }

    for (let j = 0; j < rowsToDelete.length; j++) {
        sheet.deleteRow(rowsToDelete[j]);
    }
}

function processEmailQueueNew(sheet) {
    let data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    let updates = [];
    let statuses = [];
    for (let i = 0; i < data.length; i++) {
        statuses.push([data[i][5]]);
    }

    for (let i = 1; i < data.length; i++) {
        if (data[i][5] === "Pending") {
            // simulate sending email
            statuses[i][0] = "Sent";
        }
    }

    // batch update statuses
    sheet.getRange(1, 6, statuses.length, 1).setValues(statuses);

    let rowsToDelete = [];
    for (let i = statuses.length - 1; i >= 1; i--) {
        let status = String(statuses[i][0] || "");
        if (status === "Sent" || status.indexOf("Failed") > -1) {
            rowsToDelete.push(i + 1);
        }
    }

    for (let j = 0; j < rowsToDelete.length; j++) {
        sheet.deleteRow(rowsToDelete[j]);
    }
}

// Data generator
let emailQueueData = [
    ["Timestamp", "To", "Subject", "Body", "Options", "Status"]
];
for(let i = 0; i < 50; i++) {
    emailQueueData.push([new Date(), "test@test.com", "Test", "Body", "{}", i % 2 === 0 ? "Pending" : "Sent"]);
}

let oldSheet = new MockSheet("Email Queue", emailQueueData);
runBenchmark("processEmailQueue (Original)", () => {
    processEmailQueueOld(oldSheet);
});

let newSheet = new MockSheet("Email Queue", emailQueueData);
runBenchmark("processEmailQueue (Optimized)", () => {
    processEmailQueueNew(newSheet);
});

// -------------------------------------------------------------
// Benchmarks for cancelMySubDuty
// -------------------------------------------------------------
function cancelMySubDutyOld(sheet, absenceId, period, userName) {
    let data = sheet.getDataRange().getValues();
    let targetUserName = userName.toLowerCase();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(absenceId)) {
            let subColumnIndex = 10 + parseInt(period) - 1;
            let assignedSub = String(sheet.getRange(i + 1, subColumnIndex).getValue() || "").trim();
            if (assignedSub.toLowerCase() === targetUserName) {
                sheet.getRange(i + 1, subColumnIndex).setValue("");
                return { success: true };
            }
        }
    }
}

function cancelMySubDutyNew(sheet, absenceId, period, userName) {
    let data = sheet.getDataRange().getValues();
    let targetUserName = userName.toLowerCase();
    let targetRow = -1;
    let targetCol = -1;

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(absenceId)) {
            let subColumnIndex = 10 + parseInt(period) - 1;
            let assignedSub = String(data[i][subColumnIndex - 1] || "").trim();
            if (assignedSub.toLowerCase() === targetUserName) {
                targetRow = i + 1;
                targetCol = subColumnIndex;
                break;
            }
        }
    }

    if (targetRow !== -1) {
        sheet.getRange(targetRow, targetCol).setValue("");
        return { success: true };
    }
}

let absenceData = [
    ["ID", "Timestamp", "Email", "Date", "Periods", "Reason", "Duration", "Urgency", "Instructions", "Sub1", "Sub2", "Sub3", "Sub4", "Sub5", "Sub6", "Sub7", "Sub8", "Status"]
];
for(let i = 0; i < 50; i++) {
    let row = [`ID_${i}`, new Date(), "teacher@test.com", "2023-01-01", "1", "Personal", "Full Day", "Standard", "", "sub_user", "", "", "", "", "", "", "", "Active"];
    absenceData.push(row);
}

let oldAbsenceSheet = new MockSheet("Absence Requests", absenceData);
runBenchmark("cancelMySubDuty (Original)", () => {
    cancelMySubDutyOld(oldAbsenceSheet, "ID_49", 2, "sub_user"); // period 2 -> index 10 + 2 - 1 = 11, data idx 10
});

let newAbsenceSheet = new MockSheet("Absence Requests", absenceData);
runBenchmark("cancelMySubDuty (Optimized)", () => {
    cancelMySubDutyNew(newAbsenceSheet, "ID_49", 2, "sub_user");
});
