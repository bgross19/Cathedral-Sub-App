const vm = require('vm');
const fs = require('fs');

const code = fs.readFileSync('code.gs', 'utf8');

const mockDatesData = {
    "2023-10-16": "Green",
    "2023-10-17": "Blue",
    "2023-10-18": "Gold",
    "2023-10-19": "Green",
    "2023-10-20": "Blue",
    "2023-10-25": "Green"
};

const mockSS = {
    getSheetByName: (name) => {
        if (name === "Staff Roster") return {
            getDataRange: () => ({
                getValues: () => [
                    ["Name", "Email", "Role", "Duty"],
                    ["Principal Bob", "bob@school.edu", "Principal", ""],
                    ["Teacher Ann", "ann@school.edu", "Teacher", ""],
                    ["Teacher Dan", "dan@school.edu", "Teacher", "2"],
                    ["Sub Sally", "sally@school.edu", "Substitute", ""]
                ]
            })
        };
        if (name === "Absence Requests") return {
            getDataRange: () => ({
                getValues: () => [
                    // ID, Timestamp, Email, Date, Periods, Reason, Duration, Urgency, Instructions, P1, P2, P3, P4, P5, P6, P7, P8, P0, Adv, Status
                    ["Headers...", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
                    ["1", "ts", "ann@school.edu", "2023-10-17", "1, 2", "Sick", "Half Day (AM)", "High", "", "Sub Sally", "Teacher Dan - Duty", "", "", "", "", "", "", "", "", "Active"],
                    ["2", "ts", "dan@school.edu", "2023-10-19", "3, 4", "Personal", "Full Day", "Low", "", "", "", "Teacher Ann", "Sub Sally", "", "", "", "", "", "", "Active"],
                    ["3", "ts", "ann@school.edu", "2023-10-25", "1, 2", "Sick", "Full Day", "High", "", "Sub Sally", "Teacher Dan", "", "", "", "", "", "", "", "", "Active"] // Next week
                ]
            })
        };
        if (name === "Settings") return {
            getDataRange: () => ({
                getValues: () => [
                    ["Setting Name", "Setting Value"],
                    ["Green Day Pay Rate", "10"],
                    ["Blue/Gold Day Pay Rate", "20"],
                    ["App URL", "http://app.test"]
                ]
            })
        };
        if (name === "Dates") return {
            getDataRange: () => ({
                getValues: () => [
                    ["Date", "Day Type"],
                    ["2023-10-16", "Green"],
                    ["2023-10-17", "Blue"],
                    ["2023-10-18", "Gold"],
                    ["2023-10-19", "Green"],
                    ["2023-10-20", "Blue"],
                    ["2023-10-25", "Green"]
                ]
            })
        }
        return null;
    }
};

const context = {
    console: console,
    SpreadsheetApp: {
        getActiveSpreadsheet: () => mockSS,
        openById: () => mockSS
    },
    CacheService: {
        getScriptCache: () => ({
            get: () => null,
            put: () => {}
        })
    },
    getSS: () => mockSS,
    // Do not override getSettings or getDatesData, let the code use the native ones with mock SS and Cache
    DEFAULT_APP_URL: "http://app.test",
    getSubColumnIndex: (p) => {
        if (p === '1') return 10;
        if (p === '2') return 11;
        if (p === '3') return 12;
        if (p === '4') return 13;
        return 0;
    },
    Utilities: {
        formatDate: (d, tz, format) => {
            if (!d || isNaN(d.getTime())) return "";
            if (format === "MMM d") {
                const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                return months[d.getMonth()] + " " + d.getDate();
            }
            if (format === "yyyy-MM-dd") {
                const pad = (n) => n < 10 ? '0' + n : n;
                return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
            }
            return d.toString();
        }
    },
    Session: {
        getScriptTimeZone: () => 'GMT'
    }
};

vm.createContext(context);
vm.runInContext(code, context);

try {
    // 2023-10-20 is a Friday
    let html = context.generatePrincipalsDigestHTML(new Date("2023-10-20T12:00:00Z"));
    if (html.includes("Teacher Ann</strong>: 0.5 days and 2 periods")) console.log("Ann absences OK");
    if (html.includes("Teacher Dan</strong>: 1 days and 2 periods")) console.log("Dan absences OK");
    if (html.includes("Teacher Dan</strong>: covered 1 class</li>")) console.log("Dan coverage OK");
    if (html.includes("Teacher Ann</strong>: covered 1 class ($10.00 extra pay)</li>")) console.log("Ann coverage OK");
    if (html.includes("Sub Sally</strong>: covered 2 classes</li>")) console.log("Sally coverage OK");

    // Test with string date format "2023-10-20"
    let htmlStr = context.generatePrincipalsDigestHTML("2023-10-20");
    if (htmlStr.includes("Teacher Ann</strong>: 0.5 days and 2 periods") && htmlStr.includes("Teacher Dan</strong>: 1 days and 2 periods")) {
        console.log("String date parsing OK");
    } else {
        console.error("String date parsing failed!");
    }
} catch(e) {
    console.error(e);
}
