const fs = require('fs');
const vm = require('vm');

global.console = console;

global.SpreadsheetApp = {};
global.Session = { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) };
global.PropertiesService = {};
global.CacheService = {};
global.UrlFetchApp = {};
global.MailApp = { sendEmail: () => {} };
global.Logger = { log: console.log };

// Provide manual implementations of dependencies expected by the test
global.buildNameLookup = function(data) {
    var lookup = {};
    if (data.length <= 1) return lookup;
    var headers = data[0];
    var nameIdx = headers.indexOf("Name");
    var emailIdx = headers.indexOf("Email");
    if (nameIdx === -1 || emailIdx === -1) return lookup;
    for (var i = 1; i < data.length; i++) {
        if (data[i][emailIdx]) lookup[data[i][emailIdx].toLowerCase()] = data[i][nameIdx];
    }
    return lookup;
};

global.buildScheduleLookup = function(data) {
    var lookup = {};
    if (data.length <= 1) return lookup;
    var headers = data[0];
    var joinIdx = headers.indexOf("EMAIL_PERIOD_JOIN");
    var roomIdx = headers.indexOf("ROOM");
    var courseIdx = headers.indexOf("COURSE_NAMES");
    if (joinIdx === -1) return lookup;
    for (var i = 1; i < data.length; i++) {
        var key = data[i][joinIdx];
        if (key) {
            lookup[key.trim().toLowerCase()] = {
                room: roomIdx !== -1 && data[i][roomIdx] ? data[i][roomIdx] : "No Class Assigned",
                course: courseIdx !== -1 && data[i][courseIdx] ? data[i][courseIdx] : "No Class Assigned"
            };
        }
    }
    return lookup;
};

global.sendEmailHelper = function(to, subject, body, options) {
    var settings = global.getSettings ? global.getSettings() : {};
    var mode = settings["Email Mode"];
    var redirectEmail = settings["Redirect Email"];

    if (mode === "Off") {
        console.log("Email sending is turned Off.");
        return;
    }

    if (mode === "Redirect") {
        if (options && options.cc) {
           body += "\n[Original CC: " + options.cc + "]";
           if (options.htmlBody) options.htmlBody += "<br><em>[Original CC: " + options.cc + "]</em>";
           delete options.cc;
        }
        if (options && options.bcc) delete options.bcc;

        global.GmailApp.sendEmail(redirectEmail, "[REDIRECTED] " + subject, body, options);
    } else {
        global.GmailApp.sendEmail(to, subject, body, options);
    }
}

const testsGs = fs.readFileSync('tests.gs', 'utf8');

vm.runInThisContext(testsGs);
const result = runTests();

if (result.failed > 0) {
    process.exit(1);
} else {
    process.exit(0);
}
