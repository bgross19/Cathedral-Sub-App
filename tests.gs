
/**
 * Advanced Custom Testing Framework
 */
var TestRunner = (function() {
  var testsPassed = 0;
  var testsFailed = 0;
  var failedTestDetails = [];
  var currentSuite = "";

  function assert(condition, message) {
    if (!condition) {
      throw new Error(message || "Assertion failed");
    }
  }

  function describe(suiteName, fn) {
    var previousSuite = currentSuite;
    currentSuite = suiteName;
    console.log("\n📘 Suite: " + suiteName);
    try {
      fn();
    } catch (e) {
      console.log("❌ Suite Failed to Execute: " + suiteName + " - " + e.message);
    }
    currentSuite = previousSuite;
  }

  function it(testName, fn) {
    var fullName = (currentSuite ? currentSuite + " > " : "") + testName;
    try {
      fn();
      testsPassed++;
      console.log("  ✅ " + testName);
    } catch (e) {
      testsFailed++;
      console.log("  ❌ " + testName + " - " + e.message);
      failedTestDetails.push({ name: fullName, error: e.message });
    }
  }

  function getSummary() {
    console.log("\n----------------------------------------");
    console.log("Test Summary: " + testsPassed + " passed, " + testsFailed + " failed");
    if (testsFailed > 0) {
      console.log("Failed Tests:");
      for (var j = 0; j < failedTestDetails.length; j++) {
        console.log("  - " + failedTestDetails[j].name + ": " + failedTestDetails[j].error);
      }
    }
    console.log("----------------------------------------");
    return { passed: testsPassed, failed: testsFailed };
  }

  function resetCounts() {
    testsPassed = 0;
    testsFailed = 0;
    failedTestDetails = [];
    currentSuite = "";
  }

  return {
    describe: describe,
    it: it,
    assert: assert,
    getSummary: getSummary,
    resetCounts: resetCounts
  };
})();

// Provide globals for easy access
var describe = TestRunner.describe;
var it = TestRunner.it;
var assert = TestRunner.assert;

/**
 * Test function to fetch the master schedule from PowerSchool and dump it into a temporary sheet.
 *
 * NOTE FOR POWERSCHOOL ADMIN:
 * Because PowerSchool's default API doesn't expose a clean, single endpoint for this,
 * you will need to create a "PowerQuery" plugin in PowerSchool with the endpoint path:
 * /ws/schema/query/com.cathedral.subapp.masterschedule
 *
 * You can base the PowerQuery on this provided SQL:
 * WITH DistinctClasses AS (
 *     -- Step 1: Get a clean list with only ONE row per teacher, per period, per course, per room
 *     SELECT DISTINCT
 *         t.LASTFIRST,
 *         t.EMAIL_Addr,
 *         REPLACE(cc.expression, '(A)', '') AS period,
 *         t.EMAIL_Addr || '-' || REPLACE(cc.expression, '(A)', '') AS email_period_join,
 *         c.course_name,
 *         s.room,  -- Added Room Number here
 *         t.id AS teacher_id
 *     FROM cc cc
 *     JOIN courses c
 *         ON c.course_number = cc.course_number
 *     JOIN teachers t
 *         ON cc.TEACHERID = t.id
 *     JOIN sections s                     -- New JOIN for the Sections table
 *         ON cc.sectionid = s.id          -- Linking the enrollment to the specific section
 *     WHERE cc.termid = :target_term      -- Use a parameterized term ID
 * )
 * -- Step 2: Combine the co-seated courses from that clean list
 * SELECT
 *     LASTFIRST,
 *     EMAIL_Addr,
 *     period,
 *     email_period_join,
 *     room,  -- Pulling the Room Number through to the final result
 *     LISTAGG(course_name, ' / ') WITHIN GROUP (ORDER BY course_name) AS course_names,
 *     MAX(teacher_id) AS teacher_id
 * FROM DistinctClasses
 * GROUP BY
 *     LASTFIRST,
 *     EMAIL_Addr,
 *     period,
 *     email_period_join,
 *     room   -- Grouping by Room Number as well
 * ORDER BY LASTFIRST
 */
function testPowerSchoolMasterScheduleFetch() {
  const token = getPowerSchoolToken();
  if (!token) {
    Logger.log("Failed to get PowerSchool token.");
    return;
  }

  const settings = typeof getSettings === "function" ? getSettings() : {};
  const rawUrl = settings['PS_URL'];
  const POWERSCHOOL_URL = rawUrl ? rawUrl.trim().replace(/\/$/, '') : null;

  if (!POWERSCHOOL_URL) {
    Logger.log("Missing PS_URL property.");
    return;
  }

  // The placeholder PowerQuery endpoint.
  // Update this if you name your PowerQuery differently.
  // By default, PowerSchool PowerQueries limit results to 100 records.
  // Appending ?pagesize=0 instructs it to return all records at once.
  const endpoint = "/ws/schema/query/com.cathedral.subapp.masterschedule?pagesize=0";

  const options = {
    method: "POST", // PowerQueries require POST, even for retrieving data
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    // PowerQueries require a JSON payload, even if empty, when using POST.
    payload: JSON.stringify({}),
    muteHttpExceptions: true
  };

  let responseText;
  let statusCode;
  try {
    let url = POWERSCHOOL_URL + endpoint;
    Logger.log('Fetching URL: ' + url);
    const response = UrlFetchApp.fetch(url, options);
    statusCode = response.getResponseCode();
    responseText = response.getContentText();
  } catch (error) {
    Logger.log("API Fetch Error: " + error.toString());
    responseText = error.toString();
    statusCode = "ERROR";
  }

  // Now write to Google Sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "PS Master Schedule Test";
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  // Set Headers
  sheet.appendRow(["LASTFIRST", "EMAIL_ADDR", "PERIOD", "ROOM", "COURSE_NAMES", "TERM"]);

  // If we have an error code (like 404 because the query doesn't exist yet)
  if (statusCode !== 200) {
    sheet.appendRow(["API ERROR", "Status Code: " + statusCode, responseText, "", "", ""]);
    Logger.log("API returned status: " + statusCode);
    return;
  }

  try {
    const json = JSON.parse(responseText);
    // PowerQueries usually return data in a `record` array.
    const records = json.record || [];

    if (records.length === 0) {
      sheet.appendRow(["NO DATA RETURNED", JSON.stringify(json), "", "", "", ""]);
    } else {
      const rows = records.map(r => {
        let periodRaw = String(r.period || r.PERIOD || "");
        let periodMatch = periodRaw.match(/\d+/);
        let periodClean = periodMatch ? periodMatch[0] : periodRaw.trim();

        return [
          r.lastfirst || r.LASTFIRST || "",
          r.email_addr || r.EMAIL_ADDR || "",
          periodClean,
          r.room || r.ROOM || r.room_number || r.ROOM_NUMBER || "",
          r.course_names || r.COURSE_NAMES || r.course_name || r.COURSE_NAME || "",
          "Target Term" // The actual term ID should be mapped if returned by the API
        ];
      });
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }
  } catch (parseError) {
    sheet.appendRow(["JSON PARSE ERROR", parseError.toString(), responseText, "", "", ""]);
    Logger.log("Failed to parse JSON: " + parseError.toString());
  }
}

function runTests() {
  TestRunner.resetCounts();

  // Mocks Setup
  var originalGetSettings = typeof getSettings !== 'undefined' ? getSettings : null;
  var originalGmailAppSendEmail;
  if (typeof GmailApp !== 'undefined' && GmailApp.sendEmail) {
    originalGmailAppSendEmail = GmailApp.sendEmail;
  } else {
    originalGmailAppSendEmail = function() {};
    if (typeof global !== 'undefined') global.GmailApp = { sendEmail: originalGmailAppSendEmail };
  }

  var mockSendEmailCalls = [];
  var mockLogs = [];
  var originalConsoleLog = console.log;

  function resetMocks() {
    mockSendEmailCalls = [];
    mockLogs = [];
    if (typeof GmailApp !== 'undefined') {
        GmailApp.sendEmail = function(to, subject, body, options) {
          mockSendEmailCalls.push({ to: to, subject: subject, body: body, options: options });
        };
    }
    console.log = function(msg) {
      mockLogs.push(msg);
      originalConsoleLog(msg); // still output to console
    };
  }

  describe("Email Helper Tests", function() {
    it("sendEmailHelper - Live mode", function() {
      resetMocks();
      getSettings = function() {
        return { "Email Mode": "Live", "Redirect Email": "bgross@example.com" };
      };

      sendEmailHelper("test@example.com", "Test Subject", "Test Body", { cc: "cc@example.com" });

      assert(mockSendEmailCalls.length === 1, "Expected sendEmail to be called once");
      var call = mockSendEmailCalls[0];
      assert(call.to === "test@example.com", "Expected original recipient");
      assert(call.subject === "Test Subject", "Expected original subject");
      assert(call.body === "Test Body", "Expected original body");
      assert(call.options && call.options.cc === "cc@example.com", "Expected CC to be preserved");
    });

    it("sendEmailHelper - Off mode", function() {
      resetMocks();
      getSettings = function() {
        return { "Email Mode": "Off", "Redirect Email": "bgross@example.com" };
      };

      sendEmailHelper("test@example.com", "Test Subject", "Test Body", { cc: "cc@example.com" });

      assert(mockSendEmailCalls.length === 0, "Expected sendEmail to NOT be called");
      var hasSuppressionLog = mockLogs.some(function(log) { return log.indexOf("Email sending is turned Off") !== -1; });
      assert(hasSuppressionLog, "Expected suppression log");
    });

    it("sendEmailHelper - Redirect mode without options", function() {
      resetMocks();
      getSettings = function() {
        return { "Email Mode": "Redirect", "Redirect Email": "redirect@example.com" };
      };

      sendEmailHelper("test@example.com", "Test Subject", "Test Body");

      assert(mockSendEmailCalls.length === 1, "Expected sendEmail to be called once");
      var call = mockSendEmailCalls[0];
      assert(call.to === "redirect@example.com", "Expected redirect recipient");
      assert(call.subject === "[REDIRECTED] Test Subject", "Expected modified subject");
      assert(call.body === "Test Body", "Expected unmodified body");
      assert(call.options === undefined, "Expected no options");
    });

    it("sendEmailHelper - Redirect mode with options (CC and BCC)", function() {
      resetMocks();
      getSettings = function() {
        return { "Email Mode": "Redirect", "Redirect Email": "redirect@example.com" };
      };

      var options = { cc: "cc@example.com", bcc: "bcc@example.com", htmlBody: "<p>Test</p>" };
      sendEmailHelper("test@example.com", "Test Subject", "Test Body", options);

      assert(mockSendEmailCalls.length === 1, "Expected sendEmail to be called once");
      var call = mockSendEmailCalls[0];
      assert(call.to === "redirect@example.com", "Expected redirect recipient");
      assert(call.subject === "[REDIRECTED] Test Subject", "Expected modified subject");
      assert(call.body.indexOf("[Original CC: cc@example.com]") !== -1, "Expected original CC in body");
      assert(call.options.htmlBody.indexOf("<em>[Original CC: cc@example.com]</em>") !== -1, "Expected original CC in htmlBody");
      assert(call.options.cc === undefined, "Expected CC to be removed from options");
      assert(call.options.bcc === undefined, "Expected BCC to be removed from options");
    });
  });

  describe("notifyAdminOfError", function() {
    it("should send email to admin using GmailApp", function() {
      resetMocks();

      // Mock global dependencies
      var originalSession = typeof Session !== 'undefined' ? Session : null;
      var mockEmail = "user@example.com";
      if (typeof global !== 'undefined') {
        global.Session = {
          getActiveUser: function() {
            return {
              getEmail: function() {
                return mockEmail;
              }
            };
          }
        };
      }

      getSettings = function() {
        return { "Redirect Email": "admin@example.com" };
      };

      var error = new Error("Test error message");
      error.stack = "Test stack trace";

      notifyAdminOfError("TestFunction", error);

      assert(mockSendEmailCalls.length === 1, "Expected sendEmail to be called once");
      var call = mockSendEmailCalls[0];
      assert(call.to === "admin@example.com", "Expected admin email");
      assert(call.subject === "Critical App Error: TestFunction", "Expected error subject");
      assert(call.body.indexOf("Function: TestFunction") !== -1, "Expected function name in body");
      assert(call.body.indexOf("User: user@example.com") !== -1, "Expected user email in body");
      assert(call.body.indexOf("Error Message: Test error message") !== -1, "Expected error message in body");
      assert(call.body.indexOf("Stack Trace:\nTest stack trace") !== -1, "Expected stack trace in body");
      assert(call.options === undefined, "Expected no options passed to GmailApp.sendEmail for this function");

      // Restore session
      if (typeof global !== 'undefined') {
        global.Session = originalSession;
      }
    });

    it("should not send email if admin email is missing", function() {
      resetMocks();

      getSettings = function() {
        return { "Redirect Email": "" };
      };

      var error = new Error("Test error message");

      notifyAdminOfError("TestFunction", error);

      assert(mockSendEmailCalls.length === 0, "Expected no email to be sent");
    });
  });

  describe("Lookup Builders", function() {
    it("buildNameLookup - Empty array", function() {
      var result = buildNameLookup([]);
      assert(Object.keys(result).length === 0, "Expected empty object");
    });

    it("buildNameLookup - Only header row", function() {
      var result = buildNameLookup([["Name", "Email", "Duty"]]);
      assert(Object.keys(result).length === 0, "Expected empty object");
    });

    it("buildScheduleLookup - Empty array", function() {
      var result = buildScheduleLookup([]);
      assert(Object.keys(result).length === 0, "Expected empty object");
    });

    it("buildScheduleLookup - Missing JOIN key", function() {
      var data = [
        ["ROOM", "COURSE_NAMES"],
        ["101", "Math"]
      ];
      var result = buildScheduleLookup(data);
      assert(Object.keys(result).length === 0, "Expected empty object when JOIN key is missing");
    });

    it("buildScheduleLookup - Missing Room/Course headers", function() {
      var data = [
        ["EMAIL_PERIOD_JOIN"],
        ["test@example.com_1"]
      ];
      var result = buildScheduleLookup(data);
      assert(result["test@example.com_1"] !== undefined, "Expected key to exist");
      assert(result["test@example.com_1"].room === "No Class Assigned", "Expected fallback for missing room header");
      assert(result["test@example.com_1"].course === "No Class Assigned", "Expected fallback for missing course header");
    });

    it("buildScheduleLookup - Happy path", function() {
      var data = [
        ["EMAIL_PERIOD_JOIN", "ROOM", "COURSE_NAMES"],
        ["user1@example.com_1", "101", "Math"],
        ["user2@example.com_2", "102", "Science"]
      ];
      var result = buildScheduleLookup(data);
      assert(result["user1@example.com_1"] !== undefined, "Expected key user1@example.com_1");
      assert(result["user1@example.com_1"].room === "101", "Expected room 101");
      assert(result["user1@example.com_1"].course === "Math", "Expected course Math");
      assert(result["user2@example.com_2"] !== undefined, "Expected key user2@example.com_2");
      assert(result["user2@example.com_2"].room === "102", "Expected room 102");
      assert(result["user2@example.com_2"].course === "Science", "Expected course Science");
    });

    it("buildScheduleLookup - Case/Whitespace handling", function() {
      var data = [
        ["EMAIL_PERIOD_JOIN", "ROOM", "COURSE_NAMES"],
        ["  USER@EXAMPLE.COM_1  ", "101", "Math"]
      ];
      var result = buildScheduleLookup(data);
      assert(result["user@example.com_1"] !== undefined, "Expected lowercased and trimmed key");
    });

    it("buildScheduleLookup - Empty/falsy room/course values", function() {
      var data = [
        ["EMAIL_PERIOD_JOIN", "ROOM", "COURSE_NAMES"],
        ["user1@example.com_1", "", false],
        ["user2@example.com_2", null, undefined]
      ];
      var result = buildScheduleLookup(data);
      assert(result["user1@example.com_1"].room === "No Class Assigned", "Expected fallback for empty string room");
      assert(result["user1@example.com_1"].course === "No Class Assigned", "Expected fallback for false course");
      assert(result["user2@example.com_2"].room === "No Class Assigned", "Expected fallback for null room");
      assert(result["user2@example.com_2"].course === "No Class Assigned", "Expected fallback for undefined course");
    });
  });

  // Teardown
  if (originalGetSettings) getSettings = originalGetSettings;
  if (typeof GmailApp !== 'undefined') {
    GmailApp.sendEmail = originalGmailAppSendEmail;
  }
  console.log = originalConsoleLog;

  return TestRunner.getSummary();
}

/**
 * Temporary benchmark function to measure execution time of enqueueEmail.
 */
function benchmarkEnqueueEmail() {
  var start = new Date().getTime();
  for (var i = 0; i < 100; i++) {
    enqueueEmail("test@example.com", "Benchmark Test", "This is a benchmark test.");
  }
  var end = new Date().getTime();
  Logger.log("Execution time for 100 iterations: " + (end - start) + " ms");
}
