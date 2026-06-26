import re

def main():
    with open('tests.gs', 'r') as f:
        content = f.read()

    new_test_framework = """
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
    console.log("\\n📘 Suite: " + suiteName);
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
    console.log("\\n----------------------------------------");
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
"""

    with open('tests.gs', 'w') as f:
        f.write(new_test_framework)

if __name__ == "__main__":
    main()
