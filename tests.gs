function runTests() {
  var testsPassed = 0;
  var testsFailed = 0;
  var failedTestDetails = [];

  // Mocks
  var originalGetSettings = getSettings;
  var originalGmailAppSendEmail;
  if (typeof GmailApp !== 'undefined' && GmailApp.sendEmail) {
    originalGmailAppSendEmail = GmailApp.sendEmail;
  } else {
    // If running outside Apps Script or GmailApp is not available
    originalGmailAppSendEmail = function() {};
    global.GmailApp = { sendEmail: originalGmailAppSendEmail };
  }
  var originalConsoleLog = console.log;

  var mockSendEmailCalls = [];
  var mockLogs = [];

  function resetMocks() {
    mockSendEmailCalls = [];
    mockLogs = [];
    GmailApp.sendEmail = function(to, subject, body, options) {
      mockSendEmailCalls.push({ to: to, subject: subject, body: body, options: options });
    };
    console.log = function(msg) {
      mockLogs.push(msg);
    };
  }

  function assert(condition, message) {
    if (!condition) {
      throw new Error(message);
    }
  }

  // Test Suite
  var tests = [
    {
      name: "sendEmailHelper - Live mode",
      run: function() {
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
      }
    },
    {
      name: "sendEmailHelper - Off mode",
      run: function() {
        getSettings = function() {
          return { "Email Mode": "Off", "Redirect Email": "bgross@example.com" };
        };

        sendEmailHelper("test@example.com", "Test Subject", "Test Body", { cc: "cc@example.com" });

        assert(mockSendEmailCalls.length === 0, "Expected sendEmail to NOT be called");
        assert(mockLogs.length === 1, "Expected a log message");
        assert(mockLogs[0].indexOf("Email sending is turned Off") !== -1, "Expected suppression log");
      }
    },
    {
      name: "sendEmailHelper - Redirect mode without options",
      run: function() {
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
      }
    },
    {
      name: "sendEmailHelper - Redirect mode with options (CC and BCC)",
      run: function() {
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
      }
    }
  ];

  // Run tests
  for (var i = 0; i < tests.length; i++) {
    resetMocks();
    try {
      tests[i].run();
      testsPassed++;
      originalConsoleLog("✅ Passed: " + tests[i].name);
    } catch (e) {
      testsFailed++;
      originalConsoleLog("❌ Failed: " + tests[i].name + " - " + e.message);
      failedTestDetails.push({ name: tests[i].name, error: e.message });
    }
  }

  // Restore mocks
  getSettings = originalGetSettings;
  if (typeof GmailApp !== 'undefined') {
    GmailApp.sendEmail = originalGmailAppSendEmail;
  }
  console.log = originalConsoleLog;

  // Summary
  console.log("----------------------------------------");
  console.log("Test Summary: " + testsPassed + " passed, " + testsFailed + " failed");
  if (testsFailed > 0) {
    console.log("Failed Tests:");
    for (var j = 0; j < failedTestDetails.length; j++) {
      console.log("  - " + failedTestDetails[j].name + ": " + failedTestDetails[j].error);
    }
  }
  console.log("----------------------------------------");

  return {
      passed: testsPassed,
      failed: testsFailed
  };
}
