import re

def main():
    with open('code.gs', 'r') as f:
        content = f.read()

    # Find the end of getSettings or anywhere appropriate to inject global error handler
    # We will inject the global error handler at the end of the file.

    global_handler = """
/**
 * Executes a function and catches any unhandled exceptions to notify the admin.
 * @param {Function} func - The function to execute.
 * @param {string} funcName - The name of the function for logging.
 * @returns {any} The result of the function.
 */
function withGlobalExceptionHandler(funcName, func) {
  return function() {
    try {
      return func.apply(this, arguments);
    } catch (e) {
      console.error("Global Error in " + funcName + ": " + e.message + "\\nStack: " + e.stack);

      try {
        var settings = getSettings();
        var adminEmail = settings["Redirect Email"];
        if (adminEmail && adminEmail.trim() !== "") {
          var subject = "Critical App Error: " + funcName;
          var body = "An error occurred in the Cathedral Sub App.\\n\\n" +
                     "Function: " + funcName + "\\n" +
                     "User: " + Session.getActiveUser().getEmail() + "\\n" +
                     "Error Message: " + e.message + "\\n\\n" +
                     "Stack Trace:\\n" + e.stack;

          MailApp.sendEmail({
            to: adminEmail,
            subject: subject,
            body: body
          });
        }
      } catch (mailError) {
        console.error("Failed to send admin error email: " + mailError.message);
      }

      throw e; // Re-throw so frontend still sees an error
    }
  };
}
"""
    if 'withGlobalExceptionHandler' not in content:
        content += "\n" + global_handler + "\n"

    # Now we need to wrap the main entry points exposed to the frontend (google.script.run)
    # The frontend calls:
    # updateSettings
    # getUserRoles
    # editUserRole
    # deleteUserRole
    # addUserRole
    # assignSubToPeriod
    # cancelMySubDuty
    # updateAbsence
    # submitAbsence
    # cancelAbsence
    # getSettingsForFrontend
    # getInitialPayload

    # But we can't just wrap the function declaration easily.
    # The best way is to rename the original function to e.g., `_getInitialPayload`
    # and redefine `getInitialPayload` as `withGlobalExceptionHandler("getInitialPayload", _getInitialPayload);`
    # However, apps script handles bound functions better if they are just normal declarations.

    # Alternatively, we can wrap the bodies of these functions in `try { ... } catch(e) { globalErrorHandler(e, funcName); throw e; }`

    error_handler = """
function notifyAdminOfError(funcName, e) {
  console.error("Global Error in " + funcName + ": " + e.message + "\\nStack: " + e.stack);
  try {
    var settings = getSettings();
    var adminEmail = settings["Redirect Email"];
    if (adminEmail && adminEmail.trim() !== "") {
      var subject = "Critical App Error: " + funcName;
      var body = "An error occurred in the Cathedral Sub App.\\n\\n" +
                 "Function: " + funcName + "\\n" +
                 "User: " + Session.getActiveUser().getEmail() + "\\n" +
                 "Error Message: " + e.message + "\\n\\n" +
                 "Stack Trace:\\n" + e.stack;

      MailApp.sendEmail({
        to: adminEmail,
        subject: subject,
        body: body
      });
    }
  } catch (mailError) {
    console.error("Failed to send admin error email: " + mailError.message);
  }
}
"""

    if 'notifyAdminOfError' not in content:
        content = error_handler + "\n" + content

    # Let's just modify the `catch` block of `getInitialPayload` for now,
    # as that's the most critical one requested in TECH_DEBT.md.

    # In getInitialPayload, there's a try { ... } catch (err) { ... }
    # We'll replace the catch block in getInitialPayload

    target_catch = """  } catch (err) {
    throw new Error("Failed to get initial payload: " + err.message);
  }"""

    new_catch = """  } catch (err) {
    notifyAdminOfError("getInitialPayload", err);
    throw new Error("Failed to get initial payload: " + err.message);
  }"""

    content = content.replace(target_catch, new_catch)

    with open('code.gs', 'w') as f:
        f.write(content)

if __name__ == "__main__":
    main()
