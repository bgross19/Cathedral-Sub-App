import re

def main():
    with open('code.gs', 'r') as f:
        content = f.read()

    # Modify getInitialPayload to use the Data Fetch Window setting.
    # It reads getSettings() anyway in the UI data sections? Wait, does getInitialPayload read settings?
    # No, it reads absence requests.
    # Let's see getInitialPayload logic:
    # `var today = new Date();`
    # `var todayTime = today.getTime();`
    # `var settings = getSettings();` -> Let's add this.
    # `var fetchWindow = parseInt(settings["Data Fetch Window (Days)"]) || 30;`
    # `var cutoffDate = new Date(todayTime - (fetchWindow * 24 * 60 * 60 * 1000));`
    # And then add `if (rowDate < cutoffDate) continue;` to the loops.

    # First, let's look at getInitialPayload

    # We will inject the logic at the top of getInitialPayload
    target_start = """    var targetEndToday = new Date(today);
    targetEndToday.setHours(23, 59, 59, 999);"""

    new_start = """    var targetEndToday = new Date(today);
    targetEndToday.setHours(23, 59, 59, 999);

    var settings = getSettings();
    var fetchWindowDays = parseInt(settings["Data Fetch Window (Days)"]);
    if (isNaN(fetchWindowDays)) fetchWindowDays = 30; // default to 30 days
    var cutoffDate = new Date(today);
    cutoffDate.setDate(today.getDate() - fetchWindowDays);
    cutoffDate.setHours(0, 0, 0, 0);"""

    content = content.replace(target_start, new_start)

    # Now in the loop for absences (My Absences / Sub Duties)
    # The loop starts with:
    #       for (var i = 1; i < absenceData.length; i++) {
    #         var row = absenceData[i];
    #         if (String(row[17] || "").trim() === "Canceled") continue;
    #
    #         var dateVal = row[3];
    #         if (!dateVal) continue;
    #         var rowDate = new Date(dateVal);
    #         if (isNaN(rowDate.getTime())) continue;

    # Add if (rowDate < cutoffDate) continue;
    target_loop1 = """        var rowDate = new Date(dateVal);
        if (isNaN(rowDate.getTime())) continue;"""
    new_loop1 = """        var rowDate = new Date(dateVal);
        if (isNaN(rowDate.getTime())) continue;
        if (rowDate < cutoffDate) continue;"""

    content = content.replace(target_loop1, new_loop1)

    # The Admin Data loop also needs this.
    #       for (var i = 1; i < absenceData.length; i++) {
    #         var row = absenceData[i];
    #         if (String(row[17] || "").trim() === "Canceled") continue;
    #
    #         var rowTeacherEmail = String(row[2]).toLowerCase().trim();

    target_loop2 = """      // Admin Dashboard Data
      var adminData = [];
      for (var i = 1; i < absenceData.length; i++) {
        var row = absenceData[i];
        if (String(row[17] || "").trim() === "Canceled") continue;

        var rowTeacherEmail = String(row[2]).toLowerCase().trim();
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
        var dateStr = row[3];
        var dateObj = new Date(dateStr);"""

    new_loop2 = """      // Admin Dashboard Data
      var adminData = [];
      for (var i = 1; i < absenceData.length; i++) {
        var row = absenceData[i];
        if (String(row[17] || "").trim() === "Canceled") continue;

        var dateStr = row[3];
        var dateObj = new Date(dateStr);
        if (!isNaN(dateObj.getTime()) && dateObj < cutoffDate) continue;

        var rowTeacherEmail = String(row[2]).toLowerCase().trim();
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;"""

    content = content.replace(target_loop2, new_loop2)

    # HR Data loop
    #       for (var i = 1; i < absenceData.length; i++) {
    #         var row = absenceData[i];
    #         if (String(row[17] || "").trim() === "Canceled") continue;
    #
    #         var rowTeacherEmail = String(row[2]).toLowerCase().trim();
    #         var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
    #         var dateStr = row[3];
    #         var dateObj = new Date(dateStr);

    target_loop3 = """      for (var i = 1; i < absenceData.length; i++) {
        var row = absenceData[i];
        if (String(row[17] || "").trim() === "Canceled") continue;

        var rowTeacherEmail = String(row[2]).toLowerCase().trim();
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;
        var dateStr = row[3];
        var dateObj = new Date(dateStr);"""

    new_loop3 = """      for (var i = 1; i < absenceData.length; i++) {
        var row = absenceData[i];
        if (String(row[17] || "").trim() === "Canceled") continue;

        var dateStr = row[3];
        var dateObj = new Date(dateStr);
        if (!isNaN(dateObj.getTime()) && dateObj < cutoffDate) continue;

        var rowTeacherEmail = String(row[2]).toLowerCase().trim();
        var teacherName = nameLookup[rowTeacherEmail] || rowTeacherEmail;"""

    content = content.replace(target_loop3, new_loop3)

    with open('code.gs', 'w') as f:
        f.write(content)

if __name__ == "__main__":
    main()
