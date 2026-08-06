1.  **Update `assignSubToPeriod` in `code.gs` to check role.**
    *   Currently, the function checks the `SubstituteAvailability` sheet for *anyone* assigned to cover a class.
    *   We need to pull the newly assigned person's role from the Staff Roster.
    *   If their role contains "Substitute", proceed with the existing `SubstituteAvailability` sheet check.
    *   If their role *does not* contain "Substitute", they are likely a Teacher. We should fetch their `teacherSchedule` from the master schedule data.
    *   If the requested `period` is in their `teacherSchedule` (using `getScheduleJoinPeriod` to account for 0/Advisory), throw the `"Sub not listed as available, proceed?"` warning.
    *   If the requested `period` is not in their schedule, they are free, so bypass the warning and allow the assignment.
2.  **Complete pre commit steps**
    *   Ensure proper testing, verifications, reviews, and reflections are done via `pre_commit_instructions`.
3.  **Submit the change.**
    *   Commit the code with a descriptive message and submit.
