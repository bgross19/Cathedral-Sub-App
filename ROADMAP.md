# Cathedral Sub Coverage App Roadmap

This roadmap outlines the next phase of feature development and infrastructure improvements for the Cathedral Sub Coverage application. Having addressed the majority of our initial technical debt, these items focus on proactive user experience enhancements, administrative accountability, and backend performance.

## 1. Automated Background Tasks (Cache Warming)
Currently, the Master Schedule is fetched from PowerSchool and cached for 6 hours. If the cache expires, the next user to load the app experiences a significant delay while the data is fetched.
*   **Action Item:** Implement a time-based trigger (Cron Job) via Google Apps Script that runs in the background (e.g., every 2 hours).
*   **Benefit:** This script will silently refresh the Master Schedule cache before it expires. Users will always experience immediate load times, as they will never be the ones waiting for the fresh PowerSchool API response.

## 2. Modern UI/UX Enhancements
While the application is fully functional, the user interface relies on some legacy browser behaviors that interrupt workflow. Furthermore, as data grows, viewing on small screens can be challenging.
*   **Action Item - Toast Notifications:** Replace all native browser `alert()` and `confirm()` dialogs. We will implement a custom "Toast" notification system for non-blocking success/error messages (e.g., "✅ Substitute Assigned!") and styled modal dialogs for critical confirmations (e.g., deleting a request).
*   **Action Item - Mobile Responsiveness Audit:** Conduct a focused review of the application on mobile viewports. We will implement responsive design patterns—such as card-based layouts for smaller screens instead of horizontal scrolling tables—ensuring the Admin and HR dashboards are easily usable on phones.
*   **Benefit:** A polished, modern app feel that doesn't freeze the screen with alerts, and allows administrators to easily manage coverage on-the-go from their mobile devices.

## 3. Advanced Data Exporting
Administrators and HR currently rely on a "Copy this data" button to move information to their clipboards for pasting into Excel.
*   **Action Item:** Add dedicated "Export to CSV" and potentially "Export to PDF" functionality directly within the Admin, HR, and At-A-Glance dashboards.
*   **Benefit:** Provides a standard, foolproof way to generate reports and save physical or digital records of sub coverage and payroll data without relying on clipboard formatting.

## 4. Comprehensive Audit Logging
Currently, while we track the current state of absence requests, we do not have a historical log of who made specific changes and when.
*   **Action Item:** Create a hidden "Audit Log" Google Sheet. Update all backend data mutation functions (assigning a sub, cancelling a request, changing a user's role, etc.) to append a row to this sheet.
*   **Data Captured:** Timestamp, Actor (User Email), Action Type (e.g., "ASSIGN_SUB"), Target ID (Absence ID/Period), and Details ("Assigned John Doe to Period 3").
*   **Benefit:** Essential for accountability and debugging. If a sub assignment is accidentally changed or a request deleted, administrators can consult the log to see exactly who performed the action and when.
