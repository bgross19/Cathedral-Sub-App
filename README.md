# Cathedral Sub App

## Project Overview
The Cathedral Sub App is a custom Google Apps Script web application designed to manage substitute teacher requests, coverage assignments, and HR dashboard reporting. It integrates directly with a Google Sheet as its primary database and interfaces with the PowerSchool API to pull live Master Schedule data.

## Prerequisites
Before you begin, ensure you have the following:
*   A **Google Workspace account** with permissions to create Google Apps Script projects and Google Sheets.
*   **PowerSchool Admin access** to create and configure a custom PowerQuery endpoint and generate OAuth API credentials (Client ID and Client Secret).

## Installation & Deployment Instructions
1.  **Create a New Google Sheet:** Create a new, blank Google Sheet. This will act as the database for the application.
2.  **Open Apps Script Editor:** In the Google Sheet, go to `Extensions` > `Apps Script`.
3.  **Push or Copy Source Code:** Push the `.gs` and `.html` files using `clasp` (or copy them into the Apps Script editor):
    *   **Backend (`.gs`):** `core.gs`, `Setupdatabase.gs`, `powerschool.gs`, `absences.gs`, `sub_matching.gs`, `payload.gs`, `notifications.gs`, `admin_hr.gs`, `tests.gs`
    *   **Frontend (`.html`):** `Index.html`, `Views_Dashboards.html`, `Views_Modals.html`, `JS_Core.html`, `JS_Absences.html`, `JS_Admin.html`, `JS_HR.html`, `JS_Settings.html`
4.  **Deploy as Web App:**
    *   Click on the `Deploy` button in the top right corner and select `New deployment`.
    *   Select `Web app` as the type.
    *   Set the "Execute as" option to `User accessing the web app` (or `Me` depending on your organization's preference).
    *   Set "Who has access" to your organization or `Anyone` (with Google account).
    *   Click `Deploy` and grant the necessary permissions when prompted.
    *   Copy the Web App URL provided.

## Quick Start Guide (For First-Time Administrators)

### 1. Database Setup
Once you have pasted the code into the Apps Script editor, you need to initialize the database:
1.  Open `Setupdatabase.gs` in the editor.
2.  Select the `setupDatabase` function from the dropdown in the toolbar.
3.  Click the **Run** button. This will automatically create the necessary sheets and default settings in your connected Google Sheet.

### 2. Configure PowerSchool API
The application requires live Master Schedule data from PowerSchool.
1.  **Create PowerQuery:** In PowerSchool, create a custom PowerQuery plugin with the endpoint path: `/ws/schema/query/com.cathedral.subapp.masterschedule`. (Refer to `powerschool.gs` for the exact SQL query required).
2.  **Generate Credentials:** Generate OAuth client credentials (Client ID and Secret) in PowerSchool.
3.  **Update App Settings:** Go to the "Settings" sheet in your Google Sheet (or use the Admin Dashboard in the web app) and update the following settings:
    *   `PS_CLIENT_ID`: Your PowerSchool Client ID.
    *   `PS_CLIENT_SECRET`: Your PowerSchool Client Secret.
    *   `PS_URL`: The base URL of your PowerSchool instance (e.g., `https://powerschool.yourschool.edu`).
    *   `Term ID`: The current academic term ID (e.g., `3601`).

### 3. Initial Configuration
To start using the app, you need to populate the initial data:
*   **Staff Roster:** Go to the `Staff Roster` sheet and add your staff members (Name, Email, Duty). Alternatively, you can use the bulk upload feature in the Admin Settings Dashboard.
*   **Pay Periods:** Go to the `PayPeriods` sheet (or use the HR Dashboard) to configure the pay periods (Period Number, Start Date, End Date) for HR reporting.

### 4. Cache Warming Trigger
To optimize performance, the app caches the PowerSchool Master Schedule.
1. Open `powerschool.gs` in the editor.
2. Select the `setupCacheWarmingTrigger` function.
3. Click the **Run** button to create a time-driven trigger that refreshes the cache daily between 1 AM and 2 AM.

### 5. Email Queue Trigger
To enable the application to process and send automated email notifications in the background:
1. Open `notifications.gs` in the editor.
2. Select the `setupEmailQueueTrigger` function.
3. Click the **Run** button to create a time-driven trigger that processes the `Email Queue` sheet every minute.

### 6. Principal's Digest Trigger
To enable the application to send automated weekly summaries to principals:
1. Open `notifications.gs` in the editor.
2. Select the `setupPrincipalsDigestTrigger` function.
3. Click the **Run** button to create a time-driven trigger that runs every Friday around 12 PM.

## Architecture & Code Structure
The application is modularized into focused backend `.gs` files and frontend `.html` partials:
*   **`core.gs`**: Web App entry (`doGet`, `include`), sheet/settings caching, RBAC, `CacheService` chunking, date/color normalizers, and period bell time helpers.
*   **`Setupdatabase.gs`**: Initializes the Google Sheet database tabs (`Absence Requests`, `Staff Roster`, `SubstituteAvailability`, `Dates`, `PayPeriods`, `Payroll`, `Settings`, `Email Queue`, `Audit Log`).
*   **`powerschool.gs`** & **`powerschool-plugin/`**: Handles OAuth 2.0 authentication and PowerQuery sync (`com.cathedral.subapp.masterschedule`, `com.cathedral.subapp.bellschedule`, `com.cathedral.subapp.bellscheduletimes`) with date-aware Semester 1/2 (`3601`/`3603`) schedules.
*   **`absences.gs`**, **`sub_matching.gs`**, **`payload.gs`**: Absence request lifecycle, `LockService` concurrency, smart sub suggestions (`getSuggestedSubs`), and batched UI payloads.
*   **`notifications.gs`**: Transactional emails, `Email Queue` background delivery, `.ics` exports, and zero-click `CalendarApp` sub duty event sync (`syncSubDutyCalendarEvent`).
*   **`admin_hr.gs`**: Settings CRUD, Staff Roster management, Pay Period approval & `Payroll` sheet snapshot logging, Bell Schedule management, and historical data archiving.
*   **`tests.gs`**: Built-in unit & integration test suite (`runAllUnitTests`).
*   **`Index.html`**, **`Views_*.html`**, **`JS_*.html`**: Responsive Tailwind CSS views, modals, and client-side controllers.
