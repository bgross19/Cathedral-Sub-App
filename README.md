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
3.  **Copy Source Code:** Copy the contents of the `.gs` and `.html` files from this repository into the Apps Script editor. Specifically:
    *   `Code.gs` -> `Code.gs`
    *   `Setupdatabase.gs` -> `Setupdatabase.gs`
    *   `powerschool.gs` -> `powerschool.gs`
    *   `Index.html` -> `Index.html`
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
    *   `Term ID`: The current academic term ID (e.g., `3503`).

### 3. Initial Configuration
To start using the app, you need to populate the initial data:
*   **Staff Roster:** Go to the `Staff Roster` sheet and add your staff members (Name, Email, Duty). Alternatively, you can use the bulk upload feature in the Admin Settings Dashboard.
*   **User Roles:** Go to the `User Roles` sheet and assign roles to users (Admin, HR, Sub Coordinator, Principal). Give yourself the `Admin` role to access the Settings Dashboard.
*   **Pay Periods:** Go to the `PayPeriods` sheet (or use the HR Dashboard) to configure the pay periods (Period Number, Start Date, End Date) for HR reporting.

### 4. Cache Warming Trigger
To optimize performance, the app caches the PowerSchool Master Schedule.
1. Open `powerschool.gs` in the editor.
2. Select the `setupCacheWarmingTrigger` function.
3. Click the **Run** button to create a time-driven trigger that refreshes the cache daily between 1 AM and 2 AM.

## Architecture & Code Structure
The application is structured into the following key files:
*   **`Index.html`**: The frontend UI of the web application. It handles routing, rendering dashboards, modals, and user interactions using vanilla JavaScript and Tailwind CSS.
*   **`code.gs`**: The main backend logic. It contains functions for handling absence requests, assigning subs, fetching payload data, validating permissions, and processing email notifications.
*   **`powerschool.gs`**: Handles all interactions with the PowerSchool API, including OAuth authentication, querying the Master Schedule via PowerQuery, and caching the results.
*   **`Setupdatabase.gs`**: A one-time utility script used to initialize the Google Sheet with the required tabs (Absence Requests, Staff Roster, User Roles, etc.) and default headers.
