# Technical Debt and Potential Issues

## 1. Duplicated Frontend Code
The `Index.html` file and `script.js` file have significant duplication of JavaScript logic. It seems the frontend logic was copied from `Index.html` into `script.js` but both files retain the `<script>` blocks. Any changes to the frontend must currently be made in both places, which is highly error-prone. This needs to be refactored so `Index.html` properly sources `script.js` or uses Apps Script's `HtmlService.createHtmlOutputFromFile()` includes.

## 2. Redundant Backend Data Loading Functions
Several backend functions are now obsolete because their functionality was consolidated into `getInitialPayload()`. These functions still exist in the frontend and backend, adding bloat:
* `getStaffList()` and `loadStaffList()`
* `getTodaysOpenJobsData()` and `loadTodaysOpenJobs()`
* `getQuickCoverData()` and `loadQuickCover()`
* `getAdminDashboardData()` and `loadAdminDashboardData()`
* `getMySubDuties()` and `loadMySubDuties()`
* `getMyAbsences()` and `loadMyAbsences()`
* `getHRDashboardData()` and `loadHRDashboardData()`
(Note: some load functions might still be needed for manual refreshes, but they currently just call the old individual backend functions which adds `google.script.run` overhead).

## 3. Lack of True Testing Framework
The current testing setup (`tests.gs`) uses a custom built assert function and mocks services manually. A proper testing framework (like Jest or Clasp with TypeScript) would be far more robust and easier to maintain.

## 4. Hardcoded Term ID in PowerSchool Query
The PowerSchool SQL query in the comment of `testPowerSchoolMasterScheduleFetch` has `termid = 3503` hardcoded. While `getMasterScheduleData` uses a dynamic Term ID, the comment and documentation reflect a hardcoded value, which could mislead future developers.

## 5. Potential Cache Expiration Issue
The Master Schedule cache in `getMasterScheduleData` is set to 21600 seconds (6 hours). If a teacher changes rooms or classes mid-day, the change won't reflect in the system for up to 6 hours unless the cache is manually cleared.

## 6. Large Data Transfer on Initial Load
`getInitialPayload()` fetches *all* admin, HR, and quick cover data on initial load. As the `Absences` sheet grows over the year, this payload will become massive, leading to slow startup times. Pagination or lazy-loading (e.g., only loading this week's absences initially) should be implemented.
## 7. Repeated Master Schedule Fetches
Several functions in `code.gs` call `getMasterScheduleData()`. While it is cached, pulling a 100KB JSON string from `CacheService` and parsing it via `JSON.parse` is still an expensive operation. If multiple operations happen in a single execution (which `getInitialPayload` avoids but others might not), it will be repeatedly parsed.

## 8. Missing Scripting for UI Updates
When `adminData` is updated, the frontend needs to manually ensure that all related dashboards (At A Glance, Period View, Request View) are refreshed. A state management library or more robust pub-sub system would reduce the risk of views falling out of sync.
## 9. Inconsistent usage of `getSheetOrThrow`
The `code.gs` file has `getSheetOrThrow` and `getSheetCaseInsensitiveOrThrow` defined, but many functions (like `getMyAbsences`, `getMySubDuties`, `getTodaysOpenJobsData`, `getInitialPayload`, etc.) still use `ss.getSheetByName()`. This should be standardized across the codebase so that missing sheets throw explicit, traceable errors instead of causing silent failures later in the execution.

## 10. Lack of Global Exception Handling / Observability
Currently, some blocks use `try-catch`, and some errors are logged to `console.error` (which maps to Stackdriver/Apps Script dashboard logs), but there is no centralized alerting. If `getInitialPayload` fails entirely, the user sees a raw error string, and there's no proactive notification sent to an administrator.
