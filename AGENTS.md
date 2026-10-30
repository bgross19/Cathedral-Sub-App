# Agent Rules & Project Conventions

## Git & Pull Request Workflow
- **Always use a branch and Pull Request instead of pushing directly to `main`.**
  1. Create a descriptive branch from `main` before committing changes (e.g., `feat/<topic>` or `fix/<topic>`).
  2. Push the branch to `origin` (`git push -u origin <branch-name>`).
  3. Open a Pull Request against `main` using `gh pr create` with a clear summary of changes and verification steps so there is a permanent record on GitHub.
  4. Push the updated `.gs` and `.html` files directly to Google Apps Script and update the live Web App deployment by running `/Users/bgross/.gemini/antigravity/scratch/.tools/bin/push-gas --deploy "<summary>"`.
- **Local Tooling Paths**:
  - `git`: `/Users/bgross/.gemini/antigravity/scratch/.tools/env/bin/git`
  - `gh`: `/Users/bgross/.gemini/antigravity/scratch/.tools/env/bin/gh`
  - `clasp`: `/Users/bgross/.gemini/antigravity/scratch/.tools/env/bin/clasp` (auth file at `/Users/bgross/.gemini/antigravity/scratch/.tools/.clasprc.json`)
  - `push-gas`: `/Users/bgross/.gemini/antigravity/scratch/.tools/bin/push-gas` (pushes all `.gs`, `.html`, and `appsscript.json` files to Google Apps Script `@HEAD`; pass `--deploy "Description"` to also create a new version and update the live Web App deployment `AKfycbwKZrBo4R-9O97aVNCjOHk9PddWCb6XNKviDS1lj4nNc49khl3T9OL8pGUDa7E1XE0`).
  - JS syntax & unit test runner: `/System/Library/Frameworks/JavaScriptCore.framework/Versions/Current/Helpers/jsc`

## Domain Conventions (Cathedral Sub App)
- **Advisory / Period 10**:
  - Always store and process Advisory internally as `"10"` (sheet values, checkbox values, period sorting at `1.5`, and `parsePeriod`).
  - Always display `"10"` as `"Advisory"` in user-facing UI and emails using the display helpers (`formatPeriodDisplay`, `formatPeriodLabelDisplay`, `formatPeriodsListDisplay` in `core.gs`, and `formatPeriodLabel`, `formatPeriodValue`, `formatPeriodsList` in `JS_Core.html`).
