# Agent Rules & Project Conventions

## Git & Pull Request Workflow
- **Always use a branch and Pull Request instead of pushing directly to `main`.**
  1. Create a descriptive branch from `main` before committing changes (e.g., `feat/<topic>` or `fix/<topic>`).
  2. Push the branch to `origin` (`git push -u origin <branch-name>`).
  3. Open a Pull Request against `main` using `gh pr create` with a clear summary of changes and verification steps so there is a permanent record on GitHub.
- **Local Tooling Paths**:
  - `git`: `/Users/bgross/.gemini/antigravity/scratch/.tools/env/bin/git`
  - `gh`: `/Users/bgross/.gemini/antigravity/scratch/.tools/env/bin/gh`
  - JS syntax & unit test runner: `/System/Library/Frameworks/JavaScriptCore.framework/Versions/Current/Helpers/jsc`

## Domain Conventions (Cathedral Sub App)
- **Advisory / Period 10**:
  - Always store and process Advisory internally as `"10"` (sheet values, checkbox values, period sorting at `1.5`, and `parsePeriod`).
  - Always display `"10"` as `"Advisory"` in user-facing UI and emails using the display helpers (`formatPeriodDisplay`, `formatPeriodLabelDisplay`, `formatPeriodsListDisplay` in `core.gs`, and `formatPeriodLabel`, `formatPeriodValue`, `formatPeriodsList` in `JS_Core.html`).
