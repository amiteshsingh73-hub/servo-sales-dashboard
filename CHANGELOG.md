# Changelog - Servo Sales Dashboard

The version is shown at the bottom of the sidebar and on the **Data & checks** tab (`hodash/version.py`).

## v3.1.1 - 27-Oct-2026 - stability release
**Crash-proofing**
- Whole-page safety net: `app.py` now runs the dashboard (`dashboard.py`) inside a guard, so even an error outside a tab shows a short message with a *Try again* button instead of a broken page. Every tab keeps its own guard as before.
- Every caught error is also written to `Sales Review/logs/dashboard_errors.log` (with the technical details) so it can be checked later without the dashboard open.
- Sales files with a changed layout no longer stop the dashboard: missing optional columns in the Secondary report (SSF, district, sub-group, sales area, taxable value, ...) or the Primary report (stockist name, description) are filled with blanks instead of rejecting the whole file.
- A broken, locked or half-synced **Target** workbook no longer blocks loading; a newer Target file that is unreadable or laid out wrongly falls back to the previous good one (with a warning).
- A broken Material / Pack master no longer blocks loading (the sales files' own product columns are used).
- The Data folder being missing (OneDrive not signed in) or a file disappearing while OneDrive replaces it is handled like an empty folder / a later re-check.

**Automatic file pickup**
- New **📤 Add sales files** box in the sidebar (and on the start screen when no data is loaded): drop Primary / Secondary / Target / Material / Pack files; each is recognised from its column headings, checked, and saved into `Data/` with a clear name. Duplicates are ignored. The dashboard refreshes itself.
- Files that could not be opened because OneDrive was still syncing them are retried automatically every 90 seconds for up to 15 minutes (before, they stayed skipped until another file changed).
- Files dropped straight into `Data/` (Finder, OneDrive web, e-mail attachment saved there) are still picked up within 20 seconds.

**Testing** (run after any change - see MAINTENANCE.md)
- `tests/smoke_test.py` - every tab under 18 period / filter combinations.
- `tests/data_scenarios.py` - 12 messy data folders (empty, missing, corrupt, half-synced, missing columns, new half month, no last-year history, one stockist).
- `tests/fuzz_widgets.py` - changes every filter / selector on every tab and re-renders everything after each change.
- `selfcheck.py` - messy file merges.

## v3.1.0 - Oct-2026
Premium & Customers tab (segment / period / customer / product-series filters incl. HYPER), LRS Format tab (SSR-wise Excel, potential from market share), Review Letters tab, Area Report, Bazaar counter tiers with counter lists, chart label fixes.

## v3.0 - 19-Oct-2026
New look (KPI cards, target tracker, pill tabs), Bazaar analysis tab, per-tab crash guards, smoke test, published on GitHub.
