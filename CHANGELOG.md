# Changelog - Servo Sales Dashboard

The version is shown at the bottom of the sidebar and on the **Data & checks** tab (`hodash/version.py`).

## v3.2.1 - 03-Nov-2026 - works without AI or internet
- `Repair Sales Review.command`: rebuilds both Python environments from the offline copy; puts the old one back if the rebuild fails.
- `requirements-lock.txt` pins all 44 libraries (before: only the 8 main ones, so a rebuild could pick up untested versions). Offline copies in `~/.local/wheels/` (sales_review_ho, sales_master); launchers install from there first, internet only as fallback, and use Python 3.11 explicitly.
- Month / fiscal-year labels no longer crash on a decimal or blank month (Explore and Data & checks errors of 02-Nov).
- `START HERE.txt` in Sales Review: one-page use and troubleshooting guide.

## v3.2.0 - 02-Nov-2026 - complete HO deck
**PPT Studio fills the whole HO LSE-interaction template** (`hodash/ho_deck.py`)
- All 37 slides of the HO standard template, in the template's order: cover date, area map, industry size (IOCL row in MT), highlights, stockist profile, Target vs Achievement (15 parameters), primary by segment / stockist / stockist target-vs-achievement by year / premium / category / Hyper series / SERVO & Genuine oil by brand / product group / pack / premium plan / bucket grades, secondary by channel / RO / Bazaar / OEM brand / OEM penetration / Fleet, Bazaar counters SSR- and district-wise, counter details (serviced, new, inactive, nil selling, net addition), sales per counter, the two Bazaar tracking slides (volume / invoice / SKU buckets), IG / ServoXpert. Activity slides (marketing, inspection, development, visits, initiatives, concerns) are kept for manual entry. Your own "SSR Wise Channel Wise" slide is added after the Bazaar slide.
- Slides are found by title and columns by heading, tables are resized to the real stockists + Total, the HO formatting of every cell is kept.
- **Start from your last filled deck**: numbers are refreshed, everything typed by hand is kept (universe, plans, profile, typed premium / product targets, older years), and "Apr-Sep" style period labels on refreshed slides follow the new period.
- Numbers checked against the 04-Oct-2026 deck: identical on every data slide except grease (sales file has 5.1 MT for FY25-26).
- Built on a button click (no slow-down elsewhere); report of what was filled / left for hand entry; Excel copy of all tables.

**Fixes**
- A blank-dated row in a newly added sales file turned the year column into decimals and broke the Explore and Data & checks tabs - fixed in the loader (years / months always whole numbers).
- Bazaar counter counts leave out ServoXpert workshops (their volume stays in Bazaar) - as in the HO deck.
- "Servosystem HLP 68" counts as Other Industrial, not System 68 - as in the HO deck.
- Primary keeps SAP net weight (MT) for MT figures.
- selfcheck builds its test months from the latest Primary file.

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
