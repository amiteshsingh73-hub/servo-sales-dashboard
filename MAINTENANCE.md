# Maintenance guide - Servo Sales Dashboard (works without any AI)

## Monthly routine (5 minutes)
1. Export SAP Primary and DMS Secondary reports. Save into `Sales Review/Data/` named `Primary Sales <Mon><YY>.xlsx` / `Secondary Sales <Mon><YY>.xlsx`.
2. Update the Target workbook in `Data/` (one file only; newest `Target*` is used).
3. Start the dashboard: double-click `Start Sales Review.command` (or `run_dashboard.command`).
4. Open the **Data & checks** tab: confirm file counts, latest invoice date and zero warnings.
5. Review, then use **PPT Studio** to download the HO-template deck.

## Once a year (April): roll over the year
- `hodash/config.py` -> `TARGET_SHEET` is currently `"FY2526 Base Target"`. When the new year's target block sheet exists in the Target workbook, change the name here. If targets look wrong or stale, this is the first thing to check.
- Replace `HO PPT format LSE STANDARD TEMPLATE ....pptx` in `Data/` with the new template.
- `STOCKIST_ALIASES`, `CHANNEL_MAP`, `PACK_GROUPS`: edit in `hodash/config.py` when a stockist, customer group or pack size is added.

## Where things live
| Need | File |
|---|---|
| Business rules, mappings, paths | `hodash/config.py` |
| Reading/merging Excel files | `hodash/loader.py` |
| Target parsing | `hodash/targets.py` |
| KPIs and maths | `hodash/metrics.py` |
| Bazaar (counter) analysis | `hodash/bazaar.py`, `bazaar_ui.py` |
| PPT/Excel export | `hodash/pptx_export.py` |
| Incentive calculator | `hodash/schemes.py`, `scheme_rules.py` (rules in `Scheme Rules.xlsx`) |
| Screens | `app.py`, `hodash/ui.py` |

## Health checks (run after ANY change or after a Python/library update)
```
cd "Sales Review/sales_dashboard"
~/.local/venvs/sales_review_ho/bin/python selfcheck.py          # messy-file handling, should end ALL CHECKS PASSED
~/.local/venvs/sales_review_ho/bin/python tests/smoke_test.py   # renders every tab, ~2 min, should end ALL SCENARIOS PASSED
```
Last verified: 19-Oct-2026, both passed.

## Where the Python environment lives
`~/.local/venvs/sales_review_ho` (outside OneDrive on purpose: ~15,000 small files in the folder made OneDrive and Spotlight slow the Mac). The launchers create it automatically on first run if missing. It is rebuilt from `requirements.txt`; nothing in it is precious.

## If it will not start
- Look at `Sales Review/logs/sales_dashboard.log` (last lines show the error).
- Rebuild the environment (safe, data is untouched):
  ```
  cd "Sales Review/sales_dashboard"; rm -rf ~/.local/venvs/sales_review_ho
  python3 -m venv ~/.local/venvs/sales_review_ho && ~/.local/venvs/sales_review_ho/bin/pip install -r requirements.txt
  ```
- Python 3.11 is required (install from python.org if the Mac has been replaced).
- A tab shows "section hit unexpected data": a file layout changed. Compare the column names with `loader.py`; selfcheck helps isolate it.

## Backups
- Double-click `backup.command`: zips everything (the Python environment is not in the folder at all) into `~/Documents/SalesReview_Backups/`. Do this monthly and before editing code.
- Code is also under git (`git log`, `git checkout <file>` restores a file). Sales data is deliberately not in git.
- OneDrive keeps version history for the Data files (right-click -> Version history).

## Safety rules
- Keep sales data inside IOCL's network; do not host publicly.
- Edit one thing, run the two checks, then continue. Keep a `*_backup_*.py` copy of any config you change (these are git-ignored).
