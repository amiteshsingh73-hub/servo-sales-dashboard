# Maintenance guide - Servo Sales Dashboard (works without any AI)

## Monthly routine (5 minutes)
1. Export SAP Primary and DMS Secondary reports. Either drop them on **📤 Add sales files** in the dashboard sidebar (they are recognised, checked and saved into `Data/` for you), or save them into `Sales Review/Data/` yourself named `Primary Sales <Mon><YY>.xlsx` / `Secondary Sales <Mon><YY>.xlsx`. Either way the open dashboard refreshes within 20 seconds; nothing else is needed (no AI, no internet).
2. Update the Target workbook in `Data/` (one file only; newest `Target*` is used).
3. Start the dashboard: double-click `Start Sales Review.command` (or `run_dashboard.command`).
4. Open the **Data & checks** tab: confirm file counts, latest invoice date and zero warnings.
5. Review, then **PPT Studio** → *Start from* (HO template, or last month's filled deck to keep what you typed) → **Build the HO deck** → download. The status table lists the slides left for hand entry. When HO sends a new template, put it in `Data/` (name `HO PPT format*TEMPLATE*.pptx`); slides are found by title and columns by heading, so small layout changes need no code change.

## Review Letters (monthly SSR Bazaar review letter + mail)
- Tab **✉️ Review Letters**. Defaults to the latest complete month in the Primary/Secondary files; covers Apr → that month vs LY (like-for-like) and target, with the month itself highlighted.
- The first time a new complete month shows up, the letters for all SSRs are written automatically to `Sales Review/Reports/Review Letters <Mon-YY>/` (Word letter + lost/declining counter list per SSR). Delete that folder to have it regenerated; or use **Save** / **ZIP** in the tab after editing the settings.
- Tone is set by the data: Strong (Bazaar < 95% of target or below LY), Watch (latest month > 10% below LY, or lost counters > 2× new), Good (appreciation).
- Ref. no., date, address, previous-letter reference, reply days and new-counter target are editable in "Letter settings". Default addresses: `LETTER_ADDRESS` in `hodash/config.py`.
- Letterhead: put `Letterhead*.docx` (IOCL letterhead with header/footer) in `Data/` and the letters are printed on it. Mail addresses: `Data/SSR Contacts.xlsx` (To / CC, Signature sheet).
- Code: `hodash/review_letter.py` (numbers, wording, Word file, mail) and `hodash/review_letter_ui.py` (screen).

## Area report (segment / customer / area / period workbook)
- Double-click `Build Area Report.command` after dropping a new Secondary file (and a newer `*Customer Master*.xlsx`) in `Data/`. ~20 s; output in `Sales Review/Reports/`.
- Area mapping lives in `Data/Area Master.xlsx` (pincode -> Area -> Zone; customer pincode overrides for ROs). Fill the yellow cells and rebuild. Seed values: `hodash/areas.py` (only used when that file does not exist).
- The script builds the file in Excel's local folder (~/Library/Containers/com.microsoft.Excel/Data/tmp/area_report), lets Excel calculate it there, then copies it into Reports. Never let Excel recalc inside OneDrive: older cloud copies come back over the new file. `--no-excel` skips the recalc.

## Once a year (April): roll over the year
- `hodash/config.py` -> `TARGET_SHEET` is `"FY2526 Base Target"`. Checked 20-Oct-2026: despite its name, that sheet holds the **FY 26-27** targets (FY25-26 sales as "BASE", months Apr'26-Mar'27), so it is correct for this year. Next April, point it at whichever sheet holds the FY 27-28 blocks.
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
~/.local/venvs/sales_review_ho/bin/python selfcheck.py               # messy-file merges, should end ALL CHECKS PASSED
~/.local/venvs/sales_review_ho/bin/python tests/smoke_test.py        # every tab x 18 filter combinations, ~3 min -> ALL SCENARIOS PASSED
~/.local/venvs/sales_review_ho/bin/python tests/data_scenarios.py    # 12 messy Data folders, ~2 min -> ALL DATA SCENARIOS PASSED
~/.local/venvs/sales_review_ho/bin/python tests/fuzz_widgets.py --quick   # changes every filter on every tab, ~25 min -> NO CRASHES FOUND
```
All of them work on copies / read-only, never on your real files. Last verified: 27-Oct-2026 (v3.1.1), all passed.

## Version
Shown at the bottom of the sidebar and on Data & checks. Lives in `hodash/version.py`; every release gets a line in `CHANGELOG.md` and a git tag (`git tag v3.1.1`).

## If a tab shows "could not be shown for this selection"
The rest of the dashboard keeps working. The full error is in `Sales Review/logs/dashboard_errors.log` (newest at the bottom) - usually a sales file whose column headings changed. Compare the headings with `REQUIRED` in `hodash/loader.py`.

## Where the Python environment lives
`~/.local/venvs/sales_review_ho` (outside OneDrive on purpose: ~15,000 small files in the folder made OneDrive and Spotlight slow the Mac). The launchers create it automatically on first run if missing. It is rebuilt from `requirements.txt`; nothing in it is precious.

## If it will not start
- Look at `Sales Review/logs/ho_dashboard.log` (HO dashboard, port 8502) or `sales_dashboard.log` (port 8501) - the last lines show the error - and `dashboard_errors.log`.
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
