# Servo Sales Dashboard (Primary, Secondary & Bazaar)

Streamlit dashboard for Servo lubricant sales reviews: target tracking, channel performance, a dedicated **Bazaar (retail counter) analysis** tab, insights, pivot, scheme calculator and HO-template PPT export. It runs fully offline on local Excel files.

> **No data is stored in this repository.** Sales, target and master workbooks stay in your local `Data` folder (see `.gitignore`).

## Tabs
| Tab | What it answers |
|---|---|
| 🏠 Overview | KPI cards with trend sparklines, target tracker (primary / secondary / Bazaar / RO / OEM), channel cards, auto-written highlights |
| 🛒 Bazaar | Counter health (new / reactivated / growing / declining / lost), dormant-counter call list, counter tiers, SSF scorecard, sub-group & district, product / category / pack mix, cross-sell gaps, top counters, counter 360° |
| 🏭 Primary / 🚚 Secondary | HO-template tables and charts |
| ⚖️ Primary vs Secondary | stock build at stockists |
| 💡 Insights | volume bridge, run-rate to target, achievement heat-map, counter & product movers, realisation |
| 🔎 Explore / 🧮 Pivot | click-to-drill questions and a free pivot builder |
| 🎯 Schemes · 📑 PPT Studio · 🗂️ Data & checks | incentive calculator, PPTX/Excel export, file and data checks |

## Tests
- `~/.local/venvs/sales_review_ho/bin/python selfcheck.py` - messy file drops (title rows, serial dates, corrupt/locked files).
- `~/.local/venvs/sales_review_ho/bin/python tests/smoke_test.py` - renders every tab headless under 15 period / filter combinations and fails if any section errors.

Every tab runs inside a guard (`hodash/ui.py`): if one section hits unexpected data it shows a small notice and the rest of the dashboard keeps working.


Start: double-click `run_dashboard.command` (or `~/.local/venvs/sales_review_ho/bin/streamlit run app.py`).

- Watches the **Sales Review** folder (change in `hodash/config.py` or env `SALES_DATA_DIR`).
- Drop `Primary*.xlsx` / `Secondary*.xlsx` (monthly or full-year). Overlapping months never double count; the file with the latest invoice date for a month wins. Page refreshes itself within ~20s.
- `Material master*.xlsx` / `Pack Master*.xlsx` (optional) improve category / segment mapping; without them it uses the Primary file's own Category/Segment columns and parses pack size from product names.
- Targets: read directly from the Target workbook (`Target*.xlsx`, sheet set in `config.py` -> `TARGET_SHEET`): primary blocks per stockist and secondary blocks per stockist x channel. Edit that workbook; the dashboard refreshes. A warning appears if a stockist's months don't add to the stated total.
- Business rules (ClearBlue excluded, ServoXpert in Bazaar, Dealer = RO, stockist names, pack/product groups): `hodash/config.py`, `hodash/loader.py`.
- PPT Studio tab: pick slides -> downloads PPTX on the HO template + Excel of the same tables.

## Sharing / deployment
- Just you: `Sales Dashboard.app` (in the Sales Review folder) - one click.
- Colleagues on the office network/VPN: run `share_on_network.command` on the PC that holds the files and send the printed link. They always see the latest data because the page watches the same folder.
- Always-on (team) use: run the same command on an office server / shared PC (e.g. as a login item or with `launchctl`), or package with Docker (`pip install -r requirements.txt && streamlit run app.py`). Keep it inside IOCL's network - do not put the sales files on public hosting (e.g. Streamlit Community Cloud) without IT approval.

## Independence
No AI service, internet call or cloud dependency at runtime: only Python libraries (streamlit, pandas, plotly, python-pptx, openpyxl, python-calamine) reading your local Excel files.

## If something looks wrong
Run `~/.local/venvs/sales_review_ho/bin/python selfcheck.py` - it simulates messy drops (extra pivot sheets, title rows, text/serial dates, corrupt or locked files) on a temporary copy and reports PASS/FAIL. Unreadable files are skipped and listed as warnings on the dashboard instead of stopping it.
