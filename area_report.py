"""Secondary sales workbook: segment-wise, customer-wise, area-wise and period-wise.

Run:  ~/.local/venvs/sales_review_ho/bin/python area_report.py      (or double-click `Build Area Report.command`)

Reads from Sales Review/Data:
  Secondary*.xlsx                 DMS secondary sales (same merge + ClearBlue rule as the HO dashboard)
  *Customer Master*.xlsx          Lubes Customer Master Report (pincode, status, mobile)
  Area Master.xlsx                pincode -> area -> zone, plus customer pincode overrides (created on first run)
Writes Sales Review/Reports/Secondary Sales - Segment Area Customer <Mon-YY>.xlsx

All summary tabs are SUMIFS/COUNTIFS formulas over the 'Data' and 'Customers' tabs, driven by the
From/To months on the Dashboard, so one workbook answers any period.
"""
from __future__ import annotations
import re, shutil, sys, time
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
from hodash import config as C, loader  # noqa: E402
from hodash.areas import PINCODES, RO_SEED, ZONE_ORDER  # noqa: E402

AREA_MASTER = C.DATA_DIR / "Area Master.xlsx"
OUT_DIR = C.DATA_DIR.parent / "Reports"
EXCEL_WORK = Path.home() / "Library/Containers/com.microsoft.Excel/Data/tmp/area_report"
UNMAPPED = "Unmapped"

# ---------------------------------------------------------------- styles
FONT = "Arial"
NAVY, ORANGE, GREY, LIGHT = "1F3864", "F37021", "F2F2F2", "DCE6F1"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=16, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=10, italic=True, color="595959")
F_SEC = Font(name=FONT, size=12, bold=True, color=NAVY)
F_INPUT = Font(name=FONT, size=11, bold=True, color="0000FF")
F_KPI = Font(name=FONT, size=18, bold=True, color=NAVY)
F_KPIL = Font(name=FONT, size=9, color="595959")
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_TOT = PatternFill("solid", fgColor=LIGHT)
FILL_ZONE = PatternFill("solid", fgColor="FCE4D6")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_KPI = PatternFill("solid", fgColor=GREY)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
KL = '#,##0.0;[Red]-#,##0.0;"-"'
LIT = '#,##0;[Red]-#,##0;"-"'
PCT = '0.0%;[Red]-0.0%;"-"'
GR = '+0.0%;[Red]-0.0%;0.0%'
CNT = '#,##0;-#,##0;"-"'
DATE, MON = "dd-mmm-yy", "mmm-yy"
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def fy_label(fy: int) -> str:
    return f"FY{(fy - 1) % 100:02d}-{fy % 100:02d}"


def clean_id(x) -> str:
    s = str(x).strip()
    return s[:-2] if s.endswith(".0") else s


def clean_pin(x) -> str | None:
    s = clean_id(x) if x is not None and not (isinstance(x, float) and pd.isna(x)) else ""
    return s if re.fullmatch(r"[1-9]\d{5}", s) else None


# ---------------------------------------------------------------- inputs
def load_secondary():
    files = sorted(loader.discover()["secondary"], key=lambda p: p.stat().st_mtime)
    frames, problems, clearblue = [], [], 0.0
    for f in files:
        raw, prob = loader.read_table(f, "secondary")
        if prob:
            problems.append(prob); continue
        sku = loader._col(raw, "SKU Name").astype(str).str.upper()
        qty = loader._num_col(loader._col(raw, "Product Quantity(L/Kg)")).fillna(0)
        clearblue += float(qty[sku.str.contains("CLEARBLUE", regex=False)].sum())
        frames.append((f.name, loader.norm_secondary(raw)))
    if not frames:
        sys.exit("No Secondary*.xlsx file found in " + str(C.DATA_DIR))
    sec = loader._merge_sources(frames)
    sec["customer_id"] = sec.customer_id.map(clean_id)
    sec["litres"] = sec.kl * 1000
    sec["month"] = sec.date.values.astype("datetime64[M]")
    sec["ssr"] = sec.ssf_name.replace({"-": "Not assigned", "nan": "Not assigned"}).str.strip()
    sec["customer"] = sec.customer.replace({"nan": None, "None": None})
    return sec, [f.name for f in files], problems, clearblue


def load_customer_master():
    files = [f for f in C.DATA_DIR.iterdir() if re.search(r"customer master", f.name, re.I)
             and f.suffix.lower() in (".xlsx", ".xls") and not f.name.startswith("~$")]
    if not files:
        return pd.DataFrame(columns=["customer_id"]), None
    f = max(files, key=lambda p: p.stat().st_mtime)
    raw = pd.read_excel(f, header=None, dtype=str, engine="calamine")
    h = next(i for i in range(min(25, len(raw))) if "Customer Id" in raw.iloc[i].astype(str).str.strip().tolist())
    m = raw.iloc[h + 1:].copy(); m.columns = [str(c).strip() for c in raw.iloc[h]]
    m = m.dropna(subset=["Customer Id"])
    m["customer_id"] = m["Customer Id"].map(clean_id)
    m["pin"] = m["Postal Code"].map(clean_pin)
    m = m.sort_values("Customer Status").drop_duplicates("customer_id", keep="first")   # ACTIVE first
    return m.set_index("customer_id"), f.name


def load_area_master(sec, cm):
    """Reads Area Master.xlsx (creates it on first run) and appends rows that need the user's input."""
    pins, locs, changed = {}, {}, False
    if AREA_MASTER.exists():
        wb = load_workbook(AREA_MASTER, data_only=True)
        for r in wb["Pincode Area"].iter_rows(min_row=2, values_only=True):
            p = clean_pin(r[0])
            if p: pins[p] = (str(r[1] or "").strip(), str(r[2] or "").strip(), str(r[3] or "").strip())
        for r in wb["Customer Location"].iter_rows(min_row=2, values_only=True):
            if r[0]:
                locs[clean_id(r[0])] = {"name": r[1], "segment": r[2], "stockist": r[3], "pin": clean_pin(r[4]),
                                        "area": str(r[5] or "").strip(), "zone": str(r[6] or "").strip(), "note": r[7]}
    else:
        changed = True
        pins = {p: (a, z, "") for p, (a, z) in PINCODES.items()}
        for cid, (p, nm) in RO_SEED.items():
            locs[cid] = {"name": nm, "segment": "RO", "stockist": "", "pin": p, "area": "", "zone": "",
                         "note": "Seeded from FY25-26 'Taluka wise sales of SSR doshi.xlsx'"}
    # pincodes seen in the master that are not mapped yet
    for p in sorted(set(cm["pin"].dropna()) - set(pins)) if len(cm) else []:
        pins[p] = ("", "", "NEW - fill Area and Zone"); changed = True
    # customers with recent sales but no usable pincode
    recent = sec[sec.fy >= sec.fy.max() - 1]
    last = recent.sort_values("date").groupby("customer_id").tail(1).set_index("customer_id")
    for cid, r in last.iterrows():
        has_pin = cid in cm.index and pd.notna(cm.at[cid, "pin"])
        if not has_pin and cid not in locs:
            locs[cid] = {"name": r.customer, "segment": r.channel, "stockist": r.stockist, "pin": None, "area": "",
                         "zone": "", "note": "NEW - no pincode in Customer Master: fill Pincode (or Area + Zone)"}
            changed = True
    if changed:
        try:
            write_area_master(pins, locs)
        except PermissionError:
            print("! Area Master.xlsx is open in Excel - new rows were not added this time.")
    return pins, locs


def write_area_master(pins, locs):
    wb = Workbook()
    ws = wb.active; ws.title = "How to use"
    notes = [
        ("Area Master - edit this file, then rebuild the report", F_TITLE),
        ("", None),
        ("Pincode Area: one row per pincode. Area = locality shown in the report, Zone = sales territory.", F_BASE),
        ("   Rows marked 'NEW' are pincodes found in the Customer Master that are not mapped yet - fill Area and Zone.", F_BASE),
        ("Customer Location: customers whose pincode is missing in the Customer Master (all ROs, a few counters).", F_BASE),
        ("   Fill Pincode, or type Area and Zone directly. Area/Zone typed here win over the pincode mapping.", F_BASE),
        ("   You can also add any customer here to move it to a different area.", F_BASE),
        ("Yellow cells are the ones to fill. Keep the column order. Zone names used today:", F_BASE),
    ] + [("   " + z, F_BASE) for z in ZONE_ORDER if z != UNMAPPED]
    for i, (t, f) in enumerate(notes, 1):
        ws.cell(i, 1, t).font = f or F_BASE
    ws.column_dimensions["A"].width = 110

    def sheet(title, headers, rows, fill_cols, widths):
        s = wb.create_sheet(title)
        s.append(headers)
        for c in s[1]:
            c.font, c.fill, c.alignment = F_HEAD, FILL_HEAD, CENTER
        for r in rows:
            s.append(list(r))
            for ci in fill_cols:
                c = s.cell(s.max_row, ci)
                if c.value in (None, ""): c.fill = FILL_INPUT
        for i, w in enumerate(widths, 1):
            s.column_dimensions[L(i)].width = w
        for row in s.iter_rows():
            for c in row:
                if c.font != F_HEAD: c.font = F_BASE
        s.freeze_panes = "A2"; s.auto_filter.ref = s.dimensions
        return s

    zrank = {z: i for i, z in enumerate(ZONE_ORDER)}
    prow = sorted(pins.items(), key=lambda kv: (kv[1][0] == "", zrank.get(kv[1][1], 99), kv[0]))
    sheet("Pincode Area", ["Pincode", "Area", "Zone", "Note"], [(p, a, z, n) for p, (a, z, n) in prow], [2, 3], [10, 34, 40, 40])
    lrow = sorted(locs.items(), key=lambda kv: (bool(kv[1]["pin"] or kv[1]["area"]), str(kv[1]["segment"]), str(kv[1]["name"])))
    sheet("Customer Location", ["Customer Id", "Customer Name", "Segment", "SSR", "Pincode", "Area (optional)", "Zone (optional)", "Note"],
          [(k, v["name"], v["segment"], v["stockist"], v["pin"], v["area"], v["zone"], v["note"]) for k, v in lrow],
          [5], [18, 40, 10, 24, 10, 26, 34, 60])
    for row in wb["Customer Location"].iter_rows(min_row=2):
        row[0].number_format = "@"
    wb.save(AREA_MASTER)


# ---------------------------------------------------------------- model
def build_model():
    sec, files, problems, clearblue = load_secondary()
    cm, cm_file = load_customer_master()
    pins, locs = load_area_master(sec, cm)

    last = sec.sort_values("date").groupby("customer_id").tail(1).set_index("customer_id")
    cust = pd.DataFrame(index=last.index)
    mname = cm["Customer Name"] if len(cm) else pd.Series(dtype=str)
    cust["name"] = [mname.get(i) if pd.notna(mname.get(i)) else (n or i) for i, n in zip(last.index, last.customer)]
    cust["segment"] = last.channel
    cust["group"] = last.cust_group.replace({"Servoxpert": "ServoXpert", "Nayara Ro": "Nayara RO"})
    cust["subseg"] = last.cust_subgroup.replace({"-": "Not specified", "nan": "Not specified"})
    cust["stockist"], cust["ssr"], cust["district"] = last.stockist, last.ssr, last.district
    pin, area, zone, src = [], [], [], []
    for cid in cust.index:
        o = locs.get(cid, {})
        p = o.get("pin") or (cm.at[cid, "pin"] if cid in cm.index else None)
        p = p if isinstance(p, str) else None
        a, z = (o.get("area") or ""), (o.get("zone") or "")
        s = "Area Master" if (o.get("pin") or a) else ("Customer Master" if p else "No pincode")
        if not a and p and pins.get(p, ("", ""))[0]:
            a, z = pins[p][0], z or pins[p][1]
        if not a:
            a = f"Pincode {p}" if p else f"No pincode ({cust.at[cid, 'district']})"
        pin.append(p or ""); area.append(a); zone.append(z or UNMAPPED); src.append(s)
    cust["pin"], cust["area"], cust["zone"], cust["loc_src"] = pin, area, zone, src
    if len(cm):
        cust["status"] = [cm.at[i, "Customer Status"] if i in cm.index else "" for i in cust.index]
        cust["mobile"] = [cm.at[i, "Mobile Number"] if i in cm.index and pd.notna(cm.at[i, "Mobile Number"]) else "" for i in cust.index]
        cust.loc[cust.status.eq("") & cust.segment.eq("RO"), "status"] = "RO (not in master)"
    else:
        cust["status"], cust["mobile"] = "", ""
    g = sec.groupby("customer_id").date
    cust["first"], cust["last"] = g.min(), g.max()
    fyv = sec.pivot_table(index="customer_id", columns="fy", values="litres", aggfunc="sum").fillna(0)
    cust["sortkey"] = fyv.iloc[:, -2:].sum(axis=1) if fyv.shape[1] > 1 else fyv.iloc[:, -1]
    zrank = {z: i for i, z in enumerate(ZONE_ORDER)}
    cust["zr"] = cust.zone.map(zrank).fillna(98)
    cust = cust.sort_values(["zr", "area", "sortkey"], ascending=[True, True, False])

    data = (sec.groupby(["customer_id", "month", "stockist", "ssr"], as_index=False).litres.sum())
    data = data[data.litres.abs() > 1e-9]
    for col in ("name", "segment", "subseg", "zone", "area", "pin"):
        data[col] = data.customer_id.map(cust[col])
    data["fy"] = data.month.dt.year + (data.month.dt.month >= 4)
    data["q"] = "Q" + (((data.month.dt.month - 4) % 12) // 3 + 1).astype(str)
    data = data.sort_values(["month", "segment", "zone", "area", "customer_id"]).reset_index(drop=True)

    return dict(sec=sec, data=data, cust=cust, pins=pins, files=files, problems=problems, clearblue=clearblue,
                cm_file=cm_file, cm=cm, last_date=sec.date.max())


# ---------------------------------------------------------------- workbook helpers
def head(ws, row, col, labels, height=30):
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BOX
    ws.row_dimensions[row].height = height


def put(ws, row, col, value, fmt=None, bold=False, fill=None, align=None):
    c = ws.cell(row, col, value)
    c.font = F_BOLD if bold else F_BASE
    c.border = BOX
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if align: c.alignment = align
    return c


def title(ws, text, sub=None):
    ws["B1"] = text; ws["B1"].font = F_TITLE
    ws.row_dimensions[1].height = 26
    if sub:
        ws["B2"] = sub; ws["B2"].font = F_SUB
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2


def period_line(ws, cell="B3"):
    ws[cell] = ('="Period: "&TEXT(PFrom,"mmm-yy")&" to "&TEXT(PTo,"mmm-yy")&"   vs  same period last year: "'
                '&TEXT(LFrom,"mmm-yy")&" to "&TEXT(LTo,"mmm-yy")&"   (change the period on the Dashboard)"')
    ws[cell].font = Font(name=FONT, size=10, bold=True, color=ORANGE)


def growth_cf(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=["0.1"], fill=PatternFill("solid", fgColor="E2EFDA")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThanOrEqual", formula=["-0.1"], fill=PatternFill("solid", fgColor="FCE4E4")))


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def sumifs(crit: str, period: str = "CY", div=True) -> str:
    """crit: extra ',range,criteria' pairs. period CY uses PFrom..PTo, LY uses LFrom..LTo."""
    a, b = ("PFrom", "PTo") if period == "CY" else ("LFrom", "LTo")
    f = f'SUMIFS(dL,dMonth,">="&{a},dMonth,"<="&{b}{crit})'
    return f"{f}/1000" if div else f


def gr(cy, ly):
    return f'=IF({ly}=0,"",{cy}/{ly}-1)'


# ---------------------------------------------------------------- workbook
def build_workbook(M) -> Path:
    data, cust, sec = M["data"], M["cust"], M["sec"]
    segs = C.CHANNEL_ORDER + sorted(set(cust.segment) - set(C.CHANNEL_ORDER))
    last_date = M["last_date"]
    cur_fy = int(last_date.year + (last_date.month >= 4))
    fys = sorted(data.fy.unique())
    months = pd.date_range(data.month.min(), data.month.max(), freq="MS")
    zones = [z for z in ZONE_ORDER if z in set(cust.zone)] + sorted(set(cust.zone) - set(ZONE_ORDER))
    stockists = sec.groupby("stockist").litres.sum().sort_values(ascending=False).index.tolist()
    ssrs = data[data.fy >= cur_fy - 1].groupby("ssr").litres.sum().sort_values(ascending=False).index.tolist()
    ssrs += sorted(set(data.ssr) - set(ssrs))
    areas = cust.drop_duplicates(["zone", "area"])[["zone", "area"]]
    areas["zr"] = areas.zone.map({z: i for i, z in enumerate(zones)})
    areas = areas.sort_values(["zr", "area"])

    wb = Workbook()
    ws_d = wb.active; ws_d.title = "Dashboard"
    names = ["Period Trend", "Area Wise", "Area x SSF", "SSF Wise", "Sub-segment", "Customers", "Action List",
             "Data", "Area Master", "Data Check", "Notes", "Lists"]
    S = {n: wb.create_sheet(n) for n in names}
    S["Dashboard"] = ws_d
    wb.calculation.fullCalcOnLoad = True

    # ---------- Lists (hidden)
    ls = S["Lists"]
    ls["A1"], ls["C1"] = "Months", "FY"
    for i, m in enumerate(months, 2):
        ls.cell(i, 1, m.to_pydatetime()).number_format = MON
    for i, f in enumerate(fys, 2):
        ls.cell(i, 3, int(f))
    ls.sheet_state = "hidden"

    # ---------- Data
    d = S["Data"]
    dcols = ["Month", "FY", "Qtr", "Customer Id", "Customer Name", "Segment", "Sub-segment", "SSR", "SSF",
             "Zone", "Area", "Pincode", "Litres"]
    head(d, 1, 1, dcols, 22)
    for r in data.itertuples(index=False):
        d.append([r.month.to_pydatetime(), fy_label(r.fy), r.q, r.customer_id, r.name, r.segment, r.subseg,
                  r.stockist, r.ssr, r.zone, r.area, r.pin, round(float(r.litres), 3)])
    n = len(data) + 1
    for row in d.iter_rows(min_row=2, max_row=n, min_col=1, max_col=1):
        row[0].number_format = MON
    for row in d.iter_rows(min_row=2, max_row=n, min_col=13, max_col=13):
        row[0].number_format = "#,##0.0"
    d.freeze_panes = "A2"; d.auto_filter.ref = f"A1:M{n}"
    widths(d, {"A": 9, "B": 9, "C": 5, "D": 18, "E": 36, "F": 8, "G": 26, "H": 22, "I": 18, "J": 30, "K": 28, "L": 9, "M": 11})
    for nm, col in (("dMonth", "A"), ("dFY", "B"), ("dCust", "D"), ("dSeg", "F"), ("dSub", "G"), ("dStk", "H"),
                    ("dSSR", "I"), ("dZone", "J"), ("dArea", "K"), ("dL", "M")):
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"Data!${col}$2:${col}${n}")

    # ---------- Dashboard inputs + names
    ws = ws_d
    title(ws, "Secondary Sales - Segment, Area & Customer Review",
          f"Mumbai Lubes Reseller SA  |  {', '.join(stockists)}  |  ClearBlue (DEF) excluded, ServoXpert in Bazaar, Dealer = RO  |  "
          f"data up to {last_date:%d-%b-%Y}")
    ws["B4"] = "SELECT PERIOD"; ws["B4"].font = F_SEC
    ws["B5"], ws["B6"] = "From month", "To month"
    ws["C5"] = datetime(cur_fy - 1, 4, 1); ws["C6"] = months[-1].to_pydatetime()
    for c in ("C5", "C6"):
        ws[c].number_format, ws[c].font, ws[c].fill, ws[c].border = MON, F_INPUT, FILL_INPUT, BOX
        ws[c].alignment = Alignment(horizontal="center")
    ws["B5"].font = ws["B6"].font = F_BOLD
    ws["E5"], ws["E6"] = "Compared with", ""
    ws["E5"].font = F_BOLD
    ws["F5"], ws["F6"] = "=LFrom", "=LTo"
    for c in ("F5", "F6"):
        ws[c].number_format, ws[c].font, ws[c].border = MON, F_BOLD, BOX
        ws[c].alignment = Alignment(horizontal="center")
    ws["H5"] = "Pick From / To from the drop-down (yellow cells). Every tab follows this period, except Period Trend (pick the FY there)."
    ws["H6"] = "Default = current FY to date vs the same months last year (like-for-like)."
    ws["H5"].font = ws["H6"].font = F_SUB
    dv = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${len(months) + 1}", allow_blank=False)
    ws.add_data_validation(dv); dv.add("C5"); dv.add("C6")
    for nm, ref in (("PFrom", "$C$5"), ("PTo", "$C$6")):
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"Dashboard!{ref}")
    wb.defined_names["LFrom"] = DefinedName("LFrom", attr_text="EDATE(Dashboard!$C$5,-12)")
    wb.defined_names["LTo"] = DefinedName("LTo", attr_text="EDATE(Dashboard!$C$6,-12)")
    wb.defined_names["DataDate"] = DefinedName("DataDate", attr_text="Notes!$C$4")

    # ---------- Customers (needed by dashboard counters)
    cs = S["Customers"]
    title(cs, "Customer-wise Sales (litres)", "One row per customer billed by the SSRs since "
          f"{months[0]:%b-%y}. Use the filter buttons: Segment, SSR, SSF, Zone, Area, Status.")
    period_line(cs)
    fy_cols = [f"{fy_label(f)}{' (to ' + last_date.strftime('%b-%y') + ')' if f == cur_fy else ''}" for f in fys]
    ccols = (["Customer Id", "Customer Name", "Segment", "Customer Group", "Sub-segment", "SSR", "SSF (latest bill)",
              "Zone", "Area", "Pincode", "Location from", "Master Status", "Mobile"] + fy_cols +
             ["Selected period (L)", "Same period LY (L)", "Change (L)", "Growth %", "Status", "First bill", "Last bill",
              "Days since last bill"])
    HR = 5
    head(cs, HR, 1, ccols, 42)
    nfy = len(fys)
    cFY0 = 14
    cCY, cLY = cFY0 + nfy, cFY0 + nfy + 1
    cCH, cGR, cST, cFB, cLB, cDS = cLY + 1, cLY + 2, cLY + 3, cLY + 4, cLY + 5, cLY + 6
    r0 = HR + 1
    for i, (cid, c) in enumerate(cust.iterrows()):
        r = r0 + i
        vals = [cid, c["name"], c.segment, c.group, c.subseg, c.stockist, c.ssr, c.zone, c.area, c.pin, c.loc_src,
                c.status, c.mobile]
        for j, v in enumerate(vals, 1):
            put(cs, r, j, v)
        cs.cell(r, 1).number_format = "@"
        for k, f in enumerate(fys):
            put(cs, r, cFY0 + k, f'=SUMIFS(dL,dCust,$A{r}&"*",dFY,"{fy_label(f)}")', LIT)
        crit = f',dCust,$A{r}&"*"'
        put(cs, r, cCY, "=" + sumifs(crit, "CY", False), LIT, bold=True)
        put(cs, r, cLY, "=" + sumifs(crit, "LY", False), LIT)
        cy, ly = f"{L(cCY)}{r}", f"{L(cLY)}{r}"
        put(cs, r, cCH, f"={cy}-{ly}", LIT)
        put(cs, r, cGR, gr(cy, ly), GR)
        fb = f"{L(cFB)}{r}"
        put(cs, r, cST, f'=IF(AND(ROUND({cy},1)=0,ROUND({ly},1)=0),"No sales",IF(ROUND({ly},1)<=0,IF({fb}>=PFrom,"New","Re-activated"),'
                         f'IF(ROUND({cy},1)<=0,"Lost",IF({cy}/{ly}-1>=0.1,"Growing",IF({cy}/{ly}-1<=-0.1,"Declining","Stable")))))',
            align=Alignment(horizontal="center"))
        put(cs, r, cFB, c["first"].to_pydatetime(), DATE)
        put(cs, r, cLB, c["last"].to_pydatetime(), DATE)
        put(cs, r, cDS, f"=DataDate-{L(cLB)}{r}", CNT)
    cn = r0 + len(cust) - 1
    cs.freeze_panes = cs.cell(r0, 3)
    cs.auto_filter.ref = f"A{HR}:{L(cDS)}{cn}"
    growth_cf(cs, f"{L(cGR)}{r0}:{L(cGR)}{cn}")
    for st, col in (("Lost", "FCE4E4"), ("Declining", "FFF2CC"), ("New", "E2EFDA"), ("Growing", "E2EFDA")):
        cs.conditional_formatting.add(f"{L(cST)}{r0}:{L(cST)}{cn}",
                                      CellIsRule(operator="equal", formula=[f'"{st}"'], fill=PatternFill("solid", fgColor=col)))
    cs.conditional_formatting.add(f"{L(cCY)}{r0}:{L(cCY)}{cn}", DataBarRule(start_type="num", start_value=0, end_type="max", color="9BC2E6"))
    widths(cs, {"A": 17, "B": 36, "C": 8, "D": 12, "E": 22, "F": 20, "G": 17, "H": 26, "I": 24, "J": 8, "K": 13, "L": 12, "M": 12})
    for k in range(nfy):
        cs.column_dimensions[L(cFY0 + k)].width = 11
    for col, w in ((cCY, 12), (cLY, 12), (cCH, 10), (cGR, 9), (cST, 12), (cFB, 10), (cLB, 10), (cDS, 9)):
        cs.column_dimensions[L(col)].width = w
    for nm, col in (("cSeg", 3), ("cSub", 5), ("cStk", 6), ("cSSR", 7), ("cZone", 8), ("cArea", 9), ("cCY", cCY),
                    ("cLY", cLY), ("cStatus", cST)):
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"Customers!${L(col)}${r0}:${L(col)}${cn}")

    # ---------- Dashboard body
    period_line(ws, "B3")
    # KPI tiles
    kpis = [("Sales, selected period (KL)", f"={sumifs('')}", KL), ("Same period last year (KL)", f"={sumifs('', 'LY')}", KL),
            ("Growth", '=IF(C10=0,"",B10/C10-1)', GR), ("Counters billed", '=COUNTIFS(cCY,">0")', CNT),
            ("Counters billed LY", '=COUNTIFS(cLY,">0")', CNT), ("New + re-activated", '=COUNTIFS(cStatus,"New")+COUNTIFS(cStatus,"Re-activated")', CNT),
            ("Lost counters", '=COUNTIFS(cStatus,"Lost")', CNT)]
    ws["B8"] = "HEADLINE"; ws["B8"].font = F_SEC
    for i, (lab, f, fmt) in enumerate(kpis):
        col = 2 + i
        a = ws.cell(9, col, lab); a.font, a.fill, a.alignment = F_KPIL, FILL_KPI, CENTER
        b = ws.cell(10, col, f); b.font, b.fill, b.number_format, b.alignment = F_KPI, FILL_KPI, fmt, Alignment(horizontal="center")
    ws.row_dimensions[9].height = 28; ws.row_dimensions[10].height = 30

    r = 12
    ws.cell(r, 2, "SEGMENT-WISE").font = F_SEC
    hdr = ["Segment", "Sales (KL)", "LY (KL)", "Change (KL)", "Growth %", "Share %", "Counters billed", "Counters LY",
           "Avg L / counter", "New + re-act.", "Lost", "Growing", "Declining"]
    head(ws, r + 1, 2, hdr)
    s0 = r + 2
    for i, sg in enumerate(segs):
        rr = s0 + i
        put(ws, rr, 2, sg, bold=True)
        put(ws, rr, 3, f'={sumifs(f",dSeg,$B{rr}")}', KL)
        put(ws, rr, 4, f'={sumifs(f",dSeg,$B{rr}", "LY")}', KL)
        put(ws, rr, 5, f"=C{rr}-D{rr}", KL)
        put(ws, rr, 6, gr(f"C{rr}", f"D{rr}"), GR)
        put(ws, rr, 7, f"=IF($C${s0 + len(segs)}=0,\"\",C{rr}/$C${s0 + len(segs)})", PCT)
        put(ws, rr, 8, f'=COUNTIFS(cSeg,$B{rr},cCY,">0")', CNT)
        put(ws, rr, 9, f'=COUNTIFS(cSeg,$B{rr},cLY,">0")', CNT)
        put(ws, rr, 10, f'=IF(H{rr}=0,"",C{rr}*1000/H{rr})', LIT)
        put(ws, rr, 11, f'=COUNTIFS(cSeg,$B{rr},cStatus,"New")+COUNTIFS(cSeg,$B{rr},cStatus,"Re-activated")', CNT)
        for k, stt in enumerate(("Lost", "Growing", "Declining")):
            put(ws, rr, 12 + k, f'=COUNTIFS(cSeg,$B{rr},cStatus,"{stt}")', CNT)
    tr = s0 + len(segs)
    put(ws, tr, 2, "Total", bold=True, fill=FILL_TOT)
    for col in range(3, 15):
        cl = L(col)
        f = {6: gr(f"C{tr}", f"D{tr}"), 7: "=IF(C{0}=0,\"\",1)".format(tr), 10: f'=IF(H{tr}=0,"",C{tr}*1000/H{tr})'}.get(col, f"=SUM({cl}{s0}:{cl}{tr - 1})")
        put(ws, tr, col, f, {6: GR, 7: PCT, 3: KL, 4: KL, 5: KL, 10: LIT}.get(col, CNT), bold=True, fill=FILL_TOT)
    growth_cf(ws, f"F{s0}:F{tr}")
    seg_rows = (s0, tr - 1)

    r = tr + 2
    ws.cell(r, 2, "SSR x SEGMENT (KL)").font = F_SEC
    head(ws, r + 1, 2, ["SSR"] + segs + ["Total", "LY", "Change", "Growth %", "Share %"])
    k0 = r + 2
    nseg = len(segs)
    for i, st in enumerate(stockists):
        rr = k0 + i
        put(ws, rr, 2, st, bold=True)
        for j, sg in enumerate(segs):
            put(ws, rr, 3 + j, f'={sumifs(f",dStk,$B{rr},dSeg,{L(3 + j)}${k0 - 1}")}', KL)
        tc = 3 + nseg
        put(ws, rr, tc, f"=SUM({L(3)}{rr}:{L(tc - 1)}{rr})", KL, bold=True)
        put(ws, rr, tc + 1, f'={sumifs(f",dStk,$B{rr}", "LY")}', KL)
        put(ws, rr, tc + 2, f"={L(tc)}{rr}-{L(tc + 1)}{rr}", KL)
        put(ws, rr, tc + 3, gr(f"{L(tc)}{rr}", f"{L(tc + 1)}{rr}"), GR)
        put(ws, rr, tc + 4, f'=IF({L(tc)}${k0 + len(stockists)}=0,"",{L(tc)}{rr}/{L(tc)}${k0 + len(stockists)})', PCT)
    kt = k0 + len(stockists)
    put(ws, kt, 2, "Total", bold=True, fill=FILL_TOT)
    tc = 3 + nseg
    for col in range(3, tc + 5):
        cl = L(col)
        f = {tc + 3: gr(f"{L(tc)}{kt}", f"{L(tc + 1)}{kt}"), tc + 4: f'=IF({L(tc)}{kt}=0,"",1)'}.get(col, f"=SUM({cl}{k0}:{cl}{kt - 1})")
        put(ws, kt, col, f, GR if col == tc + 3 else PCT if col == tc + 4 else KL, bold=True, fill=FILL_TOT)
    growth_cf(ws, f"{L(tc + 3)}{k0}:{L(tc + 3)}{kt}")

    r = kt + 2
    ws.cell(r, 2, "ZONE SNAPSHOT (KL) - detail on Area Wise").font = F_SEC
    head(ws, r + 1, 2, ["Zone"] + segs + ["Total", "LY", "Change", "Growth %", "Share %"])
    z0 = r + 2
    for i, z in enumerate(zones):
        rr = z0 + i
        put(ws, rr, 2, z, bold=True)
        for j, sg in enumerate(segs):
            put(ws, rr, 3 + j, f'={sumifs(f",dZone,$B{rr},dSeg,{L(3 + j)}${z0 - 1}")}', KL)
        put(ws, rr, tc, f"=SUM({L(3)}{rr}:{L(tc - 1)}{rr})", KL, bold=True)
        put(ws, rr, tc + 1, f'={sumifs(f",dZone,$B{rr}", "LY")}', KL)
        put(ws, rr, tc + 2, f"={L(tc)}{rr}-{L(tc + 1)}{rr}", KL)
        put(ws, rr, tc + 3, gr(f"{L(tc)}{rr}", f"{L(tc + 1)}{rr}"), GR)
        put(ws, rr, tc + 4, f'=IF({L(tc)}${z0 + len(zones)}=0,"",{L(tc)}{rr}/{L(tc)}${z0 + len(zones)})', PCT)
    zt = z0 + len(zones)
    put(ws, zt, 2, "Total", bold=True, fill=FILL_TOT)
    for col in range(3, tc + 5):
        cl = L(col)
        f = {tc + 3: gr(f"{L(tc)}{zt}", f"{L(tc + 1)}{zt}"), tc + 4: f'=IF({L(tc)}{zt}=0,"",1)'}.get(col, f"=SUM({cl}{z0}:{cl}{zt - 1})")
        put(ws, zt, col, f, GR if col == tc + 3 else PCT if col == tc + 4 else KL, bold=True, fill=FILL_TOT)
    growth_cf(ws, f"{L(tc + 3)}{z0}:{L(tc + 3)}{zt}")
    ws.conditional_formatting.add(f"{L(tc)}{z0}:{L(tc)}{zt - 1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="F4B183"))

    ch = BarChart(); ch.type = "bar"; ch.style = 10; ch.title = "Segment: selected period vs LY (KL)"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=seg_rows[0] - 1, max_row=seg_rows[1]), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=seg_rows[0], max_row=seg_rows[1]))
    ch.height, ch.width = 7.5, 14
    ws.add_chart(ch, "P12")
    ch2 = BarChart(); ch2.type = "col"; ch2.style = 10; ch2.title = "Zone: selected period vs LY (KL)"
    ch2.add_data(Reference(ws, min_col=tc, max_col=tc + 1, min_row=z0 - 1, max_row=zt - 1), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=2, min_row=z0, max_row=zt - 1))
    ch2.height, ch2.width = 9, 22
    ws.add_chart(ch2, f"P{z0 - 2}")
    widths(ws, {"B": 38, **{L(c): 12 for c in range(3, 15)}})
    ws.freeze_panes = "A8"

    # ---------- Period Trend
    pt = S["Period Trend"]
    title(pt, "Period-wise Trend (KL)", "Month, quarter, half-year and full-year view for one financial year vs the year before. "
          "Not tied to the Dashboard period.")
    pt["B4"] = "Financial year (ending)"; pt["B4"].font = F_BOLD
    pt["C4"] = cur_fy
    pt["C4"].font, pt["C4"].fill, pt["C4"].border, pt["C4"].alignment = F_INPUT, FILL_INPUT, BOX, Alignment(horizontal="center")
    pt["D4"] = '="= FY"&RIGHT(C4-1,2)&"-"&RIGHT(C4,2)&"  vs  FY"&RIGHT(C4-2,2)&"-"&RIGHT(C4-1,2)'
    pt["D4"].font = Font(name=FONT, size=10, bold=True, color=ORANGE)
    dv2 = DataValidation(type="list", formula1=f"=Lists!$C$2:$C${len(fys) + 1}", allow_blank=False)
    pt.add_data_validation(dv2); dv2.add("C4")
    hr = 6
    pt.cell(hr, 3, '="FY"&RIGHT($C$4-1,2)&"-"&RIGHT($C$4,2)').font = F_SEC
    pt.cell(hr, 3 + nseg + 1, '="FY"&RIGHT($C$4-2,2)&"-"&RIGHT($C$4-1,2)&" (LY)"').font = F_SEC
    hdr = ["Month"] + segs + ["Total"] + segs + ["Total LY", "Growth %", "Cum. FY", "Cum. LY", "Cum. growth %"]
    head(pt, hr + 1, 2, hdr)
    m0 = hr + 2
    totc = 3 + nseg               # CY total column
    lyc0 = totc + 1               # first LY segment column
    lytot = lyc0 + nseg
    for i in range(12):
        rr = m0 + i
        put(pt, rr, 2, f"=DATE($C$4-1,{4 + i},1)", MON, bold=True)
        for j, sg in enumerate(segs):
            put(pt, rr, 3 + j, f'=SUMIFS(dL,dMonth,$B{rr},dSeg,"{sg}")/1000', KL)
            put(pt, rr, lyc0 + j, f'=SUMIFS(dL,dMonth,EDATE($B{rr},-12),dSeg,"{sg}")/1000', KL)
        put(pt, rr, totc, f"=SUM({L(3)}{rr}:{L(totc - 1)}{rr})", KL, bold=True)
        put(pt, rr, lytot, f"=SUM({L(lyc0)}{rr}:{L(lytot - 1)}{rr})", KL, bold=True)
        put(pt, rr, lytot + 1, f'=IF(OR({L(lytot)}{rr}=0,$B{rr}>DataDate),"",{L(totc)}{rr}/{L(lytot)}{rr}-1)', GR)
        put(pt, rr, lytot + 2, f'=IF($B{rr}>DataDate,"",SUM({L(totc)}${m0}:{L(totc)}{rr}))', KL)
        put(pt, rr, lytot + 3, f'=IF($B{rr}>DataDate,"",SUM({L(lytot)}${m0}:{L(lytot)}{rr}))', KL)
        put(pt, rr, lytot + 4, f'=IF(OR({L(lytot + 3)}{rr}="",N({L(lytot + 3)}{rr})=0),"",{L(lytot + 2)}{rr}/{L(lytot + 3)}{rr}-1)', GR)
    me = m0 + 11
    growth_cf(pt, f"{L(lytot + 1)}{m0}:{L(lytot + 1)}{me}")
    growth_cf(pt, f"{L(lytot + 4)}{m0}:{L(lytot + 4)}{me}")
    q0 = me + 2
    head(pt, q0, 2, ["Period"] + segs + ["Total"] + segs + ["Total LY", "Growth %"])
    blocks = [("Q1 (Apr-Jun)", 0, 3), ("Q2 (Jul-Sep)", 3, 3), ("Q3 (Oct-Dec)", 6, 3), ("Q4 (Jan-Mar)", 9, 3),
              ("H1 (Apr-Sep)", 0, 6), ("H2 (Oct-Mar)", 6, 6), ("Full year", 0, 12)]
    for i, (lab, st_, ln) in enumerate(blocks):
        rr = q0 + 1 + i
        bold = lab == "Full year"
        put(pt, rr, 2, lab, bold=True, fill=FILL_TOT if lab.startswith(("H", "Full")) else None)
        for col in range(3, lytot + 1):
            put(pt, rr, col, f"=SUM({L(col)}{m0 + st_}:{L(col)}{m0 + st_ + ln - 1})", KL, bold=bold,
                fill=FILL_TOT if lab.startswith(("H", "Full")) else None)
        put(pt, rr, lytot + 1, f'=IF(OR({L(lytot)}{rr}=0,$B${m0 + st_ + ln - 1}>DataDate),"",{L(totc)}{rr}/{L(lytot)}{rr}-1)', GR, bold=bold)
    growth_cf(pt, f"{L(lytot + 1)}{q0 + 1}:{L(lytot + 1)}{q0 + len(blocks)}")
    pt.cell(q0 + len(blocks) + 1, 2, "Growth is shown only for complete periods (months up to the latest data date). The Cum. columns give the like-for-like year-to-date.").font = F_SUB

    y0 = q0 + len(blocks) + 3
    pt.cell(y0, 2, "MULTI-YEAR VIEW (KL)").font = F_SEC
    head(pt, y0 + 1, 2, ["Financial year"] + segs + ["Total", "Growth %", "Counters billed"])
    for i, f in enumerate(fys):
        rr = y0 + 2 + i
        lab = fy_label(f) + (f" (to {last_date:%b-%y})" if f == cur_fy else "") + (f" (from {months[0]:%b-%y})" if f == fys[0] and months[0].month != 4 else "")
        put(pt, rr, 2, lab, bold=True)
        for j, sg in enumerate(segs):
            put(pt, rr, 3 + j, f'=SUMIFS(dL,dFY,"{fy_label(f)}",dSeg,"{sg}")/1000', KL)
        put(pt, rr, totc, f"=SUM({L(3)}{rr}:{L(totc - 1)}{rr})", KL, bold=True)
        put(pt, rr, totc + 1, "" if i == 0 or f == cur_fy else gr(f"{L(totc)}{rr}", f"{L(totc)}{rr - 1}"), GR)
        cc = cFY0 + i
        put(pt, rr, totc + 2, f'=COUNTIFS(Customers!${L(cc)}${r0}:${L(cc)}${cn},">0")', CNT)
    pt.cell(y0 + 2 + len(fys), 2, "No growth is shown for the current part-year - see the Cum. columns above or the Dashboard for like-for-like.").font = F_SUB
    lc = LineChart(); lc.title = "Monthly sales (KL): selected FY vs LY"; lc.style = 12; lc.height, lc.width = 8, 18
    lc.add_data(Reference(pt, min_col=totc, min_row=hr + 1, max_row=me), titles_from_data=True)
    lc.add_data(Reference(pt, min_col=lytot, min_row=hr + 1, max_row=me), titles_from_data=True)
    lc.set_categories(Reference(pt, min_col=2, min_row=m0, max_row=me))
    pt.add_chart(lc, f"B{y0 + len(fys) + 4}")
    widths(pt, {"B": 22, **{L(c): 10.5 for c in range(3, lytot + 6)}})
    pt.freeze_panes = pt.cell(m0, 3)

    # ---------- Area Wise
    aw = S["Area Wise"]
    title(aw, "Area-wise Sales (KL)", "Area = post-office locality of the customer's pincode, Zone = sales territory (mapping in Data/Area Master.xlsx). "
          "SSR = stockist; Main SSR = SSR with the most sales in the area. Click + / - on the left to open or close a zone.")
    period_line(aw)
    acols = (["Zone / Area", "Main SSR"] + segs + ["Total", "LY", "Change", "Growth %", "Share %", "Counters billed", "Counters LY",
              "Avg L / counter"] + stockists + ["Pincodes"])
    head(aw, 5, 2, acols)
    tcA = 4 + nseg
    k1 = tcA + 8                       # first SSR (stockist) column
    pcol = k1 + len(stockists)          # pincodes column
    for c0, c1, lab in ((4, tcA - 1, "SALES BY SEGMENT (KL)"), (tcA, tcA + 7, "TOTAL"), (k1, pcol - 1, "SALES BY SSR (KL)")):
        aw.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c1)
        c = aw.cell(4, c0, lab); c.font, c.fill, c.alignment = F_HEAD, PatternFill("solid", fgColor=ORANGE), CENTER
    rr = 6
    zone_rows = []
    pins_by_area = cust[cust.pin.ne("")].groupby(["zone", "area"]).pin.agg(lambda s: ", ".join(sorted(set(s))))
    grand = None
    area_rows_all = []
    for z in zones:
        zr = rr
        zone_rows.append(zr)
        put(aw, zr, 2, z, bold=True, fill=FILL_ZONE)
        za = areas[areas.zone == z].area.tolist()
        rr += 1
        for a in za:
            put(aw, rr, 2, "   " + a)
            aw.cell(rr, 2).alignment = Alignment(indent=1)
            put(aw, rr, pcol, pins_by_area.get((z, a), ""))
            aw.cell(rr, pcol).font = Font(name=FONT, size=8, color="808080")
            for j, st in enumerate(stockists):
                put(aw, rr, k1 + j, f'={sumifs(f",dZone,{chr(34)}{z}{chr(34)},dArea,{chr(34)}{a}{chr(34)},dStk,{L(k1 + j)}$5")}', KL)
            for j, sg in enumerate(segs):
                put(aw, rr, 4 + j, f'={sumifs(f",dZone,{chr(34)}{z}{chr(34)},dArea,{chr(34)}{a}{chr(34)},dSeg,{L(4 + j)}$5")}', KL)
            put(aw, rr, tcA, f"=SUM({L(4)}{rr}:{L(tcA - 1)}{rr})", KL, bold=True)
            put(aw, rr, tcA + 1, f'={sumifs(f",dZone,{chr(34)}{z}{chr(34)},dArea,{chr(34)}{a}{chr(34)}", "LY")}', KL)
            put(aw, rr, tcA + 5, f'=COUNTIFS(cZone,"{z}",cArea,"{a}",cCY,">0")', CNT)
            put(aw, rr, tcA + 6, f'=COUNTIFS(cZone,"{z}",cArea,"{a}",cLY,">0")', CNT)
            aw.row_dimensions[rr].outlineLevel = 1
            area_rows_all.append(rr)
            rr += 1
        put(aw, zr, pcol, "", fill=FILL_ZONE)
        for col in list(range(4, tcA + 2)) + [tcA + 5, tcA + 6] + list(range(k1, pcol)):
            put(aw, zr, col, f"=SUM({L(col)}{zr + 1}:{L(col)}{rr - 1})" if za else 0, CNT if col in (tcA + 5, tcA + 6) else KL, bold=True, fill=FILL_ZONE)
    grand = rr
    put(aw, grand, 2, "Total", bold=True, fill=FILL_TOT)
    put(aw, grand, pcol, "", fill=FILL_TOT)
    for col in list(range(4, tcA + 2)) + [tcA + 5, tcA + 6] + list(range(k1, pcol)):
        put(aw, grand, col, "=" + "+".join(f"{L(col)}{z}" for z in zone_rows), CNT if col in (tcA + 5, tcA + 6) else KL, bold=True, fill=FILL_TOT)
    for x in zone_rows + area_rows_all + [grand]:
        fill = FILL_ZONE if x in zone_rows else FILL_TOT if x == grand else None
        b = x in zone_rows or x == grand
        put(aw, x, tcA + 2, f"={L(tcA)}{x}-{L(tcA + 1)}{x}", KL, bold=b, fill=fill)
        put(aw, x, tcA + 3, gr(f"{L(tcA)}{x}", f"{L(tcA + 1)}{x}"), GR, bold=b, fill=fill)
        put(aw, x, tcA + 4, f'=IF(${L(tcA)}${grand}=0,"",{L(tcA)}{x}/${L(tcA)}${grand})', PCT, bold=b, fill=fill)
        put(aw, x, tcA + 7, f'=IF({L(tcA + 5)}{x}=0,"",{L(tcA)}{x}*1000/{L(tcA + 5)}{x})', LIT, bold=b, fill=fill)
        kr = f"${L(k1)}{x}:${L(pcol - 1)}{x}"
        put(aw, x, 3, f'=IF(MAX({kr})<=0,"",INDEX(${L(k1)}$5:${L(pcol - 1)}$5,MATCH(MAX({kr}),{kr},0)))', bold=b, fill=fill)
    growth_cf(aw, f"{L(tcA + 3)}6:{L(tcA + 3)}{grand}")
    aw.sheet_properties.outlinePr.summaryBelow = False
    widths(aw, {"B": 36, "C": 22, **{L(c): 10.5 for c in range(4, pcol)}, L(pcol): 18})
    aw.conditional_formatting.add(f"{L(k1)}6:{L(pcol - 1)}{grand - 1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="F4B183"))
    aw.freeze_panes = "D6"

    # ---------- Area x SSR
    ax = S["Area x SSF"]
    title(ax, "Area x SSF (KL)", "Which SSF (stockist salesman) sells where: secondary sales in the selected period by the SSF on the invoice. "
          "For the area-to-SSR (stockist) split see Area Wise.")
    period_line(ax)
    head(ax, 5, 2, ["Zone / Area"] + ssrs + ["Total", "LY", "Growth %", "Main SSF"], 42)
    tcX = 3 + len(ssrs)
    rr = 6
    zr_x = []
    for z in zones:
        zr = rr; zr_x.append(zr)
        put(ax, zr, 2, z, bold=True, fill=FILL_ZONE)
        za = areas[areas.zone == z].area.tolist()
        rr += 1
        for a in za:
            put(ax, rr, 2, "   " + a)
            for j, s_ in enumerate(ssrs):
                put(ax, rr, 3 + j, f'={sumifs(f",dZone,{chr(34)}{z}{chr(34)},dArea,{chr(34)}{a}{chr(34)},dSSR,{L(3 + j)}$5")}', KL)
            put(ax, rr, tcX + 1, f'={sumifs(f",dZone,{chr(34)}{z}{chr(34)},dArea,{chr(34)}{a}{chr(34)}", "LY")}', KL)
            ax.row_dimensions[rr].outlineLevel = 1
            rr += 1
        for col in list(range(3, tcX)) + [tcX + 1]:
            put(ax, zr, col, f"=SUM({L(col)}{zr + 1}:{L(col)}{rr - 1})" if za else 0, KL, bold=True, fill=FILL_ZONE)
    gx = rr
    put(ax, gx, 2, "Total", bold=True, fill=FILL_TOT)
    for col in list(range(3, tcX)) + [tcX + 1]:
        put(ax, gx, col, "=" + "+".join(f"{L(col)}{z}" for z in zr_x), KL, bold=True, fill=FILL_TOT)
    for x in range(6, gx + 1):
        b = x in zr_x or x == gx
        fill = FILL_ZONE if x in zr_x else FILL_TOT if x == gx else None
        put(ax, x, tcX, f"=SUM({L(3)}{x}:{L(tcX - 1)}{x})", KL, bold=True, fill=fill)
        put(ax, x, tcX + 2, gr(f"{L(tcX)}{x}", f"{L(tcX + 1)}{x}"), GR, bold=b, fill=fill)
        rngs = f"$C{x}:${L(tcX - 1)}{x}"
        put(ax, x, tcX + 3, f'=IF({L(tcX)}{x}=0,"",INDEX($C$5:${L(tcX - 1)}$5,MATCH(MAX({rngs}),{rngs},0)))', bold=b, fill=fill)
    growth_cf(ax, f"{L(tcX + 2)}6:{L(tcX + 2)}{gx}")
    ax.conditional_formatting.add(f"C6:{L(tcX - 1)}{gx - 1}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, color="1F3864", bold=True)))
    ax.sheet_properties.outlinePr.summaryBelow = False
    widths(ax, {"B": 36, **{L(c): 11 for c in range(3, tcX + 3)}, L(tcX + 3): 18})
    ax.freeze_panes = "C6"

    # ---------- SSR Wise
    sw = S["SSF Wise"]
    title(sw, "SSF-wise Sales (KL)", "SSF = SSR (stockist) salesman on the invoice. Counters = customers whose latest bill was by that SSF.")
    period_line(sw)
    main_stk = data.groupby(["ssr", "stockist"]).litres.sum().reset_index().sort_values("litres").groupby("ssr").tail(1).set_index("ssr").stockist
    head(sw, 5, 2, ["SSF", "SSR"] + segs + ["Total", "LY", "Change", "Growth %", "Share %", "Counters billed", "Counters LY"])
    tcS = 4 + nseg
    for i, s_ in enumerate(ssrs):
        rr = 6 + i
        put(sw, rr, 2, s_, bold=True); put(sw, rr, 3, main_stk.get(s_, ""))
        for j, sg in enumerate(segs):
            put(sw, rr, 4 + j, f'={sumifs(f",dSSR,$B{rr},dSeg,{L(4 + j)}$5")}', KL)
        put(sw, rr, tcS, f"=SUM({L(4)}{rr}:{L(tcS - 1)}{rr})", KL, bold=True)
        put(sw, rr, tcS + 1, f'={sumifs(f",dSSR,$B{rr}", "LY")}', KL)
        put(sw, rr, tcS + 5, f'=COUNTIFS(cSSR,$B{rr},cCY,">0")', CNT)
        put(sw, rr, tcS + 6, f'=COUNTIFS(cSSR,$B{rr},cLY,">0")', CNT)
    st_ = 6 + len(ssrs)
    put(sw, st_, 2, "Total", bold=True, fill=FILL_TOT); put(sw, st_, 3, "", fill=FILL_TOT)
    for col in list(range(4, tcS + 2)) + [tcS + 5, tcS + 6]:
        put(sw, st_, col, f"=SUM({L(col)}6:{L(col)}{st_ - 1})", CNT if col > tcS + 1 else KL, bold=True, fill=FILL_TOT)
    for x in range(6, st_ + 1):
        fill = FILL_TOT if x == st_ else None
        put(sw, x, tcS + 2, f"={L(tcS)}{x}-{L(tcS + 1)}{x}", KL, fill=fill, bold=x == st_)
        put(sw, x, tcS + 3, gr(f"{L(tcS)}{x}", f"{L(tcS + 1)}{x}"), GR, fill=fill, bold=x == st_)
        put(sw, x, tcS + 4, f'=IF(${L(tcS)}${st_}=0,"",{L(tcS)}{x}/${L(tcS)}${st_})', PCT, fill=fill, bold=x == st_)
    growth_cf(sw, f"{L(tcS + 3)}6:{L(tcS + 3)}{st_}")
    r2 = st_ + 3
    sw.cell(r2 - 1, 2, "SSF x ZONE (KL, selected period)").font = F_SEC
    head(sw, r2, 2, ["SSF"] + zones + ["Total"], 42)
    for i, s_ in enumerate(ssrs):
        rr = r2 + 1 + i
        put(sw, rr, 2, s_, bold=True)
        for j, z in enumerate(zones):
            put(sw, rr, 3 + j, f'={sumifs(f",dSSR,$B{rr},dZone,{L(3 + j)}${r2}")}', KL)
        put(sw, rr, 3 + len(zones), f"=SUM(C{rr}:{L(2 + len(zones))}{rr})", KL, bold=True)
    sz = r2 + 1 + len(ssrs)
    put(sw, sz, 2, "Total", bold=True, fill=FILL_TOT)
    for col in range(3, 4 + len(zones)):
        put(sw, sz, col, f"=SUM({L(col)}{r2 + 1}:{L(col)}{sz - 1})", KL, bold=True, fill=FILL_TOT)
    sw.conditional_formatting.add(f"C{r2 + 1}:{L(2 + len(zones))}{sz - 1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="9BC2E6"))
    widths(sw, {"B": 22, "C": 22, **{L(c): 11.5 for c in range(4, max(tcS + 7, 4 + len(zones)))}})
    sw.freeze_panes = "D6"

    # ---------- Sub-segment
    ss = S["Sub-segment"]
    title(ss, "Sub-segment Sales (KL)", "Customer Sub Group within each segment (e.g. OEM by brand dealership, Bazaar by shop type).")
    period_line(ss)
    head(ss, 5, 2, ["Segment", "Sub-segment", "Sales (KL)", "LY (KL)", "Change", "Growth %", "Share in segment", "Counters billed", "Counters LY", "Avg L / counter"])
    subs = cust.groupby(["segment", "subseg"]).size().reset_index()
    subs["sr"] = subs.segment.map({s: i for i, s in enumerate(segs)})
    rr = 6
    for sg in segs:
        part = subs[subs.segment == sg]
        if part.empty: continue
        hrow = rr
        put(ss, hrow, 2, sg, bold=True, fill=FILL_ZONE); put(ss, hrow, 3, "All", bold=True, fill=FILL_ZONE)
        rr += 1
        for sub in part.subseg:
            put(ss, rr, 2, sg); put(ss, rr, 3, sub)
            put(ss, rr, 4, f'={sumifs(f",dSeg,$B{rr},dSub,$C{rr}")}', KL)
            put(ss, rr, 5, f'={sumifs(f",dSeg,$B{rr},dSub,$C{rr}", "LY")}', KL)
            put(ss, rr, 9, f'=COUNTIFS(cSeg,$B{rr},cSub,$C{rr},cCY,">0")', CNT)
            put(ss, rr, 10, f'=COUNTIFS(cSeg,$B{rr},cSub,$C{rr},cLY,">0")', CNT)
            ss.row_dimensions[rr].outlineLevel = 1
            rr += 1
        for col in (4, 5, 9, 10):
            put(ss, hrow, col, f"=SUM({L(col)}{hrow + 1}:{L(col)}{rr - 1})", KL if col < 6 else CNT, bold=True, fill=FILL_ZONE)
        for x in range(hrow, rr):
            fill = FILL_ZONE if x == hrow else None
            put(ss, x, 6, f"=D{x}-E{x}", KL, fill=fill, bold=x == hrow)
            put(ss, x, 7, gr(f"D{x}", f"E{x}"), GR, fill=fill, bold=x == hrow)
            put(ss, x, 8, f'=IF(D${hrow}=0,"",D{x}/D${hrow})', PCT, fill=fill, bold=x == hrow)
            put(ss, x, 11, f'=IF(I{x}=0,"",D{x}*1000/I{x})', LIT, fill=fill, bold=x == hrow)
    growth_cf(ss, f"G6:G{rr}")
    ss.sheet_properties.outlinePr.summaryBelow = False
    widths(ss, {"B": 10, "C": 32, **{L(c): 11.5 for c in range(4, 12)}})
    ss.freeze_panes = "D6"

    # ---------- Action List (static ranking for the default period)
    al = S["Action List"]
    a_from, a_to = pd.Timestamp(cur_fy - 1, 4, 1), months[-1]
    l_from, l_to = a_from - pd.DateOffset(years=1), a_to - pd.DateOffset(years=1)
    cyv = data[(data.month >= a_from) & (data.month <= a_to)].groupby("customer_id").litres.sum()
    lyv = data[(data.month >= l_from) & (data.month <= l_to)].groupby("customer_id").litres.sum()
    act = cust.assign(cy=cyv.reindex(cust.index).fillna(0), ly=lyv.reindex(cust.index).fillna(0))
    title(al, "Action List - lost and declining counters",
          f"Fixed list for {a_from:%b-%y} to {a_to:%b-%y} vs {l_from:%b-%y} to {l_to:%b-%y} (rebuild the report to refresh). "
          "For any other period filter the Customers tab on Status.")
    cols_a = ["Customer Id", "Customer Name", "Segment", "SSF (latest bill)", "SSR", "Zone", "Area", "Mobile",
              "LY (L)", "This year (L)", "Drop (L)", "Last bill"]

    def block(r, label, df):
        al.cell(r, 2, label).font = F_SEC
        head(al, r + 1, 2, cols_a)
        for i, (cid, c) in enumerate(df.iterrows()):
            x = r + 2 + i
            for j, v in enumerate([cid, c["name"], c.segment, c.ssr, c.stockist, c.zone, c.area, c.mobile], 2):
                put(al, x, j, v)
            put(al, x, 10, round(c.ly), LIT); put(al, x, 11, round(c.cy), LIT)
            put(al, x, 12, f"=J{x}-K{x}", LIT, bold=True)
            put(al, x, 13, c["last"].to_pydatetime(), DATE)
        return r + 2 + len(df) + 2

    lost = act[(act.ly > 0.5) & (act.cy <= 0.5)].sort_values("ly", ascending=False)
    dec = act[(act.cy > 0.5) & (act.ly > 0) & (act.cy < act.ly * 0.9)].assign(drop=lambda x: x.ly - x.cy).sort_values("drop", ascending=False).head(75)
    nxt = block(4, f"LOST: billed last year, nothing this year ({len(lost)} counters, {lost.ly.sum() / 1000:,.1f} KL last year)", lost)
    block(nxt, f"DECLINING: top {len(dec)} counters by volume drop (down 10% or more)", dec)
    widths(al, {"B": 17, "C": 36, "D": 8, "E": 17, "F": 20, "G": 26, "H": 24, "I": 12, "J": 10, "K": 11, "L": 10, "M": 10})
    al.freeze_panes = "D4"

    # ---------- Area Master (reference)
    am = S["Area Master"]
    title(am, "Area Master used in this report", f"Edit Data/{AREA_MASTER.name} (not this tab) and rebuild to change any mapping.")
    head(am, 4, 2, ["Pincode", "Area", "Zone", "Customers"])
    pc = cust[cust.pin.ne("")].groupby("pin").size()
    prow = sorted(((p, a, z) for p, (a, z, _) in M["pins"].items() if a), key=lambda t: (ZONE_ORDER.index(t[2]) if t[2] in ZONE_ORDER else 99, t[0]))
    for i, (p, a, z) in enumerate(prow):
        x = 5 + i
        put(am, x, 2, p); put(am, x, 3, a); put(am, x, 4, z); put(am, x, 5, int(pc.get(p, 0)), CNT)
    widths(am, {"B": 10, "C": 34, "D": 40, "E": 11})
    am.freeze_panes = "B5"

    # ---------- Data Check
    dc = S["Data Check"]
    title(dc, "Data Check", "Reconciliation and items that need a mapping fix.")
    raw_total = float(sec.litres.sum())
    rows = [
        ("Secondary files read", ", ".join(M["files"])),
        ("Customer Master file", M["cm_file"] or "NOT FOUND - areas only from Area Master"),
        ("Area Master file", f"Data/{AREA_MASTER.name}"),
        ("Invoice dates covered", f"{sec.date.min():%d-%b-%Y} to {sec.date.max():%d-%b-%Y}"),
        ("Secondary litres after ClearBlue exclusion (source)", round(raw_total, 1)),
        ("Litres in the Data tab", "=SUM(dL)"),
        ("Difference (must be 0)", "=ROUND(C10-C9,1)"),
        ("ClearBlue (DEF) litres excluded", round(M["clearblue"], 1)),
        ("Customers", len(cust)),
        ("Customers with an area (pincode mapped)", int((~cust.zone.eq(UNMAPPED)).sum())),
        ("Customers without an area", int(cust.zone.eq(UNMAPPED).sum())),
        ("Litres without an area, current FY", round(float(data[(data.zone == UNMAPPED) & (data.fy == cur_fy)].litres.sum()), 1)),
        ("Share of current-FY litres without an area", "=IF(SUMIFS(dL,dFY,\"" + fy_label(cur_fy) + "\")=0,\"\",C16/SUMIFS(dL,dFY,\"" + fy_label(cur_fy) + "\"))"),
    ] + [("File warning", p) for p in M["problems"]]
    head(dc, 4, 2, ["Check", "Value"])
    for i, (k, v) in enumerate(rows):
        x = 5 + i
        put(dc, x, 2, k, bold=True); put(dc, x, 3, v, LIT if isinstance(v, float) else None)
    dc["C17"].number_format = PCT
    dc["C10"].number_format = LIT
    dc.conditional_formatting.add("C11", CellIsRule(operator="notEqual", formula=["0"], fill=PatternFill("solid", fgColor="FF9999")))
    um = cust[cust.zone.eq(UNMAPPED)].copy()
    um["l12"] = data[data.month > months[-1] - pd.DateOffset(months=12)].groupby("customer_id").litres.sum().reindex(um.index).fillna(0)
    um = um.sort_values("l12", ascending=False)
    x0 = 6 + len(rows) + 1
    dc.cell(x0, 2, f"CUSTOMERS WITHOUT AN AREA ({len(um)}) - fill their Pincode in Data/{AREA_MASTER.name}, sheet 'Customer Location'").font = F_SEC
    head(dc, x0 + 1, 2, ["Customer Id", "Customer Name", "Segment", "SSR", "District", "Pincode found", "Litres last 12 months"])
    for i, (cid, c) in enumerate(um.iterrows()):
        x = x0 + 2 + i
        for j, v in enumerate([cid, c["name"], c.segment, c.stockist, c.district, c.pin], 2):
            put(dc, x, j, v)
        put(dc, x, 8, round(c.l12), LIT)
    widths(dc, {"B": 48, "C": 40, "D": 10, "E": 24, "F": 18, "G": 13, "H": 14})

    # ---------- Notes
    nt = S["Notes"]
    title(nt, "Notes & definitions")
    nt["B4"], nt["C4"] = "Latest invoice date in data", last_date.to_pydatetime()
    nt["C4"].number_format = DATE
    nt["B5"], nt["C5"] = "Report built on", datetime.now()
    nt["C5"].number_format = "dd-mmm-yy hh:mm"
    notes = [
        ("Source", "DMS Secondary Sales Detail Report (stockist to counter). Same file merge rules as the HO dashboard: overlapping months are not double counted."),
        ("Volume", "Product Quantity (L/Kg). Summary tabs in KL (thousand litres), Customers and Action List in litres. Returns are netted."),
        ("Excluded", "IOC ClearBlue (DEF) - not a lubricant."),
        ("Segment", "From the invoice's Customer Group: Bazaar = Bazaar + ServoXpert + Others; RO = Dealer (IOC retail outlets) + Nayara RO; OEM; Fleet. "
                    "Each customer is placed in the segment of its latest invoice."),
        ("Sub-segment", "Customer Sub Group on the latest invoice (e.g. MARUTI - ARENA, SPARE PART SHOP, IOC-NON KSK)."),
        ("Area / Zone", "Customer pincode (Lubes Customer Master) mapped to a locality (Area) and a sales territory (Zone) via Data/Area Master.xlsx. "
                        "ROs are not in the Customer Master: their pincodes come from the Area Master 'Customer Location' sheet."),
        ("SSR", "Stockist (Doshi, AMCO, Madhuban). Area Wise shows each area's sales by SSR and its Main SSR (largest sales in the period)."),
        ("SSF", "SSF Name on the invoice (the SSR's salesman). Area x SSF and SSF Wise use the SSF on each invoice; the Customers tab shows the SSF of the latest bill."),
        ("Period", "Dashboard From/To months drive every tab except Period Trend. LY = the same months one year earlier (like-for-like)."),
        ("Counters billed", "Customers with net positive volume in the period."),
        ("Status", "New = first ever bill falls in the period; Re-activated = billed now, nothing in the LY period but billed before; Lost = billed in LY period, "
                   "nothing now; Growing / Declining = +/-10% or more vs LY; Stable = within +/-10%."),
        ("Data tab", "One row per customer x month x SSR x SSF - use it for your own pivot tables."),
        ("Refresh", "Drop the new Secondary file (and a newer Customer Master if any) in Data, then double-click 'Build Area Report.command' in sales_dashboard."),
    ]
    for i, (k, v) in enumerate(notes):
        x = 7 + i
        nt.cell(x, 2, k).font = F_BOLD
        c = nt.cell(x, 3, v); c.font = F_BASE; c.alignment = Alignment(wrap_text=True, vertical="top")
        nt.row_dimensions[x].height = 30
    widths(nt, {"B": 26, "C": 120})

    # fonts everywhere that are still default
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                if c.value is not None and (c.font is None or c.font.name != FONT):
                    c.font = Font(name=FONT, size=c.font.size if c.font and c.font.size else 10, bold=c.font.b if c.font else False,
                                  color=c.font.color if c.font else None)
    S["Dashboard"].sheet_properties.tabColor = ORANGE
    for n_ in ("Data", "Area Master", "Data Check", "Notes"):
        S[n_].sheet_properties.tabColor = "A6A6A6"
    wb.active = 0

    OUT_DIR.mkdir(exist_ok=True)
    # Built and recalculated outside OneDrive (Excel's local folder), then copied in once: when Excel opens or saves a
    # file inside OneDrive it syncs it its own way and older cloud copies come back over the new report.
    work = EXCEL_WORK if EXCEL_WORK.parent.parent.exists() else OUT_DIR
    work.mkdir(parents=True, exist_ok=True)
    tmp = work / f"_building {datetime.now():%Y%m%d-%H%M%S}.xlsx"
    wb.save(tmp)
    return tmp, OUT_DIR / f"Secondary Sales - Segment Area Customer {last_date:%b-%y}.xlsx"


def excel_recalc(path: Path) -> bool:
    """Lets Microsoft Excel calculate and save the file so it also shows numbers in previews. Optional."""
    import subprocess, time
    script = (f'tell application "Microsoft Excel"\n open POSIX file "{path}"\n delay 3\n set wb to active workbook\n'
              ' calculate full\n save wb\n close wb saving no\nend tell')
    for _ in range(2):
        try:
            if subprocess.run(["osascript", "-e", script], capture_output=True, timeout=300).returncode == 0:
                return True
        except Exception:
            pass
        time.sleep(8)
    return False


def main():
    M = build_model()
    for old in [*OUT_DIR.glob("_building *.xlsx"), *EXCEL_WORK.glob("_building *.xlsx")]:      # leftovers of earlier runs
        try:
            old.unlink()
        except OSError:
            pass
    tmp, out = build_workbook(M)
    if "--no-excel" not in sys.argv:
        print("Excel recalculation:", "done" if excel_recalc(tmp) else "skipped (Excel will calculate when you open the file)")
    try:
        time.sleep(2)
        shutil.copyfile(tmp, out)
        tmp.unlink()
    except OSError:
        out = tmp
        print("! The previous report is open in Excel - close it and rebuild. New copy kept as " + tmp.name)
    cust = M["cust"]
    print(f"Saved: {out}")
    print(f"Customers: {len(cust)}  |  without area: {(cust.zone == UNMAPPED).sum()}  |  data rows: {len(M['data'])}")
    for p in M["problems"]:
        print("! " + p)
    return out


if __name__ == "__main__":
    main()
