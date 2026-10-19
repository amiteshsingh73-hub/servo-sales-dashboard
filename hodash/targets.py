"""Reads primary + secondary targets straight from the Target workbook (no intermediate file).

Layout understood (as in 'Target File with Sales FY2627.xlsx', sheet 'FY2526 Base Target'):
  * Block header row:  SAP Code | Name of Stockist | [Group] | APR'26 ... MAR'27 | FY 26-27
      - with a 'Group' column -> secondary targets, one row per channel (Bazar/RO/OEM/Fleet)
      - without it            -> primary target, one row per stockist
  * Fallback for a stockist with no primary block: the 'Month | BASE | TARGET | Sales' summary table
    whose stockist names (DOSHI / AMCO / MADHUBAN) sit on the row above.
If a stockist has several primary blocks the first one in the sheet wins.
"""
from __future__ import annotations
import re
import numpy as np
import pandas as pd
from . import config as C

MON = r"^([A-Za-z]{3,4})\s*['’]\s*(\d{2})$"
MON_NO = {m: i for i, m in enumerate(["APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "JAN", "FEB", "MAR"], 1)}
CH = {"bazar": "Bazaar", "bazaar": "Bazaar", "baazaar": "Bazaar", "ro": "RO", "oem": "OEM", "fleet": "Fleet"}
COLS = ["fy", "fm", "stockist_code", "kind", "channel", "target"]


def _num(x):
    try:
        v = float(x)
        return v if v == v else None
    except (TypeError, ValueError):
        return None


def _month_cols(row):
    """Column indexes of the 12 month labels (APR'yy..MAR'yy+1) in a header row, and the fiscal year."""
    cols = []
    for j, v in enumerate(row):
        m = re.match(MON, str(v).strip())
        if m and m.group(1).upper()[:3] in MON_NO:
            cols.append((j, MON_NO[m.group(1).upper()[:3]], int(m.group(2))))
    if len(cols) < 12: return None, None
    cols = cols[:12]
    if [c[1] for c in cols] != list(range(1, 13)): return None, None
    return [c[0] for c in cols], 2000 + cols[0][2] + 1


def find_target_file():
    fs = [f for f in C.DATA_DIR.glob("*") if f.is_file() and not f.name.startswith(("~$", "."))
          and f.suffix.lower() in (".xlsx", ".xlsm") and re.match(C.TARGET_PATTERN.replace("*", ".*"), f.name, re.I)]
    return max(fs, key=lambda p: p.stat().st_mtime) if fs else None


def parse(path, name_to_code: dict[str, int]):
    """Returns (DataFrame[COLS], warnings[list[str]])."""
    rows, warn = [], []
    if path is None:
        return pd.DataFrame(columns=COLS), ["No target file found (name it 'Target*.xlsx' in the data folder)."]
    xl = pd.ExcelFile(path, engine="calamine")
    sheet = C.TARGET_SHEET if C.TARGET_SHEET in xl.sheet_names else next(
        (s for s in xl.sheet_names if "base target" in s.lower() and s != "Base Target"), None)
    if sheet is None:
        return pd.DataFrame(columns=COLS), [f"Sheet '{C.TARGET_SHEET}' not found in {path.name}."]
    sh = xl.parse(sheet, header=None)
    n = len(sh)
    seen_p, seen_s, fy_found = set(), set(), None

    def add(code, kind, ch, vals, fy):
        for fm, v in enumerate(vals, 1):
            v = _num(v)
            if v is not None:
                rows.append({"fy": fy, "fm": fm, "stockist_code": int(code), "kind": kind, "channel": ch, "target": v})

    for i in range(n):
        r = sh.iloc[i].tolist()
        mc, fy = _month_cols(r)
        if not mc: continue
        fy_found = fy
        is_sec = any(str(v).strip().lower() == "group" for v in r[:mc[0]])
        total_col = mc[-1] + 1
        k = i + 1
        while k < n:
            d = sh.iloc[k].tolist()
            code = _num(d[0])
            if code is None:
                if is_sec and str(d[2]).strip().lower().startswith("total"):   # 'Total Secondary Sale' row -> skip
                    k += 1; continue
                break
            if is_sec:
                ch = CH.get(str(d[2]).strip().lower())
                if ch and (int(code), ch) not in seen_s:
                    seen_s.add((int(code), ch)); add(code, "Secondary", ch, [d[j] for j in mc], fy)
            elif int(code) not in seen_p:
                seen_p.add(int(code)); vals = [d[j] for j in mc]; add(code, "Primary", "", vals, fy)
                stated = _num(d[total_col]) if total_col < len(d) else None
                s = sum(_num(v) or 0 for v in vals)
                if stated and abs(s - stated) > 0.5:
                    warn.append(f"Primary {int(code)}: months add to {s:,.0f} but the sheet total says {stated:,.0f}.")
            k += 1
    # fallback: summary table 'Month | BASE | TARGET | Sales' under a row of stockist names
    for i in range(1, n):
        r = sh.iloc[i].tolist()
        for j, v in enumerate(r):
            if str(v).strip().lower() != "month": continue
            names = sh.iloc[i - 1].tolist()
            for jn, nm in enumerate(names):
                if not isinstance(nm, str) or not nm.strip() or jn < j: continue
                code = next((c for t, c in name_to_code.items() if nm.strip().lower() in t.lower()), None)
                if code is None or code in seen_p: continue
                tcol = next((c for c in range(jn, min(jn + 4, len(r))) if str(r[c]).strip().lower() == "target"), None)
                if tcol is None: continue
                vals = [sh.iloc[i + 1 + m, tcol] for m in range(12)] if i + 12 < n else []
                if len(vals) == 12 and all(_num(x) is not None for x in vals):
                    fyy = fy_found or 2027
                    seen_p.add(code); add(code, "Primary", "", vals, fyy)
                    stated = _num(sh.iloc[i + 13, tcol]) if i + 13 < n else None
                    s = sum(_num(x) for x in vals)
                    if stated and abs(s - stated) > 0.5:
                        warn.append(f"Primary target for {nm.strip()} (summary table, sheet '{sheet}'): months add to {s:,.0f} "
                                    f"but the stated total is {stated:,.0f}. Check the monthly cells.")
    t = pd.DataFrame(rows, columns=COLS)
    if t.empty: warn.append("No targets recognised in the target file.")
    return t, warn
