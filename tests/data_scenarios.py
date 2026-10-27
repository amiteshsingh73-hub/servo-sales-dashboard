"""Crash test with messy data folders: renders the whole dashboard against temporary copies of the Data folder
(empty, missing files, corrupt files, missing columns, a half month, no history...). Your real files are only read.
Run:  ~/.local/venvs/sales_review_ho/bin/python tests/data_scenarios.py
"""
import os, shutil, sys, tempfile, time
from pathlib import Path
import pandas as pd

APP_DIR = Path(__file__).resolve().parent.parent
SRC = Path(os.environ.get("SALES_DATA_DIR", APP_DIR.parent / "Data")).resolve()
os.environ["SALES_REVIEW_NO_AUTOSAVE"] = "1"
sys.path.insert(0, str(APP_DIR))


def _first(pattern):
    m = sorted(SRC.glob(pattern))
    return m[0] if m else None


PRI, SEC, TGT = _first("Primary*"), _first("Secondary*"), _first("Target*")
EXTRA = [f for f in SRC.iterdir() if f.is_file() and not f.name.startswith(("Primary", "Secondary", "Target", "~$", "."))]


def build(tmp: Path, name: str):
    """Fill tmp (an empty data folder) for scenario `name`."""
    from hodash import loader
    cp = lambda f: f and shutil.copy(f, tmp / f.name)
    full = lambda: [cp(PRI), cp(SEC), cp(TGT)] + [cp(f) for f in EXTRA]
    if name == "empty folder": return
    if name == "data folder missing": tmp.rmdir(); return
    if name == "only primary": cp(PRI); return
    if name == "only secondary": cp(SEC); return
    if name == "no target, no masters": cp(PRI); cp(SEC); return
    full()
    if name == "full": return
    if name == "junk files":
        (tmp / "Primary corrupt.xlsx").write_text("not excel")
        (tmp / "Secondary half-synced.xlsx").write_bytes(b"PK\x03\x04" + b"\0" * 200)
        (tmp / "~$Primary Sales.XLSX").write_text("lock")
        (tmp / "Target half-synced.xlsx").write_bytes(b"PK\x03\x04" + b"\0" * 300)
        pd.DataFrame({"x": [1, 2]}).to_excel(tmp / "Target wrong layout.xlsx", index=False)
        pd.DataFrame({"y": [1]}).to_excel(tmp / "Secondary wrong.xlsx", index=False)
        return
    raw_p, _ = loader.read_table(tmp / PRI.name, "primary")
    raw_s, _ = loader.read_table(tmp / SEC.name, "secondary")
    pdate = pd.to_datetime(raw_p["Inv.Date"], dayfirst=True, errors="coerce")
    sdate = pd.to_datetime(raw_s["Invoice Date"], dayfirst=True, errors="coerce")
    if name in ("new half month (both)", "new month primary only"):
        last = pdate.max()
        nxt = raw_p[(pdate > last - pd.Timedelta(days=30)) & (pdate <= last - pd.Timedelta(days=20))].copy()
        nxt["Inv.Date"] = (pd.to_datetime(nxt["Inv.Date"], dayfirst=True) + pd.DateOffset(months=1)).dt.strftime("%d.%m.%Y")
        nxt.to_csv(tmp / "Primary Sales next.csv", index=False)
        if name == "new half month (both)":
            s = raw_s[(sdate > sdate.max() - pd.Timedelta(days=30)) & (sdate <= sdate.max() - pd.Timedelta(days=20))].copy()
            s["Invoice Date"] = (pd.to_datetime(s["Invoice Date"], dayfirst=True) + pd.DateOffset(months=1)).dt.strftime("%d-%m-%Y")
            s.to_csv(tmp / "Secondary Sales next.csv", index=False)
        return
    if name == "secondary missing optional columns":
        (tmp / SEC.name).unlink()
        drop = [c for c in raw_s.columns if c in ("SSF Name", "SSF Code", "Taxable Amount (Rs.)", "Customer District", "District",
                                                     "Customer Sub Group", "Sales Area", "Customer Sub Class", "Order Type")]
        raw_s.drop(columns=drop).to_csv(tmp / "Secondary Sales.csv", index=False)
        return
    if name == "only current year (no LY)":
        (tmp / PRI.name).unlink(); (tmp / SEC.name).unlink()
        raw_p[pdate >= pdate.max() - pd.DateOffset(months=5)].to_csv(tmp / "Primary Sales.csv", index=False)
        raw_s[sdate >= sdate.max() - pd.DateOffset(months=5)].to_csv(tmp / "Secondary Sales.csv", index=False)
        return
    if name == "one stockist only":
        (tmp / PRI.name).unlink(); (tmp / SEC.name).unlink()
        code = raw_p["Payer"].iloc[0] if "Payer" in raw_p else None
        raw_p[raw_p["Payer"] == code].to_csv(tmp / "Primary Sales.csv", index=False)
        raw_s[pd.to_numeric(raw_s["Stockist Code"], errors="coerce") == pd.to_numeric(code, errors="coerce")].to_csv(tmp / "Secondary Sales.csv", index=False)
        return
    raise ValueError(name)


SCENARIOS = ["full", "data folder missing", "empty folder", "only primary", "only secondary", "no target, no masters", "junk files",
             "new half month (both)", "new month primary only", "secondary missing optional columns",
             "only current year (no LY)", "one stockist only"]


NO_DASH = {"empty folder", "data folder missing", "only primary", "only secondary"}   # these must show the 'add files' screen


def run(name):
    from streamlit.testing.v1 import AppTest
    tmp = Path(tempfile.mkdtemp()) / "Data"; tmp.mkdir()
    try:
        build(tmp, name)
        os.environ["SALES_DATA_DIR"] = str(tmp)
        for m in [m for m in sys.modules if m.startswith("hodash")]: del sys.modules[m]     # config re-reads SALES_DATA_DIR
        at = AppTest.from_file(str(APP_DIR / "app.py"), default_timeout=900).run()
        exc = [e.message for e in at.exception]
        oops = [m.value for m in at.markdown if "class='oops'" in str(m.value)]
        errs = [e.value for e in at.error]
        warns = " ".join(str(w.value) for w in at.warning)
        ntabs = len(at.tabs)
        if name not in NO_DASH and ntabs == 0:
            exc.append("dashboard did not open although usable Primary + Secondary files were present: " + (str(errs[0])[:150] if errs else ""))
        if name == "junk files" and "Using " not in warns:
            exc.append("targets were lost: the good Target workbook was not used when a newer one was broken")
        return exc + [o[:200] for o in oops], errs, ntabs
    finally:
        shutil.rmtree(tmp.parent, ignore_errors=True)


def main():
    bad = 0
    for name in (sys.argv[1:] or SCENARIOS):
        t = time.time()
        try:
            fails, errs, ntabs = run(name)
        except Exception as e:  # noqa: BLE001
            fails, errs, ntabs = [f"harness: {type(e).__name__}: {e}"], [], 0
        ok = not fails
        bad += not ok
        note = f"{ntabs} tabs" + (f" | message: {str(errs[0])[:110]}" if errs else "")
        print(f"{'PASS' if ok else 'FAIL'}  {name:38s} {time.time() - t:5.1f}s  {note}")
        for f in fails[:3]: print("      ", str(f)[:300])
    print("\nALL DATA SCENARIOS PASSED" if not bad else f"\n{bad} DATA SCENARIO(S) FAILED")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
