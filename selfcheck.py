"""Safety check: simulates messy file drops in a temporary copy of your folder and confirms the dashboard copes.
Run:  ~/.local/venvs/sales_review_ho/bin/python selfcheck.py       (your real files are only read, never changed)
"""
import os, shutil, sys, tempfile
from pathlib import Path
import pandas as pd

_root = Path(__file__).resolve().parent.parent
SRC = Path(os.environ.get("SALES_DATA_DIR", _root / "Data" if (_root / "Data").is_dir() else _root))
tmp = Path(tempfile.mkdtemp())
for pat in ("Primary Sales*.XLSX", "Secondary Sales*.xlsx", "Target File*.xlsx"):
    for f in SRC.glob(pat): shutil.copy(f, tmp / f.name)
os.environ["SALES_DATA_DIR"] = str(tmp)
sys.path.insert(0, str(Path(__file__).parent))
from hodash import loader  # noqa: E402

ok = True
def check(name, cond, extra=""):
    global ok; ok &= bool(cond); print(("PASS " if cond else "FAIL ") + name, extra)

base = loader.load_all(); b0 = base.primary.kl.sum()
check("real files load", len(base.primary) and len(base.secondary) and not base.info["problems"])

raw = loader.read_table(next(tmp.glob("Primary Sales*")), "primary")[0]
last = pd.to_datetime(raw["Inv.Date"]).max()
new = raw[pd.to_datetime(raw["Inv.Date"]) >= last.replace(day=1)].copy()
oct_, nov = new.head(30).copy(), new.head(10).copy()
oct_["Inv.Date"] = (pd.to_datetime(oct_["Inv.Date"]) + pd.DateOffset(months=1)).dt.strftime("%d-%m-%Y")      # text dd-mm-yyyy, title rows above
with pd.ExcelWriter(tmp / "Primary Sales Oct.xlsx") as w:
    pd.DataFrame([["Primary report"], [""]]).to_excel(w, sheet_name="Data", header=False, index=False)
    oct_.to_excel(w, sheet_name="Data", startrow=2, index=False)
nov["Inv.Date"] = (pd.to_datetime(nov["Inv.Date"]) + pd.DateOffset(months=2) - pd.Timestamp("1899-12-30")).dt.days  # Excel serials
nov.to_csv(tmp / "Primary Sales Nov.csv", index=False)
pd.DataFrame({"x": [1]}).to_excel(tmp / "Primary wrong.xlsx", index=False)
(tmp / "Primary corrupt.xlsx").write_text("not excel")
(tmp / "~$Primary Sales.XLSX").write_text("lock")

d = loader.load_all()
exp = b0 + oct_.KL.sum() + nov.KL.sum()
check("monthly files (text dates, serial dates, title rows, csv) merged", abs(d.primary.kl.sum() - exp) < 0.01, f"{d.primary.kl.sum():.1f} vs {exp:.1f}")
exp_last = max((pd.to_datetime(oct_["Inv.Date"], format="%d-%m-%Y")).max(),
               (pd.Timestamp("1899-12-30") + pd.to_timedelta(nov["Inv.Date"], unit="D")).max(), last)
check("no date flipped (ISO / day-first / serial)", d.primary.date.max() == exp_last, f"{d.primary.date.max():%d-%b-%Y} vs {exp_last:%d-%b-%Y}")
check("bad / corrupt files skipped with messages, no crash", len(d.info["problems"]) == 2, f"{len(d.info['problems'])} problems")
check("lock file ignored", all("~$" not in f for f in d.info["files"]["primary"]))

(tmp / "Primary Sales.XLSX").unlink()
for f in tmp.glob("Primary Sales*"): f.unlink()
try:
    e = loader.load_all(); check("no usable primary file -> clean result, no crash", e.primary.empty)
except Exception as ex:
    check("no usable primary file -> clean result, no crash", False, repr(ex))
shutil.rmtree(tmp, ignore_errors=True)
print("\nALL CHECKS PASSED" if ok else "\nSOME CHECKS FAILED"); sys.exit(0 if ok else 1)
