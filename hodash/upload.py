"""Add sales files from the dashboard: each uploaded workbook is recognised from its columns (Primary / Secondary /
Target / Material master / Pack master), checked, and saved into the Data folder under a clear name. The dashboard then
picks it up like any file dropped into the folder. Nothing leaves the computer.
"""
from __future__ import annotations
import hashlib, os, re, tempfile, time
from pathlib import Path
import pandas as pd

from . import config as C, loader

KINDS = {"primary": "Primary Sales", "secondary": "Secondary Sales", "material": "Material master", "pack": "Pack Master"}
OK_EXT = {".xlsx", ".xlsm", ".xls", ".csv"}


def _md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def _existing_hashes(folder: Path) -> dict[str, str]:
    out = {}
    for f in folder.iterdir() if folder.is_dir() else []:
        if f.is_file() and f.suffix.lower() in OK_EXT and not f.name.startswith(("~$", ".")):
            try:
                if f.stat().st_size < 200_000_000: out[_md5(f.read_bytes())] = f.name
            except OSError:
                pass
    return out


def _is_target(path: Path) -> bool:
    if path.suffix.lower() not in (".xlsx", ".xlsm"): return False
    try:
        names = pd.ExcelFile(path, engine="calamine").sheet_names
    except Exception:
        return False
    return C.TARGET_SHEET in names or any("base target" in s.lower() for s in names)


def _months(df: pd.DataFrame, kind: str) -> str:
    col = next((c for c in (["Inv.Date", "Invoice Date", "Billing Date"] if kind == "primary" else ["Invoice Date"]) if c in df), None)
    if col is None: return ""
    d = loader._to_datetime(df[col]).dropna()
    if d.empty: return ""
    a, b = d.min(), d.max()
    return a.strftime("%b%y") if (a.year, a.month) == (b.year, b.month) else f"{a:%b%y}-{b:%b%y}"


def save_upload(name: str, data: bytes, folder: Path | None = None) -> dict:
    """Recognise and store one uploaded file. Returns {ok, kind, saved_as, message}."""
    folder = Path(folder or C.DATA_DIR)
    ext = Path(name).suffix.lower()
    if ext not in OK_EXT:
        return dict(ok=False, kind=None, saved_as=None, message=f"{name}: only Excel (.xlsx/.xls/.xlsm) or .csv files can be added.")
    if not data:
        return dict(ok=False, kind=None, saved_as=None, message=f"{name}: the file is empty.")
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return dict(ok=False, kind=None, saved_as=None, message=f"Data folder {folder} is not available ({e}).")
    dup = _existing_hashes(folder).get(_md5(data))
    if dup:
        return dict(ok=True, kind=None, saved_as=dup, message=f"{name}: already in the Data folder as '{dup}' - nothing to do.")

    tmpdir = Path(tempfile.mkdtemp())
    probe = tmpdir / f"upload{ext}"
    probe.write_bytes(data)
    try:
        kind, detail = None, ""
        for k in ("primary", "secondary"):
            df, prob = loader.read_table(probe, k)
            if df is not None and len(df):
                kind, detail = k, f"{len(df):,} rows" + (f", {_months(df, k)}" if _months(df, k) else "")
                break
        if kind is None and _is_target(probe):
            kind, detail = "target", "target workbook"
        if kind is None:
            for k in ("material", "pack"):
                df, _ = loader.read_table(probe, k)
                if df is not None and len(df):
                    kind, detail = k, f"{len(df):,} rows"; break
        if kind is None:
            return dict(ok=False, kind=None, saved_as=None,
                        message=(f"{name}: not recognised. It needs the usual column headings of the SAP Primary report "
                                 "(Inv.Date, KL, Payer, Material), the DMS Secondary Sales Detail Report (Invoice Date, "
                                 "Product Quantity(L/Kg), Stockist Code, SKU Code) or the Target workbook. Nothing was saved."))
        stem = re.sub(r"[^\w\-. ()']", "", Path(name).stem).strip() or "upload"
        prefix = "Target File" if kind == "target" else KINDS[kind]
        base = stem if stem.lower().startswith(prefix.split()[0].lower()) else f"{prefix} - {stem}"
        stamp = time.strftime("%d%b%y %H%M")
        final = folder / f"{base} (added {stamp}){ext}"
        n = 2
        while final.exists():
            final = folder / f"{base} (added {stamp}) {n}{ext}"; n += 1
        part = folder / f".{final.name}.part"            # hidden while writing, so a half-written file is never read
        part.write_bytes(data)
        os.replace(part, final)
        return dict(ok=True, kind=kind, saved_as=final.name, message=f"{name}: added as {prefix.split()[0]} file '{final.name}' ({detail}).")
    except Exception as e:  # noqa: BLE001
        return dict(ok=False, kind=None, saved_as=None, message=f"{name}: could not be added ({type(e).__name__}: {e}). Nothing was saved.")
    finally:
        try:
            probe.unlink(); tmpdir.rmdir()
        except OSError:
            pass
