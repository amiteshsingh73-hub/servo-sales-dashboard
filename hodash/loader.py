"""Finds the sales files, reads + normalises them, and merges monthly drops.

Merge rule: for every calendar month the file whose data for that month reaches the latest
invoice date wins (ties: most rows, then newest file). So you can drop a monthly file, a fuller
re-export of the same month, or a whole-year file - overlaps never double count.
"""
from __future__ import annotations
import hashlib, re
from dataclasses import dataclass, field
from pathlib import Path
import numpy as np
import pandas as pd
from . import config as C, targets


# ---------------------------------------------------------------- file discovery
def _match(pattern: str) -> list[Path]:
    rx = re.compile("^" + re.escape(pattern).replace(r"\*", ".*") + "$", re.I)
    out = []
    for f in C.DATA_DIR.iterdir():
        if f.is_file() and not f.name.startswith(("~$", ".")) and f.suffix.lower() in C.EXTENSIONS and rx.match(f.name):
            out.append(f)
    return sorted(out)


def discover() -> dict[str, list[Path]]:
    return {k: _match(p) for k, p in C.PATTERNS.items()}


def signature() -> str:
    """Cheap fingerprint of all watched files (+targets). Changes when a file is added/replaced."""
    parts = []
    for k, fs in discover().items():
        for f in fs:
            st = f.stat()
            parts.append(f"{k}|{f.name}|{st.st_size}|{int(st.st_mtime)}")
    tf = targets.find_target_file()
    if tf:
        st = tf.stat(); parts.append(f"t|{tf.name}|{st.st_size}|{int(st.st_mtime)}")
    return hashlib.md5("\n".join(parts).encode()).hexdigest()


# ---------------------------------------------------------------- reading
REQUIRED = {
    "primary": [["Inv.Date", "Invoice Date", "Billing Date"], ["KL", "Billing qty in SKU"], ["Payer", "Stockist Code"], ["Material"]],
    "secondary": [["Invoice Date"], ["Product Quantity(L/Kg)"], ["Stockist Code"], ["SKU Code"]],
    "material": [["PROD CODE"], ["Category"]],
    "pack": [["SAP PACK CODE"]],
}


def _norm(x) -> str:
    return re.sub(r"[^a-z0-9]", "", str(x).lower())


def _has(cols, groups) -> bool:
    cn = {_norm(c) for c in cols}
    return all(any(_norm(a) in cn for a in g) for g in groups)


def _find_header(raw: pd.DataFrame, groups):
    """Row index (within the first 25 rows) whose cells contain every required column, else None."""
    for i in range(min(25, len(raw))):
        if _has(raw.iloc[i].tolist(), groups): return i
    return None


def read_table(path: Path, kind: str):
    """Reads the sheet of `path` that actually holds the data (other sheets such as pivots are ignored, the header
    row may sit below a title). Returns (DataFrame | None, problem message | None)."""
    groups = REQUIRED[kind]
    try:
        if path.suffix.lower() == ".csv":
            sheets = {"csv": pd.read_csv(path, header=None, low_memory=False)}
        else:
            xl = pd.ExcelFile(path, engine="calamine")
            sheets = {n: xl.parse(n, header=None) for n in xl.sheet_names}
    except Exception as e:
        return None, f"{path.name}: could not be opened ({type(e).__name__}: {e}). Is it open/locked in Excel or still syncing?"
    best = None
    for name, raw in sheets.items():
        h = _find_header(raw, groups)
        if h is None: continue
        df = raw.iloc[h + 1:].copy(); df.columns = [str(c).strip() for c in raw.iloc[h]]
        df = df.dropna(how="all").reset_index(drop=True)
        if best is None or len(df) > len(best[1]): best = (name, df)
    if best is None:
        need = "; ".join(" / ".join(g) for g in groups)
        return None, (f"{path.name}: no sheet with the expected {kind} columns ({need}). "
                      f"Sheets seen: {', '.join(map(str, sheets))}. File skipped.")
    return best[1], None


def _read(path: Path, sheet=0) -> pd.DataFrame:      # kept for simple callers
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, low_memory=False)
    return pd.read_excel(path, sheet_name=sheet, engine="calamine")


def _col(df: pd.DataFrame, *names):
    """Case/space-insensitive column lookup."""
    norm = {re.sub(r"[^a-z0-9]", "", c.lower()): c for c in df.columns}
    for n in names:
        k = re.sub(r"[^a-z0-9]", "", n.lower())
        if k in norm:
            return df[norm[k]]
    return None


def _num_col(x) -> pd.Series:
    if x is None: return pd.Series(dtype=float)
    if x.dtype == object: x = x.astype(str).str.replace(",", "", regex=False)
    return pd.to_numeric(x, errors="coerce")


def _clean_name(x: pd.Series) -> pd.Series:
    return x.astype(str).str.strip().str.replace(r"[\s.,]+$", "", regex=True)


# ---------------------------------------------------------------- masters
def load_masters(files: dict[str, list[Path]], primary_raw: list[pd.DataFrame], problems: list):
    prod = pd.DataFrame(columns=["key", "m_product", "m_category", "m_segment", "m_spec"])
    if files["material"]:
        mm, prob = read_table(files["material"][-1], "material")
        if prob: problems.append(prob + " (optional file)")
        mm = mm if mm is not None else pd.DataFrame({"PROD CODE": [], "Category": []})
        base = pd.DataFrame({
            "m_product": _col(mm, "Product Name"), "m_category": _col(mm, "Category"),
            "m_segment": _col(mm, "Segment(MAT-1)"), "m_spec": _col(mm, "Spec(MAT-2)")})
        keys = []
        for c in ("PROD CODE", "Text"):
            s = _col(mm, c)
            if s is not None:
                k = base.copy(); k["key"] = pd.to_numeric(s, errors="coerce"); keys.append(k)
        prod = pd.concat(keys).dropna(subset=["key"]).astype({"key": "int64"}).drop_duplicates("key")
    # products only known from the primary file's own lookup columns (e.g. codes missing in the master)
    extra = []
    for d in primary_raw:
        cols = [_col(d, "Mat code"), _col(d, "Description"), _col(d, "Product cat"),
                _col(d, "Segment(MAT-1)"), _col(d, "Spec(MAT-2)")]
        if all(c is not None for c in cols):
            e = pd.DataFrame({"key": pd.to_numeric(cols[0], errors="coerce"), "m_product": None,
                              "m_category": cols[2], "m_segment": cols[3], "m_spec": cols[4]}).dropna(subset=["key"])
            extra.append(e.astype({"key": "int64"}).drop_duplicates("key"))
    if extra:
        e = pd.concat(extra).drop_duplicates("key")
        prod = pd.concat([prod, e[~e.key.isin(prod.key)]])
    pack = pd.DataFrame(columns=["pk", "pack_cat", "pack_desc"])
    if files["pack"]:
        pm, prob = read_table(files["pack"][-1], "pack")
        if prob: problems.append(prob + " (optional file)")
        pm = pm if pm is not None else pd.DataFrame({"SAP PACK CODE": []})
        pack = pd.DataFrame({"pk": pd.to_numeric(_col(pm, "SAP PACK CODE"), errors="coerce"),
                             "pack_cat": _col(pm, "Category -Pack (ltr/Kg", "Category-Pack"),
                             "pack_desc": _col(pm, "PACK DESCRIPTION")}).dropna(subset=["pk"])
        pack["pk"] = pack.pk.astype("int64"); pack = pack.drop_duplicates("pk")
    return prod, pack


def _enrich(df: pd.DataFrame, prod: pd.DataFrame, pack: pd.DataFrame, fallback_pack=None) -> pd.DataFrame:
    df = df.merge(prod, left_on="pcode", right_on="key", how="left").drop(columns="key")
    df = df.merge(pack, on="pk", how="left")
    if "pack_cat_file" in df:
        df["pack_cat"] = df["pack_cat"].fillna(df["pack_cat_file"]); df = df.drop(columns="pack_cat_file")
    df["category"] = df.m_category.fillna("Unmapped").str.title()
    df["segment"] = df.m_segment.fillna("Unmapped")
    df["spec"] = df.m_spec.fillna("Unmapped")
    df["product"] = df.m_product.fillna(df["product_raw"]).astype(str).str.strip()
    df["pack_cat"] = df.pack_cat.fillna("Unmapped")
    desc = df.pack_desc.where(df.pack_desc.notna(), df["product_raw"])
    unit = desc.map(_unit_litres)
    df["pack_group"] = [pack_group(u) for u in unit]
    fb = df.pack_cat.map(C.PACK_GROUPS).fillna(df.pack_cat)
    df["pack_group"] = df["pack_group"].fillna(fb)
    df["product_group"] = [product_group(s, sp, p) for s, sp, p in zip(df.segment, df.spec, df["product"])]
    df["hseg"] = np.where(df.segment.eq("GENUINE OILS"), "Co-branded (Genuine)",
                          np.where(df.segment.eq("INDUSTRIAL"), "Servo - Industrial", "Servo - Automotive"))
    return df.drop(columns=["m_product", "m_category", "m_segment", "m_spec"])


def _unit_litres(desc):
    """(size, unit) of one container from a pack description such as '20X1L', '4X6L', '7.5 L', '1X200KG'."""
    ms = re.findall(r"(?:(\d+)X)?([\d.]+)(L|KG|ML)", str(desc).upper().replace(" ", ""))
    if not ms: return None
    _, num, unit = ms[-1]
    try:
        v = float(num)
    except ValueError:
        return None
    return (v / 1000 if unit == "ML" else v, unit)


def pack_group(u):
    if u is None or u != u: return None
    v, unit = u
    if v < 0.25: return "Pouch <250 ml"
    if v <= 1.0: return "250 ml to 1 ltr"
    if v <= 6: return "> 1 ltr to 6 ltr"
    if v < 15: return "> 6 ltr and < 15 ltr"
    if v <= 25: return "> 15 ltr to 25 ltr"
    if v <= 26: return "26 ltr"
    if v <= 50: return "50 ltr"
    return "182, 175 kg, 200 kg" if unit == "KG" else "210, 208,205,200 ltr"   # drums / barrels


def product_group(seg: str, spec: str, name: str) -> str:
    seg, spec, name = str(seg), str(spec), str(name).upper()
    if seg == "MCO":
        return "2T Oil" if spec == "2T" else "4T Oil"
    if seg in ("PCMO", "DEO", "AGRI"):
        return {"PCMO": "PCMO", "DEO": "DEO", "AGRI": "Agri"}[seg]
    if seg == "GEAR": return "Gear Oil"
    if seg == "GREASE": return "Grease"
    if seg in ("HBF & Cool", "Transmission", "EV"): return "Speciality"
    if seg == "GENUINE OILS": return "Genuine Oil"
    if seg == "INDUSTRIAL":
        if "HYDRASHAKTI" in spec.upper() or "HYDRASHAKTI" in name: return "Hydrashakti 68"
        if "SYSTEM" in name: return "System 68"
        return "Other Industrial"
    return "Other"


# ---------------------------------------------------------------- normalise
def _to_datetime(col: pd.Series) -> pd.Series:
    """Dates as real datetimes, ISO text (2026-09-12), Excel serial numbers, or dd-mm-yyyy text.
    ISO is tried first so 2026-09-12 is never read as 9-Dec; day-first only applies to non-ISO text."""
    if pd.api.types.is_datetime64_any_dtype(col):
        return col
    num = pd.to_numeric(col, errors="coerce")
    if num.notna().mean() > 0.9 and num.dropna().between(20000, 80000).all():       # Excel serial dates
        return pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
    out = pd.to_datetime(col, errors="coerce", format="ISO8601")
    miss = out.isna() & col.notna()
    if miss.any():
        out[miss] = pd.to_datetime(col[miss].astype(str).str.strip(), errors="coerce", format="mixed", dayfirst=True)
    return out


def _dates(df: pd.DataFrame, col) -> pd.DataFrame:
    if col is None:
        raise ValueError("no date column found")
    d = _to_datetime(col)
    df["date"] = d.dt.normalize()
    df["mi"] = d.dt.year * 12 + d.dt.month - 1                      # absolute month index
    df["fy"] = d.dt.year + (d.dt.month >= C.FY_START_MONTH).astype(int)  # FY ending year (FY26-27 -> 2027)
    df["fm"] = (d.dt.month - C.FY_START_MONTH) % 12 + 1               # fiscal month 1..12 (Apr=1)
    return df


def _stockist(df: pd.DataFrame, code: pd.Series, name: pd.Series) -> pd.DataFrame:
    df["stockist_code"] = pd.to_numeric(code, errors="coerce").fillna(0).astype("int64")
    nm = _clean_name(name)
    df["stockist"] = [C.STOCKIST_ALIASES.get(c, n) for c, n in zip(df.stockist_code, nm)]
    return df


def norm_primary(raw: pd.DataFrame) -> pd.DataFrame:
    df = pd.DataFrame(index=raw.index)
    _dates(df, _col(raw, "Inv.Date", "Invoice Date", "Billing Date"))
    _stockist(df, _col(raw, "Payer", "Stockist Code"), _col(raw, "Name", "Stockist Name"))
    kl = _col(raw, "KL")
    if kl is None:
        kl = _num_col(_col(raw, "Billing qty in SKU")) / 1000
    df["kl"] = _num_col(kl).fillna(0.0)
    mat = pd.to_numeric(_col(raw, "Material"), errors="coerce").fillna(0).astype("int64")
    df["pcode"], df["pk"] = mat // 1000, mat % 1000
    df["product_raw"] = _col(raw, "Description")
    pc = _col(raw, "Pack Cat")
    df["pack_cat_file"] = pc if pc is not None else np.nan
    df["invoice"] = _col(raw, "Invoice No").astype(str) if _col(raw, "Invoice No") is not None else ""
    # keep the file's own lookup columns for products the master doesn't know
    for tgt, src in (("f_cat", "Product cat"), ("f_seg", "Segment(MAT-1)"), ("f_spec", "Spec(MAT-2)")):
        s = _col(raw, src); df[tgt] = s if s is not None else np.nan
    return df.dropna(subset=["date"])


def norm_secondary(raw: pd.DataFrame) -> pd.DataFrame:
    df = pd.DataFrame(index=raw.index)
    _dates(df, _col(raw, "Invoice Date"))
    _stockist(df, _col(raw, "Stockist Code"), _col(raw, "Stockist Name"))
    df["sales_area"] = _col(raw, "Sales Area")
    df["kl"] = _num_col(_col(raw, "Product Quantity(L/Kg)")).fillna(0.0) / 1000
    sku = pd.to_numeric(_col(raw, "SKU Code"), errors="coerce").fillna(0).astype("int64")
    df["pcode"], df["pk"] = sku // 1000, sku % 1000
    df["product_raw"] = _col(raw, "SKU Name")
    df["sku_name"] = df["product_raw"].astype(str)
    amt = _col(raw, "Taxable Amount (Rs.)")
    df["amount"] = _num_col(amt).fillna(0.0) if amt is not None else 0.0
    ssf = _col(raw, "SSF Code")
    df["ssf_code"] = ssf.astype(str).where(ssf.notna(), "-") if ssf is not None else "-"
    ssfn = _col(raw, "SSF Name")
    df["ssf_name"] = ssfn.astype(str).where(ssfn.notna(), "-") if ssfn is not None else "-"
    df["cust_subclass"] = (_col(raw, "Customer Sub Class").fillna("-").astype(str) if _col(raw, "Customer Sub Class") is not None else "-")
    for cc_, nm_ in (("Cat3 Agreement Sales Claimable Discount (Rs.)", "d_cat3"), ("HO Promotional Disc (Rs.)", "d_ho_promo"),
                     ("HO Sec Sales Discount (Rs.)", "d_ho_sec"), ("OEM Discount (Rs.)", "d_oem"), ("Manual Discount (Rs.)", "d_manual")):
        v_ = _col(raw, cc_)
        df[nm_] = (-_num_col(v_).fillna(0.0)) if v_ is not None else 0.0     # stored positive = discount given
    inv = _col(raw, "Invoice Number")
    df["invoice"] = inv.astype(str) if inv is not None else ""
    df["customer_id"] = _col(raw, "Customer Id").astype(str)
    df["customer"] = _col(raw, "Customer Name").astype(str).str.strip()
    grp = _col(raw, "Customer Group").astype(str).str.upper().str.strip()
    df["channel"] = grp.map(C.CHANNEL_MAP).fillna("Other")
    df["cust_group"] = grp.str.title()
    df["cust_subgroup"] = _col(raw, "Customer Sub Group").fillna("-").astype(str)
    df["district"] = (_col(raw, "Customer District") if _col(raw, "Customer District") is not None else _col(raw, "District")).fillna("-").astype(str).str.title()
    df["order_type"] = _col(raw, "Order Type").fillna("-").astype(str)
    for pat in C.EXCLUDE_SKU_PATTERNS:
        df = df[~df.sku_name.str.upper().str.contains(pat.upper(), regex=False)]
    return df.dropna(subset=["date"])


# ---------------------------------------------------------------- merge monthly drops
def _merge_sources(frames: list[tuple[str, pd.DataFrame]]) -> pd.DataFrame:
    """frames: [(file_name, df)] each with 'mi' and 'date'. Latest/fullest file per month wins."""
    if len(frames) == 1:
        f = frames[0][1].copy(); f["source"] = frames[0][0]; return f
    stats = []
    for name, d in frames:
        g = d.groupby("mi").agg(last=("date", "max"), rows=("date", "size")).reset_index()
        g["file"] = name; stats.append(g)
    st = pd.concat(stats)
    order = {n: i for i, (n, _) in enumerate(frames)}           # later in the (mtime-sorted) list = newer
    st["ord"] = st.file.map(order)
    win = st.sort_values(["mi", "last", "rows", "ord"]).groupby("mi").tail(1).set_index("mi").file
    parts = []
    for name, d in frames:
        keep = d[d.mi.map(win) == name].copy(); keep["source"] = name; parts.append(keep)
    return pd.concat(parts, ignore_index=True)


@dataclass
class Data:
    primary: pd.DataFrame
    secondary: pd.DataFrame
    targets: pd.DataFrame            # columns: fy, fm, stockist_code, target
    info: dict = field(default_factory=dict)


def load_all() -> Data:
    files = discover()
    files_sorted = {k: sorted(v, key=lambda p: p.stat().st_mtime) for k, v in files.items()}
    problems: list[str] = []
    raw_p, raw_s = [], []
    for kind, lst, fs in (("primary", raw_p, files_sorted["primary"]), ("secondary", raw_s, files_sorted["secondary"])):
        for f in fs:
            df, prob = read_table(f, kind)
            if prob: problems.append(prob)
            else: lst.append((f.name, df))
    prod, pack = load_masters(files_sorted, [d for _, d in raw_p], problems)
    out = {}
    for key, raws, fn in (("primary", raw_p, norm_primary), ("secondary", raw_s, norm_secondary)):
        if not raws:
            out[key] = pd.DataFrame(); continue
        normed = []
        for n, d in raws:
            try:
                nd = fn(d)
                if nd.empty: problems.append(f"{n}: no rows with a valid date - file skipped."); continue
                normed.append((n, nd))
            except Exception as e:
                problems.append(f"{n}: could not be processed ({type(e).__name__}: {e}) - file skipped.")
        if not normed:
            out[key] = pd.DataFrame(); continue
        merged = _merge_sources(normed)
        merged = _enrich(merged, prod, pack)
        if key == "primary":   # fall back to file's own lookup cols for unmapped materials
            um = merged.segment.eq("Unmapped")
            if um.any():
                merged.loc[um, "category"] = merged.loc[um, "f_cat"].fillna("Unmapped").astype(str).str.title()
                merged.loc[um, "segment"] = merged.loc[um, "f_seg"].fillna("Unmapped")
                merged.loc[um, "spec"] = merged.loc[um, "f_spec"].fillna("Unmapped")
                merged["product_group"] = [product_group(s, sp, p) for s, sp, p in zip(merged.segment, merged.spec, merged["product"])]
                merged["hseg"] = np.where(merged.segment.eq("GENUINE OILS"), "Co-branded (Genuine)",
                                          np.where(merged.segment.eq("INDUSTRIAL"), "Servo - Industrial", "Servo - Automotive"))
            merged = merged.drop(columns=["f_cat", "f_seg", "f_spec"])
        # category cleanup: HO deck names
        merged["category"] = merged["category"].replace({"Unmapped": "Unmapped"})
        out[key] = merged.reset_index(drop=True)
    if "stockist" in out["primary"] and "sales_area" in out["secondary"]:
        amap = out["secondary"].dropna(subset=["sales_area"]).drop_duplicates("stockist_code").set_index("stockist_code").sales_area
        out["primary"]["sales_area"] = out["primary"].stockist_code.map(amap).fillna("Unknown")
    info = {
        "problems": problems,
        "files": {k: [f.name for f in v] for k, v in files_sorted.items()},
        "primary_last": out["primary"].date.max() if len(out["primary"]) else None,
        "secondary_last": out["secondary"].date.max() if len(out["secondary"]) else None,
        "data_dir": str(C.DATA_DIR),
        "unmapped_primary": float(out["primary"].loc[out["primary"].segment.eq("Unmapped"), "kl"].sum()) if len(out["primary"]) else 0,
        "unmapped_secondary": float(out["secondary"].loc[out["secondary"].segment.eq("Unmapped"), "kl"].sum()) if len(out["secondary"]) else 0,
    }
    names = {}
    for df in (out["primary"], out["secondary"]):
        if len(df): names.update(dict(zip(df.stockist, df.stockist_code)))
    tdf, twarn = targets.parse(targets.find_target_file(), names)
    info["target_file"] = targets.find_target_file().name if targets.find_target_file() else None
    info["target_warnings"] = twarn
    return Data(out["primary"], out["secondary"], tdf, info)
