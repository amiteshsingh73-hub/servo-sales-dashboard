"""Free-form pivot over the primary / secondary data (no dependency on the pre-built HO tables)."""
from __future__ import annotations
import numpy as np
import pandas as pd
from .periods import MONTHS, fy_label, mi_name

TIME_DIMS = ["Fiscal year", "Quarter", "Month", "Fiscal month"]
COMMON = {"Stockist": "stockist", "Sales area": "sales_area", "Category": "category", "Segment (MAT-1)": "segment",
          "Spec (MAT-2)": "spec", "Product group (HO)": "product_group", "HO segment": "hseg", "Product": "product",
          "Pack group (HO)": "pack_group", "Pack category": "pack_cat"}
SEC_ONLY = {"Channel": "channel", "Customer group": "cust_group", "Customer sub-group": "cust_subgroup",
            "District": "district", "Customer": "customer", "Order type": "order_type"}
MEASURES = {   # label -> (key, additive, secondary_only, number format)
    "Volume (KL)": ("kl", True, False, "kl"),
    "Invoice lines": ("lines", True, False, "int"),
    "Invoices": ("invoices", False, False, "int"),
    "Active counters": ("counters", False, True, "int"),
    "Taxable value (₹)": ("amount", True, True, "int"),
    "Realisation (₹/L)": ("real", False, True, "dec"),
}
SHOW_AS = ["Value", "% of grand total", "% of row total", "% of column total", "% change vs previous column"]


def dims_for(kind: str) -> dict[str, str]:
    d = {t: t for t in TIME_DIMS}
    d.update(COMMON)
    if kind == "Secondary": d.update(SEC_ONLY)
    return d


def measures_for(kind: str) -> list[str]:
    return [m for m, (_, _, sec, _) in MEASURES.items() if kind == "Secondary" or not sec]


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    """Adds the time dimensions (labels) and fills blanks so every dimension can be grouped."""
    d = df.copy()
    d["Fiscal year"] = d.fy.map(fy_label)
    d["Quarter"] = "Q" + ((d.fm - 1) // 3 + 1).astype(str)
    d["Month"] = d.mi.map(mi_name)
    d["Fiscal month"] = d.fm.map(lambda m: MONTHS[int(m) - 1])
    return d


_ORDER = {"Fiscal year": lambda d: d.drop_duplicates("Fiscal year").sort_values("fy")["Fiscal year"].tolist(),
          "Quarter": lambda d: ["Q1", "Q2", "Q3", "Q4"],
          "Month": lambda d: d.drop_duplicates("Month").sort_values("mi")["Month"].tolist(),
          "Fiscal month": lambda d: MONTHS}


def _labels(d: pd.DataFrame, dim_label: str, col: str, present) -> list:
    present = set(present)
    if dim_label in _ORDER:
        return [x for x in _ORDER[dim_label](d) if x in present]
    return sorted(present, key=lambda x: str(x))


def _agg(d: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
    spec = {"kl": ("kl", "sum"), "lines": ("kl", "size")}
    if "invoice" in d: spec["invoices"] = ("invoice", "nunique")
    if "customer_id" in d: spec["counters"] = ("customer_id", "nunique")
    if "amount" in d: spec["amount"] = ("amount", "sum")
    g = d.groupby(keys, observed=True, dropna=False) if keys else d.groupby(np.zeros(len(d)))
    out = g.agg(**spec)
    if "amount" in out: out["real"] = (out.amount / (out.kl * 1000)).replace([np.inf, -np.inf], np.nan)
    return out


def build(df: pd.DataFrame, rows: list[str], cols: list[str], measures: list[str], dims: dict[str, str],
          show_as="Value", totals=True, top_n=0, sort_total=True):
    """Returns (table, info). `df` must come from prepare()."""
    if df.empty or not measures: return pd.DataFrame(), {}
    d = df.copy()
    rc = [dims[r] if dims[r] in d else r for r in rows]            # actual column names
    cc = [dims[c] if dims[c] in d else c for c in cols]
    for c in rc + cc: d[c] = d[c].astype(object).where(d[c].notna(), "-").astype(str)
    keys = rc + cc
    body_all = _agg(d, keys) if keys else _agg(d, [])
    row_tot = _agg(d, rc) if rc else _agg(d, [])
    col_tot = _agg(d, cc) if cc else None
    grand = _agg(d, []).iloc[0]
    n_rows_total = len(row_tot) if rc else 1

    # row order / top-N
    if rc:
        if sort_total or len(rc) > 1:
            order = row_tot[MEASURES[measures[0]][0]].sort_values(ascending=False).index
        else:
            lab = _labels(d, rows[0], rc[0], row_tot.index)
            order = pd.Index(lab, name=rc[0])
        if top_n and len(order) > top_n: order = order[:top_n]
    else:
        order = None
    # column order
    col_order = None
    if cc:
        if len(cc) == 1:
            col_order = _labels(d, cols[0], cc[0], col_tot.index)
        else:
            col_order = list(col_tot.index.sort_values())

    blocks, keys_ = [], []
    for m in measures:
        mk, additive, _, _ = MEASURES[m]
        if mk not in body_all: continue
        if cc:
            w = body_all[mk].unstack(cc) if rc else body_all[mk].to_frame().T
            if not rc: w.index = ["All"]
            w = w.reindex(columns=col_order)
        else:
            w = body_all[mk].to_frame("Value") if rc else pd.DataFrame({"Value": [body_all[mk].iloc[0]]}, index=["All"])
        if order is not None: w = w.reindex(order)
        rt = (row_tot[mk].reindex(w.index) if rc else pd.Series([grand[mk]], index=w.index))
        if totals and cc:
            w["Total"] = rt.values
        gt = None
        if totals and rc:
            ct = (col_tot[mk].reindex(col_order) if cc else pd.Series([grand[mk]], index=["Value"]))
            gt = list(ct.values) + ([grand[mk]] if cc else [])
        # show-as
        if show_as != "Value" and additive:
            num = w.copy()
            data_cols = [c for c in w.columns if c != "Total"]
            if show_as == "% of grand total":
                num = w / grand[mk] * 100
            elif show_as == "% of row total":
                rtv = row_tot[mk].reindex(w.index) if rc else pd.Series([grand[mk]], index=w.index)
                num = w.div(rtv, axis=0) * 100
            elif show_as == "% of column total":
                base = pd.Series({c: (col_tot[mk].get(c, np.nan) if cc else grand[mk]) for c in data_cols})
                num = w[data_cols].div(base, axis=1) * 100
                if totals and cc: num["Total"] = w["Total"] / grand[mk] * 100
            elif show_as == "% change vs previous column":
                num = w[data_cols].pct_change(axis=1) * 100
                if totals and cc: num["Total"] = np.nan
            w = num
            if totals and rc and gt is not None:
                gt = None
        if totals and rc and gt is not None:
            w.loc["Total"] = gt
        blocks.append(w); keys_.append(m)
    if not blocks: return pd.DataFrame(), {}
    if len(blocks) == 1:
        t = blocks[0]
    else:
        t = pd.concat(blocks, axis=1, keys=keys_)
    t.index.name = " / ".join(rows) if rows else None
    if len(rows) > 1 and isinstance(t.index[0], tuple) if len(t) else False:
        t.index = pd.MultiIndex.from_tuples([x if isinstance(x, tuple) else (x,) + ("",) * (len(rows) - 1) for x in t.index], names=rows)
    pct_mode = show_as != "Value"
    return t, {"rows_total": n_rows_total, "rows_shown": len(order) if order is not None else 1, "pct": pct_mode,
               "measures": measures, "grand": {m: grand.get(MEASURES[m][0]) for m in measures}}


def fmt_value(v, kind="kl", pct=False):
    if v is None or (isinstance(v, float) and np.isnan(v)): return ""
    if pct: return f"{v:.1f}%"
    if kind == "kl": return f"{v:,.1f}" if abs(v) < 100 else f"{v:,.0f}"
    if kind == "dec": return f"{v:,.1f}"
    return f"{v:,.0f}"


def measure_kind(col, measures) -> str:
    m = col[0] if isinstance(col, tuple) else (measures[0] if measures else "Volume (KL)")
    return MEASURES.get(m, ("", 0, 0, "kl"))[3]


def to_text(t: pd.DataFrame, info: dict) -> pd.DataFrame:
    """Formatted strings (for display / PPT)."""
    out = t.copy().astype(object)
    for c in t.columns:
        k = measure_kind(c, info.get("measures", []))
        out[c] = [fmt_value(v, k, info.get("pct", False)) for v in t[c]]
    return out


def heat_style(t: pd.DataFrame, text: pd.DataFrame, rgb=(11, 31, 102)):
    """Styler with per-column shading (no matplotlib needed). Total row/column left unshaded."""
    def colors(_):
        css = pd.DataFrame("", index=t.index, columns=t.columns)
        body = t.drop(index=[i for i in t.index if (i == "Total" or (isinstance(i, tuple) and i[0] == "Total"))], errors="ignore")
        for c in t.columns:
            cname = c[-1] if isinstance(c, tuple) else c
            if cname == "Total": continue
            v = body[c].astype(float)
            lo, hi = v.min(), v.max()
            if not np.isfinite(lo) or hi == lo: continue
            a = ((v - lo) / (hi - lo)).fillna(0) * .8
            def one(x):
                r, g, b = (int(round(255 * (1 - x) + k * x)) for k in rgb)      # opaque blend with white
                return f"background-color: rgb({r},{g},{b}); color: {'#ffffff' if x > .45 else '#111111'}"
            css.loc[v.index, c] = [one(x) for x in a]
        return css
    return text.style.apply(colors, axis=None)
