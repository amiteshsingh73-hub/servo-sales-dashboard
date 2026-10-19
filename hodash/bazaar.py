"""Bazaar (retail counter) analytics: counter health, tiers, dormancy, sales-force scorecard, cross-sell gaps.
All functions take the secondary frame already narrowed to Bazaar + sidebar filters and return plain DataFrames,
so they can be tested without Streamlit.
"""
from __future__ import annotations
import numpy as np
import pandas as pd
from .metrics import sel, growth, pct
from .periods import Window

# Monthly litres per counter that define the tiers (edit to suit the market).
TIERS = [("Platinum", 500), ("Gold", 200), ("Silver", 75), ("Bronze", 0)]
TIER_ORDER = [t for t, _ in TIERS]
STATUS_ORDER = ["New", "Reactivated", "Growing", "Stable", "Declining", "Lost"]
FOCUS_GROUPS = ["4T Oil", "PCMO", "DEO", "Gear Oil", "Grease", "2T Oil", "Agri", "Genuine Oil"]


def months_in(w: Window) -> float:
    """Length of a window in months (fractional when the data stops mid-month)."""
    return max(((w.end - w.start).days + 1) / 30.4375, 1 / 30.4375)


def _by_counter(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["kl", "amount", "invoices", "skus", "groups", "last"])
    return df.groupby("customer_id").agg(kl=("kl", "sum"), amount=("amount", "sum"), invoices=("invoice", "nunique"),
                                         skus=("product", "nunique"), groups=("product_group", "nunique"), last=("date", "max"))


def counter_master(bz: pd.DataFrame) -> pd.DataFrame:
    """Latest known name / stockist / SSF / district / sub-group of every counter."""
    if bz.empty:
        return pd.DataFrame(columns=["Customer", "Stockist", "SSF", "District", "Sub-group"])
    d = bz.sort_values("date").groupby("customer_id").last()
    return pd.DataFrame({"Customer": d.customer, "Stockist": d.stockist, "SSF": d.ssf_name, "District": d.district,
                         "Sub-group": d.cust_subgroup})


def kpis(bz: pd.DataFrame, ws: dict) -> dict:
    out = {}
    for k in ("cy", "ly"):
        d = sel(bz, ws[k]); d_pos = d[d.kl > 0]
        n = d_pos.customer_id.nunique(); m = months_in(ws[k])
        kl_ = float(d.kl.sum())
        prem = float(d[d.category.isin(["Premium"])].kl.sum())
        bc = _by_counter(d_pos)
        out[k] = dict(kl=kl_, counters=n, lpcm=kl_ * 1000 / n / m if n else np.nan,
                      rpl=float(d.amount.sum()) / (kl_ * 1000) if kl_ > 0 else np.nan,
                      prem_share=pct(prem, kl_), skus=float(bc.skus.mean()) if len(bc) else np.nan,
                      freq=float(bc.invoices.mean()) / m if len(bc) else np.nan,
                      skus_total=d_pos["product"].nunique())
    return out


def monthly(bz: pd.DataFrame, fy: int) -> pd.DataFrame:
    d = bz[bz.fy == fy]
    g = d.groupby("fm").agg(kl=("kl", "sum"))
    g["counters"] = d[d.kl > 0].groupby("fm").customer_id.nunique()
    return g.reindex(range(1, 13))


def lifecycle(bz: pd.DataFrame, ws: dict, band=10.0) -> pd.DataFrame:
    """One row per counter with CY / LY volume and a status:
    New (first ever purchase in CY), Reactivated (bought before, not in LY), Growing / Stable / Declining (±band %), Lost."""
    cy, ly = sel(bz, ws["cy"]).groupby("customer_id").kl.sum(), sel(bz, ws["ly"]).groupby("customer_id").kl.sum()
    t = pd.DataFrame({"cy": cy, "ly": ly}).fillna(0.0)
    t = t[(t.cy > 0) | (t.ly > 0)]
    if t.empty:
        return t.assign(status=pd.Series(dtype=str), delta=pd.Series(dtype=float))
    first = bz[bz.kl > 0].groupby("customer_id").date.min()
    g = (t.cy / t.ly.where(t.ly > 0) - 1) * 100
    st = np.select([(t.ly <= 0) & (first.reindex(t.index) >= ws["cy"].start), (t.ly <= 0), t.cy <= 0, g > band, g < -band],
                   ["New", "Reactivated", "Lost", "Growing", "Declining"], "Stable")
    t["status"] = st
    t["delta"] = t.cy - t.ly
    return t


def lifecycle_summary(lc: pd.DataFrame) -> pd.DataFrame:
    if lc.empty: return pd.DataFrame(columns=["Counters", "CY (KL)", "LY (KL)", "Change (KL)"])
    s = lc.groupby("status").agg(Counters=("cy", "size"), cy=("cy", "sum"), ly=("ly", "sum"), delta=("delta", "sum"))
    s = s.reindex([x for x in STATUS_ORDER if x in s.index])
    return s.rename(columns={"cy": "CY (KL)", "ly": "LY (KL)", "delta": "Change (KL)"})


def tiers(bz: pd.DataFrame, w: Window) -> pd.DataFrame:
    """Counters bucketed by average litres / month in the window."""
    d = sel(bz, w)
    c = d.groupby("customer_id").kl.sum()
    c = c[c > 0]
    if c.empty: return pd.DataFrame(columns=["tier", "lpm", "kl"])
    lpm = c * 1000 / months_in(w)
    tier = pd.Series("Bronze", index=c.index)
    for name, lo in reversed(TIERS[:-1]):
        tier[lpm >= lo] = name
    return pd.DataFrame({"tier": tier, "lpm": lpm, "kl": c})


def tier_summary(tr: pd.DataFrame) -> pd.DataFrame:
    if tr.empty: return pd.DataFrame()
    s = tr.groupby("tier").agg(Counters=("kl", "size"), KL=("kl", "sum"), avg=("lpm", "mean")).reindex(TIER_ORDER).dropna(how="all")
    s["% Counters"] = s.Counters / s.Counters.sum() * 100
    s["% Volume"] = s.KL / s.KL.sum() * 100
    s = s.rename(columns={"avg": "Avg L / month"})
    rng = {n: (f"≥ {lo:,} L/mo" if i == 0 else f"{lo:,}–{TIERS[i - 1][1]:,} L/mo" if lo else f"< {TIERS[i - 1][1]:,} L/mo") for i, (n, lo) in enumerate(TIERS)}
    s.insert(0, "Band", [rng[i] for i in s.index])
    return s


def dormant(bz: pd.DataFrame, as_of: pd.Timestamp, min_days=30, lookback_months=6) -> pd.DataFrame:
    """Counters that bought in the last `lookback_months` (before the gap) but have not billed for `min_days`.
    Sorted by the volume at risk (their average monthly litres)."""
    if bz.empty or pd.isna(as_of): return pd.DataFrame()
    start = as_of - pd.DateOffset(months=lookback_months)
    d = bz[(bz.date > start) & (bz.date <= as_of) & (bz.kl > 0)]
    last_any = bz[bz.kl > 0].groupby("customer_id").date.max()
    g = d.groupby("customer_id").agg(kl=("kl", "sum"), bills=("invoice", "nunique"))
    g["last"] = last_any.reindex(g.index)
    g["days"] = (as_of - g["last"]).dt.days
    g = g[g.days >= min_days]
    g["Avg L / month"] = g.kl * 1000 / lookback_months
    master = counter_master(bz)
    out = master.reindex(g.index).join(g)
    out = out.rename(columns={"last": "Last bill", "days": "Days since bill", "bills": f"Bills (last {lookback_months}m)",
                              "kl": f"KL (last {lookback_months}m)"})
    out = out.sort_values("Avg L / month", ascending=False)
    out["Last bill"] = pd.to_datetime(out["Last bill"]).dt.strftime("%d-%b-%y")
    return out


def scorecard(bz: pd.DataFrame, ws: dict, dim: str, lc: pd.DataFrame | None = None) -> pd.DataFrame:
    """Per member of `dim` (SSF, district, sub-group, stockist): volume, growth, counters, productivity, realisation."""
    cy, ly = sel(bz, ws["cy"]), sel(bz, ws["ly"])
    m = months_in(ws["cy"])
    if cy.empty and ly.empty: return pd.DataFrame()
    a = cy.groupby(dim).agg(cy=("kl", "sum"), amt=("amount", "sum"))
    a["counters"] = cy[cy.kl > 0].groupby(dim).customer_id.nunique()
    b = ly.groupby(dim).agg(ly=("kl", "sum"))
    b["counters_ly"] = ly[ly.kl > 0].groupby(dim).customer_id.nunique()
    t = a.join(b, how="outer").fillna(0.0)
    t = t[(t.cy.abs() > 1e-9) | (t.ly.abs() > 1e-9)]
    if t.empty: return pd.DataFrame()
    t["% Growth"] = [growth(x, y) for x, y in zip(t.cy, t.ly)]
    t["% Share"] = t.cy / t.cy.sum() * 100 if t.cy.sum() else 0.0
    t["L / counter / month"] = (t.cy * 1000 / t.counters.replace(0, np.nan) / m)
    t["₹ / L"] = t.amt / (t.cy * 1000).replace(0, np.nan)
    if lc is not None and len(lc):
        home = counter_master(bz)[{"ssf_name": "SSF", "district": "District", "cust_subgroup": "Sub-group", "stockist": "Stockist"}.get(dim, "SSF")]
        l2 = lc.join(home.rename("k"))
        t["New"] = l2[l2.status == "New"].groupby("k").size().reindex(t.index).fillna(0)
        t["Lost"] = l2[l2.status == "Lost"].groupby("k").size().reindex(t.index).fillna(0)
    t = t.sort_values("cy", ascending=False)
    t = t.rename(columns={"cy": "CY (KL)", "ly": "LY (KL)", "counters": "Active counters", "counters_ly": "Counters LY"}).drop(columns="amt")
    t.index.name = None
    return t


def penetration(bz: pd.DataFrame, w: Window, groups=None) -> pd.DataFrame:
    """% of active counters that bought each product group in the window."""
    d = sel(bz, w); d = d[d.kl > 0]
    n = d.customer_id.nunique()
    if not n: return pd.DataFrame()
    g = d.groupby("product_group").agg(buyers=("customer_id", "nunique"), kl=("kl", "sum"))
    g["% of counters"] = g.buyers / n * 100
    g["Non-buyers"] = n - g.buyers
    if groups: g = g.reindex([x for x in groups if x in g.index] + [x for x in g.index if x not in groups])
    return g.sort_values("% of counters", ascending=False)


def cross_sell(bz: pd.DataFrame, w: Window, group: str, top=30) -> pd.DataFrame:
    """Biggest active counters in the window that did not buy `group` at all - the easiest place to add a line."""
    d = sel(bz, w); d = d[d.kl > 0]
    if d.empty: return pd.DataFrame()
    buyers = set(d[d.product_group == group].customer_id)
    tot = d.groupby("customer_id").agg(kl=("kl", "sum"), groups=("product_group", "nunique"))
    tot = tot[~tot.index.isin(buyers)].sort_values("kl", ascending=False).head(top)
    tot["L / month"] = tot.kl * 1000 / months_in(w)
    out = counter_master(bz).reindex(tot.index).join(tot)
    return out.rename(columns={"kl": "Total KL (window)", "groups": "Product groups bought"})


def counter_history(bz: pd.DataFrame, cid: str) -> pd.DataFrame:
    d = bz[bz.customer_id == cid]
    return d.groupby("mi").agg(kl=("kl", "sum"), amount=("amount", "sum"), invoices=("invoice", "nunique")).sort_index()
