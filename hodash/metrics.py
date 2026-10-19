from __future__ import annotations
import numpy as np
import pandas as pd
from .periods import Period, Window, windows, fy_label, MONTHS, mi_of, month_start


def sel(df: pd.DataFrame, w: Window) -> pd.DataFrame:
    return df[(df.date >= w.start) & (df.date <= w.end)]


def kl(df, w) -> float:
    return float(sel(df, w).kl.sum()) if len(df) else 0.0


def growth(cy, ly):
    return (cy / ly - 1) * 100 if ly and abs(ly) > 1e-9 else np.nan


def pct(a, b):
    return a / b * 100 if b and abs(b) > 1e-9 else np.nan


def fmt_kl(x, d=0):
    return "-" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:,.{d}f}"


def fmt_pct(x, d=0, sign=False):
    if x is None or (isinstance(x, float) and np.isnan(x)): return "-"
    return f"{x:+,.{d}f}%" if sign else f"{x:,.{d}f}%"


# ------------------------------------------------------------------ targets
def _tsel(targets, fy, fms, kind, channel, codes=None):
    if targets is None or targets.empty: return None
    t = targets[(targets.fy == fy) & targets.fm.isin(list(fms)) & (targets.kind == kind)]
    if channel is not None: t = t[t.channel == channel]
    if codes is not None: t = t[t.stockist_code.isin(list(codes))]
    return t


def target_for(targets, codes, fy: int, fms, kind="Primary", channel=None) -> float:
    t = _tsel(targets, fy, fms, kind, channel, codes)
    return float(t.target.sum()) if t is not None and len(t) else np.nan


def target_by_stockist(targets, fy, fms, kind="Primary", channel=None) -> pd.Series:
    t = _tsel(targets, fy, fms, kind, channel)
    return t.groupby("stockist_code").target.sum() if t is not None and len(t) else pd.Series(dtype=float)


def target_by_channel(targets, fy, fms, codes=None) -> pd.Series:
    t = _tsel(targets, fy, fms, "Secondary", None, codes)
    return t.groupby("channel").target.sum() if t is not None and len(t) else pd.Series(dtype=float)


def monthly_target(targets, fy, kind="Primary", channel=None, codes=None) -> pd.Series:
    t = _tsel(targets, fy, range(1, 13), kind, channel, codes)
    return t.groupby("fm").target.sum() if t is not None and len(t) else pd.Series(dtype=float)


# ------------------------------------------------------------------ monthly trend
def monthly(df: pd.DataFrame, fy: int) -> pd.Series:
    d = df[df.fy == fy]
    return d.groupby("fm").kl.sum().reindex(range(1, 13))


# ------------------------------------------------------------------ dimension compare
def compare(df: pd.DataFrame, dim: str | list, wins: dict[str, Window], order=None) -> pd.DataFrame:
    """CY / LY / Prev KL by dimension, with growth and share."""
    cols = {}
    for k, w in wins.items():
        cols[k] = sel(df, w).groupby(dim).kl.sum()
    t = pd.DataFrame(cols).fillna(0.0)
    t["growth"] = [growth(a, b) for a, b in zip(t.cy, t.ly)]
    t["growth_prev"] = [growth(a, b) for a, b in zip(t.cy, t.prev)]
    t["share"] = t.cy / t.cy.sum() * 100 if t.cy.sum() else 0
    t["delta"] = t.cy - t.ly
    if order:
        idx = [o for o in order if o in t.index] + [i for i in t.index if i not in order]
        t = t.loc[idx]
    else:
        t = t.sort_values("cy", ascending=False)
    return t


# ------------------------------------------------------------------ customer KPIs (secondary)
def customer_kpis(sec: pd.DataFrame, full_sec: pd.DataFrame, wins: dict[str, Window], p: Period) -> dict:
    """Active / new / not-transacted counters + churn volume, per the report-card in the recording."""
    cy, ly = sel(sec, wins["cy"]), sel(sec, wins["ly"])
    cy, ly = cy[cy.kl > 0], ly[ly.kl > 0]
    a_cy, a_ly = set(cy.customer_id), set(ly.customer_id)
    hist_start = full_sec.date.min() if len(full_sec) else None
    first = full_sec[full_sec.kl > 0].groupby("customer_id").date.min()
    # 'new' is only meaningful once there is a full year of history before the period starts
    new_ok = hist_start is not None and wins["cy"].start >= hist_start + pd.DateOffset(years=1)
    new_ids = set(first[(first >= wins["cy"].start) & (first <= wins["cy"].end)].index) & a_cy if new_ok else set()
    lost = a_ly - a_cy
    months = max(p.n_months, 1)

    def per_counter(c, ch):
        d = c[c.channel == ch]
        n = d.customer_id.nunique()
        return d.kl.sum() * 1000 / n / months if n else np.nan

    return {
        "active": len(a_cy), "active_ly": len(a_ly),
        "new": len(new_ids) if new_ok else None,
        "new_kl": float(cy[cy.customer_id.isin(new_ids)].kl.sum()) if new_ok else None,
        "lost": len(lost), "lost_kl": float(ly[ly.customer_id.isin(lost)].kl.sum()),
        "pcs": per_counter(cy, "Bazaar"), "pcs_ly": per_counter(ly, "Bazaar"),
        "ro_ppt": per_counter(cy, "RO"), "ro_ppt_ly": per_counter(ly, "RO"),
        "bazaar_counters": cy[cy.channel == "Bazaar"].customer_id.nunique(),
        "ro_outlets": cy[cy.channel == "RO"].customer_id.nunique(),
    }


# ------------------------------------------------------------------ HO-template table
def ho_table(df: pd.DataFrame, dim: str, p: Period, last_date, targets: pd.DataFrame | None = None,
             order=None, apportion=True, dim_is_stockist=False, extra_total=True, rename_dim=None,
             tkind="Primary", tchannel=None) -> pd.DataFrame:
    """Mirrors the HO standard-template tables (slides 7-19):
    FY history | CAGR | Avg Vol/Mth | Target | Pro-rata target | Ach CP | Ach LY | % Tgt ach | % Growth | Balance."""
    w = windows(p, last_date)
    fy = p.fy
    hist_years = [y for y in (fy - 3, fy - 2, fy - 1) if len(df) and (df.fy == y).any()]
    keys = list(df[dim].dropna().unique()) if len(df) else []
    if order: keys = [o for o in order if o in keys] + [k for k in keys if k not in order]
    else:
        base = df[df.fy == fy - 1].groupby(dim).kl.sum()
        cur = df[df.fy == fy].groupby(dim).kl.sum()
        keys = sorted(keys, key=lambda k: -(base.get(k, 0) + cur.get(k, 0)))
    cp_lbl, ly_lbl = f"Ach {w['cy'].label}", f"Ach {w['ly'].label}"
    rows = []
    scope = sorted(df.stockist_code.unique()) if len(df) else []
    tgt_total_annual = target_for(targets, scope, fy, range(1, 13), tkind, tchannel)
    tgt_total_pro = target_for(targets, scope, fy, p.fms, tkind, tchannel)
    lyshare = df[df.fy == fy - 1].groupby(dim).kl.sum()
    lyshare = lyshare / lyshare.sum() if lyshare.sum() else lyshare
    ts_year = target_by_stockist(targets, fy, range(1, 13), tkind, tchannel) if dim_is_stockist else None
    ts_pro = target_by_stockist(targets, fy, p.fms, tkind, tchannel) if dim_is_stockist else None
    by_ch = dim == "channel" and tkind == "Secondary"
    tc_year = target_by_channel(targets, fy, range(1, 13), scope) if by_ch else None
    tc_pro = target_by_channel(targets, fy, p.fms, scope) if by_ch else None
    code_of = df.drop_duplicates(dim).set_index(dim).stockist_code if dim_is_stockist and len(df) else None
    for k in keys:
        d = df[df[dim] == k]
        r = {"": k}
        for y in hist_years: r[fy_label(y, True)] = float(d[d.fy == y].kl.sum())
        yrs = [r[fy_label(y, True)] for y in hist_years]
        r["CAGR %"] = ((yrs[-1] / yrs[0]) ** (1 / (len(yrs) - 1)) - 1) * 100 if len(yrs) > 1 and yrs[0] > 0 and yrs[-1] > 0 else np.nan
        r["Avg Vol/Mth"] = (yrs[-1] / 12) if yrs else np.nan
        if dim_is_stockist and len(ts_year):
            code = code_of[k]; ta, tp = ts_year.get(code, np.nan), ts_pro.get(code, np.nan)
        elif by_ch and len(tc_year):
            ta, tp = tc_year.get(k, np.nan), tc_pro.get(k, np.nan)
        elif apportion and not np.isnan(tgt_total_annual):
            sh = lyshare.get(k, 0); ta, tp = tgt_total_annual * sh, tgt_total_pro * sh
        else:
            ta = tp = np.nan
        r[f"Target {fy_label(fy, True)}"] = ta
        r["Pro-rata target"] = tp
        cyv, lyv = kl(d, w["cy"]), kl(d, w["ly"])
        r[cp_lbl], r[ly_lbl] = cyv, lyv
        r["% Tgt ach"] = pct(cyv, tp)
        r["% Growth"] = growth(cyv, lyv)
        r["Balance to target"] = (ta - cyv) if not np.isnan(ta) else np.nan
        rows.append(r)
    t = pd.DataFrame(rows)
    if extra_total and len(t):
        tot = {"": "Total"}
        for c in t.columns[1:]:
            tot[c] = t[c].sum(min_count=1)
        yrs = [tot[fy_label(y, True)] for y in hist_years]
        tot["CAGR %"] = ((yrs[-1] / yrs[0]) ** (1 / (len(yrs) - 1)) - 1) * 100 if len(yrs) > 1 and yrs[0] > 0 else np.nan
        tp, ta = tot["Pro-rata target"], tot[f"Target {fy_label(fy, True)}"]
        tot["% Tgt ach"] = pct(tot[cp_lbl], tp); tot["% Growth"] = growth(tot[cp_lbl], tot[ly_lbl])
        t = pd.concat([t, pd.DataFrame([tot])], ignore_index=True)
    if rename_dim: t = t.rename(columns={"": rename_dim})
    else: t = t.rename(columns={"": dim})
    return t


def drop_empty_cols(t: pd.DataFrame) -> pd.DataFrame:
    return t.dropna(axis=1, how="all")


# ------------------------------------------------------------------ analytics helpers
def bridge(df, dim, wins) -> pd.DataFrame:
    """LY -> CY volume bridge: contribution of each member of `dim` to the change."""
    t = compare(df, dim, wins)
    return t[["ly", "cy", "delta"]].sort_values("delta", ascending=False)


def counter_moves(sec, wins) -> pd.DataFrame:
    """Per counter CY vs LY (named), for gainers / decliners / new / lost lists."""
    a = sel(sec, wins["cy"]).groupby("customer_id").agg(Customer=("customer", "last"), Channel=("channel", "last"),
                                                       Stockist=("stockist", "last"), cy=("kl", "sum"))
    b = sel(sec, wins["ly"]).groupby("customer_id").agg(Customer2=("customer", "last"), Channel2=("channel", "last"),
                                                       Stockist2=("stockist", "last"), ly=("kl", "sum"))
    t = a.join(b, how="outer")
    t["Customer"] = t.Customer.fillna(t.Customer2); t["Channel"] = t.Channel.fillna(t.Channel2); t["Stockist"] = t.Stockist.fillna(t.Stockist2)
    t = t.drop(columns=["Customer2", "Channel2", "Stockist2"]).fillna({"cy": 0.0, "ly": 0.0})
    t["delta"] = t.cy - t.ly
    return t.reset_index()


def run_rate(df, targets, fy, fm_end, last_date, codes=None, kind="Primary", channel=None) -> pd.DataFrame:
    """Per stockist: YTD, annual target, balance, required vs current run-rate, projection."""
    d = df[(df.fy == fy) & (df.fm <= fm_end)]
    ytd = d.groupby("stockist_code").kl.sum()
    ann = target_by_stockist(targets, fy, range(1, 13), kind, channel)
    pro = target_by_stockist(targets, fy, range(1, fm_end + 1), kind, channel)
    names = df.drop_duplicates("stockist_code").set_index("stockist_code").stockist
    elapsed = fm_end
    if last_date is not None and pd.notna(last_date):
        end = last_date
        from .periods import mi_of, month_end, month_start
        mi = mi_of(fy, fm_end)
        if end.year * 12 + end.month - 1 == mi and end < month_end(mi):
            elapsed = fm_end - 1 + end.day / month_end(mi).day
    left = 12 - elapsed
    rows = []
    for c in sorted(set(ytd.index) | set(ann.index)):
        a, y = ann.get(c, np.nan), float(ytd.get(c, 0.0))
        cur = y / elapsed if elapsed else np.nan
        rows.append({"Stockist": names.get(c, str(c)), "YTD": y, "Pro-rata target": pro.get(c, np.nan), "% Tgt ach": pct(y, pro.get(c, np.nan)),
                     "Annual target": a, "Balance": a - y if a == a else np.nan,
                     "Required / month": (a - y) / left if a == a and left > 0 else np.nan,
                     "Current run-rate / month": cur, "Projected FY": y + cur * left if cur == cur else np.nan})
    t = pd.DataFrame(rows)
    if len(t):
        tot = {"Stockist": "Total"}
        for c in t.columns[1:]: tot[c] = t[c].sum(min_count=1)
        tot["% Tgt ach"] = pct(tot["YTD"], tot["Pro-rata target"])
        t = pd.concat([t, pd.DataFrame([tot])], ignore_index=True)
        t["Projected % of target"] = [pct(p, a) for p, a in zip(t["Projected FY"], t["Annual target"])]
    return t


def month_heat(df, targets, fy, kind="Primary", channel=None) -> pd.DataFrame:
    """Stockist x fiscal month: actual / target %."""
    act = df[df.fy == fy].groupby(["stockist_code", "fm"]).kl.sum().unstack().reindex(columns=range(1, 13))
    t = _tsel(targets, fy, range(1, 13), kind, channel)
    if t is None or t.empty: return pd.DataFrame()
    tg = t.groupby(["stockist_code", "fm"]).target.sum().unstack().reindex(columns=range(1, 13))
    names = df.drop_duplicates("stockist_code").set_index("stockist_code").stockist
    r = (act / tg * 100).reindex(tg.index)
    r.index = [names.get(i, str(i)) for i in r.index]
    return r.where(act.reindex(tg.index).notna().values)
