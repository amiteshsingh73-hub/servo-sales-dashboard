"""Scheme eligibility data + approximate incentive calculator.

Works on any period for which the rule book has rows. Months that already have sales data use actuals; months
without data use editable assumptions (target x achievement %, LY x growth, or a reference quarter's behaviour).
Results are APPROXIMATE (see the 'Not modelled' note in the rule book) - the circular and SAP/SDMS claims are final.
"""
from __future__ import annotations
from dataclasses import dataclass
import numpy as np
import pandas as pd
from . import metrics as M, scheme_rules as R
from .periods import month_start, month_end, mi_of, mi_name

ELIG = ["Mass", "Aspirational", "Premium", "Silver", "Industrial"]        # grades on which MPI/QPI/Early Bird are paid
CATS = ["Premium", "Aspirational", "Silver", "Mass", "Industrial", "Special"]
PAY = {"Early Bird": "In the SAP primary invoice (ZEBD) - monthly",
       "MPI": "SAP rebate module - after the month",
       "QPI": "SAP rebate module - after the quarter",
       "Categorisation": "Through SAP - after the quarter",
       "Mass trade discount": "SAP rebate module - after the scheme period",
       "Metro": "Applied on DLP in the invoice",
       "SSF": "SAP rebate (quarterly) after RCL verifies SSF payment proof",
       "RPI": "SDMS claim - after the quarter", "Festival Bonanza": "SDMS claim - after the quarter",
       "PPI": "SDMS claim - after the quarter", "RGI": "SDMS claim - after the quarter",
       "ServoXpert": "SDMS claim - after the quarter", "Servo Excellence": "SAP rebate module - after Sep'26"}


# ---------------------------------------------------------------- periods
def period_dates(pid: str):
    _, s, e, _, _ = R.PERIODS[pid]
    return pd.Timestamp(s), pd.Timestamp(e)


def period_months(pid: str) -> list[int]:
    s, e = period_dates(pid)
    a, b = s.year * 12 + s.month - 1, e.year * 12 + e.month - 1
    return list(range(a, b + 1))


def fy_fm(mi: int):
    y, m0 = divmod(mi, 12); m = m0 + 1
    return (y + 1 if m >= 4 else y), (m - 4) % 12 + 1


def shift_period(pid: str, k: int) -> str:
    ids = list(R.PERIODS)
    i = ids.index(pid) + k
    return ids[i] if 0 <= i < len(ids) else None


def window_before(start: pd.Timestamp, months: int):
    s = start - pd.DateOffset(months=months)
    return s, start - pd.Timedelta(days=1)


def months_of(a: pd.Timestamp, b: pd.Timestamp):
    return list(range(a.year * 12 + a.month - 1, b.year * 12 + b.month))


def slots(rs: R.RuleSet):
    return [(int(rs.f(f"eb{i}_from", 0)), int(rs.f(f"eb{i}_to", 0)), rs.f(f"eb{i}_rate", 0.0)) for i in range(1, 5)]


def slot_index(day: int, sl) -> int:
    for i, (a, b, _) in enumerate(sl):
        if a <= day <= b: return i
    return len(sl) - 1


# ---------------------------------------------------------------- primary profiles
def _p(pri, code, a, b):
    d = pri[(pri.stockist_code == code) & (pri.date >= a) & (pri.date <= b)]
    return d


def profile(pri, code, a, b, sl):
    """Category / Hydrashakti / early-bird slot shares for a stockist over a window."""
    d = _p(pri, code, a, b)
    tot = d.kl.sum()
    out = {"total": tot, "cat": {c: 0.0 for c in CATS}, "slot": [0.0] * 4, "hs": 0.0}
    if tot <= 0: return out
    for c in CATS: out["cat"][c] = float(d[d.category == c].kl.sum()) / tot
    out["hs"] = float(d[d.product_group == "Hydrashakti 68"].kl.sum()) / tot
    el = d[d.category.isin(ELIG)]
    if len(el):
        idx = el.date.dt.day.map(lambda x: slot_index(x, sl))
        v = el.groupby(idx.values).kl.sum()
        tv = v.sum()
        out["slot"] = [float(v.get(i, 0.0)) / tv for i in range(4)] if tv else [0.0] * 4
    return out


# ---------------------------------------------------------------- assumption grid
GRID_COLS = ["Primary KL", "Secondary KL", "Bazaar sec KL", "Premium %", "Aspirational %", "Silver %", "Mass %", "Industrial %",
             "Special/Hydrashakti %", "EB 1st slot %", "EB 2nd slot %", "EB 3rd slot %", "EB 4th slot %"]
BASIS = ["Target x achievement %", "Last year same month x (1 + growth %)", "Reference-quarter monthly average x (1 + growth %)"]


def build_grid(pri, sec, targets, rs: R.RuleSet, pid, codes, names, basis, ach_pct, growth_pct, ref_kind, last_p, last_s):
    """One editable row per stockist x month of the quarter."""
    sl = slots(rs)
    qs, qe = period_dates(pid)
    ref_a, ref_b = (qs - pd.DateOffset(months=3), qs - pd.Timedelta(days=1)) if ref_kind == "Previous quarter" else (qs - pd.DateOffset(years=1), qe - pd.DateOffset(years=1))
    rows = []
    for code in codes:
        ref = profile(pri, code, ref_a, ref_b, sl)
        ref_mean = ref["total"] / 3 if ref["total"] else np.nan
        sref = sec[(sec.stockist_code == code) & (sec.date >= ref_a) & (sec.date <= ref_b)]
        baz_share = float(sref[sref.cust_group.str.upper() == "BAZAAR"].kl.sum() / sref.kl.sum()) if sref.kl.sum() > 0 else 0.5
        for mi in period_months(pid):
            ms, me = month_start(mi), month_end(mi)
            fy, fm = fy_fm(mi)
            act = _p(pri, code, ms, me)
            status = "actual" if (pd.notna(last_p) and last_p >= me) else ("assumed" if (pd.isna(last_p) or last_p < ms) else "partial")
            tgt = M.target_for(targets, [code], fy, [fm], "Primary")
            ly_m = pri[(pri.stockist_code == code) & (pri.date >= ms - pd.DateOffset(years=1)) & (pri.date <= me - pd.DateOffset(years=1))].kl.sum()
            if basis == BASIS[0]: guess = tgt * ach_pct / 100 if tgt == tgt else ly_m * (1 + growth_pct / 100)
            elif basis == BASIS[1]: guess = ly_m * (1 + growth_pct / 100)
            else: guess = ref_mean * (1 + growth_pct / 100)
            guess = 0.0 if guess != guess else float(guess)
            if status == "actual":
                tot = float(act.kl.sum()); pr = profile(pri, code, ms, me, sl)
                cat, slot = pr["cat"], pr["slot"]
            elif status == "assumed":
                tot = guess; cat, slot = ref["cat"], ref["slot"]
            else:   # partial month: actual so far + remainder at reference mix
                a_tot = float(act.kl.sum()); tot = max(guess, a_tot); rem = tot - a_tot
                pa = profile(pri, code, ms, me, sl)
                cat = {c: ((pa["cat"][c] * a_tot) + ref["cat"][c] * rem) / tot if tot else 0 for c in CATS}
                slot = ref["slot"]
            # secondary
            fy_, fm_ = fy, fm
            s_act = sec[(sec.stockist_code == code) & (sec.date >= ms) & (sec.date <= me)]
            s_tgt = M.target_for(targets, [code], fy_, [fm_], "Secondary")
            s_ly = sec[(sec.stockist_code == code) & (sec.date >= ms - pd.DateOffset(years=1)) & (sec.date <= me - pd.DateOffset(years=1))].kl.sum()
            s_status = "actual" if (pd.notna(last_s) and last_s >= me) else "assumed"
            if s_status == "actual":
                s_tot = float(s_act.kl.sum()); s_baz = float(s_act[s_act.cust_group.str.upper() == "BAZAAR"].kl.sum())
            else:
                s_tot = (s_tgt * ach_pct / 100) if (basis == BASIS[0] and s_tgt == s_tgt) else s_ly * (1 + growth_pct / 100)
                s_tot = 0.0 if s_tot != s_tot else float(s_tot); s_baz = s_tot * baz_share
            row = {"Stockist": names[code], "code": code, "Month": mi_name(mi), "mi": mi, "Status": status if status != "actual" or s_status == "actual" else "actual",
                   "Primary KL": round(tot, 2), "Secondary KL": round(s_tot, 2), "Bazaar sec KL": round(s_baz, 2)}
            for c, col in zip(CATS, ["Premium %", "Aspirational %", "Silver %", "Mass %", "Industrial %", "Special/Hydrashakti %"]):
                row[col] = round(cat[c] * 100, 1)
            for i, col in enumerate(["EB 1st slot %", "EB 2nd slot %", "EB 3rd slot %", "EB 4th slot %"]):
                row[col] = round(slot[i] * 100, 1)
            row["Primary target"] = round(tgt, 2) if tgt == tgt else np.nan
            row["Secondary target"] = round(s_tgt, 2) if s_tgt == s_tgt else np.nan
            rows.append(row)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- counters (secondary)
def _bazaar(sec):
    return sec[sec.cust_group.str.upper() == "BAZAAR"]


def ssf_table(sec, window, rs: R.RuleSet, codes=None, base_date="2026-03-31"):
    a, b = window
    bz = _bazaar(sec)
    if codes is not None: bz = bz[bz.stockist_code.isin(codes)]
    base_ids = bz[(bz.date <= pd.Timestamp(base_date)) & (bz.kl > 0)].groupby("ssf_code").customer_id.nunique()
    inw = bz[(bz.date >= a) & (bz.date <= b)]
    billed = inw[inw.kl > 0].groupby("ssf_code").customer_id.nunique()
    first = bz[bz.kl > 0].groupby("customer_id").date.min()
    # new counters: first ever sale inside the window and >= min litres in some month of the window
    mon = inw.assign(mi=inw.date.dt.year * 12 + inw.date.dt.month).groupby(["ssf_code", "customer_id", "mi"]).kl.sum().reset_index()
    ok = mon[mon.kl * 1000 >= rs.f("ssf_new_min_l", 50)].customer_id.unique()
    newids = set(first[(first >= a) & (first <= b)].index) & set(ok)
    new = inw[inw.customer_id.isin(newids)].groupby("ssf_code").customer_id.nunique()
    vol = inw[inw.product_group != "Hydrashakti 68"].groupby("ssf_code").kl.sum()
    nm = bz.groupby("ssf_code").agg(SSF=("ssf_name", "last"), stockist=("stockist", "last"), code=("stockist_code", "last"))
    t = nm.join(pd.DataFrame({"Base counters (31-Mar-26)": base_ids, "Counters billed": billed, "New counters": new, "Bazaar KL (excl. HS68)": vol})).fillna(0)
    t = t[(t["Base counters (31-Mar-26)"] > 0) | (t["Counters billed"] > 0)]
    t["Coverage %"] = np.where(t["Base counters (31-Mar-26)"] > 0, t["Counters billed"] / t["Base counters (31-Mar-26)"] * 100, np.nan)
    return t.reset_index().rename(columns={"index": "ssf_code"}).sort_values(["stockist", "SSF"])


def counter_frame(sec, window, ly_window, rs: R.RuleSet, codes=None):
    """Per bazaar counter: quarter volume, premium SKUs/litres, LY base and the RPI slab it falls in."""
    a, b = window
    bz = _bazaar(sec)
    if codes is not None: bz = bz[bz.stockist_code.isin(codes)]
    d = bz[(bz.date >= a) & (bz.date <= b) & (bz.kl > 0)]
    nonhs = d[d.product_group != "Hydrashakti 68"]
    g = nonhs.groupby("customer_id").agg(Customer=("customer", "last"), Stockist=("stockist", "last"), code=("stockist_code", "last"),
                                         SSF=("ssf_name", "last"), litres=("kl", lambda x: x.sum() * 1000))
    pm = nonhs[nonhs.category == "Premium"]
    g["prem_litres"] = pm.groupby("customer_id").kl.sum() * 1000
    g["prem_skus"] = pm.groupby("customer_id").sku_name.nunique()
    g = g.fillna({"prem_litres": 0, "prem_skus": 0})
    ly = bz[(bz.date >= ly_window[0]) & (bz.date <= ly_window[1]) & (bz.product_group != "Hydrashakti 68")].groupby("customer_id").kl.sum() * 1000
    g["LY litres"] = ly.reindex(g.index).fillna(0)
    since = pd.Timestamp(rs.v("rpi_new_since", "2026-01-01"))
    recent = set(bz[(bz.date >= since) & (bz.date < a) & (bz.kl > 0)].customer_id)
    g["New/zero-base"] = (~g.index.isin(recent)) | (g["LY litres"] <= 0)
    return g.reset_index()


def rpi_apply(cf: pd.DataFrame, rs: R.RuleSet, scale=1.0):
    """Reseller Performance Incentive (Q3 onward) on a counter frame. `scale` multiplies volumes (projection)."""
    out = cf.copy()
    if out.empty:
        return out.assign(slab=[], rate=[], eligible_l=[], amount=[])
    for c in ("litres", "prem_litres", "LY litres"): out[c] = out[c] * scale
    cap = rs.f("rpi_cap_l", 1000)
    slabs, rates = [], []
    for r in out.to_dict("records"):
        rate, slab = 0.0, "not eligible"
        if r["New/zero-base"]:
            for i in (1, 2, 3):
                if r["litres"] > rs.f(f"rpi_n{i}_vol") and r["prem_skus"] >= rs.f(f"rpi_n{i}_sku") and r["prem_litres"] >= rs.f(f"rpi_n{i}_prem"):
                    rate, slab = rs.f(f"rpi_n{i}_rate"), f"new slab {i}"; break
        else:
            if (r["litres"] >= r["LY litres"] * (1 + rs.f("rpi_e_growth", 0.2)) and r["prem_skus"] >= rs.f("rpi_e_sku")
                    and r["prem_litres"] >= rs.f("rpi_e_prem")):
                rate, slab = rs.f("rpi_e_rate"), "existing growth"
        slabs.append(slab); rates.append(rate)
    out["slab"], out["rate"] = slabs, rates
    out["eligible_l"] = np.minimum(out["litres"], cap) if cap else out["litres"]
    out["amount"] = out["eligible_l"] * out["rate"]
    return out


def festival_ro(sec, window, apr_window, ly_apr_window, rs: R.RuleSet, codes=None, scale=1.0, cy_pre=None):
    """Servo Festival Bonanza: IOC RO outlets ('Dealer'). Growth is measured Apr-Dec vs LY Apr-Dec."""
    ro = sec[sec.cust_group.str.upper() == "DEALER"]
    if codes is not None: ro = ro[ro.stockist_code.isin(codes)]
    a, b = window
    q = ro[(ro.date >= a) & (ro.date <= b) & (ro.kl > 0) & (ro.product_group != "Hydrashakti 68")]
    g = q.groupby("customer_id").agg(Outlet=("customer", "last"), Stockist=("stockist", "last"), q_l=("kl", lambda x: x.sum() * 1000))
    g["prem_l"] = q[q.category == "Premium"].groupby("customer_id").kl.sum() * 1000
    g["honda_l"] = q[q.spec.astype(str).str.upper() == "HONDA"].groupby("customer_id").kl.sum() * 1000
    g = g.fillna(0)
    pre = ro[(ro.date >= apr_window[0]) & (ro.date < a) & (ro.product_group != "Hydrashakti 68")].groupby("customer_id").kl.sum() * 1000
    ly = ro[(ro.date >= ly_apr_window[0]) & (ro.date <= ly_apr_window[1]) & (ro.product_group != "Hydrashakti 68")].groupby("customer_id").kl.sum() * 1000
    g["Apr-pre-quarter l"] = pre.reindex(g.index).fillna(0)
    g["LY Apr-Dec l"] = ly.reindex(g.index).fillna(0)
    for c in ("q_l", "prem_l", "honda_l"): g[c] = g[c] * scale
    g["CY Apr-Dec l"] = g["Apr-pre-quarter l"] + g["q_l"]
    g["growth"] = np.where(g["LY Apr-Dec l"] > 0, g["CY Apr-Dec l"] / g["LY Apr-Dec l"] - 1, np.nan)
    rates, slab = [], []
    for r in g.to_dict("records"):
        gr = r["growth"]
        if gr == gr and gr > rs.f("fb_hi_growth", .15) and r["prem_l"] >= rs.f("fb_hi_prem_l", 40): rates.append(rs.f("fb_hi_rate")); slab.append("higher (>15% & premium litres)")
        elif gr == gr and gr > rs.f("fb_lo_growth", .10) and r["honda_l"] >= rs.f("fb_lo_honda_l", 40): rates.append(rs.f("fb_lo_rate")); slab.append("lower (>10% & Honda 4T litres)")
        else: rates.append(0.0); slab.append("not eligible")
    g["rate"], g["slab"] = rates, slab
    g["amount"] = g["q_l"] * g["rate"]
    return g.reset_index()


# ---------------------------------------------------------------- stockist calculations
@dataclass
class Inputs:
    pid: str
    dlp: dict                      # Rs/L by category {"Premium":..,"Aspirational":..,"Silver":..,"Mass":..}
    metro: dict                    # stockist code -> bool
    cat_override: dict             # code -> "Auto"/"Premium"/"Aspirational"/"None"
    ssf_assume: pd.DataFrame | None = None     # per SSF expected new counters / coverage % (projection)
    ref_kind: str = "Previous quarter"


def _val(r, key):
    return r[key] if key in r and r[key] == r[key] else 0.0


def calc_stockists(grid: pd.DataFrame, pri, sec, targets, rs: R.RuleSet, inp: Inputs, last_p, last_s):
    """Returns list of result rows (dicts) for the monthly/quarterly stockist schemes."""
    out = []
    pid = inp.pid
    qs, qe = period_dates(pid)
    sl = slots(rs)
    basis_tag = lambda mrows: ("Actual" if (mrows.Status == "actual").all() else ("Assumed" if (mrows.Status == "assumed").all() else "Actual + assumed"))
    for code, g in grid.groupby("code", sort=False):
        name = g.Stockist.iloc[0]
        tag = basis_tag(g)
        vol, el_vol, hs_vol, cat_kl = {}, {}, {}, {c: 0.0 for c in CATS}
        mrows = []
        for r in g.to_dict("records"):
            p = r["Primary KL"]
            sh = {c: _val(r, col) / 100 for c, col in zip(CATS, ["Premium %", "Aspirational %", "Silver %", "Mass %", "Industrial %", "Special/Hydrashakti %"])}
            tot_sh = sum(sh.values()) or 1.0
            sh = {k: v / tot_sh for k, v in sh.items()}
            ck = {c: p * sh[c] for c in CATS}
            elig = sum(ck[c] for c in ELIG)
            eb = np.array([_val(r, f"EB {o} slot %") for o in ("1st", "2nd", "3rd", "4th")], dtype=float)
            eb = eb / eb.sum() if eb.sum() > 0 else np.array([.4, .2, .2, .2])
            hs = ck["Special"]                      # Special category = Hydrashakti 68
            mrows.append({**r, "ck": ck, "elig": elig, "eb": eb, "hs": hs})
            for c in CATS: cat_kl[c] += ck[c]
        P = sum(m["Primary KL"] for m in mrows); ELG = sum(m["elig"] for m in mrows); HS = sum(m["hs"] for m in mrows)
        S = sum(m["Secondary KL"] for m in mrows)
        tP = sum(m["Primary target"] for m in mrows if m["Primary target"] == m["Primary target"])
        tS = sum(m["Secondary target"] for m in mrows if m["Secondary target"] == m["Secondary target"])
        has_pt = any(m["Primary target"] == m["Primary target"] for m in mrows)
        has_st = any(m["Secondary target"] == m["Secondary target"] for m in mrows)

        def add(scheme, cond, measured, required, met, rate, volume, amount, note="", basis=None):
            out.append({"Stockist": name, "code": code, "Scheme": scheme, "Condition": cond, "Measured": measured, "Required": required,
                        "Eligible": "Yes" if met else "No", "Rate": rate, "Volume (KL)": volume, "Approx. amount (Rs)": amount,
                        "Payable": PAY.get(scheme.split(" - ")[0], ""), "Basis": basis or tag, "Note": note})

        # ---- Early Bird (per month, per slot)
        if rs.on("eb_on"):
            amt, ekl = 0.0, 0.0
            for m in mrows:
                for i, (_, _, rate) in enumerate(sl):
                    amt += m["elig"] * m["eb"][i] * rate * 1000; ekl += m["elig"] * m["eb"][i] * (1 if rate > 0 else 0)
            add("Early Bird", "Eligible primary billed in the early days of each month (Rs/L by invoice day)",
                f"{ekl:,.1f} KL billed in paying slots", "-", ekl > 0, "Rs 0-1.5/L", ELG, amt, "Slots: " + ", ".join(f"{a}-{b}: Rs {r:g}" for a, b, r in sl))
        # ---- MPI (monthly)
        if rs.on("mpi_on"):
            cap = rs.v("mpi_hs_cap"); amt, okm, notes = 0.0, [], []
            for m in mrows:
                tgt = m["Primary target"]; hs_sh = m["hs"] / m["Primary KL"] if m["Primary KL"] else 0
                met = (tgt == tgt) and m["Primary KL"] >= tgt and (cap is None or hs_sh < float(cap))
                if met: amt += m["elig"] * rs.f("mpi_rate") * 1000; okm.append(m["Month"])
                elif tgt == tgt: notes.append(f"{m['Month']}: {m['Primary KL']:,.0f}/{tgt:,.0f} KL" + (f", HS {hs_sh:.0%}" if cap is not None else ""))
            add("MPI", f"Monthly primary target achieved" + (f" and Hydrashakti < {float(cap):.0%}" if cap is not None else ""),
                f"{len(okm)} of {len(mrows)} months met" + (f" ({', '.join(okm)})" if okm else ""), "each month", len(okm) > 0, f"Rs {rs.f('mpi_rate'):g}/L",
                sum(m["elig"] for m in mrows if m["Month"] in okm), amt, "Missed: " + "; ".join(notes) if notes else "")
        # ---- QPI
        if rs.on("qpi_on"):
            cap = rs.v("qpi_hs_cap"); hs_sh = HS / P if P else 0
            hs_ok = cap is None or hs_sh < float(cap)
            met_p = has_pt and P >= tP and hs_ok
            amt_p = ELG * rs.f("qpi_pri_rate") * 1000 if met_p else 0.0
            add("QPI - primary", "Quarterly primary target achieved" + (f" and Hydrashakti < {float(cap):.0%} of primary" if cap is not None else ""),
                f"{P:,.0f} KL (Hydrashakti {hs_sh:.0%})", f"{tP:,.0f} KL" if has_pt else "no target", met_p, f"Rs {rs.f('qpi_pri_rate'):g}/L", ELG if met_p else 0.0, amt_p,
                "" if met_p else (f"Short by {tP - P:,.0f} KL" if has_pt and P < tP else ("Hydrashakti share too high" if not hs_ok else "")))
            met_s = met_p and has_st and S >= tS
            add("QPI - secondary", "Quarterly SECONDARY target achieved (primary target also needed)", f"{S:,.0f} KL", f"{tS:,.0f} KL" if has_st else "no target",
                met_s, f"Rs {rs.f('qpi_sec_rate'):g}/L", ELG if met_s else 0.0, ELG * rs.f("qpi_sec_rate") * 1000 if met_s else 0.0,
                "" if met_s else (f"Short by {tS - S:,.0f} KL" if has_st and S < tS else "Needs primary target first"))
        # ---- categorisation (previous period)
        if rs.on("cat_on"):
            ca = qs - pd.DateOffset(months=int(rs.f("cat_months", 3)))
            pr = _p(pri, code, ca, qs - pd.Timedelta(days=1)); tot_ref = pr.kl.sum(); prem_ref = pr[pr.category == "Premium"].kl.sum()
            thr_p = min(rs.f("cat_prem_share") * tot_ref, rs.f("cat_prem_kl")); thr_a = min(rs.f("cat_asp_share") * tot_ref, rs.f("cat_asp_kl"))
            auto = "Premium" if prem_ref > thr_p else ("Aspirational" if prem_ref > thr_a else "None")
            cat = inp.cat_override.get(code, "Auto"); cat = auto if cat in (None, "Auto") else cat
            rate = {"Premium": rs.f("cat_prem_rate"), "Aspirational": rs.f("cat_asp_rate"), "None": 0.0}[cat]
            base = cat_kl["Aspirational"] + cat_kl["Premium"]
            add("Categorisation", f"Premium-grade share in {ca:%b'%y}-{(qs - pd.Timedelta(days=1)):%b'%y}: > lower of {rs.f('cat_prem_share'):.0%}/{rs.f('cat_prem_kl'):g} KL = Premium; > {rs.f('cat_asp_share'):.0%}/{rs.f('cat_asp_kl'):g} KL = Aspirational",
                f"Premium grade {prem_ref:,.1f} KL of {tot_ref:,.0f} KL ({(prem_ref / tot_ref if tot_ref else 0):.1%}) -> {cat}", f"> {thr_p:,.1f} KL (Premium) / > {thr_a:,.1f} KL (Aspirational)",
                cat != "None", f"Rs {rate:g}/L on Aspirational+Premium", base if cat != "None" else 0.0, base * rate * 1000 if cat != "None" else 0.0,
                "Override applied" if inp.cat_override.get(code, "Auto") not in (None, "Auto") else "", basis="Actual (previous period)" if tot_ref > 0 else "No data")
        # ---- Mass trade discount
        if rs.on("mass_on"):
            prem_q = cat_kl["Premium"]; tier, pct = 0, rs.f("mass_base_pct")
            for i in (1, 2, 3):
                sh_, kl_ = rs.f(f"mass_t{i}_share"), rs.v(f"mass_t{i}_kl")
                thr = min(sh_ * P, float(kl_)) if kl_ is not None else sh_ * P
                if prem_q > thr: tier, pct = i, rs.f(f"mass_t{i}_pct"); break
            met = has_pt and P >= tP and tier > 0
            diff = max(pct - rs.f("mass_base_pct"), 0.0)
            d_mass = inp.dlp.get("Mass", 0.0)
            amt = cat_kl["Mass"] * 1000 * d_mass * diff if (met and d_mass > 0) else (np.nan if (met and diff > 0) else 0.0)
            add("Mass trade discount", "Premium-grade share of quarter primary decides extra trade discount on MASS grades (primary target mandatory)",
                f"Premium {prem_q:,.1f} KL = {(prem_q / P if P else 0):.1%} of primary -> tier {tier or '-'} ({pct:.1%})", "Primary target achieved" if has_pt else "no target", met,
                f"{diff:.1%} of DLP above base", cat_kl["Mass"], amt, "Enter Mass DLP (Rs/L) to value this" if (met and diff > 0 and d_mass <= 0) else "")
        # ---- Metro
        if rs.on("metro_on") and inp.metro.get(code, True):
            vol_m = cat_kl["Premium"] + cat_kl["Aspirational"] + cat_kl["Silver"]
            amt = sum(cat_kl[c] * 1000 * inp.dlp.get(c, 0.0) for c in ("Premium", "Aspirational", "Silver")) * rs.f("metro_pct")
            need = any(inp.dlp.get(c, 0.0) <= 0 for c in ("Premium", "Aspirational", "Silver"))
            add("Metro", "Godown in metro municipal limits (Mumbai/Navi Mumbai/Thane): extra % on Premium, Aspirational & Silver", "Metro stockist", "-", True,
                f"{rs.f('metro_pct'):.1%} of DLP", vol_m, amt if not need else (amt if amt > 0 else np.nan), "Enter DLP (Rs/L) for Premium/Aspirational/Silver" if need else "")
        # ---- H1-only: Excellence
        if rs.on("exc_on"):
            exc_p = pri[(pri.stockist_code == code) & (pri.date >= qs.replace(month=4, day=1)) & (pri.date <= qe) & (~pri.category.isin(["Special"])) &
                        (~pri.segment.isin(["GENUINE OILS"])) & (~pri.spec.astype(str).str.contains("Hydraulic", case=False))].kl.sum()
            h1_p = pri[(pri.stockist_code == code) & (pri.date >= qs.replace(month=4, day=1)) & (pri.date <= qe)].kl.sum()
            h1_t = M.target_for(targets, [code], 2027, range(1, 7), "Primary")
            met = h1_t == h1_t and h1_p >= h1_t
            add("Servo Excellence", "Apr-Sep primary target achieved: Rs/L on primary excl. Hydrashakti 68, hydraulic oils & OEM grades", f"{h1_p:,.0f} KL", f"{h1_t:,.0f} KL" if h1_t == h1_t else "no target",
                met, f"Rs {rs.f('exc_rate'):g}/L", exc_p if met else 0.0, exc_p * rs.f("exc_rate") * 1000 if met else 0.0, "Counted on Apr-Sep actuals", basis="Actual")
    return out


# ---------------------------------------------------------------- reseller-level + orchestration
def h1_reseller(sec, window, ly_window, rs: R.RuleSet, codes=None, scale=1.0):
    """PPI / RGI (bazaar) and ServoXpert slabs from circular 01 (Apr-Sep'26)."""
    rows = []
    a, b = window
    d = sec[(sec.date >= a) & (sec.date <= b) & (sec.kl > 0)]
    ly = sec[(sec.date >= ly_window[0]) & (sec.date <= ly_window[1])]
    if codes is not None: d, ly = d[d.stockist_code.isin(codes)], ly[ly.stockist_code.isin(codes)]
    bz, bzl = d[d.cust_group.str.upper() == "BAZAAR"], ly[ly.cust_group.str.upper() == "BAZAAR"]
    if rs.on("ppi_on"):
        pm = bz[bz.category == "Premium"].groupby(["stockist", "customer_id"]).kl.sum() * 1000 * scale
        ok = pm[pm > rs.f("ppi_min_l", 50)]
        for st_, v in ok.groupby(level=0).sum().items(): rows.append({"Stockist": st_, "Scheme": "PPI", "Volume (KL)": v / 1000, "Rate": rs.f("ppi_rate"), "Approx. amount (Rs)": v * rs.f("ppi_rate"), "Counters": int((ok.index.get_level_values(0) == st_).sum())})
    if rs.on("rgi_on"):
        cur = bz[bz.product_group != "Hydrashakti 68"].groupby(["stockist", "customer_id"]).kl.sum() * 1000 * scale
        base = bzl[bzl.product_group != "Hydrashakti 68"].groupby(["stockist", "customer_id"]).kl.sum() * 1000
        base = base.reindex(cur.index).fillna(0).where(lambda x: x > 0, rs.f("rgi_nil_base_l", 120))
        ok = cur[cur > base * (1 + rs.f("rgi_growth", .10))]
        for st_, v in ok.groupby(level=0).sum().items(): rows.append({"Stockist": st_, "Scheme": "RGI", "Volume (KL)": v / 1000, "Rate": rs.f("rgi_rate"), "Approx. amount (Rs)": v * rs.f("rgi_rate"), "Counters": int((ok.index.get_level_values(0) == st_).sum())})
    if rs.on("sx_on"):
        sx = d[(d.cust_group.str.upper() == "SERVOXPERT") & d.category.isin(["Aspirational", "Premium", "Silver"])]
        v = sx.groupby(["stockist", "customer_id"]).kl.sum() * 1000 * scale
        for (st_, cid), x in v.items():
            r_ = rs.f("sx_s2_rate") if x > rs.f("sx_s2_min", 300) else (rs.f("sx_s1_rate") if x >= rs.f("sx_s1_min", 150) else 0.0)
            if r_ > 0: rows.append({"Stockist": st_, "Scheme": "ServoXpert", "Volume (KL)": x / 1000, "Rate": r_, "Approx. amount (Rs)": x * r_, "Counters": 1})
    if not rows: return pd.DataFrame(columns=["Stockist", "Scheme", "Volume (KL)", "Rate", "Approx. amount (Rs)", "Counters"])
    t = pd.DataFrame(rows)
    return t.groupby(["Stockist", "Scheme"], as_index=False).agg({"Volume (KL)": "sum", "Approx. amount (Rs)": "sum", "Counters": "sum", "Rate": "mean"})


def run_all(pri, sec, targets, rs: R.RuleSet, grid, inp: Inputs, last_p, last_s, name_by_code):
    """Full calculation. Returns dict(results, ssf, rpi, festival, notes)."""
    pid = inp.pid
    qs, qe = period_dates(pid)
    sec_complete = pd.notna(last_s) and last_s >= qe
    res = pd.DataFrame(calc_stockists(grid, pri, sec, targets, rs, inp, last_p, last_s))
    notes = []
    # reference window for counter-level schemes
    if sec_complete:
        win, ly_win, ref_note = (qs, qe), (qs - pd.DateOffset(years=1), qe - pd.DateOffset(years=1)), "actual"
    else:
        win = (qs - pd.DateOffset(months=3), qs - pd.Timedelta(days=1)) if inp.ref_kind == "Previous quarter" else (qs - pd.DateOffset(years=1), qe - pd.DateOffset(years=1))
        ly_win = (win[0] - pd.DateOffset(years=1), win[1] - pd.DateOffset(years=1))
        ref_note = f"projected from {win[0]:%b'%y}-{win[1]:%b'%y} counters, scaled to your assumed volumes"
    extra_rows, ssf_out, rpi_out, fb_out = [], pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    scale_b, scale_t = {}, {}
    for code, g in grid.groupby("code", sort=False):
        s_ref = sec[(sec.stockist_code == code) & (sec.date >= win[0]) & (sec.date <= win[1])]
        ref_b = s_ref[s_ref.cust_group.str.upper() == "BAZAAR"].kl.sum(); ref_t = s_ref.kl.sum()
        scale_b[code] = 1.0 if sec_complete else (g["Bazaar sec KL"].sum() / ref_b if ref_b > 0 else 0.0)
        scale_t[code] = 1.0 if sec_complete else (g["Secondary KL"].sum() / ref_t if ref_t > 0 else 0.0)
    codes = list(grid.code.unique())
    if rs.on("ssf_on"):
        st_ = ssf_table(sec, win, rs, codes)
        if not st_.empty:
            st_["scale"] = st_.code.map(scale_b).fillna(1.0)
            if not sec_complete and inp.ssf_assume is not None and len(inp.ssf_assume):
                am = inp.ssf_assume.set_index("ssf_code")
                st_["New counters"] = st_.ssf_code.map(am["Expected new counters"]).fillna(st_["New counters"])
                st_["Coverage %"] = st_.ssf_code.map(am["Expected coverage %"]).fillna(st_["Coverage %"])
            st_["Volume (KL)"] = st_["Bazaar KL (excl. HS68)"] * st_["scale"]
            st_["Add target met"] = st_["New counters"] >= rs.f("ssf_min_add", 12)
            st_["Coverage met"] = st_["Coverage %"] >= rs.f("ssf_cover_pct", .8) * 100
            st_["Approx. amount (Rs)"] = st_["Volume (KL)"] * 1000 * (st_["Add target met"] * rs.f("ssf_add_rate") + st_["Coverage met"] * rs.f("ssf_cover_rate"))
            ssf_out = st_
            for stk, gg in st_.groupby("stockist"):
                code = gg.code.iloc[0]
                extra_rows.append({"Stockist": stk, "code": code, "Scheme": "SSF", "Condition": f"Per SSF: >= {rs.f('ssf_min_add'):g} new counters and/or >= {rs.f('ssf_cover_pct'):.0%} of 31-Mar-26 bazaar counters billed (Rs {rs.f('ssf_add_rate'):g} each, on SSF bazaar volume excl. HS68)",
                    "Measured": f"{int(gg['Add target met'].sum())} of {len(gg)} SSFs met addition; {int(gg['Coverage met'].sum())} met coverage", "Required": "per SSF", "Eligible": "Yes" if (gg["Add target met"] | gg["Coverage met"]).any() else "No",
                    "Rate": "Rs 1 + Rs 1", "Volume (KL)": gg["Volume (KL)"].sum(), "Approx. amount (Rs)": gg["Approx. amount (Rs)"].sum(), "Payable": PAY["SSF"],
                    "Basis": "Actual" if sec_complete else "Assumed/projected", "Note": ref_note})
    if rs.on("rpi_on"):
        cf = counter_frame(sec, win, ly_win, rs, codes)
        if not cf.empty:
            cf["scale"] = cf.code.map(scale_b).fillna(1.0)
            parts = []
            for code, c in cf.groupby("code"): parts.append(rpi_apply(c, rs, c.scale.iloc[0]))
            rp = pd.concat(parts) if parts else cf
            rpi_out = rp
            for stk, gg in rp.groupby("Stockist"):
                q = gg[gg.rate > 0]
                extra_rows.append({"Stockist": stk, "code": gg.code.iloc[0], "Scheme": "RPI", "Condition": "Bazaar counters: litre slab + premium-grade SKUs/litres (new/zero-base) or >=20% growth (existing); max 1,000 L/counter",
                    "Measured": f"{len(q):,} of {len(gg):,} counters qualify", "Required": "per counter", "Eligible": "Yes" if len(q) else "No", "Rate": "Rs 3.00-3.50",
                    "Volume (KL)": q.eligible_l.sum() / 1000, "Approx. amount (Rs)": q.amount.sum(), "Payable": PAY["RPI"], "Basis": "Actual" if sec_complete else "Projected", "Note": ref_note})
    if rs.on("fb_on"):
        apr = (pd.Timestamp(year=qs.year if qs.month >= 4 else qs.year - 1, month=4, day=1), qs - pd.Timedelta(days=1))
        fb_win = win
        ly_apr = (apr[0] - pd.DateOffset(years=1), qe - pd.DateOffset(years=1))
        parts = []
        for code in codes:
            parts.append(festival_ro(sec, fb_win, apr, ly_apr, rs, [code], scale=scale_t[code]).assign(code=code))
        fb_out = pd.concat([p for p in parts if len(p)]) if any(len(p) for p in parts) else pd.DataFrame()
        if len(fb_out):
            for stk, gg in fb_out.groupby("Stockist"):
                q = gg[gg.rate > 0]
                extra_rows.append({"Stockist": stk, "code": gg.code.iloc[0], "Scheme": "Festival Bonanza", "Condition": "IOC RO: Apr-Dec growth >15% with 40 L premium (Rs 3) or >10% with 40 L Honda 4T (Rs 2)",
                    "Measured": f"{len(q):,} of {len(gg):,} ROs qualify", "Required": "per outlet", "Eligible": "Yes" if len(q) else "No", "Rate": "Rs 2-3",
                    "Volume (KL)": q.q_l.sum() / 1000, "Approx. amount (Rs)": q.amount.sum(), "Payable": PAY["Festival Bonanza"], "Basis": "Actual" if sec_complete else "Projected", "Note": ref_note})
    h1 = pd.DataFrame()
    if rs.on("ppi_on") or rs.on("rgi_on") or rs.on("sx_on"):
        sc = np.mean(list(scale_t.values())) if scale_t else 1.0
        h1 = h1_reseller(sec, win, ly_win, rs, codes, sc)
        for r in h1.to_dict("records"):
            code = next((c for c, n in name_by_code.items() if n == r["Stockist"]), None)
            extra_rows.append({"Stockist": r["Stockist"], "code": code, "Scheme": r["Scheme"], "Condition": {"PPI": "Bazaar counters: premium-grade litres > 50/quarter (Rs 5/L)", "RGI": ">10% growth vs LY same quarter (Rs 1/L)", "ServoXpert": "ServoXpert counters: 150-300 L Rs 1; >300 L Rs 2"}[r["Scheme"]],
                "Measured": f"{int(r['Counters']):,} counters qualify", "Required": "per counter", "Eligible": "Yes", "Rate": f"Rs {r['Rate']:g}", "Volume (KL)": r["Volume (KL)"],
                "Approx. amount (Rs)": r["Approx. amount (Rs)"], "Payable": PAY[r["Scheme"]], "Basis": "Actual" if sec_complete else "Projected", "Note": ref_note})
    # SDMS-booked discounts (information)
    for code in codes:
        d = sec[(sec.stockist_code == code) & (sec.date >= win[0]) & (sec.date <= win[1])]
        amt = (d.d_ho_promo.sum() + d.d_ho_sec.sum() + d.d_oem.sum()) * scale_t.get(code, 1.0)
        extra_rows.append({"Stockist": name_by_code[code], "code": code, "Scheme": "SDMS product-discount claims (info)", "Condition": "HO promotional + HO secondary + OEM/Fleet claimable discounts booked in SDMS (reseller product schemes, customer-specific discounts)",
            "Measured": f"{d.kl.sum() * scale_t.get(code, 1.0):,.0f} KL secondary", "Required": "-", "Eligible": "-", "Rate": "SKU-wise", "Volume (KL)": d.kl.sum() * scale_t.get(code, 1.0),
            "Approx. amount (Rs)": amt, "Payable": "SDMS claim (recovery of discounts given to customers) - not an incentive", "Basis": "Actual" if sec_complete else "Projected from " + f"{win[0]:%b'%y}-{win[1]:%b'%y}", "Note": "Kept out of the incentive total"})
    allr = pd.concat([res, pd.DataFrame(extra_rows)], ignore_index=True) if extra_rows else res
    return {"results": allr, "ssf": ssf_out, "rpi": rpi_out, "festival": fb_out, "window": win, "complete": sec_complete}
