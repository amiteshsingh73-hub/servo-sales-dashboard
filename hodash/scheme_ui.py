"""Streamlit screens for the scheme eligibility data + incentive calculator."""
from __future__ import annotations
import io, re
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from . import config as C, metrics as M, schemes as S, scheme_rules as R
from .periods import mi_name

NAVY, ORANGE, GREY = "#0B1F66", "#E8710A", "#B0B7C3"
CAT = ["#0B1F66", "#E8710A", "#2A9D8F", "#8E6BBF", "#C0392B", "#7F8C8D", "#4C78A8", "#F2B134", "#3B7D23", "#B5446E", "#00A6D6", "#8C6D31"]


def rs_fmt(x, d=0):
    if x is None or (isinstance(x, float) and np.isnan(x)): return "needs input"
    if abs(x) >= 1e5: return f"₹ {x / 1e5:,.2f} L"
    return f"₹ {x:,.{d}f}"


def _table(df, height="content"):
    d = df.copy()
    for c in d.columns:
        if pd.api.types.is_float_dtype(d[c]): d[c] = d[c].map(lambda v: "" if pd.isna(v) else (f"{v:,.1f}" if abs(v) < 1000 else f"{v:,.0f}"))
    st.dataframe(d, hide_index=True, width="stretch", height=height)


def default_period(last_p):
    for pid in R.PERIODS:
        if pd.isna(last_p) or S.period_dates(pid)[1] > last_p: return pid
    return list(R.PERIODS)[-1]


def circular_library():
    rows = []
    if not C.SCHEME_DIR.exists(): return pd.DataFrame()
    for f in sorted(C.SCHEME_DIR.rglob("*.pdf")):
        n = f.stem.upper()
        per = []
        if re.search(r"APR.?26.?-?JUNE|APR.?26", n): per.append("Q1")
        if re.search(r"JULY|SEPT|AUG|JUL", n): per.append("Q2")
        if re.search(r"OCT|DEC|SEPT.?26 ?\(|OCT.?26", n): per.append("Q3")
        if re.search(r"MARCH.?27|OCT.?26.?-?MARCH", n): per.append("Q4")
        if re.search(r"2026-27|FY|ANNEX", n) and not per: per.append("FY")
        rows.append({"Folder": f.parent.name, "Circular": f.name, "Likely period": ", ".join(dict.fromkeys(per)) or "see circular",
                     "Modified": pd.Timestamp(f.stat().st_mtime, unit="s").strftime("%d-%b-%Y")})
    return pd.DataFrame(rows)


def render(data, pri_all, sec_all, last_p, last_s):
    targets = data.targets
    rules_df = R.load_rules()
    if "rb_df" not in st.session_state or st.session_state.get("rb_sig") != str(C.RULES_FILE.stat().st_mtime):
        st.session_state["rb_sig"] = str(C.RULES_FILE.stat().st_mtime)
    names = dict(zip(pri_all.stockist_code, pri_all.stockist))
    names.update({c: n for c, n in zip(sec_all.stockist_code, sec_all.stockist) if c not in names})

    st.markdown("#### Scheme eligibility & incentive calculator")
    pids = list(R.PERIODS)
    c1, c2, c3 = st.columns([1.4, 2, 2])
    pid = c1.selectbox("Scheme period", pids, index=pids.index(default_period(last_p)), format_func=lambda p: R.PERIODS[p][0], key="sc_pid")
    rs = R.RuleSet(rules_df, pid)
    qs, qe = S.period_dates(pid)
    c2.markdown(f"**Circular(s):** {R.PERIODS[pid][3]}")
    codes_all = sorted(names)
    sel_codes = c3.multiselect("Stockists", codes_all, default=codes_all, format_func=lambda c: names[c], key="sc_codes")
    if rs.status == "provisional":
        st.warning(f"No circular has been issued for {R.PERIODS[pid][0]} yet: the rules shown are copied from the Oct-Dec'26 circular (01C) and are **provisional**. "
                   f"Update them in the 'Rules & circulars' tab when the circular arrives.")
    if not sel_codes: st.stop()

    t_el, t_calc, t_rules = st.tabs(["1 · Eligibility data", "2 · Incentive calculator", "3 · Rules & circulars"])

    # ============================================================ 1. eligibility data
    with t_el:
        st.caption("Every number the schemes test, for the selected quarter (actuals where sales data exists). Download all of it as one Excel workbook.")
        sl = S.slots(rs)
        qm = S.period_months(pid)
        pr = pri_all[pri_all.stockist_code.isin(sel_codes)]
        sc = sec_all[sec_all.stockist_code.isin(sel_codes)]
        # --- scoreboard
        rows = []
        for code in sel_codes:
            for mi in qm:
                ms, me = S.month_start(mi), S.month_end(mi)
                d = S._p(pr, code, ms, me); fy, fm = S.fy_fm(mi)
                tot = d.kl.sum(); prem = d[d.category == "Premium"].kl.sum(); hs = d[d.product_group == "Hydrashakti 68"].kl.sum()
                elig = d[d.category.isin(S.ELIG)].kl.sum()
                ds_ = sc[(sc.stockist_code == code) & (sc.date >= ms) & (sc.date <= me)]
                tp = M.target_for(targets, [code], fy, [fm], "Primary"); ts = M.target_for(targets, [code], fy, [fm], "Secondary")
                rows.append({"Stockist": names[code], "Month": mi_name(mi), "Primary KL": tot, "Primary target": tp, "Primary ach %": M.pct(tot, tp),
                             "Eligible KL (excl. Special)": elig, "Premium KL": prem, "Premium % of primary": prem / tot * 100 if tot else np.nan,
                             "Hydrashakti KL": hs, "Hydrashakti %": hs / tot * 100 if tot else np.nan,
                             "Secondary KL": ds_.kl.sum(), "Secondary target": ts, "Secondary ach %": M.pct(ds_.kl.sum(), ts)})
        score = pd.DataFrame(rows)
        # --- early bird
        eb = []
        for code in sel_codes:
            for mi in qm:
                d = S._p(pr, code, S.month_start(mi), S.month_end(mi)); d = d[d.category.isin(S.ELIG)]
                idx = d.date.dt.day.map(lambda x: S.slot_index(x, sl)) if len(d) else pd.Series(dtype=int)
                v = d.groupby(idx.values).kl.sum() if len(d) else pd.Series(dtype=float)
                eb.append({"Stockist": names[code], "Month": mi_name(mi), **{f"Days {a}-{b} (Rs {r:g}/L) KL": float(v.get(i, 0.0)) for i, (a, b, r) in enumerate(sl)}})
        eb = pd.DataFrame(eb)
        # --- categorisation
        ca = qs - pd.DateOffset(months=int(rs.f("cat_months", 3))); cb = qs - pd.Timedelta(days=1)
        cat = []
        for code in sel_codes:
            d = S._p(pr, code, ca, cb); tot = d.kl.sum(); prem = d[d.category == "Premium"].kl.sum()
            tp_, ta_ = min(rs.f("cat_prem_share") * tot, rs.f("cat_prem_kl")), min(rs.f("cat_asp_share") * tot, rs.f("cat_asp_kl"))
            cat.append({"Stockist": names[code], "Basis period": f"{ca:%b'%y}-{cb:%b'%y}", "Total primary KL": tot, "Premium-grade KL": prem, "Premium share %": prem / tot * 100 if tot else np.nan,
                        "Premium threshold KL": tp_, "Aspirational threshold KL": ta_, "Category earned": "Premium" if prem > tp_ else ("Aspirational" if prem > ta_ else "None"),
                        "KL short of next category": max((tp_ - prem) if prem <= tp_ and prem > ta_ else (ta_ - prem if prem <= ta_ else 0), 0)})
        cat = pd.DataFrame(cat)
        # --- counters (reference window: the quarter if complete, else previous quarter)
        sec_complete = pd.notna(last_s) and last_s >= qe
        win = (qs, qe) if sec_complete else (qs - pd.DateOffset(months=3), qs - pd.Timedelta(days=1))
        ly_win = (win[0] - pd.DateOffset(years=1), win[1] - pd.DateOffset(years=1))
        ssf = S.ssf_table(sc, win, rs, sel_codes)
        cf = S.rpi_apply(S.counter_frame(sc, win, ly_win, rs, sel_codes), rs) if rs.on("rpi_on") else S.counter_frame(sc, win, ly_win, rs, sel_codes)
        apr = (pd.Timestamp(year=qs.year if qs.month >= 4 else qs.year - 1, month=4, day=1), qs - pd.Timedelta(days=1))
        ro = pd.concat([S.festival_ro(sc, win, apr, (apr[0] - pd.DateOffset(years=1), qe - pd.DateOffset(years=1)), rs, [c]) for c in sel_codes]) if rs.on("fb_on") else pd.DataFrame()
        sets = {"Stockist scoreboard (monthly)": score, "Early Bird slots": eb, "Stockist categorisation test": cat,
                "SSF counters (addition & coverage)": ssf.drop(columns=[c for c in ("code", "scale") if c in ssf.columns]),
                "Bazaar counters (RPI inputs)": cf.drop(columns=[c for c in ("code",) if c in cf.columns]), "RO outlets (Festival Bonanza inputs)": ro.drop(columns=[c for c in ("code",) if c in ro.columns]) if len(ro) else ro}
        pick = st.radio("Dataset", list(sets), horizontal=True, key="el_ds")
        note = {"Stockist scoreboard (monthly)": "Target achievement, Hydrashakti share and premium-grade share drive MPI, QPI, Mass tiers and categorisation.",
                "Early Bird slots": "Eligible primary (Mass, Aspirational, Premium, Silver, Industrial) by invoice day.",
                "Stockist categorisation test": "Premium-grade upliftment of the look-back period decides the Rs/L benefit for the quarter.",
                "SSF counters (addition & coverage)": f"Counters as on 31-Mar-26 vs billed in {win[0]:%b'%y}-{win[1]:%b'%y}; new counters need >= {rs.f('ssf_new_min_l'):g} L in a month.",
                "Bazaar counters (RPI inputs)": f"Counter-level volumes for {win[0]:%b'%y}-{win[1]:%b'%y} (excl. Hydrashakti 68), premium SKUs/litres and LY base; slab shown if RPI rules apply.",
                "RO outlets (Festival Bonanza inputs)": "RO outlets (customer group Dealer): Apr-Dec growth, premium-grade and Honda 4T litres."}[pick]
        st.caption(note)
        view = sets[pick]
        if len(view): _table(view.round(2), height=420)
        else: st.info("No rows for this selection / scheme not active in this period.")
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            for k, v in sets.items():
                if len(v): v.to_excel(xw, sheet_name=re.sub(r"[^A-Za-z0-9 ]", "", k)[:31], index=False)
        st.download_button("⬇ Download all eligibility data (Excel)", buf.getvalue(), f"Scheme_eligibility_{pid}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("Need a different cut (by SKU, customer, district…)? Use the **Explore** and **Pivot (advanced)** tabs on the same data.")

    # ============================================================ 2. calculator
    with t_calc:
        sec_complete = pd.notna(last_s) and last_s >= qe
        pri_complete = pd.notna(last_p) and last_p >= qe
        if pri_complete and sec_complete:
            st.success(f"{R.PERIODS[pid][0]} is complete in the sales data - amounts below are calculated from **actuals**.")
        elif (pd.notna(last_p) and last_p >= qs):
            st.info(f"Part of {R.PERIODS[pid][0]} is in the sales data (till {last_p:%d-%b}). Elapsed months use actuals; the rest uses your assumptions.")
        else:
            st.info(f"{R.PERIODS[pid][0]} has no sales yet, so volumes are **assumptions** (set below). Eligibility baselines (previous-period premium share, counters, LY growth) still come from actual data.")
        a1, a2, a3, a4 = st.columns(4)
        basis = a1.selectbox("Volume assumption", S.BASIS, key="sc_basis", help="How to fill months that have no sales yet. You can overwrite any cell in the grid below.")
        ach = a2.number_input("Achievement of target %", 0, 300, 100, 5, key="sc_ach")
        gr = a3.number_input("Growth % (for LY / reference bases)", -50, 200, 10, 5, key="sc_gr")
        ref_kind = a4.selectbox("Reference for mix & counters", ["Previous quarter", "Same quarter last year"], key="sc_ref",
                                help="Category mix, early-bird slot pattern and counter behaviour are taken from this period.")
        d1, d2, d3, d4 = st.columns(4)
        dlp = {"Premium": d1.number_input("Avg DLP Rs/L - Premium", 0.0, 2000.0, 0.0, 10.0, key="dlp_p"),
               "Aspirational": d2.number_input("Avg DLP Rs/L - Aspirational", 0.0, 2000.0, 0.0, 10.0, key="dlp_a"),
               "Silver": d3.number_input("Avg DLP Rs/L - Silver", 0.0, 2000.0, 0.0, 10.0, key="dlp_s"),
               "Mass": d4.number_input("Avg DLP Rs/L - Mass", 0.0, 2000.0, 0.0, 10.0, key="dlp_m")}
        st.caption("DLP is needed only for % based schemes (Metro 0.5%, Mass-grade trade discount). Primary sales files carry no value, so enter your average list price per litre; until then those rows show 'needs input'.")
        with st.expander("Stockist settings: metro status and category override"):
            m_cols = st.columns(len(sel_codes))
            metro, cato = {}, {}
            for col, code in zip(m_cols, sel_codes):
                metro[code] = col.checkbox(f"{names[code][:22]} - metro godown", True, key=f"metro_{code}")
                cato[code] = col.selectbox("Category", ["Auto", "Premium", "Aspirational", "None"], key=f"cato_{code}")
        grid0 = S.build_grid(pri_all, sec_all, targets, rs, pid, sel_codes, names, basis, ach, gr, ref_kind, last_p, last_s)
        sig = f"{pid}|{basis}|{ach}|{gr}|{ref_kind}|{sel_codes}|{last_p}|{last_s}"
        if st.session_state.get("grid_sig") != sig:
            st.session_state["grid_sig"] = sig; st.session_state["grid_df"] = grid0
        st.markdown("**Assumptions by stockist and month** (edit any cell - the results below update)")
        cfg = {c: st.column_config.NumberColumn(c, format="%.1f", min_value=0) for c in S.GRID_COLS}
        edited = st.data_editor(st.session_state["grid_df"].drop(columns=["code", "mi"]), column_config=cfg, hide_index=True, width="stretch", key=f"grid_{abs(hash(sig))}",
                                disabled=["Stockist", "Month", "Status", "Primary target", "Secondary target"])
        grid = edited.copy(); grid["code"] = st.session_state["grid_df"]["code"].values; grid["mi"] = st.session_state["grid_df"]["mi"].values
        bad = grid.groupby("Stockist").apply(lambda g: g[["Premium %", "Aspirational %", "Silver %", "Mass %", "Industrial %", "Special/Hydrashakti %"]].sum(axis=1).sub(100).abs().max())
        if (bad > 3).any(): st.caption("⚠ Category shares for some rows do not add to 100% - they are rescaled automatically.")

        ssf_assume = None
        if not sec_complete:
            ref_win = (qs - pd.DateOffset(months=3), qs - pd.Timedelta(days=1)) if ref_kind == "Previous quarter" else (qs - pd.DateOffset(years=1), qe - pd.DateOffset(years=1))
            base_ssf = S.ssf_table(sec_all, ref_win, rs, sel_codes)
            if len(base_ssf):
                with st.expander("SSF assumptions (new counters and coverage expected per SSF)", expanded=False):
                    ssf_a = pd.DataFrame({"ssf_code": base_ssf.ssf_code, "Stockist": base_ssf.stockist, "SSF": base_ssf.SSF,
                                          "Reference new counters": base_ssf["New counters"], "Reference coverage %": base_ssf["Coverage %"].round(0),
                                          "Expected new counters": base_ssf["New counters"], "Expected coverage %": base_ssf["Coverage %"].round(0)})
                    ssf_assume = st.data_editor(ssf_a, hide_index=True, width="stretch", key=f"ssfa_{abs(hash(sig))}",
                                                disabled=["ssf_code", "Stockist", "SSF", "Reference new counters", "Reference coverage %"], column_config={"ssf_code": None})
        inp = S.Inputs(pid, dlp, metro, cato, ssf_assume, ref_kind)
        out = S.run_all(pri_all, sec_all, targets, rs, grid, inp, last_p, last_s, names)
        res = out["results"].copy()
        info_mask = res.Scheme.str.contains("(info)", regex=False)
        inc = res[~info_mask].copy()
        inc["amt"] = inc["Approx. amount (Rs)"]
        tot_by = inc.groupby("Stockist").agg(total=("amt", lambda x: x.sum(min_count=0)), unvalued=("amt", lambda x: int(x.isna().sum())))
        grand = inc.amt.sum(); unv = int(inc.amt.isna().sum())

        st.markdown("##### Approximate incentive")
        kc = st.columns(len(tot_by) + 1)
        kc[0].metric("Total (all selected)", rs_fmt(grand), f"{unv} item(s) need DLP" if unv else None, delta_color="off")
        for col, (stk, r) in zip(kc[1:], tot_by.iterrows()): col.metric(stk[:24], rs_fmt(r.total))
        pay = inc.groupby("Payable").agg(amt=("amt", "sum"), unv=("amt", lambda x: int(x.isna().sum()))).sort_values("amt", ascending=False)
        a, b = st.columns([3, 2])
        pv = inc.pivot_table(index="Stockist", columns="Scheme", values="amt", aggfunc="sum").fillna(0)
        fig = go.Figure()
        for j, c in enumerate(pv.columns):
            fig.add_bar(x=pv.index, y=pv[c], name=c, marker_color=CAT[j % len(CAT)])
        fig.update_layout(barmode="stack", title="Approx. incentive by scheme (₹)", height=380, margin=dict(l=10, r=10, t=60, b=10), legend=dict(orientation="h", y=-0.2),
                          plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        a.plotly_chart(fig, width="stretch")
        b.markdown("**When it is paid**")
        pt = pay.reset_index()
        pt["Approx. ₹"] = [rs_fmt(a) + (f"  (+{u} needs DLP)" if u else "") for a, u in zip(pt.amt, pt.unv)]
        _table(pt.rename(columns={"Payable": "Mechanism / timing"})[["Mechanism / timing", "Approx. ₹"]])
        w0, w1 = out["window"]
        basis_txt = "actual data" if out["complete"] else f"projected from {w0:%b'%y}-{w1:%b'%y} counters, scaled to your assumed volumes"
        st.caption(f"Counter-level schemes (SSF, RPI, Festival Bonanza) basis: {basis_txt}.")

        st.markdown("##### Scheme by scheme")
        show = res.drop(columns=["code"]).copy()
        show["Approx. amount (Rs)"] = show["Approx. amount (Rs)"].map(lambda v: rs_fmt(v) if True else v)
        show["Volume (KL)"] = show["Volume (KL)"].map(lambda v: f"{v:,.1f}" if v == v else "")
        stk_f = st.multiselect("Show stockists", sorted(show.Stockist.unique()), default=sorted(show.Stockist.unique()), key="sc_show")
        st.dataframe(show[show.Stockist.isin(stk_f)][["Stockist", "Scheme", "Condition", "Measured", "Required", "Eligible", "Rate", "Volume (KL)", "Approx. amount (Rs)", "Payable", "Basis", "Note"]],
                     hide_index=True, width="stretch", height=520)

        # gap analysis
        st.markdown("##### What it takes to earn more")
        gaps = []
        for code, g in grid.groupby("code", sort=False):
            P = g["Primary KL"].sum(); tP = g["Primary target"].sum(min_count=1)
            prem = (g["Primary KL"] * g["Premium %"] / 100).sum(); hs = (g["Primary KL"] * g["Special/Hydrashakti %"] / 100).sum()
            cap = rs.v("qpi_hs_cap")
            if tP == tP and P < tP: gaps.append({"Stockist": names[code], "Lever": "Primary target", "Gap": f"{tP - P:,.1f} KL more primary needed to hit the quarterly target (unlocks QPI/MPI/Mass tiers)"})
            if cap is not None and P and hs / P >= float(cap): gaps.append({"Stockist": names[code], "Lever": "Hydrashakti share", "Gap": f"Hydrashakti is {hs / P:.0%} of primary (cap {float(cap):.0%}). To qualify for MPI/QPI either shift ~{hs - float(cap) * P:,.0f} KL from Hydrashakti to other grades, or add ~{hs / float(cap) - P:,.0f} KL of other grades on top"})
            if rs.on("mass_on") and P:
                for i in (3, 2, 1):
                    kl_ = rs.v(f"mass_t{i}_kl"); thr = min(rs.f(f"mass_t{i}_share") * P, float(kl_)) if kl_ is not None else rs.f(f"mass_t{i}_share") * P
                    if prem <= thr:
                        gaps.append({"Stockist": names[code], "Lever": f"Mass tier {i}", "Gap": f"{thr - prem + 0.1:,.1f} KL more premium-grade primary reaches the tier paying {rs.f(f'mass_t{i}_pct'):.1%}"}); break
        if gaps: _table(pd.DataFrame(gaps))
        else: st.write("All headline conditions are met under these assumptions.")

        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine="openpyxl") as xw:
            res.drop(columns=["code"]).to_excel(xw, sheet_name="Results", index=False)
            grid.drop(columns=["code", "mi"]).to_excel(xw, sheet_name="Assumptions", index=False)
            if len(out["ssf"]): out["ssf"].to_excel(xw, sheet_name="SSF", index=False)
            if len(out["rpi"]): out["rpi"].to_excel(xw, sheet_name="RPI counters", index=False)
            if len(out["festival"]): out["festival"].to_excel(xw, sheet_name="Festival ROs", index=False)
            rs.rows.to_excel(xw, sheet_name="Rules used", index=False)
        e1, e2 = st.columns([1, 3])
        e1.download_button("⬇ Download calculation (Excel)", buf2.getvalue(), f"Incentive_{pid}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        summ = inc.pivot_table(index="Scheme", columns="Stockist", values="amt", aggfunc="sum").fillna(0)
        summ["Total"] = summ.sum(axis=1)
        summ = summ.reset_index()
        for c in summ.columns[1:]: summ[c] = summ[c].map(lambda v: f"{v:,.0f}")
        if e2.checkbox("Use this summary in PPT Studio", False, key="sc_ppt"):
            st.session_state["pivot_ppt"] = (f"Approx. incentive {R.PERIODS[pid][0]} (Rs)", summ.set_index("Scheme")); st.session_state["pivot_ppt_src"] = "scheme"
        elif st.session_state.get("pivot_ppt_src") == "scheme" and "pivot_ppt" in st.session_state:
            del st.session_state["pivot_ppt"]
        st.caption("Approximate: based on the circular as read into the rule book and the sales data/assumptions above. Final payment is as processed in SAP/SDMS under the circular.")

    # ============================================================ 3. rules & circulars
    with t_rules:
        st.markdown(f"**Rule book:** `{C.RULES_FILE}`")
        st.caption("Edit values here (or directly in the Excel). Blank = no condition / not applicable. Shares are decimals (0.10 = 10%). Flags: 1 = active, 0 = inactive.")
        sub = rules_df[rules_df.Period == pid][["Scheme", "Key", "Description", "Value", "Unit", "Source", "Status"]].reset_index(drop=True)
        ed = st.data_editor(sub, hide_index=True, width="stretch", height=480, disabled=["Scheme", "Key", "Description", "Unit"], key=f"rules_{pid}")
        r1, r2 = st.columns([1, 3])
        if r1.button("💾 Save changes to rule book", key=f"save_{pid}"):
            full = rules_df.copy()
            m = full.Period == pid
            upd = ed.copy(); upd["Period"] = pid
            full = pd.concat([full[~m], upd.assign(Period=pid)[["Period", "Scheme", "Key", "Description", "Value", "Unit", "Source", "Status"]]], ignore_index=True)
            full["__o"] = full.Period.map({p: i for i, p in enumerate(R.PERIODS)}); full = full.sort_values("__o", kind="stable").drop(columns="__o")
            R.save_rules(full); st.success("Saved. The calculator now uses the updated rules."); st.rerun()
        if r2.button("Copy this period's rules to the next period (as provisional)", key=f"copy_{pid}"):
            nxt = S.shift_period(pid, 1)
            if nxt:
                full = rules_df.copy(); cur = full[full.Period == pid].copy()
                cur["Period"], cur["Status"] = nxt, "provisional"; cur["Source"] = f"Copied from {pid} - update when the circular is issued"
                full = pd.concat([full[full.Period != nxt], cur], ignore_index=True); R.save_rules(full); st.success(f"Copied to {nxt}."); st.rerun()
        st.markdown("##### Not modelled in the calculator")
        st.write(R.NOTES[-1])
        lib = circular_library()
        st.markdown("##### Circular library")
        if lib.empty: st.info(f"No circulars found in {C.SCHEME_DIR}")
        else:
            st.caption(f"{len(lib)} circulars in {C.SCHEME_DIR}. The calculator does not read the PDFs directly - when a new circular arrives, update the rule book above.")
            st.dataframe(lib, hide_index=True, width="stretch", height=380)
