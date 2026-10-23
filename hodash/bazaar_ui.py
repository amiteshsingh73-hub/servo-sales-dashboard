"""The 'Bazaar' tab: everything a retail (counter) sales manager reviews, on one page."""
from __future__ import annotations
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from . import bazaar as B, metrics as M, ui
from .periods import MONTHS, fy_label, mi_name

STATUS_COL = {"New": ui.TEAL, "Reactivated": "#4E79A7", "Growing": ui.GREEN, "Stable": ui.GREY, "Declining": ui.ORANGE, "Lost": ui.RED}
TIER_COL = {"Platinum": ui.NAVY, "Gold": "#D4A017", "Silver": "#8D99AE", "Bronze": "#B5651D"}


def _num(df: pd.DataFrame, show, height=None):
    """Round the numeric columns sensibly then hand over to the app's table renderer."""
    d = df.copy()
    for c in d.columns:
        if d[c].dtype.kind in "fi":
            d[c] = d[c].round(0 if any(k in str(c) for k in ("L /", "Counters", "counters", "Bills", "Days", "New", "Lost", "₹")) else 1)
    show(d, height=height or (380 if len(d) > 10 else "content"))


def _csv(df: pd.DataFrame, name: str, label="⬇ Download list (CSV)"):
    st.download_button(label, df.to_csv().encode("utf-8"), name, "text/csv", key=f"dl_{name}")


def render(bz: pd.DataFrame, ws: dict, period, fy: int, tg, codes, sec_last, show):
    if bz.empty:
        st.info("No Bazaar sales for the current sidebar filters.")
        return
    st.caption("Bazaar = counters in the 'Bazaar', 'ServoXpert' and 'Others' customer groups (secondary, ClearBlue excluded). "
               "Sidebar stockist / product / district filters apply; the channel filter does not.")
    k = B.kpis(bz, ws)
    cy, ly = k["cy"], k["ly"]
    tgt = M.target_for(tg, codes, fy, period.fms, "Secondary", "Bazaar")
    tgt_fy = M.target_for(tg, codes, fy, range(1, 13), "Secondary", "Bazaar")
    mt = B.monthly(bz, fy)
    spark_kl = mt.kl.dropna().tolist()
    spark_ct = mt.counters.dropna().tolist()

    # ------------------------------------------------------------ KPI strip
    ui.cards([
        ui.kpi("Bazaar sales", f"{cy['kl']:,.0f}", f"LY {ly['kl']:,.0f} KL", M.growth(cy["kl"], ly["kl"]), "KL", spark_kl),
        ui.kpi("Active counters", f"{cy['counters']:,}", f"LY {ly['counters']:,}", M.growth(cy["counters"], ly["counters"]), spark=spark_ct),
        ui.kpi("Throughput / counter", f"{cy['lpcm']:,.0f}" if cy["lpcm"] == cy["lpcm"] else "–",
               f"LY {ly['lpcm']:,.0f} L" if ly["lpcm"] == ly["lpcm"] else "", M.growth(cy["lpcm"], ly["lpcm"]), "L / month"),
        ui.kpi("Realisation", f"₹{cy['rpl']:,.0f}" if cy["rpl"] == cy["rpl"] else "–",
               f"LY ₹{ly['rpl']:,.0f}" if ly["rpl"] == ly["rpl"] else "", M.growth(cy["rpl"], ly["rpl"]), "/ L"),
    ])
    st.write("")
    ui.cards([
        ui.progress(f"Bazaar target · {ws['cy'].label}", cy["kl"], tgt),
        ui.kpi("Premium share", f"{cy['prem_share']:.1f}%" if cy["prem_share"] == cy["prem_share"] else "–",
               f"LY {ly['prem_share']:.1f}%" if ly["prem_share"] == ly["prem_share"] else "", None, tone="plain"),
        ui.kpi("Range selling", f"{cy['skus']:.1f}" if cy["skus"] == cy["skus"] else "–",
               f"SKUs per counter · LY {ly['skus']:.1f}" if ly["skus"] == ly["skus"] else "SKUs per counter",
               M.growth(cy["skus"], ly["skus"])),
        ui.kpi("Billing frequency", f"{cy['freq']:.2f}" if cy["freq"] == cy["freq"] else "–",
               f"bills / counter / month · LY {ly['freq']:.2f}" if ly["freq"] == ly["freq"] else "bills / counter / month",
               M.growth(cy["freq"], ly["freq"])),
    ])

    # ------------------------------------------------------------ trend
    ui.section("Monthly trend", f"{fy_label(fy)} vs {fy_label(fy - 1)}, with active counters")
    ml = B.monthly(bz, fy - 1)
    tm = M.monthly_target(tg, fy, "Secondary", "Bazaar", codes)
    in_p = [m in period.fms for m in range(1, 13)]
    fig = go.Figure()
    fig.add_bar(x=MONTHS, y=ml.kl.values, name=f"KL {fy_label(fy - 1, True)}", marker_color=ui.GREY)
    fig.add_bar(x=MONTHS, y=mt.kl.values, name=f"KL {fy_label(fy, True)}", marker_color=[ui.NAVY if i else ui.NAVY_LT for i in in_p])
    if len(tm):
        fig.add_scatter(x=MONTHS, y=tm.reindex(range(1, 13)).values, name="Target", mode="lines+markers", line=dict(color=ui.ORANGE, width=2, dash="dot"))
    fig.add_scatter(x=MONTHS, y=mt.counters.values, name="Active counters", yaxis="y2", mode="lines+markers", line=dict(color=ui.TEAL, width=2.5))
    fig.update_layout(barmode="group", title="Bazaar volume (KL) and active counters",
                      yaxis2=dict(overlaying="y", side="right", showgrid=False, title="Counters", rangemode="tozero"))
    ui.chart(ui.layout(fig, 380))
    if tgt_fy == tgt_fy and tgt_fy > 0:
        ytd = float(bz[(bz.fy == fy)].kl.sum())
        left = max(12 - period.fm_end, 0)
        req = (tgt_fy - ytd) / left if left else np.nan
        done = len(mt.kl.dropna())
        rr = ytd / done if done else np.nan
        st.caption(f"FY target {tgt_fy:,.0f} KL · done {ytd:,.0f} KL ({M.pct(ytd, tgt_fy):.0f}%)"
                   + (f" · need {req:,.0f} KL / month for the remaining {left} months vs current run-rate {rr:,.0f} KL / month." if req == req else "."))

    # ------------------------------------------------------------ counter health
    lc = B.lifecycle(bz, ws)
    ui.section("Counter health", f"each counter's {ws['cy'].label} vs {ws['ly'].label}  (Growing / Declining = more than ±10%)")
    summ = B.lifecycle_summary(lc)
    if summ.empty:
        st.info("No counter activity in this period.")
    else:
        a, b = st.columns([2, 3])
        with a:
            fig = go.Figure(go.Pie(labels=summ.index, values=summ.Counters, hole=.62, sort=False,
                                   marker=dict(colors=[STATUS_COL[s] for s in summ.index]), textinfo="value", hovertemplate="%{label}: %{value} counters<extra></extra>"))
            fig.update_layout(title="Counters by status", annotations=[dict(text=f"<b>{int(summ.Counters.sum()):,}</b><br>counters", showarrow=False, font=dict(size=15))])
            ui.chart(ui.layout(fig, 340, hover="closest"))
        with b:
            fig = go.Figure(go.Bar(x=summ.index, y=summ["Change (KL)"], marker_color=[STATUS_COL[s] for s in summ.index],
                                   text=[f"{v:+,.0f}" for v in summ["Change (KL)"]], textposition="outside"))
            fig.update_layout(title="Volume gained / lost by counter status (KL vs LY)")
            ui.chart(ui.layout(fig, 340, legend=False, hover="closest"))
        _num(summ.reset_index().rename(columns={"status": "Status"}), show)
        pick = st.selectbox("See counters in status", list(summ.index), index=list(summ.index).index("Declining") if "Declining" in summ.index else 0, key="bz_status")
        lst = B.counter_master(bz).join(lc[lc.status == pick], how="inner").sort_values("delta").rename(columns={"cy": "CY (KL)", "ly": "LY (KL)", "delta": "Change (KL)"})
        if pick in ("Growing", "New", "Reactivated"): lst = lst.sort_values("Change (KL)", ascending=False)
        _num(lst.drop(columns="status").head(50), show)
        _csv(lst.drop(columns="status"), f"Bazaar_{pick}_counters.csv")

    # ------------------------------------------------------------ dormant
    ui.section("Dormant counters — call list", f"bought in the 6 months to {sec_last:%d-%b-%Y} but no bill since")
    c1, c2 = st.columns([1, 4])
    days = c1.select_slider("No bill for at least", [15, 30, 45, 60, 90], value=30, format_func=lambda d: f"{d} days", key="bz_days")
    dm = B.dormant(bz, sec_last, days)
    if dm.empty:
        c2.success(f"No counter has gone {days}+ days without a bill. 👏")
    else:
        risk = dm["Avg L / month"].sum() / 1000
        c2.markdown(ui.kpi("Volume at risk", f"{risk:,.1f}", f"{len(dm):,} counters silent for {days}+ days — avg monthly offtake", unit="KL / month", tone="warn"), unsafe_allow_html=True)
        _num(dm.head(40), show)
        _csv(dm, f"Bazaar_dormant_{days}d.csv")

    # ------------------------------------------------------------ tiers
    ui.section("Counter tiers", "by average litres / month in the period — where the volume really sits")
    tr = B.tiers(bz, ws["cy"]); ts = B.tier_summary(tr)
    if not ts.empty:
        a, b = st.columns([3, 2])
        with a:
            fig = go.Figure()
            for t_ in ts.index:
                fig.add_bar(y=["Counters", "Volume"], x=[ts.loc[t_, "% Counters"], ts.loc[t_, "% Volume"]], name=t_, orientation="h",
                            marker_color=TIER_COL[t_], text=[f"{ts.loc[t_, '% Counters']:.0f}%", f"{ts.loc[t_, '% Volume']:.0f}%"], textposition="inside",
                            hovertemplate=f"{t_}: %{{x:.1f}}%<extra>%{{y}}</extra>")
            fig.update_layout(barmode="stack", title="Share of counters vs share of volume", xaxis=dict(ticksuffix="%", range=[0, 100.5]),
                              legend=dict(traceorder="normal"))
            ui.chart(ui.layout(fig, 260, hover="closest"))
        with b:
            st.write("")
            tsv = ts.reset_index().rename(columns={"tier": "Tier"})
            tsv["Counters"] = tsv["Counters"].astype(int)
            _num(tsv, show)
        st.caption(f"Tier bands: {', '.join(f'{n} {r}' for n, r in zip(ts.index, ts.Band))}. Bands are set in hodash/bazaar.py (TIERS).")
        pick_t = st.radio("Counters in tier", list(ts.index), horizontal=True, key="bz_tier",
                          format_func=lambda x: f"{x} ({int(ts.loc[x, 'Counters'])})")
        lst = B.counter_master(bz).join(tr[tr.tier == pick_t], how="inner").sort_values("lpm", ascending=False)
        lst = lst.rename(columns={"lpm": "Avg L / month", "kl": f"KL {ws['cy'].label}"}).drop(columns="tier")
        _num(lst, show)
        _csv(lst, f"Bazaar_{pick_t}_counters.csv")

    # ------------------------------------------------------------ sales force & sub-groups
    ui.section("Sales force scorecard (SSF)", "volume, growth, reach and productivity per field officer")
    sc = B.scorecard(bz, ws, "ssf_name", lc)
    if not sc.empty:
        fig = go.Figure()
        top = sc.head(15).iloc[::-1]
        fig.add_bar(y=top.index, x=top["LY (KL)"], name=f"LY {ws['ly'].label}", orientation="h", marker_color=ui.GREY)
        fig.add_bar(y=top.index, x=top["CY (KL)"], name=f"CY {ws['cy'].label}", orientation="h", marker_color=ui.NAVY,
                    text=[f"{v:,.0f} ({g:+.0f}%)" if g == g and abs(g) < 999 else f"{v:,.0f}" for v, g in zip(top["CY (KL)"], top["% Growth"])], textposition="outside")
        fig.update_layout(barmode="group", title="KL by SSF")
        fig.update_xaxes(range=[0, float(max(top["CY (KL)"].max(), top["LY (KL)"].max(), 1e-9)) * 1.3])
        a, b = st.columns([2, 3])
        with a: ui.chart(ui.layout(fig, max(300, 70 + 34 * len(top)), hover="closest"))
        with b: _num(sc.reset_index().rename(columns={"index": "SSF"}), show)

    ui.section("Counter type (sub-group) & geography")
    a, b = st.columns(2)
    for col, dim, ttl in ((a, "cust_subgroup", "Sub-group"), (b, "district", "District")):
        t = B.scorecard(bz, ws, dim, lc)
        with col:
            if t.empty: st.info("No data"); continue
            t = t.head(12); tc = t.head(8)           # chart: top 8 so labels stay readable; table: top 12
            xl = [ui.wrap(str(i).title(), 10) for i in tc.index]
            fig = go.Figure(go.Bar(x=xl, y=tc["CY (KL)"], marker_color=ui.NAVY, name="CY"))
            fig.add_scatter(x=xl, y=tc["LY (KL)"], mode="markers", name="LY", marker=dict(color=ui.ORANGE, size=14, symbol="line-ew-open", line=dict(width=3)))
            ui.growth_labels(fig, xl, tc["CY (KL)"], tc["LY (KL)"], tc["% Growth"])
            fig.update_layout(title=f"{ttl}: CY KL (LY marker, growth % on top)", xaxis_tickangle=0)
            ui.chart(ui.layout(fig, 400, hover="closest"))
            _num(t[["CY (KL)", "LY (KL)", "% Growth", "% Share", "Active counters", "L / counter / month", "₹ / L"]].reset_index().rename(columns={"index": ttl}), show)

    # ------------------------------------------------------------ product mix
    ui.section("Product & pack mix", "premiumisation and the grades / packs that move in Bazaar")
    a, b = st.columns(2)
    with a:
        pg = M.compare(bz, "product_group", ws).head(10)
        fig = go.Figure()
        fig.add_bar(x=pg.index, y=pg.ly, name="LY", marker_color=ui.GREY)
        fig.add_bar(x=pg.index, y=pg.cy, name="CY", marker_color=ui.NAVY, text=[f"{g:+.0f}%" if g == g and abs(g) < 999 else "" for g in pg.growth], textposition="outside")
        fig.update_layout(barmode="group", title="Product group (KL)", xaxis_tickangle=-25)
        ui.chart(ui.layout(fig, 360))
    with b:
        cat_order = ["Premium", "Aspirational", "Silver", "Mass", "Special", "Industrial"]
        cc = M.compare(bz, "category", ws)
        fig = go.Figure()
        for j, c in enumerate([c for c in cat_order if c in cc.index] + [c for c in cc.index if c not in cat_order]):
            tot_ly, tot_cy = cc.ly.sum() or 1, cc.cy.sum() or 1
            fig.add_bar(x=[f"LY {ws['ly'].label}", f"CY {ws['cy'].label}"], y=[cc.loc[c, "ly"] / tot_ly * 100, cc.loc[c, "cy"] / tot_cy * 100], name=c,
                        marker_color=ui.CAT[j % len(ui.CAT)], text=[f"{cc.loc[c, 'ly'] / tot_ly * 100:.0f}%", f"{cc.loc[c, 'cy'] / tot_cy * 100:.0f}%"], textposition="inside",
                        hovertemplate=f"{c}: %{{y:.1f}}%<extra>%{{x}}</extra>")
        fig.update_layout(barmode="stack", title="Category mix (% of volume)", yaxis=dict(ticksuffix="%", range=[0, 100]))
        ui.chart(ui.layout(fig, 360, hover="closest"))
    a, b = st.columns(2)
    with a:
        pk = M.compare(bz, "pack_group", ws)
        fig = go.Figure(go.Pie(labels=pk.index, values=pk.cy.clip(lower=0), hole=.55, sort=False, textposition="inside", textinfo="percent"))
        fig.update_layout(title="Pack mix (CY volume)", uniformtext=dict(minsize=10, mode="hide"))
        ui.chart(ui.layout(fig, 340, hover="closest"))
    with b:
        sk = M.compare(bz, "product", ws).head(12).iloc[::-1]
        fig = go.Figure()
        fig.add_bar(y=sk.index, x=sk.ly, name="LY", orientation="h", marker_color=ui.GREY)
        fig.add_bar(y=sk.index, x=sk.cy, name="CY", orientation="h", marker_color=ui.NAVY)
        fig.update_layout(barmode="group", title="Top 12 products (KL)")
        ui.chart(ui.layout(fig, 420, hover="closest"))

    # ------------------------------------------------------------ cross-sell
    ui.section("Cross-sell gaps", "share of active counters buying each product group — and the biggest counters not buying it")
    pen = B.penetration(bz, ws["cy"], B.FOCUS_GROUPS)
    if not pen.empty:
        a, b = st.columns([2, 3])
        with a:
            p_ = pen.iloc[::-1]
            fig = go.Figure(go.Bar(y=p_.index, x=p_["% of counters"], orientation="h", marker_color=ui.TEAL,
                                   text=[f"{v:.0f}%  ({int(n):,})" for v, n in zip(p_["% of counters"], p_.buyers)], textposition="outside"))
            fig.update_layout(title="Penetration: % of active counters buying", xaxis=dict(ticksuffix="%", range=[0, 115]))
            ui.chart(ui.layout(fig, max(320, 60 + 30 * len(p_)), legend=False, hover="closest"))
        with b:
            opts = [g for g in B.FOCUS_GROUPS if g in pen.index] + [g for g in pen.index if g not in B.FOCUS_GROUPS]
            grp = st.selectbox("Counters NOT buying", opts, index=opts.index("PCMO") if "PCMO" in opts else 0, key="bz_xsell")
            xs = B.cross_sell(bz, ws["cy"], grp)
            if xs.empty: st.success("Every active counter already buys this group.")
            else:
                _num(xs.drop(columns=["Sub-group"], errors="ignore").head(25), show)
                _csv(xs, f"Bazaar_crosssell_{grp}.csv")

    # ------------------------------------------------------------ top counters + 360
    ui.section("Top counters", f"{ws['cy'].label} vs {ws['ly'].label}")
    if len(lc):
        top = B.counter_master(bz).join(lc, how="inner").sort_values("cy", ascending=False).head(25)
        top["% Growth"] = [M.growth(a_, b_) for a_, b_ in zip(top.cy, top.ly)]
        top = top.rename(columns={"cy": "CY (KL)", "ly": "LY (KL)", "delta": "Change (KL)", "status": "Status"})
        _num(top, show)

    ui.section("Counter 360°", "pick any counter to see its full history")
    master = B.counter_master(bz)
    vol = bz.groupby("customer_id").kl.sum().sort_values(ascending=False)
    ids = [i for i in vol.index if i in master.index]
    if ids:
        cid = st.selectbox("Counter", ids, format_func=lambda i: f"{master.loc[i, 'Customer']} · {master.loc[i, 'District']} · {master.loc[i, 'Stockist']}", key="bz_cid")
        h = B.counter_history(bz, cid)
        d = bz[bz.customer_id == cid]
        last = d[d.kl > 0].date.max()
        ui.cards([
            ui.kpi("Lifetime volume", f"{d.kl.sum():,.2f}", f"since {d.date.min():%b-%Y}", unit="KL", tone="plain"),
            ui.kpi("Last bill", f"{last:%d-%b-%y}" if pd.notna(last) else "–", f"{(sec_last - last).days} days ago" if pd.notna(last) else "", tone="plain"),
            ui.kpi("Products bought", f"{d['product'].nunique()}", f"{d.product_group.nunique()} product groups", tone="plain"),
            ui.kpi("SSF / Sub-group", f"<span style='font-size:.6em'>{html_escape(str(master.loc[cid, 'SSF']).title())}</span>", html_escape(master.loc[cid, "Sub-group"]), tone="plain"),
        ])
        a, b = st.columns([3, 2])
        with a:
            fig = go.Figure(go.Bar(x=[mi_name(int(m)) for m in h.index], y=h.kl * 1000, marker_color=ui.NAVY))
            fig.update_layout(title="Monthly litres")
            ui.chart(ui.layout(fig, 300, legend=False, hover="closest"))
        with b:
            pg_ = d.groupby("product_group").kl.sum().sort_values(ascending=False)
            fig = go.Figure(go.Pie(labels=pg_.index, values=pg_.clip(lower=0), hole=.55))
            fig.update_layout(title="Product groups (lifetime)")
            ui.chart(ui.layout(fig, 300, hover="closest"))


def html_escape(x) -> str:
    import html
    return html.escape(str(x))
