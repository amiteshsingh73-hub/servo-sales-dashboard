"""'Premium & customers' tab: secondary sales with segment / period / customer filters, Premium CY vs LY and premium grades billed."""
from __future__ import annotations
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from . import config as C, premium as P, ui
from .metrics import growth
from .periods import mi_name


def _ok(x): return x is not None and x == x


def render(sec_all: pd.DataFrame, sec_last, default_window, show):
    if sec_all.empty:
        st.info("No secondary sales loaded."); return
    st.caption("Secondary sales (stockist to customer) from the Secondary sales file, ClearBlue excluded. "
               "Premium = products in the 'Premium' category of the Material master. Filters on this tab work on their own (the sidebar does not apply).")
    mis = sorted(int(m) for m in sec_all.mi.unique())
    d_from = default_window.start.year * 12 + default_window.start.month - 1
    d_to = default_window.end.year * 12 + default_window.end.month - 1
    d_from, d_to = (d_from if d_from in mis else mis[0]), (d_to if d_to in mis else mis[-1])

    # ------------------------------------------------------------ filters
    with st.container(border=True):
        c1, c2 = st.columns([2, 3])
        a, b = c1.select_slider("Period", options=mis, value=(d_from, d_to), format_func=lambda m: mi_name(m), key="pr_period",
                                help="Compared with the same months last year. Starts on the sidebar period.")
        segs = [s for s in sec_all.segment.dropna().astype(str).value_counts().index]
        seg = c2.multiselect("Segment", segs, placeholder="All segments", key="pr_seg")
        c3, c4, c5 = st.columns([1.2, 1.2, 3])
        stk = c3.multiselect("Stockist (SSR)", sorted(sec_all.stockist.unique()), placeholder="All", key="pr_stk")
        chn = c4.multiselect("Channel", [c for c in C.CHANNEL_ORDER if c in set(sec_all.channel)], placeholder="All", key="pr_chn")
        base = sec_all
        if stk: base = base[base.stockist.isin(stk)]
        if chn: base = base[base.channel.isin(chn)]
        names = base.sort_values("date").groupby("customer_id").agg(n=("customer", "last"), d=("district", "last"), kl=("kl", "sum")).sort_values("kl", ascending=False)
        cust = c5.multiselect("Customer name", list(names.index), format_func=lambda i: f"{names.loc[i, 'n']} · {names.loc[i, 'd']}",
                              placeholder="All customers (type to search)", key="pr_cust")
    d = base
    if seg: d = d[d.segment.astype(str).isin(seg)]
    if cust: d = d[d.customer_id.isin(cust)]
    w = P.windows_for(a, b, sec_last)
    if w["cy"].partial:
        st.info(f"Data runs till {sec_last:%d-%b-%Y}: LY is cut at the same date (like-for-like).")

    k = P.kpis(d, w); cy, ly = k["cy"], k["ly"]
    ui.cards([
        ui.kpi(f"Total sales · {w['cy'].label}", f"{cy['total']:,.1f}", f"LY {ly['total']:,.1f} KL ({w['ly'].label})", growth(cy["total"], ly["total"]), "KL"),
        ui.kpi("Premium sales", f"{cy['prem']:,.1f}", f"LY {ly['prem']:,.1f} KL", growth(cy["prem"], ly["prem"]), "KL"),
        ui.kpi("Premium share", f"{cy['share']:.1f}%" if _ok(cy["share"]) else "–", f"LY {ly['share']:.1f}%" if _ok(ly["share"]) else "", tone="plain"),
        ui.kpi("Premium grades billed", f"{cy['grades']}", f"LY {ly['grades']} · {cy['skus']} SKUs / packs CY", growth(cy["grades"], ly["grades"])),
        ui.kpi("Customers buying premium", f"{cy['prem_cust']:,}", f"of {cy['cust']:,} billed · LY {ly['prem_cust']:,}", growth(cy["prem_cust"], ly["prem_cust"])),
    ])

    # ------------------------------------------------------------ segment
    ui.section("Segment-wise", "total volume and Premium, CY vs LY")
    sg = P.by_segment(d, w)
    if sg.empty:
        st.info("No sales for these filters."); return
    body = sg[sg.Segment != "Total"]
    a_, b_ = st.columns([3, 2])
    with a_:
        fig = go.Figure()
        fig.add_bar(x=body.Segment, y=body["Total LY (KL)"], name="Total LY", marker_color=ui.GREY)
        fig.add_bar(x=body.Segment, y=body["Total CY (KL)"], name="Total CY", marker_color=ui.NAVY,
                    text=[f"{g:+.0f}%" if _ok(g) and abs(g) < 999 else "" for g in body["Total % Growth"]], textposition="outside")
        fig.add_bar(x=body.Segment, y=body["Premium LY (KL)"], name="Premium LY", marker_color="#F6C08E")
        fig.add_bar(x=body.Segment, y=body["Premium CY (KL)"], name="Premium CY", marker_color=ui.ORANGE)
        fig.update_layout(barmode="group", title="Total vs Premium volume by segment (KL)", xaxis_tickangle=-25)
        ui.chart(ui.layout(fig, 400))
    with b_:
        pg = P.premium_grades(d, w)
        if len(pg):
            top = pg.head(12).iloc[::-1]
            fig = go.Figure()
            fig.add_bar(y=top["Premium grade"], x=top["LY (KL)"], name="LY", orientation="h", marker_color=ui.GREY)
            fig.add_bar(y=top["Premium grade"], x=top["CY (KL)"], name="CY", orientation="h", marker_color=ui.ORANGE)
            fig.update_layout(barmode="group", title="Premium grades billed (KL)")
            ui.chart(ui.layout(fig, 400, hover="closest"))
        else:
            st.info("No premium grade billed in this selection.")
    show(sg)

    # ------------------------------------------------------------ customers
    ui.section("Customer-wise", "total, Premium CY vs LY, premium grades billed and segment volumes")
    cu = P.by_customer(d, w)
    q = st.radio("Show", ["All customers", "Bought premium this period", "Bought premium LY but not this period", "Never bought premium (CY or LY)"],
                 horizontal=True, key="pr_q")
    if q == "Bought premium this period": cu = cu[cu["Premium CY (KL)"] > 0]
    elif q.startswith("Bought premium LY"): cu = cu[(cu["Premium LY (KL)"] > 0) & (cu["Premium CY (KL)"] <= 0)]
    elif q.startswith("Never"): cu = cu[(cu["Premium LY (KL)"] <= 0) & (cu["Premium CY (KL)"] <= 0)]
    st.caption(f"{len(cu):,} customers · total CY {cu['Total CY (KL)'].sum():,.1f} KL · premium CY {cu['Premium CY (KL)'].sum():,.1f} KL")
    view = cu.reset_index(drop=True)
    st.dataframe(view.round(2), hide_index=True, width="stretch", height=min(560, 60 + 35 * max(len(view), 1)),
                 column_config={"Premium grades (CY)": st.column_config.TextColumn(width="large")})

    if len(pg):
        ui.section("Premium grades", "every premium grade billed in the selection")
        show(pg)

    hdr = [f"Secondary sales · {w['cy'].label} vs {w['ly'].label}", f"Segment: {', '.join(seg) or 'All'}",
           f"Stockist: {', '.join(stk) or 'All'}", f"Channel: {', '.join(chn) or 'All'}",
           f"Customers: {', '.join(names.loc[c, 'n'] for c in cust) if cust else 'All'}", f"Customer list: {q}",
           f"Data till {sec_last:%d-%b-%Y}. Premium = Material master category 'Premium'. ClearBlue excluded."]
    st.download_button("⬇ Download all tables (Excel)", P.excel({"Segment": sg, "Customers": cu, "Premium grades": pg}, hdr),
                       f"Secondary Premium Analysis {w['cy'].label.replace(chr(39), '')}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="pr_dl")
