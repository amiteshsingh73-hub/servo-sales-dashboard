"""Servo sales dashboard - Primary & Secondary performance, built for HO-template PPTs.
Run:  ./run_dashboard.command   (or  ~/.local/venvs/sales_review_ho/bin/streamlit run app.py)
"""
import glob, io, time
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from hodash import config as C, loader, metrics as M, ho_tables as H, pptx_export as X, pivot as PV, scheme_ui, ui, bazaar_ui
from hodash.periods import KINDS, MONTHS, fy_label, make_period, mi_name

st.set_page_config(page_title="Servo Sales Dashboard", page_icon="🛢️", layout="wide", initial_sidebar_state="expanded")
ui.inject_css()
NAVY, ORANGE, GREY, CAT = ui.NAVY, ui.ORANGE, ui.GREY, ui.CAT


# ------------------------------------------------------------------ data (auto-refresh on new files)
@st.cache_data(show_spinner="Reading sales files…")
def get_data(sig: str):
    return loader.load_all()


sig = loader.signature()
try:
    data = get_data(sig)
except Exception as e:      # never show a raw traceback to the user
    import traceback
    st.error("The sales files could not be loaded.")
    st.markdown(f"**{type(e).__name__}:** {e}")
    st.caption("Check that the Primary / Secondary / Target files are closed in Excel, finished syncing, and still have their original column headings. "
               "The page retries automatically every 20 seconds.")
    with st.expander("Technical details"): st.code(traceback.format_exc())
    @st.fragment(run_every=20)
    def _retry():
        if loader.signature() != sig: st.cache_data.clear(); st.rerun()
    _retry(); st.stop()


@st.fragment(run_every=20)
def watcher():
    if loader.signature() != sig:
        st.cache_data.clear(); st.rerun()


watcher()

for p_ in data.info["problems"]:
    st.warning(p_)
if data.primary.empty or data.secondary.empty:
    st.error(f"Need a readable Primary and Secondary file in {C.DATA_DIR} (names starting 'Primary' / 'Secondary'). "
             f"Currently usable: Primary {'yes' if len(data.primary) else 'NO'}, Secondary {'yes' if len(data.secondary) else 'NO'}.")
    st.stop()

pri_all, sec_all = data.primary, data.secondary
pri_last, sec_last = data.info["primary_last"], data.info["secondary_last"]

# ------------------------------------------------------------------ sidebar
sb = st.sidebar
sb.markdown("### Period")
fys = sorted(set(pri_all.fy) | set(sec_all.fy), reverse=True)
latest_fy = fys[0]
fm_labels = {i + 1: m for i, m in enumerate(MONTHS)}


def _latest_fm(y):
    fm_ = [d[d.fy == y].fm.max() for d in (pri_all, sec_all) if (d.fy == y).any()]
    return int(max(fm_)) if fm_ else 12


# When new data arrives (a later month appears) jump the selectors to it, unless the user turned this off.
stamp = f"{pri_last:%Y%m%d}|{sec_last:%Y%m%d}"
follow = sb.toggle("Follow latest data", True, key="follow", help="When a new sales file brings a later month, the period selectors move to it automatically (so the PPT Studio always reflects the latest data).")
if st.session_state.get("_stamp") != stamp:
    if st.session_state.get("_stamp") is not None and follow:
        for k_ in ("fy", "month", "till", "quarter", "half", "custom"): st.session_state.pop(k_, None)
    st.session_state["_stamp"] = stamp

fy = sb.selectbox("Financial year", fys, format_func=fy_label, key="fy")
kind = sb.selectbox("View", KINDS, index=3, key="kind")
last_fm_in_fy = _latest_fm(fy)
kw = {}
if kind == "Month":
    kw["fm"] = sb.selectbox("Month", list(fm_labels), index=last_fm_in_fy - 1, format_func=fm_labels.get, key="month")
elif kind == "Quarter":
    kw["q"] = sb.selectbox("Quarter", [1, 2, 3, 4], index=(last_fm_in_fy - 1) // 3, format_func=lambda q: f"Q{q}  ({MONTHS[3*q-3]}–{MONTHS[3*q-1]})", key="quarter")
elif kind == "Half-year":
    kw["h"] = sb.selectbox("Half", [1, 2], index=(last_fm_in_fy - 1) // 6, format_func=lambda h: f"H{h}  ({'Apr–Sep' if h == 1 else 'Oct–Mar'})", key="half")
elif kind.startswith("YTD"):
    kw["fm"] = sb.selectbox("Till month", list(fm_labels), index=last_fm_in_fy - 1, format_func=fm_labels.get, key="till")
elif kind == "Custom range":
    a, b = sb.select_slider("Months", options=list(fm_labels), value=(1, last_fm_in_fy), format_func=fm_labels.get, key="custom")
    kw["fm_from"], kw["fm_to"] = a, b
period = make_period(kind, fy, **kw)
sb.caption(f"Comparing with **LY** ({period.label(-12)}) and previous period ({period.label(-period.n_months)}).")

sb.markdown("### Filters")
SLICERS = [("Stockist", "stockist", False), ("Segment (MAT-1)", "segment", False), ("Category", "category", False),
           ("Product group", "product_group", False), ("Pack group", "pack_group", False),
           ("Channel", "channel", True), ("Customer sub-group", "cust_subgroup", True), ("District", "district", True)]
stockists = sorted(set(pri_all.stockist) | set(sec_all.stockist))
sel_st = sb.multiselect("Stockist", stockists, placeholder="All stockists", key="sl_stockist")
areas = sorted(set(sec_all.sales_area.dropna()))
sel_area = sb.multiselect("Sales area", areas, placeholder="All areas") if len(areas) > 1 else []
with sb.expander("Slice by product / channel (all tabs)", expanded=any(st.session_state.get(f"sl_{c}") for _, c, _ in SLICERS[1:])):
    for lbl_, col_, sec_only_ in SLICERS[1:]:
        opts_ = sorted(set(sec_all[col_].dropna().astype(str)) | (set() if sec_only_ else set(pri_all[col_].dropna().astype(str))))
        st.multiselect(lbl_ + (" (secondary only)" if sec_only_ else ""), opts_, placeholder="All", key=f"sl_{col_}")
apportion = sb.toggle("Apportion targets on LY mix (non-stockist tables)", True)


def filt(df, is_secondary, skip=()):
    if sel_st: df = df[df.stockist.isin(sel_st)]
    if sel_area and "sales_area" in df: df = df[df.sales_area.isin(sel_area)]
    for _, col_, sec_only_ in SLICERS[1:]:
        if col_ in skip: continue
        v_ = st.session_state.get(f"sl_{col_}")
        if v_ and (is_secondary or not sec_only_): df = df[df[col_].astype(str).isin(v_)]
    return df


pri, sec = filt(pri_all, False), filt(sec_all, True)
sliced = any(st.session_state.get(f"sl_{c}") for _, c, _ in SLICERS[1:])
tg = data.targets
if sliced:
    tg = tg.iloc[0:0]          # targets exist for whole stockists only, so achievement % is not meaningful on a product/channel slice
elif sel_st and len(tg):
    codes = pri_all[pri_all.stockist.isin(sel_st)].stockist_code.unique()
    tg = tg[tg.stockist_code.isin(codes)]
ctx = H.Ctx(pri, sec, sec_all, tg, period, pri_last, sec_last, apportion, st.session_state.get("pivot_ppt"))
wp, ws = ctx.wp, ctx.ws

sb.markdown("---")
sb.caption(f"📁 {data.info['data_dir']}")
sb.caption(f"Primary data till **{pri_last:%d-%b-%Y}**  \nSecondary data till **{sec_last:%d-%b-%Y}**")
if sb.button("🔄 Refresh now"):
    st.cache_data.clear(); st.rerun()
sb.caption("New files dropped in the folder are picked up automatically (checked every 20s).")

# ------------------------------------------------------------------ helpers
kpi = ui.kpi


def show(df: pd.DataFrame, height="content"):
    """Display a HO table with the same number formatting the PPT uses."""
    d = df.copy()
    for c in d.columns[1:]:
        if d[c].dtype.kind in "fi": d[c] = [X.fmt(c, v) for v in d[c]]
    st.dataframe(d, hide_index=True, width="stretch", height=height)


layout = ui.layout


def trend_chart(df, title, target=None):
    cy, ly = M.monthly(df, fy), M.monthly(df, fy - 1)
    in_p = [m in period.fms for m in range(1, 13)]
    fig = go.Figure()
    fig.add_bar(x=MONTHS, y=ly.values, name=fy_label(fy - 1, True), marker_color=GREY)
    fig.add_bar(x=MONTHS, y=cy.values, name=fy_label(fy, True), marker_color=[NAVY if i else "#6f7fb5" for i in in_p])
    if target is not None and len(target):
        fig.add_scatter(x=MONTHS, y=target.reindex(range(1, 13)).values, name="Target", mode="lines+markers", line=dict(color=ORANGE, width=2))
    fig.update_layout(barmode="group", title=title, bargap=.25)
    return layout(fig)


def cmp_bar(t, title, h=360, horizontal=False, w=None):
    w = w or ws
    fig = go.Figure()
    if horizontal:
        t = t.iloc[::-1]
        fig.add_bar(y=t.index, x=t.ly, name=f"LY {w['ly'].label}", marker_color=GREY, orientation="h")
        fig.add_bar(y=t.index, x=t.cy, name=f"CY {w['cy'].label}", marker_color=NAVY, orientation="h",
                    text=[f"{v:,.0f} ({g:+.0f}%)" if not np.isnan(g) and abs(g) < 999 else f"{v:,.0f}" for v, g in zip(t.cy, t.growth)], textposition="outside")
    else:
        fig.add_bar(x=t.index, y=t.ly, name=f"LY {w['ly'].label}", marker_color=GREY)
        fig.add_bar(x=t.index, y=t.cy, name=f"CY {w['cy'].label}", marker_color=NAVY,
                    text=[f"{g:+.0f}%" if not np.isnan(g) and abs(g) < 999 else "" for g in t.growth], textposition="outside")
    fig.update_layout(barmode="group", title=title, bargap=.25)
    return layout(fig, h)


# ------------------------------------------------------------------ header
ui.hero("Servo Sales Performance", period.title(),
        [", ".join(sel_st) if sel_st else "All stockists", f"Primary till {pri_last:%d-%b-%y}", f"Secondary till {sec_last:%d-%b-%y}"])
_act = [(l_, st.session_state.get(f"sl_{c_}")) for l_, c_, _ in SLICERS if st.session_state.get(f"sl_{c_}")]
if _act:
    def _clear_slicers():
        for _, c_, _ in SLICERS: st.session_state[f"sl_{c_}"] = []
    ca, cb = st.columns([6, 1])
    ca.markdown(" ".join(f"<span class='slicechip'>"
                         f"<b>{l_}</b>: {', '.join(v_[:4])}{' +' + str(len(v_) - 4) if len(v_) > 4 else ''}</span>" for l_, v_ in _act), unsafe_allow_html=True)
    cb.button("Clear all slices", on_click=_clear_slicers)
    if sliced:
        st.caption("Slices on product / channel are active: targets and achievement % are hidden (targets exist for whole stockists only). Channel / district / sub-group slices apply to secondary only.")
if wp["cy"].partial or ws["cy"].partial:
    st.info(f"Data runs till {min(pri_last, sec_last):%d-%b-%Y}: the latest month is in progress, so LY is compared like-for-like up to the same date.")

for w_ in data.info["target_warnings"]:
    st.warning("Target file: " + w_)

tab_ov, tab_bz, tab_pri, tab_sec, tab_gap, tab_ins, tab_ex, tab_pv, tab_sch, tab_ppt, tab_data = st.tabs(
    ["🏠 Overview", "🛒 Bazaar", "🏭 Primary", "🚚 Secondary", "⚖️ Primary vs Secondary", "💡 Insights", "🔎 Explore",
     "🧮 Pivot", "🎯 Schemes", "📑 PPT Studio", "🗂️ Data & checks"])

with tab_bz, ui.guard("Bazaar analysis"):
    bz_ = filt(sec_all[sec_all.channel == "Bazaar"], True, skip=("channel",))
    bazaar_ui.render(bz_, ws, period, fy, tg, sorted(bz_.stockist_code.unique()) if len(bz_) else [], sec_last, show)

# ================================================================== OVERVIEW
with tab_ov, ui.guard("Overview"):
    pc, pl, pp = M.kl(pri, wp["cy"]), M.kl(pri, wp["ly"]), M.kl(pri, wp["prev"])
    sc, sl, sp = M.kl(sec, ws["cy"]), M.kl(sec, ws["ly"]), M.kl(sec, ws["prev"])
    pcodes, scodes = sorted(pri.stockist_code.unique()), sorted(sec.stockist_code.unique())
    tp = M.target_for(tg, pcodes if len(tg) else None, fy, period.fms)
    ts_ = M.target_for(tg, scodes, fy, period.fms, "Secondary")
    k = M.customer_kpis(sec, sec_all, ws, period)
    mp_, ms_ = M.monthly(pri, fy), M.monthly(sec, fy)
    ok_ = lambda x: x is not None and x == x

    ui.cards([
        kpi("Primary sales", f"{pc:,.0f}", f"LY {pl:,.0f} · prev {pp:,.0f}", M.growth(pc, pl), "KL", mp_.dropna().tolist()),
        kpi("Secondary sales", f"{sc:,.0f}", f"LY {sl:,.0f} · prev {sp:,.0f}", M.growth(sc, sl), "KL", ms_.dropna().tolist()),
        kpi("Active counters", f"{k['active']:,}", f"LY {k['active_ly']:,}", M.growth(k["active"], k["active_ly"])),
        kpi("Primary ÷ Secondary", f"{pc / sc:.2f}x" if sc else "–", f"{pc - sc:+,.0f} KL stock build at stockists" if sc else "", tone="plain"),
    ])
    ui.section("Target tracker", f"pro-rata target for {period.label()}")
    ch = M.compare(sec, "channel", ws, order=C.CHANNEL_ORDER)
    tr_ = [ui.progress("Primary", pc, tp), ui.progress("Secondary", sc, ts_)]
    for name in C.CHANNEL_ORDER[:3]:
        cy_ = float(ch.loc[name, "cy"]) if name in ch.index else 0.0
        tr_.append(ui.progress(f"{name} (sec.)", cy_, M.target_for(tg, scodes, fy, period.fms, "Secondary", name)))
    ui.cards(tr_)

    ui.section("Channels", "secondary volume, contribution and growth vs LY")
    items = []
    for name in C.CHANNEL_ORDER:
        r = ch.loc[name] if name in ch.index else None
        items.append(kpi(name, f"{r.cy:,.0f}" if r is not None else "0", (f"{r.share:.0f}% contr. · LY {r.ly:,.0f}") if r is not None else "",
                         r.growth if r is not None and ok_(r.growth) and abs(r.growth) < 999 else None, "KL",
                         M.monthly(sec[sec.channel == name], fy).dropna().tolist()))
    items.append(kpi("Bazaar throughput", f"{k['pcs']:.0f}" if ok_(k["pcs"]) else "–",
                     f"LY {k['pcs_ly']:.0f} L" if ok_(k["pcs_ly"]) else "", M.growth(k["pcs"], k["pcs_ly"]), "L / counter / mo"))
    ui.cards(items)
    st.write("")
    prem = M.compare(pri, "category", wp)
    pr = prem.loc["Premium"] if "Premium" in prem.index else None
    ui.cards([
        kpi("New counters", f"{k['new']:,}" if k["new"] is not None else "n/a", f"{k['new_kl']:,.0f} KL" if k["new_kl"] is not None else "needs 1 yr history", tone="plain"),
        kpi("LY counters not billed", f"{k['lost']:,}", f"churn loss {k['lost_kl']:,.0f} KL", tone="warn" if k["lost"] else "plain"),
        kpi("Bazaar counters", f"{k['bazaar_counters']:,}", "active in period", tone="plain"),
        kpi("RO outlets", f"{k['ro_outlets']:,}", f"active · {k['ro_ppt']:.0f} L / outlet / mo" if ok_(k["ro_ppt"]) else "active", tone="plain"),
        kpi("Premium grade (primary)", f"{pr.cy:,.0f}" if pr is not None else "0", f"{pr.share:.1f}% of primary" if pr is not None else "",
            pr.growth if pr is not None and ok_(pr.growth) else None, "KL"),
    ])

    ui.section("Highlights", "auto-written from the numbers — ready for your PPT")
    hl = H.highlights(ctx)
    ui.highlights(hl)
    with st.expander("Copy as plain text"):
        st.code("\n".join("• " + h for h in hl), language=None)

    ui.section("Trends & mix")
    tgt_m = M.monthly_target(tg, fy, "Primary", None, pcodes)
    tgt_s = M.monthly_target(tg, fy, "Secondary", None, scodes)
    a, b = st.columns(2)
    with a: ui.chart(trend_chart(pri, "Primary sales by month (KL)", tgt_m))
    with b: ui.chart(trend_chart(sec, "Secondary sales by month (KL)", tgt_s))
    a, b = st.columns(2)
    with a: ui.chart(cmp_bar(M.compare(pri, "product_group", wp, order=C.PRODUCT_GROUP_ORDER), "Primary by product group (KL, growth % above)", w=wp))
    with b: ui.chart(cmp_bar(M.compare(sec, "channel", ws, order=C.CHANNEL_ORDER), "Secondary by channel (KL)"))

# ================================================================== PRIMARY
with tab_pri, ui.guard("Primary"):
    st.caption("Primary = billing from company to stockists. Tables follow the HO template columns.")
    for title, fn in (("Stockist wise", H.stockist_table), ("Product segment (Auto / Industrial / Co-branded)", H.segment_table),
                      ("Product group", H.product_group_table), ("Pack wise", H.pack_table),
                      ("Product category", H.category_table), ("Premium grade by stockist", H.premium_stockist_table)):
        ui.section(title)
        show(fn(ctx))
    ui.section("Top products")
    tp_ = M.compare(pri, "product", wp).head(15)
    ui.chart(cmp_bar(tp_, "Top 15 products — primary (KL)", 520, True, w=wp))

# ================================================================== SECONDARY
with tab_sec, ui.guard("Secondary"):
    st.caption("Secondary = stockist to counter (SDMS/Siebel). ClearBlue (DEF) excluded; ServoXpert counted in Bazaar; Dealer = RO.")
    ui.section("Channel")
    show(H.sec_channel_table(ctx))
    ui.section("Stockist wise")
    show(H.sec_stockist_table(ctx))
    for chn in C.CHANNEL_ORDER[:3]:
        ui.section(f"{chn} by stockist")
        show(H.sec_channel_stockist_table(ctx, chn))
    a, b = st.columns(2)
    sg = M.compare(sec, "segment", ws)
    a.plotly_chart(cmp_bar(sg, "Secondary by segment (KL)"), width="stretch")
    pr_ = M.compare(sec, "product", ws).head(12)
    fig = go.Figure(go.Pie(labels=list(pr_.index) + ["Others"], values=list(pr_.cy) + [M.compare(sec, "product", ws).cy.iloc[12:].sum()], hole=.45, marker_colors=None))
    b.plotly_chart(layout(fig.update_layout(title="Secondary product mix (CY)"), 360), width="stretch")
    a, b = st.columns(2)
    dist = M.compare(sec, "district", ws).head(12)
    a.plotly_chart(cmp_bar(dist, "Top districts (KL)", 420, True), width="stretch")
    sub = M.compare(sec, "cust_subgroup", ws).head(12)
    b.plotly_chart(cmp_bar(sub, "Customer sub-group (KL)", 420, True), width="stretch")
    ui.section("Top counters")
    cy = ws["cy"]; dd = sec[(sec.date >= cy.start) & (sec.date <= cy.end)]
    ly_ = sec[(sec.date >= ws["ly"].start) & (sec.date <= ws["ly"].end)].groupby("customer_id").kl.sum()
    top = dd.groupby(["customer_id", "customer", "channel", "stockist"]).kl.sum().reset_index().sort_values("kl", ascending=False).head(25)
    top["LY (KL)"] = top.customer_id.map(ly_).fillna(0)
    top["% Growth"] = [M.growth(a_, b_) for a_, b_ in zip(top.kl, top["LY (KL)"])]
    top = top.rename(columns={"kl": "CY (KL)", "customer": "Customer", "channel": "Channel", "stockist": "Stockist"}).drop(columns="customer_id")
    show(top)

# ================================================================== PRIMARY vs SECONDARY
with tab_gap, ui.guard("Primary vs Secondary"):
    cyp = pri[(pri.date >= wp["cy"].start) & (pri.date <= wp["cy"].end)].groupby("stockist").kl.sum()
    cys = sec[(sec.date >= ws["cy"].start) & (sec.date <= ws["cy"].end)].groupby("stockist").kl.sum()
    g = pd.DataFrame({"Primary": cyp, "Secondary": cys}).fillna(0)
    g["Ratio (Pri/Sec)"] = g.Primary / g.Secondary.replace(0, np.nan)
    g["Stock build (KL)"] = g.Primary - g.Secondary
    g = g.sort_values("Primary", ascending=False)
    fig = go.Figure()
    fig.add_bar(x=g.index, y=g.Primary, name="Primary", marker_color=NAVY)
    fig.add_bar(x=g.index, y=g.Secondary, name="Secondary", marker_color=ORANGE)
    fig.update_layout(barmode="group", title=f"Primary vs Secondary — {wp['cy'].label}")
    a, b = st.columns([3, 2])
    a.plotly_chart(layout(fig), width="stretch")
    with b:
        ui.section("Stock build / (release) at stockists")
        gg = g.reset_index().rename(columns={"stockist": "Stockist", "index": "Stockist"})
        show(gg)
    mp, ms = M.monthly(pri, fy), M.monthly(sec, fy)
    fig = go.Figure()
    fig.add_scatter(x=MONTHS, y=mp.cumsum().values, name="Primary (cum.)", line=dict(color=NAVY, width=3))
    fig.add_scatter(x=MONTHS, y=ms.cumsum().values, name="Secondary (cum.)", line=dict(color=ORANGE, width=3))
    fig.update_layout(title=f"Cumulative {fy_label(fy)} — secondary offtake vs primary billing (KL)")
    st.plotly_chart(layout(fig, 340), width="stretch")

# ================================================================== INSIGHTS
with tab_ins, ui.guard("Insights"):
    # ---- 1. growth bridge
    ui.section("What drove the change vs LY?")
    a_, b_, c_ = st.columns([2, 2, 3])
    ds = a_.radio("Sales", ["Primary", "Secondary"], horizontal=True, key="br_ds")
    dims = {"Stockist": "stockist", "Product group": "product_group", "Category": "category", "Pack group": "pack_group"}
    if ds == "Secondary": dims = {"Channel": "channel", **dims, "Customer sub-group": "cust_subgroup", "District": "district"}
    dname = b_.selectbox("Split by", list(dims), key="br_dim")
    bdf, bw = (pri, wp) if ds == "Primary" else (sec, ws)
    br = M.bridge(bdf, dims[dname], bw)
    br = pd.concat([br[br.delta > 0], br[br.delta <= 0].sort_values("delta")]) if len(br) <= 14 else pd.concat([br.head(7), br.tail(7)])
    tot_ly, tot_cy = M.kl(bdf, bw["ly"]), M.kl(bdf, bw["cy"])
    other = (tot_cy - tot_ly) - br.delta.sum()
    xs = [f"LY {bw['ly'].label}"] + list(br.index) + (["Others"] if abs(other) > 0.5 else []) + [f"CY {bw['cy'].label}"]
    ys = [tot_ly] + list(br.delta) + ([other] if abs(other) > 0.5 else []) + [tot_cy]
    ms = ["absolute"] + ["relative"] * (len(xs) - 2) + ["total"]
    fig = go.Figure(go.Waterfall(x=xs, y=ys, measure=ms, text=[f"{v:,.0f}" if m != "relative" else f"{v:+,.0f}" for v, m in zip(ys, ms)],
                                 textposition="outside", increasing=dict(marker_color=NAVY), decreasing=dict(marker_color=ORANGE),
                                 totals=dict(marker_color=GREY), connector=dict(line=dict(color="rgba(128,128,128,.4)"))))
    fig.update_layout(title=f"{ds} volume bridge by {dname.lower()} (KL)", showlegend=False)
    c_.caption("Bars show each member's contribution to the net change; the sum of the bars takes LY to CY.")
    st.plotly_chart(layout(fig, 420, legend=False), width="stretch")

    # ---- 2. run-rate to target
    ui.section("Run-rate vs target (to date, selected FY)")
    if len(tg):
        for nm, df_, kind_, last_ in (("Primary", pri, "Primary", pri_last), ("Secondary", sec, "Secondary", sec_last)):
            rr = M.run_rate(df_, tg, fy, period.fm_end, last_, kind=kind_)
            if len(rr):
                st.markdown(f"**{nm}**")
                show(rr)
        st.caption("Required / month = balance to annual target ÷ months left. Projected FY = YTD + current monthly run-rate × months left.")
    else:
        st.info("No targets recognised - keep the Target workbook in the data folder to see run-rate analysis.")

    # ---- 3. heatmap
    ui.section("Monthly achievement vs target (%)")
    hc = st.columns(2)
    for col, nm, df_, kind_ in ((hc[0], "Primary", pri, "Primary"), (hc[1], "Secondary", sec, "Secondary")):
        hm = M.month_heat(df_, tg, fy, kind_)
        if hm.empty: col.info("No targets"); continue
        z = hm.values
        fig = go.Figure(go.Heatmap(z=z, x=MONTHS, y=hm.index, zmin=50, zmax=150, zmid=100,
                                   colorscale=[[0, ORANGE], [.5, "#E9EBF5"], [1, NAVY]], colorbar=dict(title="%", thickness=10),
                                   text=[[f"{v:.0f}" if v == v else "" for v in r] for r in z], texttemplate="%{text}", hovertemplate="%{y} %{x}: %{z:.0f}%<extra></extra>"))
        col.plotly_chart(layout(fig.update_layout(title=f"{nm} — actual ÷ target", hovermode="closest"), 260, legend=False), width="stretch")

    # ---- 4. counters
    ui.section("Counter movers (secondary)")
    cm = M.counter_moves(sec, ws)
    cm = cm[cm.Channel.isin(["Bazaar", "RO", "OEM", "Fleet"])] if len(cm) else cm
    ch_f = st.multiselect("Channel", C.CHANNEL_ORDER, default=C.CHANNEL_ORDER, key="cm_ch")
    cm = cm[cm.Channel.isin(ch_f)]
    fmtc = lambda d: show(d.rename(columns={"cy": "CY (KL)", "ly": "LY (KL)", "delta": "Change (KL)"})[["Customer", "Channel", "Stockist", "LY (KL)", "CY (KL)", "Change (KL)"]])
    a_, b_ = st.columns(2)
    with a_:
        st.markdown("**Top gainers**"); fmtc(cm[cm.ly > 0].sort_values("delta", ascending=False).head(10))
        st.markdown("**New counters (no LY sales)**"); fmtc(cm[(cm.ly <= 0) & (cm.cy > 0)].sort_values("cy", ascending=False).head(10))
    with b_:
        st.markdown("**Top decliners**"); fmtc(cm[(cm.cy > 0) & (cm.ly > 0)].sort_values("delta").head(10))
        st.markdown("**Lost counters (sold LY, nothing now)**"); fmtc(cm[(cm.cy <= 0) & (cm.ly > 0)].sort_values("ly", ascending=False).head(10))
    n_new, n_lost = int(((cm.ly <= 0) & (cm.cy > 0)).sum()), int(((cm.cy <= 0) & (cm.ly > 0)).sum())
    st.caption(f"{n_new:,} new and {n_lost:,} lost counters in the selected channels (lost = bought in the LY window, not in this one).")

    # ---- 5. concentration
    a_, b_ = st.columns(2)
    cy_c = cm[cm.cy > 0].sort_values("cy", ascending=False)
    if len(cy_c):
        cum = cy_c.cy.cumsum() / cy_c.cy.sum() * 100
        x = np.arange(1, len(cum) + 1) / len(cum) * 100
        fig = go.Figure(go.Scatter(x=x, y=cum.values, mode="lines", line=dict(color=NAVY, width=3), fill="tozeroy", fillcolor="rgba(11,31,102,.08)"))
        fig.update_layout(title="Counter concentration (Pareto)", xaxis_title="% of counters (largest first)", yaxis_title="% of volume", hovermode="closest")
        a_.plotly_chart(layout(fig, 340, legend=False), width="stretch")
        top10 = cy_c.head(10).cy.sum() / cy_c.cy.sum() * 100
        n20 = int(np.ceil(len(cy_c) * .2)); s20 = cy_c.head(n20).cy.sum() / cy_c.cy.sum() * 100
        b_.markdown(kpi("Top 10 counters' share", f"{top10:.0f}%", f"of {len(cy_c):,} active counters"), unsafe_allow_html=True)
        b_.write("")
        b_.markdown(kpi("Top 20% counters' share", f"{s20:.0f}%", f"{n20:,} counters"), unsafe_allow_html=True)

    # ---- 6. product movers
    ui.section("Product movers (primary, KL change vs LY)")
    pm = M.compare(pri, "product", wp)
    pm = pm.reset_index().rename(columns={"product": "Product", "ly": "LY (KL)", "cy": "CY (KL)", "delta": "Change (KL)", "growth": "% Growth"})
    a_, b_ = st.columns(2)
    cols_ = ["Product", "LY (KL)", "CY (KL)", "Change (KL)", "% Growth"]
    a_.markdown("**Top gainers**"); a_.dataframe(pm.sort_values("Change (KL)", ascending=False).head(10)[cols_].round(1), hide_index=True, width="stretch")
    b_.markdown("**Top decliners**"); b_.dataframe(pm.sort_values("Change (KL)").head(10)[cols_].round(1), hide_index=True, width="stretch")

    # ---- 7. realisation
    ui.section("Net realisation (secondary, ₹ per litre)")
    def realisation(df, w, dim):
        d = M.sel(df, w); g = d.groupby(dim).agg(v=("amount", "sum"), q=("kl", "sum"))
        return (g.v / (g.q * 1000)).replace([np.inf, -np.inf], np.nan)
    a_, b_ = st.columns(2)
    for col, dim, ttl in ((a_, "channel", "By channel"), (b_, "product_group", "By product group")):
        cy_r, ly_r = realisation(sec, ws["cy"], dim), realisation(sec, ws["ly"], dim)
        r = pd.DataFrame({"CY ₹/L": cy_r, "LY ₹/L": ly_r}).dropna(how="all")
        if dim == "channel": r = r.reindex([c for c in C.CHANNEL_ORDER if c in r.index])
        r["% Change"] = (r["CY ₹/L"] / r["LY ₹/L"] - 1) * 100
        fig = go.Figure()
        fig.add_bar(x=r.index, y=r["LY ₹/L"], name="LY", marker_color=GREY)
        fig.add_bar(x=r.index, y=r["CY ₹/L"], name="CY", marker_color=NAVY, text=[f"{v:+.0f}%" if v == v else "" for v in r["% Change"]], textposition="outside")
        fig.update_layout(barmode="group", title=ttl)
        col.plotly_chart(layout(fig, 320), width="stretch")
    st.caption("Taxable value ÷ litres from the secondary file (after invoice-level discounts, before GST). Mix shifts between pack sizes and grades move this number.")

# ================================================================== EXPLORE
with tab_ex, ui.guard("Explore"):
    TIMEDIMS = set(PV.TIME_DIMS)
    HIER = {"Secondary": {"Channel": "Stockist", "Stockist": "Customer sub-group", "Customer sub-group": "Customer", "District": "Customer",
                          "Segment (MAT-1)": "Product group (HO)", "Product group (HO)": "Product", "Category": "Product", "Pack group (HO)": "Product"},
            "Primary": {"Stockist": "Product group (HO)", "Segment (MAT-1)": "Product group (HO)", "Product group (HO)": "Product",
                        "Category": "Product", "Pack group (HO)": "Product"}}
    SL_COL = {"Stockist": "stockist", "Segment (MAT-1)": "segment", "Category": "category", "Product group (HO)": "product_group",
              "Pack group (HO)": "pack_group", "Channel": "channel", "Customer sub-group": "cust_subgroup", "District": "district"}
    QUICK = {
        "— quick questions —": None,
        "Where did secondary growth come from? (by channel)": dict(ds="Secondary", by="Channel", split="(none)", cmp="Compare with last year", meas="Volume (KL)"),
        "Which stockist grew / fell most? (secondary)": dict(ds="Secondary", by="Stockist", split="(none)", cmp="Compare with last year", meas="Volume (KL)"),
        "Which product groups are losing volume? (primary)": dict(ds="Primary", by="Product group (HO)", split="(none)", cmp="Compare with last year", meas="Volume (KL)"),
        "Biggest districts and their growth (secondary)": dict(ds="Secondary", by="District", split="(none)", cmp="Compare with last year", meas="Volume (KL)"),
        "Monthly trend by channel (secondary)": dict(ds="Secondary", by="Month", split="Channel", cmp="Over time", meas="Volume (KL)"),
        "Monthly primary by stockist": dict(ds="Primary", by="Month", split="Stockist", cmp="Over time", meas="Volume (KL)"),
        "Realisation ₹/L by channel (secondary)": dict(ds="Secondary", by="Channel", split="(none)", cmp="Compare with last year", meas="Realisation (₹/L)"),
    }
    for k_, v_ in (("ex_ds", "Secondary"), ("ex_meas", "Volume (KL)"), ("ex_by", "Channel"), ("ex_split", "(none)"),
                   ("ex_cmp", "Compare with last year"), ("ex_drill", []), ("ex_n", 0)):
        st.session_state.setdefault(k_, v_)
    # changes requested by chart clicks / chips are applied here, before the widgets are created
    pend = st.session_state.pop("_ex_pending", None)
    if pend: st.session_state.update(pend)

    def _quick():
        q_ = QUICK.get(st.session_state.get("ex_quick"))
        if q_:
            st.session_state.update(ex_ds=q_["ds"], ex_by=q_["by"], ex_split=q_["split"], ex_cmp=q_["cmp"], ex_meas=q_["meas"], ex_drill=[])

    def _drop_drill(i_):
        st.session_state["ex_drill"] = st.session_state["ex_drill"][:i_]; st.session_state["ex_n"] += 1

    def _make_global():
        for lbl_, val_ in st.session_state["ex_drill"]:
            c_ = SL_COL.get(lbl_)
            if c_ and val_ not in st.session_state.get(f"sl_{c_}", []): st.session_state[f"sl_{c_}"] = list(st.session_state.get(f"sl_{c_}", [])) + [val_]
        st.session_state["ex_drill"] = []; st.session_state["ex_n"] += 1

    st.markdown("#### Ask a question of your sales data")
    st.selectbox("Quick questions", list(QUICK), key="ex_quick", on_change=_quick, label_visibility="collapsed")
    c1, c2, c3, c4, c5 = st.columns([1.1, 1.5, 1.5, 1.5, 1.7])
    ds_e = c1.radio("Data", ["Primary", "Secondary"], key="ex_ds")
    dims_e = PV.dims_for(ds_e)
    meas_e = c2.selectbox("Show", PV.measures_for(ds_e), key="ex_meas") if st.session_state["ex_meas"] in PV.measures_for(ds_e) else c2.selectbox("Show", PV.measures_for(ds_e))
    by_opts = list(dims_e)
    if st.session_state["ex_by"] not in by_opts: st.session_state["ex_by"] = "Stockist"
    by_e = c3.selectbox("by", by_opts, key="ex_by")
    split_opts = ["(none)"] + [d for d in dims_e if d != by_e]
    if st.session_state["ex_split"] not in split_opts: st.session_state["ex_split"] = "(none)"
    split_e = c4.selectbox("split by", split_opts, key="ex_split")
    cmp_opts = ["Compare with last year", "Compare with previous period", "Over time (all data)"]
    if st.session_state["ex_cmp"] not in cmp_opts: st.session_state["ex_cmp"] = cmp_opts[0]
    cmp_e = c5.selectbox("Period", cmp_opts, key="ex_cmp", help="The comparison uses the period chosen in the sidebar.")

    mk, additive, _, _k = PV.MEASURES[meas_e]
    base = pri if ds_e == "Primary" else sec
    d_e = PV.prepare(base)
    for lbl_, val_ in st.session_state["ex_drill"]:
        col_ = dims_e.get(lbl_, lbl_)
        d_e = d_e[d_e[col_].astype(str) == val_]
    # drill breadcrumb
    if st.session_state["ex_drill"]:
        bc = st.columns([*([2] * len(st.session_state["ex_drill"])), 2, 3])
        for i_, (l_, v_) in enumerate(st.session_state["ex_drill"]):
            bc[i_].button(f"✕ {l_}: {v_[:22]}", key=f"ex_chip_{i_}", on_click=_drop_drill, args=(i_,))
        bc[len(st.session_state["ex_drill"])].button("Use as dashboard slice", on_click=_make_global, help="Applies these selections to every tab.")
    else:
        st.caption("Tip: click a bar to drill into it. Selections appear as chips here; click a chip's ✕ to go back.")

    w_e = wp if ds_e == "Primary" else ws
    time_by = by_e in TIMEDIMS
    mode = "trend" if (time_by or cmp_e.startswith("Over time")) else ("ly" if "last year" in cmp_e else "prev")
    by_col = dims_e[by_e] if dims_e[by_e] in d_e else by_e
    split_col = None if split_e == "(none)" else (dims_e[split_e] if dims_e[split_e] in d_e else split_e)
    for c_ in [by_col] + ([split_col] if split_col else []):
        d_e[c_] = d_e[c_].astype(object).where(d_e[c_].notna(), "-").astype(str)

    if d_e.empty:
        st.warning("No data for this slice.")
    else:
        def one(frame, keys):
            return PV._agg(frame, keys)[mk] if len(frame) else pd.Series(dtype=float)

        n_ = st.session_state["ex_n"]
        pts = []
        if mode == "trend":
            order = PV._labels(d_e, by_e, by_col, d_e[by_col].unique()) if time_by else sorted(d_e[by_col].unique())
            if split_col:
                tab_ = one(d_e, [by_col, split_col]).unstack(split_col)
                tab_ = tab_.reindex(order)
                fig = go.Figure()
                for j_, c_ in enumerate(tab_.columns):
                    fig.add_bar(x=tab_.index, y=tab_[c_], name=str(c_), marker_color=CAT[j_ % len(CAT)])
                fig.update_layout(barmode="stack" if additive else "group")
                tot_ = one(d_e, [by_col]).reindex(order)
                out = tab_.copy(); out["Total"] = tot_ if additive else np.nan
            else:
                tot_ = one(d_e, [by_col]).reindex(order)
                fig = go.Figure(go.Bar(x=tot_.index, y=tot_.values, marker_color=NAVY))
                out = tot_.to_frame(meas_e)
            fig.update_layout(title=f"{meas_e} by {by_e.lower()}" + (f", split by {split_e.lower()}" if split_col else ""), xaxis_tickangle=-30,
                              hovermode="closest" if not split_col else "x unified")
            ev = st.plotly_chart(layout(fig, 420, legend=bool(split_col), hover="closest" if not split_col else "x unified"), width="stretch", on_select="rerun", selection_mode=("points",), key=f"ex_chart_{n_}")
            pts = ev.selection.points if ev and ev.selection else []
            member_key = "x"
            kpi_cy = float(one(d_e, []).iloc[0]) if len(d_e) else 0.0
            st.markdown(kpi(f"{meas_e} (selected slice, all data)", PV.fmt_value(kpi_cy, _k), f"{len(d_e):,} invoice lines"), unsafe_allow_html=True)
            show(out.reset_index().rename(columns={by_col: by_e, "index": by_e}) if False else out.reset_index().rename(columns={out.index.name or "index": by_e}))
        else:
            wk = "ly" if mode == "ly" else "prev"
            cyf, cmpf = M.sel(d_e, w_e["cy"]), M.sel(d_e, w_e[wk])
            cy, cm = one(cyf, [by_col]), one(cmpf, [by_col])
            t = pd.DataFrame({"CY": cy, "LY" if wk == "ly" else "Prev": cm}).fillna(0.0)
            cmp_name = "LY" if wk == "ly" else "Prev"
            t["Change"] = t.CY - t[cmp_name]
            t["% Growth"] = [M.growth(a_, b_) for a_, b_ in zip(t.CY, t[cmp_name])]
            if additive: t["% Share"] = t.CY / t.CY.sum() * 100 if t.CY.sum() else 0
            t = t.sort_values("CY", ascending=False)
            tot_cy = float(one(cyf, []).iloc[0]) if len(cyf) else 0.0
            tot_cm = float(one(cmpf, []).iloc[0]) if len(cmpf) else 0.0
            k1, k2, k3 = st.columns(3)
            k1.markdown(kpi(f"{meas_e} · {w_e['cy'].label}", PV.fmt_value(tot_cy, _k), "selected slice"), unsafe_allow_html=True)
            k2.markdown(kpi(f"{cmp_name} · {w_e[wk].label}", PV.fmt_value(tot_cm, _k), ""), unsafe_allow_html=True)
            k3.markdown(kpi("Change", PV.fmt_value(tot_cy - tot_cm, _k), "", M.growth(tot_cy, tot_cm)), unsafe_allow_html=True)
            top = t.head(15).iloc[::-1]
            fig = go.Figure()
            if split_col:
                sp = one(cyf, [by_col, split_col]).unstack(split_col).reindex(top.index).fillna(0)
                for j_, c_ in enumerate(sp.columns):
                    fig.add_bar(y=sp.index, x=sp[c_], name=str(c_), orientation="h", marker_color=CAT[j_ % len(CAT)])
                fig.update_layout(barmode="stack" if additive else "group")
            else:
                fig.add_bar(y=top.index, x=top[cmp_name], name=f"{cmp_name} {w_e[wk].label}", orientation="h", marker_color=GREY)
                fig.add_bar(y=top.index, x=top.CY, name=f"CY {w_e['cy'].label}", orientation="h", marker_color=NAVY,
                            text=[(f"{PV.fmt_value(v, _k)}  ({g:+.0f}%)" if g == g and abs(g) < 999 else PV.fmt_value(v, _k)) for v, g in zip(top.CY, top["% Growth"])], textposition="outside")
                fig.update_layout(barmode="group")
            fig.update_layout(title=f"{meas_e} by {by_e.lower()} — click a bar to drill in", yaxis=dict(automargin=True), hovermode="closest")
            if not split_col: fig.update_xaxes(range=[0, float(max(top.CY.max(), top[cmp_name].max(), 1e-9)) * 1.3])
            ev = st.plotly_chart(layout(fig, max(320, 60 + 36 * len(top)), True, "closest"), width="stretch", on_select="rerun", selection_mode=("points",), key=f"ex_chart_{n_}")
            pts = ev.selection.points if ev and ev.selection else []
            member_key = "y"
            if additive and len(t) > 1:
                g_ = t[t[cmp_name] > 0]
                up, dn = t.sort_values("Change").iloc[-1], t.sort_values("Change").iloc[0]
                st.info(f"**{t.index[0]}** is the largest ({t.iloc[0].CY:,.0f}, {t.iloc[0]['% Share']:.0f}% share). "
                        f"Biggest gain: **{t.sort_values('Change').index[-1]}** ({up.Change:+,.0f}); biggest fall: **{t.sort_values('Change').index[0]}** ({dn.Change:+,.0f}).")
            out = t
            show(t.reset_index().rename(columns={t.index.name or "index": by_e}))

        # ---- click to drill
        if pts and not time_by and mode != "trend":
            member = str(pts[0].get(member_key))
            nxt = HIER[ds_e].get(by_e)
            st.session_state["_ex_pending"] = {"ex_drill": st.session_state["ex_drill"] + [(by_e, member)], "ex_n": n_ + 1,
                                                **({"ex_by": nxt} if nxt else {})}
            st.rerun()
        elif pts and (time_by or mode == "trend"):
            member = str(pts[0].get(member_key))
            st.caption(f"Selected {by_e}: {member}. Choose a non-time field in “by” to drill by clicking.")

        # ---- export / PPT
        flat = out.reset_index() if out.index.name or not isinstance(out.index, pd.RangeIndex) else out
        e1, e2, e3 = st.columns([1, 1, 4])
        e1.download_button("⬇ CSV", flat.to_csv(index=False).encode("utf-8"), f"Explore_{ds_e}.csv", "text/csv")
        txt = flat.copy()
        for c_ in txt.columns[1:]:
            if pd.api.types.is_numeric_dtype(txt[c_]): txt[c_] = [X.fmt(str(c_), v) for v in txt[c_]]
        ttl_e = f"{ds_e} – {meas_e} by {by_e}" + (f" × {split_e}" if split_col else "")
        if e3.checkbox("Use this table in PPT Studio", False, key="ex_send"):
            st.session_state["pivot_ppt"] = (ttl_e, txt.set_index(txt.columns[0])); st.session_state["pivot_ppt_src"] = "explore"
        elif st.session_state.get("pivot_ppt_src") == "explore" and "pivot_ppt" in st.session_state:
            del st.session_state["pivot_ppt"]

# ================================================================== PIVOT
with tab_pv, ui.guard("Pivot"):
    st.caption("Build your own table: pick the fields for rows and columns, the measures, and filters. Works on full history; the sidebar stockist filter also applies.")
    PRESETS = {
        "— choose a quick start —": None,
        "Stockist × Month (KL)": dict(ds="Primary", rows=["Stockist"], cols=["Month"], meas=["Volume (KL)"]),
        "Product group × Fiscal year (KL)": dict(ds="Primary", rows=["Product group (HO)"], cols=["Fiscal year"], meas=["Volume (KL)"]),
        "Category × Stockist (% of column)": dict(ds="Primary", rows=["Category"], cols=["Stockist"], meas=["Volume (KL)"], show="% of column total"),
        "Channel × Month (secondary KL)": dict(ds="Secondary", rows=["Channel"], cols=["Month"], meas=["Volume (KL)"]),
        "Stockist × Channel (KL, counters)": dict(ds="Secondary", rows=["Stockist"], cols=["Channel"], meas=["Volume (KL)", "Active counters"]),
        "District × Fiscal year (secondary KL)": dict(ds="Secondary", rows=["District"], cols=["Fiscal year"], meas=["Volume (KL)"]),
        "Top customers (KL, value, ₹/L)": dict(ds="Secondary", rows=["Customer"], cols=[], meas=["Volume (KL)", "Taxable value (₹)", "Realisation (₹/L)"], top=50),
    }

    def _preset():
        p_ = PRESETS.get(st.session_state.get("pv_preset"))
        if not p_: return
        st.session_state.update(pv_ds=p_["ds"], pv_rows=p_["rows"], pv_cols=p_["cols"], pv_meas=p_["meas"],
                                pv_show=p_.get("show", "Value"), pv_top=p_.get("top", 0))
        for k_ in [k for k in st.session_state if k.startswith("pv_f_")]: del st.session_state[k_]
        st.session_state["pv_filters"] = []

    for k_, v_ in (("pv_ds", "Primary"), ("pv_rows", ["Stockist"]), ("pv_cols", ["Month"]), ("pv_meas", ["Volume (KL)"]),
                   ("pv_show", "Value"), ("pv_pmode", "All data"), ("pv_top", 0), ("pv_filters", [])):
        st.session_state.setdefault(k_, v_)
    st.selectbox("Quick start", list(PRESETS), key="pv_preset", on_change=_preset)
    c1, c2 = st.columns([1, 3])
    ds_ = c1.radio("Data", ["Primary", "Secondary"], key="pv_ds", horizontal=True)
    base_ = pri if ds_ == "Primary" else sec
    dims_ = PV.dims_for(ds_)
    meas_opts = PV.measures_for(ds_)
    for k_, opts_ in (("pv_rows", list(dims_)), ("pv_cols", list(dims_)), ("pv_meas", meas_opts)):   # drop stale picks when switching data
        if k_ in st.session_state: st.session_state[k_] = [x for x in st.session_state[k_] if x in opts_]
    rows_ = c2.multiselect("Rows", list(dims_), key="pv_rows")
    c1, c2, c3 = st.columns(3)
    cols_ = c1.multiselect("Columns", [d for d in dims_ if d not in rows_], key="pv_cols")
    meas_ = c2.multiselect("Measures", meas_opts, key="pv_meas")
    show_ = c3.selectbox("Show values as", PV.SHOW_AS, key="pv_show")

    cols_ = [c for c in cols_ if c not in rows_]
    st.markdown("**Period**")
    a_, b_ = st.columns([1, 2])
    pmode = a_.radio("Time range", ["All data", "Sidebar period", "Pick months"], key="pv_pmode", horizontal=False)
    pv_df = base_
    if pmode == "Sidebar period":
        w_ = wp if ds_ == "Primary" else ws
        pv_df = M.sel(base_, w_["cy"]); b_.info(f"{w_['cy'].label}  ({w_['cy'].start:%d-%b-%y} to {w_['cy'].end:%d-%b-%y})")
    elif pmode == "Pick months" and len(base_):
        mis = sorted(base_.mi.unique())
        lo, hi = b_.select_slider("Months", options=mis, value=(mis[max(0, len(mis) - 12)], mis[-1]), format_func=lambda m: mi_name(int(m)), key="pv_months")
        pv_df = base_[(base_.mi >= lo) & (base_.mi <= hi)]

    fcols = st.multiselect("Add filters on…", list(dims_), key="pv_filters")
    pv_p = PV.prepare(pv_df)
    if fcols:
        fcs = st.columns(min(len(fcols), 3))
        for i_, fd in enumerate(fcols):
            colname = dims_[fd] if dims_[fd] in pv_p else fd
            opts = sorted(pv_p[colname].dropna().astype(str).unique(), key=lambda x: (PV.fy_label if False else x))
            if fd == "Month": opts = pv_p.drop_duplicates("mi").sort_values("mi")["Month"].tolist()
            sel_ = fcs[i_ % 3].multiselect(fd, opts, key=f"pv_f_{fd}", placeholder="All")
            if sel_: pv_p = pv_p[pv_p[colname].astype(str).isin(sel_)]

    o1, o2, o3, o4 = st.columns(4)
    totals_ = o1.toggle("Totals", True, key="pv_tot")
    heat_ = o2.toggle("Heat-map shading", True, key="pv_heat")
    sort_ = o3.toggle("Sort rows by total", True, key="pv_sort")
    top_ = o4.number_input("Top N rows (0 = all)", 0, 5000, step=10, key="pv_top")

    if not rows_ and not cols_:
        st.info("Pick at least one field for Rows or Columns.")
    elif not meas_:
        st.info("Pick at least one measure.")
    elif pv_p.empty:
        st.warning("No data for this selection.")
    else:
        tbl, info = PV.build(pv_p, rows_, cols_, meas_, dims_, show_, totals_, int(top_), sort_)
        if tbl.empty:
            st.warning("Nothing to show.")
        else:
            if show_ != "Value" and any(not PV.MEASURES[m][1] for m in meas_):
                st.caption("Percent views apply to additive measures only (KL, lines, value); counters, invoices and ₹/L are shown as values.")
            if info["rows_total"] > info["rows_shown"]:
                st.caption(f"Showing {info['rows_shown']:,} of {info['rows_total']:,} rows (totals still include all).")
            text = PV.to_text(tbl, info)
            if heat_:
                st.dataframe(PV.heat_style(tbl, text), width="stretch", height=min(700, 60 + 35 * len(tbl)))
            else:
                st.dataframe(text, width="stretch", height=min(700, 60 + 35 * len(tbl)))

            # quick chart
            if len(meas_) == 1 and len(tbl) > 1 and show_ in ("Value",):
                ctab = tbl.drop(index="Total", errors="ignore")
                if isinstance(ctab.index, pd.MultiIndex): ctab = ctab.drop(index="Total", level=0, errors="ignore")
                ctab = ctab.head(20)
                labels = [" · ".join(map(str, i)) if isinstance(i, tuple) else str(i) for i in ctab.index]
                fig = go.Figure()
                series = [c for c in ctab.columns if c != "Total"] or list(ctab.columns)
                for j_, c in enumerate(series[:12]):
                    fig.add_bar(x=labels, y=ctab[c], name=str(c if not isinstance(c, tuple) else c[-1]), marker_color=CAT[j_ % len(CAT)])
                fig.update_layout(barmode="stack" if len(series) > 1 else "group", title=f"{meas_[0]} — {' / '.join(rows_) or 'total'}", xaxis_tickangle=-30)
                with st.expander("Chart", expanded=False):
                    st.plotly_chart(layout(fig, 400, legend=len(series) > 1), width="stretch")

            # exports
            buf_ = io.BytesIO()
            flat = tbl.copy()
            if isinstance(flat.columns, pd.MultiIndex): flat.columns = [" | ".join(map(str, c)) for c in flat.columns]
            with pd.ExcelWriter(buf_, engine="openpyxl") as xw:
                flat.to_excel(xw, sheet_name="Pivot")
            e1, e2, e3 = st.columns([1, 1, 3])
            e1.download_button("⬇ Excel", buf_.getvalue(), f"Pivot_{ds_}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            e2.download_button("⬇ CSV", flat.to_csv().encode("utf-8"), f"Pivot_{ds_}.csv", "text/csv")
            auto_t = f"{ds_} – {' × '.join(rows_) or 'Total'}{' by ' + ' × '.join(cols_) if cols_ else ''}"
            ttl = e3.text_input("Slide title (to use this table in PPT Studio)", auto_t, key=f"pv_title_{abs(hash(auto_t))}")
            if e3.checkbox("Send this pivot to PPT Studio", False, key="pv_send"):
                st.session_state["pivot_ppt"] = (ttl, text); st.session_state["pivot_ppt_src"] = "pivot"
                e3.caption("Saved - tick 'Custom pivot' in PPT Studio.")
            elif st.session_state.get("pivot_ppt_src") == "pivot" and "pivot_ppt" in st.session_state:
                del st.session_state["pivot_ppt"]

# ================================================================== SCHEMES
with tab_sch:
    try:
        scheme_ui.render(data, pri_all, sec_all, pri_last, sec_last)
    except Exception as e_:      # never break the rest of the dashboard
        import traceback
        st.error(f"The scheme module hit a problem: {type(e_).__name__}: {e_}")
        with st.expander("Technical details"): st.code(traceback.format_exc())

# ================================================================== PPT STUDIO
with tab_ppt, ui.guard("PPT Studio"):
    st.markdown("Pick the slides, then download a deck built on the **HO standard template** (cover kept, tables & charts are native/editable). "
                "Everything follows the period & filters selected in the sidebar.")
    st.success(f"Live data: Primary till **{pri_last:%d-%b-%Y}**, Secondary till **{sec_last:%d-%b-%Y}** · "
               f"files: {', '.join(data.info['files']['primary'] + data.info['files']['secondary'])}. "
               f"Slides rebuild automatically whenever a sales file changes.")
    tpl = glob.glob(str(C.DATA_DIR / C.HO_TEMPLATE_GLOB))
    if not tpl: st.warning("HO template not found in the data folder - a plain 16:9 deck will be generated.")
    slide_opts = [k for k in X.SLIDES if k != "pivot" or ctx.pivot]
    keys = st.multiselect("Slides", slide_opts, default=[k for k in X.DEFAULT_SLIDES if k in slide_opts] + (["pivot"] if ctx.pivot else []), format_func=lambda k: X.SLIDES[k][0])
    cover = st.checkbox("Keep template cover slide", True)
    c1, c2 = st.columns(2)
    if keys:
        fname = f"Sales_Review_{period.label().replace(chr(39), '')}_{fy_label(fy, True)}"
        try:
            deck = X.build_deck(ctx, keys, tpl[0] if tpl else None, cover)
            c1.download_button("⬇ Download PPTX (HO template)", deck, f"{fname}.pptx",
                               "application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary")
            if st.toggle("Also keep a PPTX file in the exports folder updated automatically", False, key="autosave",
                         help="Rewrites sales_dashboard/exports/Latest_Sales_Review.pptx every time the data or your selections change (works while the dashboard is running)."):
                out = C.APP_DIR / "exports"; out.mkdir(exist_ok=True)
                (out / "Latest_Sales_Review.pptx").write_bytes(deck)
                st.caption(f"Saved {out / 'Latest_Sales_Review.pptx'} · data till {min(pri_last, sec_last):%d-%b-%Y}")
        except Exception as e:
            c1.error(f"PPTX build failed: {e}")
        # Excel with the same tables
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            pd.DataFrame({"Highlights": H.highlights(ctx)}).to_excel(xw, sheet_name="Highlights", index=False)
            for k in keys:
                if k in ("highlights", "trend", "channel_chart", "pri_sec", "pivot"): continue
                fn = {"summary": H.summary_table, "segment": H.segment_table, "stockist": H.stockist_table, "premium": H.premium_stockist_table,
                      "pgroup": H.product_group_table, "pack": H.pack_table, "category": H.category_table, "sec_channel": H.sec_channel_table, "sec_stockist": H.sec_stockist_table,
                      "sec_ro": lambda c: H.sec_channel_stockist_table(c, "RO"), "sec_bazaar": lambda c: H.sec_channel_stockist_table(c, "Bazaar"),
                      "sec_oem": lambda c: H.sec_channel_stockist_table(c, "OEM")}[k]
                t = fn(ctx); t2 = t.copy()
                t2.to_excel(xw, sheet_name=k[:31], index=False)
        c2.download_button("⬇ Download same tables as Excel", buf.getvalue(), f"{fname}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    ui.section("Preview")
    for k in keys:
        if k in ("highlights",):
            st.markdown(f"**{X.SLIDES[k][0]}**"); st.markdown("\n".join(f"- {h}" for h in H.highlights(ctx)))

# ================================================================== DATA
with tab_data, ui.guard("Data & checks"):
    f = data.info["files"]
    ui.section("Files in use")
    for kname, lst in f.items():
        st.write(f"**{kname.title()}**: " + (", ".join(lst) if lst else "— none found —"))
    st.caption("Merge rule: per month, the file with the latest invoice date wins, so overlapping / re-exported files never double count.")
    a, b = st.columns(2)
    for col, nm, df in ((a, "Primary", pri_all), (b, "Secondary", sec_all)):
        col.markdown(f"**{nm}** — {len(df):,} rows · {df.date.min():%d-%b-%Y} to {df.date.max():%d-%b-%Y}")
        col.dataframe(df.groupby("fy").agg(KL=("kl", "sum"), Rows=("kl", "size")).round(1).rename(index=fy_label), width="stretch")
    ui.section("Checks")
    st.caption("Everything on this dashboard is computed from your Primary, Secondary and Target workbooks only (Material/Pack masters and the HO template are optional extras).")
    st.write(f"Unmapped material KL — primary: {data.info['unmapped_primary']:.1f}, secondary: {data.info['unmapped_secondary']:.1f}")
    neg = sec_all[sec_all.kl < 0].kl.sum(), pri_all[pri_all.kl < 0].kl.sum()
    st.write(f"Returns / credit notes netted — secondary: {neg[0]:.1f} KL, primary: {neg[1]:.1f} KL")
    st.write(f"Targets loaded: {len(data.targets)} monthly rows from `{data.info['target_file']}`" if len(data.targets) else "No targets file found.")
    for w_ in data.info["target_warnings"] + data.info["problems"]: st.warning(w_)
    if len(data.targets):
        ts = data.targets.groupby(["fy", "kind", "channel", "stockist_code"]).target.sum().reset_index()
        st.dataframe(ts, hide_index=True)
    with st.expander("Raw data preview"):
        which = st.radio("Dataset", ["Primary", "Secondary"], horizontal=True)
        st.dataframe((pri if which == "Primary" else sec).sort_values("date", ascending=False).head(500), width="stretch")
