"""Counters not billed since a chosen date: the 'not billed' tab's calculation and screen."""
from __future__ import annotations
import datetime as dt
import io
import pandas as pd
import streamlit as st

from . import config as C, ui


def not_billed_since(sec: pd.DataFrame, since: pd.Timestamp) -> pd.DataFrame:
    """One row per counter that was billed before `since` and has no bill on or after it, under any stockist.
    A return (negative quantity) is not a bill. Details come from the counter's last bill."""
    cols = ["customer_id", "Counter", "Channel", "Group", "Sub-group", "Stockist", "SSF", "District", "Last bill", "Days since last bill",
            "KL in 12 months before date", "KL total before date", "Bills before date"]
    bills = sec[sec.kl > 0]
    if bills.empty: return pd.DataFrame(columns=cols)
    last = bills.groupby("customer_id").date.max()
    gone = last[last < since].index
    if not len(gone): return pd.DataFrame(columns=cols)
    d = bills[bills.customer_id.isin(gone)].sort_values("date")
    info = d.groupby("customer_id").agg(Counter=("customer", "last"), Channel=("channel", "last"), Group=("cust_group", "last"),
                                        **{"Sub-group": ("cust_subgroup", "last"), "Stockist": ("stockist", "last"), "SSF": ("ssf_name", "last"),
                                           "District": ("district", "last"), "Last bill": ("date", "max"), "Bills before date": ("invoice", "nunique")})
    net = sec[sec.customer_id.isin(gone)]
    year_ago = since - pd.DateOffset(months=12)
    info["KL in 12 months before date"] = net[(net.date >= year_ago) & (net.date < since)].groupby("customer_id").kl.sum().reindex(info.index).fillna(0.0)
    info["KL total before date"] = net[net.date < since].groupby("customer_id").kl.sum().reindex(info.index).fillna(0.0)
    as_of = sec.date.max()
    info["Days since last bill"] = (as_of - info["Last bill"]).dt.days
    return info.reset_index()[cols].sort_values(["KL in 12 months before date", "KL total before date"], ascending=False, ignore_index=True)


def _excel(t: pd.DataFrame, by: pd.DataFrame, header: list[str]) -> bytes:
    buf = io.BytesIO()
    x = t.rename(columns={"customer_id": "Customer Id"}).copy()
    x["Last bill"] = x["Last bill"].dt.strftime("%d-%b-%Y")
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame({"": header}).to_excel(xw, sheet_name="Read me", index=False, header=False)
        x.round(3).to_excel(xw, sheet_name="Counters", index=False)
        by.round(3).to_excel(xw, sheet_name="By SSF", index=False)
    return buf.getvalue()


def render(sec_all: pd.DataFrame, sec_last):
    if sec_all.empty:
        st.info("No secondary sales loaded."); return
    first = sec_all.date.min().date()
    last = sec_last.date()
    st.caption(f"Counters billed at some time since {first:%d-%b-%Y} (start of the Secondary file) with no bill on or after the date you pick, under any stockist. "
               f"Data till {last:%d-%b-%Y}. The sidebar filters do not apply here.")
    c1, c2, c3 = st.columns([1.1, 1.6, 1.6])
    since = c1.date_input("Not billed since", min_value=first, max_value=last, format="DD/MM/YYYY", key="nb_since",
                          value=min(max(dt.date(2026, 1, 1), first), last), help="e.g. 01/01/2026 or 01/04/2025. Bills on this date count as billed.")
    chans = [c for c in C.CHANNEL_ORDER if c in set(sec_all.channel)]
    chan = c2.multiselect("Channel", chans, default=[c for c in chans if c == "Bazaar"], placeholder="All channels", key="nb_chan")
    stk = c3.multiselect("Stockist (SSR)", sorted(sec_all.stockist.unique()), placeholder="All stockists", key="nb_stk")
    all_t = not_billed_since(sec_all, pd.Timestamp(since))
    t = all_t
    if chan: t = t[t.Channel.isin(chan)]
    if stk: t = t[t.Stockist.isin(stk)]
    c4, c5, c6 = st.columns([1.6, 1.6, 1.1])
    ssf = c4.multiselect("SSF", sorted(t.SSF.astype(str).unique()), placeholder="All SSFs", key="nb_ssf")
    dist = c5.multiselect("District", sorted(t.District.astype(str).unique()), placeholder="All districts", key="nb_dist")
    min_kl = c6.number_input("Min litres in the 12 months before", 0, 1_000_000, 0, step=50, key="nb_min",
                             help="0 lists everyone, including counters that stopped more than a year before the date.")
    if ssf: t = t[t.SSF.astype(str).isin(ssf)]
    if dist: t = t[t.District.astype(str).isin(dist)]
    if min_kl: t = t[t["KL in 12 months before date"] * 1000 >= min_kl]

    # counters billed before the date in the same selection = the base for the %
    base = sec_all[(sec_all.kl > 0) & (sec_all.date < pd.Timestamp(since))].sort_values("date").groupby("customer_id").agg(ch=("channel", "last"), st=("stockist", "last"))
    if chan: base = base[base.ch.isin(chan)]
    if stk: base = base[base.st.isin(stk)]
    kl12 = float(t["KL in 12 months before date"].sum())
    ui.cards([
        ui.kpi("Counters not billed", f"{len(t):,}", f"since {since:%d-%b-%Y}", tone="warn" if len(t) else "good"),
        ui.kpi("Share of earlier counters", f"{len(t) / len(base) * 100:.0f}%" if len(base) else "–", f"of {len(base):,} billed before that date", tone="plain"),
        ui.kpi("Their volume, 12 months before", f"{kl12:,.1f}", "what these counters used to buy", unit="KL", tone="plain"),
    ])
    if t.empty:
        st.success("Every counter in this selection has been billed since that date." if len(base) else "No counter in this selection was billed before that date.")
        return

    ui.section("Counters", "biggest buyers first")
    show = t.drop(columns="customer_id").copy()
    show["Last bill"] = show["Last bill"].dt.strftime("%d-%b-%Y")
    for c in ("KL in 12 months before date", "KL total before date"): show[c] = show[c].round(2)
    st.dataframe(show, hide_index=True, width="stretch", height=min(560, 40 + 35 * len(show)))

    ui.section("By SSF", "who should follow up")
    by = t.groupby(["Stockist", "SSF"]).agg(Counters=("Counter", "size"), **{"KL in 12 months before date": ("KL in 12 months before date", "sum")}) \
        .reset_index().sort_values("KL in 12 months before date", ascending=False)
    st.dataframe(by.round(2), hide_index=True, width="stretch", height=min(420, 40 + 35 * len(by)))

    head = [f"Counters not billed since {since:%d-%b-%Y} (secondary data till {last:%d-%b-%Y})", f"Channel: {', '.join(chan) or 'All'}", f"Stockist: {', '.join(stk) or 'All'}",
            f"SSF: {', '.join(ssf) or 'All'}", f"District: {', '.join(dist) or 'All'}", f"Min litres in the 12 months before: {min_kl}",
            f"Only counters billed at least once since {first:%d-%b-%Y} can appear. Returns are not bills."]
    st.download_button("⬇ Download list (Excel)", _excel(t, by, head), f"Counters not billed since {since:%d-%b-%Y}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="nb_dl")
