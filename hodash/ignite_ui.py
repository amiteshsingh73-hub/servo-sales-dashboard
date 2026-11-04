"""'Servo Ignite' tab: incentive calculator for the Oct-Dec'26 scheme."""
from __future__ import annotations
import pandas as pd
import streamlit as st

from . import ignite as G, ui


def rs(x) -> str:
    if abs(x) >= 1e5: return f"₹ {x / 1e5:,.2f} L"
    return f"₹ {x:,.0f}"


def _tbl(df: pd.DataFrame, height=None):
    if df is None or df.empty:
        st.caption("Nothing to list yet."); return
    d = df.drop(columns=["customer_id"], errors="ignore").copy()
    for c in d.columns:
        if pd.api.types.is_float_dtype(d[c]): d[c] = d[c].round(0)
    st.dataframe(d, hide_index=True, width="stretch", height=height or min(420, 40 + 35 * len(d)))


def render(pri_all: pd.DataFrame, sec_all: pd.DataFrame, last_p, last_s):
    names = dict(zip(pri_all.stockist_code, pri_all.stockist))
    names.update({c: n for c, n in zip(sec_all.stockist_code, sec_all.stockist) if c not in names})
    days = int(max(min((min(last_p, G.END) - G.START).days + 1, G.QUARTER_DAYS), 0))
    st.markdown(f"**SERVO IGNITE, Oct-Dec'26**: 1.5% × weightage × Q3 primary invoice value (excl. GST & Hydrashakti). Five parameters of 20% each. "
                f"Circular {G.CIRCULAR}.")
    if days == 0:
        st.info("The scheme starts on 01-Oct-2026 and the sales files do not reach it yet.")
    elif days < G.QUARTER_DAYS:
        st.info(f"Quarter in progress: primary data till {last_p:%d-%b-%Y} ({days} of {G.QUARTER_DAYS} days), secondary till {last_s:%d-%b-%Y}. "
                "“Met” means met so far. SKU and new-counter counts can only grow; the ratio and PCS can move either way.")

    d = G.Rules()
    with st.expander("Settings: scheme values, GSS and actual invoice value", expanded=False):
        a, b, c = st.columns(3)
        min_skus = a.number_input("1. Unique SKUs needed", 1, 1000, d.min_skus, key="ig_min_skus")
        sku_l = a.number_input("1. Litres a SKU needs", 0.0, 1000.0, d.sku_min_l, key="ig_sku_l")
        new_ssr = b.number_input("2. New counters, SSR", 0, 1000, d.new_ssr, key="ig_new_ssr")
        new_gss = b.number_input("2. New counters, GSS", 0, 1000, d.new_gss, key="ig_new_gss")
        new_l = b.number_input("2. Litres a new counter needs", 0.0, 5000.0, d.new_min_l, key="ig_new_l")
        ratio = c.number_input("3. Bazaar ÷ RO needed", 0.0, 20.0, d.min_ratio, key="ig_ratio")
        pcs = c.number_input("4. PCS needed (L / counter / month)", 0.0, 2000.0, d.min_pcs, key="ig_pcs")
        a, b, c = st.columns(3)
        per = a.number_input("5. SKUs per counter", 1, 50, d.sku_per_counter, key="ig_per")
        ctr_l = a.number_input("5. Litres a SKU needs at the counter", 0.0, 500.0, d.counter_sku_min_l, key="ig_ctr_l",
                               help="The circular repeats 'SKU with min 20 Ltrs' under this parameter. Set 0 to count every SKU the counter bought.")
        rule = b.radio("5. How it is judged", ["every", "average"], key="ig_rule",
                       format_func=lambda x: "Every counter must reach it (circular wording)" if x == "every" else "Average across counters",
                       help="The circular says 'every bazaar counter has to purchase min 4 unique SKU'. Confirm with HO if the average is meant.")
        gss = set(b.multiselect("Stockists that are GSS", sorted(names), format_func=names.get, key="ig_gss"))
        price = c.number_input("Primary price as % of secondary ₹/L", 1.0, 150.0, 100.0, key="ig_price",
                               help="Used only for the value estimate. The Primary file has no invoice value, so value is estimated from the secondary ₹/L of each SKU.")
        st.markdown("**Actual primary invoice value so far (₹, excl. GST & Hydrashakti)**: type the SAP figure to replace the estimate; 0 = use the estimate.")
        cols = st.columns(max(len(names), 1))
        typed = {code: col.number_input(names[code], 0.0, 1e11, 0.0, step=100000.0, key=f"ig_val_{code}", format="%.0f")
                 for col, code in zip(cols, sorted(names, key=lambda k: str(names[k])))}
    rules = G.Rules(min_skus=int(min_skus), sku_min_l=sku_l, new_ssr=int(new_ssr), new_gss=int(new_gss), new_min_l=new_l, min_ratio=ratio, min_pcs=pcs,
                    sku_per_counter=int(per), counter_sku_min_l=ctr_l, sku_rule=rule)
    res = G.evaluate(pri_all, sec_all, names, rules, gss, typed, price, last_p, last_s)

    tot_now, tot_full, tot_max = (sum(getattr(r, k) for r in res) for k in ("incentive_now", "incentive_full", "incentive_max"))
    ui.cards([
        ui.kpi("Incentive on value so far", rs(tot_now), "at parameters met today", tone="plain"),
        ui.kpi("Full quarter at this pace", rs(tot_full), "same parameters, value at run-rate", tone="plain"),
        ui.kpi("If all 5 parameters are met", rs(tot_max), "1.5% of full-quarter value", tone="plain"),
        ui.kpi("Each parameter is worth", rs(tot_max * rules.weight_pct / 100), "across all stockists, full quarter", tone="plain"),
    ])
    if any(r.value_source == "estimate" for r in res):
        st.warning("Rupee figures are **estimates**: the Primary file has litres but no invoice value, so value = litres × the secondary ₹/L of each SKU, which is "
                   "the stockist's selling price and so likely higher than the primary invoice price. Type the SAP value in Settings for exact figures.")

    ui.section("Scorecard")
    rows = []
    for r in res:
        row = {"Stockist": r.name, "Type": r.kind}
        row.update({p["Parameter"]: ("✅ " if p["met"] else "❌ ") + p["Actual"] for p in r.params})
        row.update({"Weightage": f"{r.weight:.0%}", "Value so far": rs(r.value_to_date), "Basis": r.value_source, "Incentive so far": rs(r.incentive_now),
                    "Full quarter": rs(r.incentive_full), "If all 5 met": rs(r.incentive_max)})
        rows.append(row)
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
    note = [f"SERVO IGNITE Oct-Dec'26, circular {G.CIRCULAR}", f"Primary data till {last_p:%d-%b-%Y}, secondary till {last_s:%d-%b-%Y}.",
            "Hydrashakti excluded from every parameter and from the value.", "SKU = product code + pack size (promo and loose can of the same size are one SKU).",
            "Value basis 'estimate' = litres x secondary Rs/L of the SKU; replace with the SAP value for exact figures."]
    st.download_button("⬇ Download all (Excel)", G.excel(res, note), "Servo Ignite Oct-Dec26.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="ig_dl")

    ui.section("Stockist detail", "what is met, what is short, and who to bill next")
    for tab, r in zip(st.tabs([r.name for r in res]), res):
        with tab:
            ui.cards([ui.kpi(p["Parameter"], "✅" if p["met"] else "❌", p["Actual"] + (f" · short: {p['Gap']}" if p["Gap"] else ""), tone="good" if p["met"] else "bad")
                      for p in r.params])
            st.caption(f"Primary excl. Hydrashakti {r.excl_hydra_kl:,.1f} KL (Hydrashakti {r.hydra_kl:,.1f} KL excluded). "
                       f"Incentive: {rs(r.value_to_date)} × 1.5% × {r.weight:.0%} = {rs(r.incentive_now)}.")
            a, b = st.columns(2)
            with a:
                st.markdown("**SKUs billed in the last 12 months, short of the minimum this quarter**"); _tbl(r.sku_gap)
            with b:
                st.markdown("**New Bazaar counters this quarter**"); _tbl(r.new_list)
            st.markdown("**Bazaar counters and their SKU count** (biggest first)"); _tbl(r.counter_list)
    st.caption("How it is read: new counter = not billed by anyone from 01-Apr-2025 to 30-Sep-2026 (any product). Bazaar sales include ServoXpert; counter counts do not. "
               "RO = Dealer + Nayara RO. Primary and secondary volumes are net of returns.")
