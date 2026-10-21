"""'LRS Format' tab: the LRS Field Sales Analysis sheet for every stockist (SSR), for any month in the data."""
from __future__ import annotations
import html
import streamlit as st

from . import config as C, lrs, ui
from .periods import fy_label, mi_name

CSS = """<style>
.lrs {border-collapse:collapse;width:100%;font-size:.86rem;background:#fff;margin-bottom:4px}
.lrs td,.lrs th {border:1px solid #9AA3B5;padding:5px 7px;text-align:center;vertical-align:middle}
.lrs th {background:#DCE6F1;font-weight:700;color:#1B2440}
.lrs .ttl {font-weight:800;font-size:.98rem;background:#0B1F66;color:#fff}
.lrs .y {background:#FFF59D;font-weight:700}
.lrs .l {text-align:left}
.lrs .tot td {font-weight:800;background:#F4F6FB}
.lrs .neg {color:#C0392B;font-weight:700} .lrs .pos {color:#1A7F4B;font-weight:700}
.lrs .rm {text-align:left;font-size:.8rem;color:#2B3350}
</style>"""


def _n(x, d=2):
    return "" if x is None or x != x else f"{x:,.{d}f}"


def _pct(x):
    return "" if x is None or x != x else f"{x * 100:.1f}%"


def _html(r: dict) -> str:
    e = html.escape
    month = r["month_label"] + (f" (till {r['data_till']:%d-%b-%y})" if r["partial"] else "")
    h = [f"<table class='lrs'><tr><td colspan=9 class='ttl'>* 2 - Sales Analysis</td></tr>",
         f"<tr><td colspan=4>Name of the Stockist</td><td colspan=5 class='y'>{e(r['stockist'])}</td></tr>"]
    for key, lbl, val in (("month", "Month", month), ("cum", "Cumulative", r["cum_label"])):
        d = r[key]
        cls = "pos" if d["diff"] >= 0 else "neg"
        h += [f"<tr><td colspan=3><b>{lbl}</b></td><td colspan=6 class='y'>{e(val)}</td></tr>",
              "<tr><th>Potential (KL)</th><th>Target (KL)</th><th>Sales CY (KL)</th><th>Sales LY (KL)</th><th>(+/-)</th><th colspan=4>Remarks</th></tr>",
              f"<tr><td>{_n(d['potential']) or '–'}</td><td>{_n(d['target']) or '–'}</td><td><b>{_n(d['cy'])}</b></td><td>{_n(d['ly'])}</td>"
              f"<td class='{cls}'>{d['diff']:+,.2f}</td><td colspan=4 class='rm'>{e(d['remark'])}</td></tr>"]
    h += ["</table><table class='lrs'><tr><td colspan=9 class='ttl'>LUBES RESELLER GROUP</td></tr>",
          "<tr><td colspan=9>Analysis of Secondary Sales</td></tr>",
          f"<tr><td colspan=4>Period</td><td colspan=5 class='y'>{e(r['sec_label'])}</td></tr>",
          "<tr><th colspan=3>Channel</th><th>Total No.</th><th>Cum. No. Supplied</th><th>Sales CY</th><th>Sales LY</th><th>% of Total Sales CY</th><th>% of Total Sales LY</th></tr>"]
    t = r["sec"]
    for row in t.itertuples(index=False):
        h.append(f"<tr><td colspan=3 class='l'>{e(row.Channel)}</td><td>{row[1]:,}</td><td>{row[2]:,}</td><td>{_n(row[3])}</td><td>{_n(row[4])}</td>"
                 f"<td>{_pct(row[5])}</td><td>{_pct(row[6])}</td></tr>")
    tc, tl = t["Sales CY"].sum(), t["Sales LY"].sum()
    h.append(f"<tr class='tot'><td colspan=3>Total</td><td>{t['Total No.'].sum():,}</td><td>{t['Cum. No. Supplied'].sum():,}</td>"
             f"<td>{_n(tc)}</td><td>{_n(tl)}</td><td>{'100.0%' if tc else ''}</td><td>{'100.0%' if tl else ''}</td></tr></table>")
    return "".join(h)


def render(pri, sec, targets, pri_last, sec_last):
    st.markdown(CSS, unsafe_allow_html=True)
    mis = lrs.months_available(pri, sec)
    if not mis:
        st.info("No sales data yet."); return
    latest = min(pri_last, sec_last)
    default_mi = latest.year * 12 + latest.month - 1
    c1, c2, c3 = st.columns([1.3, 1.6, 1.6])
    mi = c1.selectbox("Month", mis, index=mis.index(default_mi) if default_mi in mis else 0, format_func=lambda m: mi_name(int(m)), key="lrs_month",
                      help="Defaults to the latest month in the Primary / Secondary files. Drop the October files in Data and October appears here automatically.")
    basis = c2.radio("Top section sales from", ["Primary", "Secondary"], horizontal=True, key="lrs_basis",
                     help="Primary = stockist billing from the Primary sales file vs primary target. Secondary = stockist's sales to counters from the Secondary file vs secondary target.")
    sp = c3.radio("Secondary table period", ["Cumulative", "Month"], horizontal=True, key="lrs_secp",
                  help="Cumulative = April to the selected month (matches 'Cum. No. Supplied').")
    pot = lrs.load_potential()
    blocks = lrs.build(pri, sec, targets, pri_last, sec_last, int(mi), basis, sp, pot)

    d1, d2, d3 = st.columns([1.4, 1.4, 3])
    tag = mi_name(int(mi)).replace("'", "")
    d1.download_button(f"⬇ Excel – all {len(blocks)} SSRs (one sheet each)", lrs.workbook_bytes(blocks), f"LRS Sales Analysis - {tag}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="lrs_dl_all")
    d2.download_button(f"⬇ {len(blocks)} separate files (ZIP)", lrs.zip_bytes(blocks), f"LRS Sales Analysis - {tag} (per SSR).zip", "application/zip", key="lrs_dl_zip")
    if not pot:
        with d3:
            st.caption(f"Potential (KL) is blank: add annual potential per stockist in **{lrs.POTENTIAL_FILE}** in the Data folder.")
            if not lrs.potential_path().exists() and st.button("Create the potential file for me", key="lrs_mk_pot"):
                names = dict(zip(pri.drop_duplicates("stockist_code").stockist_code.astype(int), pri.drop_duplicates("stockist_code").stockist))
                f = lrs.write_potential_template(names, blocks[0]["fy"])
                st.success(f"Created {f.name} - fill the 'Annual Potential (KL)' column, save, and refresh.")
    if blocks and blocks[0]["partial"]:
        st.info(f"{mi_name(int(mi))} is still in progress in the data (till {blocks[0]['data_till']:%d-%b-%Y}); LY is cut at the same date.")

    for tab, r in zip(st.tabs([r["stockist"] for r in blocks]), blocks):
        with tab:
            left, _ = st.columns([3, 1])
            left.download_button(f"⬇ Download {r['stockist']} – {r['month_label']} (Excel)", lrs.workbook_bytes([r]), lrs.file_name(r),
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"lrs_dl_{r['code']}")
            left.markdown(_html(r), unsafe_allow_html=True)
    st.caption(f"Sales from the Primary and Secondary sales files ({fy_label(blocks[0]['fy'])}), targets from the Target workbook. "
               "Secondary excludes ClearBlue (DEF). Outlet types: Non-KSK RO = Dealer group except 'IOC- KSK' sub-group; KSK RO = 'IOC- KSK'; "
               "Nayara = 'NAYARA RO' group / sub-group; Bazaar Shop = Bazaar + ServoXpert; OE Dealers = OEM. "
               "Total No. = outlets billed in the 12 months to the period end.")
