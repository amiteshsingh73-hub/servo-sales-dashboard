"""'Review Letters' tab: monthly Bazaar performance-review letter (Word) + mail for every SSR.
Picks the latest month in the sales files by default, and when a new complete month arrives it writes that month's
letters to Sales Review/Reports/Review Letters <Mon-YY>/ automatically (once)."""
from __future__ import annotations
import datetime as dt
import os, re, urllib.parse
from pathlib import Path
import pandas as pd
import streamlit as st

from . import config as C, lrs, review_letter as RL
from .periods import mi_name, month_end

OUT_ROOT = C.DATA_DIR.parent / "Reports"
SEV = {"strong": ("🔴 Strong – Bazaar below target / degrowing", "#C0392B"),
       "watch": ("🟠 Watch – slipping in the latest month", "#E8710A"),
       "good": ("🟢 Good – appreciation letter", "#1A7F4B")}


def _key(s) -> str:
    return re.sub(r"[^a-z]", "", str(s).lower())[:6]


@st.cache_data(show_spinner=False)
def contacts() -> tuple[dict, str]:
    """To / CC per stockist and the mail signature from Data/SSR Contacts.xlsx (optional)."""
    f = C.DATA_DIR / "SSR Contacts.xlsx"
    if not f.exists(): return {}, ""
    try:
        x = pd.read_excel(f, sheet_name=None, header=None)
    except Exception:
        return {}, ""
    out = {}
    c = x.get("Contacts")
    if c is not None and len(c) > 1:
        for row in c.iloc[1:].itertuples(index=False):
            to, cc = (str(v).strip() if pd.notna(v) else "" for v in (row[1], row[2]))
            out[_key(row[0])] = (to, cc)
    sig = x.get("Signature")
    lines = [str(v).strip() for v in sig.iloc[:, 0].tolist()] if sig is not None and len(sig) else []
    return out, "\n".join(l for l in lines if l and l != "nan")


def _folder(mi: int) -> Path:
    return OUT_ROOT / f"Review Letters {mi_name(mi).replace(chr(39), '-')}"


def _files(r, settings) -> dict[str, bytes]:
    tag = r["month"].replace("'", "-")
    return {f"Review Letter - {r['stockist']} - {r['span'].replace(chr(39), '-')}.docx":
            RL.letter_docx(r, settings["ref"], settings["date"], settings["address"], settings["prev"], settings["days"], settings["nc"]),
            f"Counter lists - {r['stockist']} - {tag}.xlsx": RL.counter_lists_xlsx(r)}


def _defaults(code: int, mi: int) -> dict:
    y, m = divmod(mi, 12)
    return dict(ref=f"CORP/PS/ADV/{code}/{y}/{m + 1:02d}", date=dt.date.today().strftime("%d.%m.%Y"),
                address=C.LETTER_ADDRESS.get(code, ""), prev="", days=5, nc=None)


def render(pri, sec, targets, pri_last, sec_last):
    mis = lrs.months_available(pri, sec)
    if not mis or targets is None or targets.empty:
        st.info("Needs Primary, Secondary and Target files."); return
    latest = min(pri_last, sec_last)
    latest_mi = latest.year * 12 + latest.month - 1
    complete = latest >= month_end(latest_mi).normalize()
    default_mi = latest_mi if complete else latest_mi - 1
    mis = [m for m in mis if m <= latest_mi]
    codes = sorted(set(sec.stockist_code.unique()) & set(targets.stockist_code.unique()))

    c1, c2 = st.columns([1.2, 3])
    mi = int(c1.selectbox("Review up to month", mis, index=mis.index(default_mi) if default_mi in mis else 0,
                          format_func=lambda m: mi_name(int(m)), key="rl_month",
                          help="The letter covers April to this month (cumulative) vs the same months last year and the target, "
                               "with the selected month highlighted. New monthly files bring the next month here automatically."))
    fy, fm = lrs._fy_fm(mi)
    c2.caption(f"Secondary data till **{sec_last:%d-%b-%Y}**, primary till **{pri_last:%d-%b-%Y}**. "
               + ("" if mi < latest_mi or complete else "⚠️ This month is still in progress: LY is cut at the same date."))

    reports = {code: RL.analyse(pri, sec, targets, code, fy, fm, sec_last, pri_last) for code in codes}
    reports = {k: v for k, v in reports.items() if v["bz_cy"] or v["bz_ly"]}
    con, sig = contacts()

    # settings per SSR (kept in the session; defaults filled from config)
    settings = {}
    with st.expander("Letter settings (reference no., date, address, previous letter)", expanded=False):
        for code, r in reports.items():
            d = _defaults(code, mi)
            st.markdown(f"**{r['stockist']}**")
            a, b, c, e, f, g = st.columns([1.6, .8, 1.2, 1.8, .6, .7])
            settings[code] = dict(
                ref=a.text_input("Ref. no.", d["ref"], key=f"rl_ref_{code}_{mi}"),
                date=b.text_input("Date", d["date"], key=f"rl_date_{code}"),
                address=c.text_input("Address line", d["address"], key=f"rl_addr_{code}"),
                prev=e.text_input("Previous letter (Ref. & date)", d["prev"], key=f"rl_prev_{code}", placeholder="e.g. CORP/PS/ADV/262098/2026/1 dated 02.03.2026"),
                days=int(f.number_input("Reply days", 1, 30, d["days"], key=f"rl_days_{code}")),
                nc=int(g.number_input("New counters", 0, 500, 0, key=f"rl_nc_{code}", help="0 = automatic (half the net counter loss, minimum 20)")) or None)
        st.caption(f"Default addresses come from LETTER_ADDRESS in hodash/config.py. Put a 'Letterhead*.docx' in {C.DATA_DIR.name}/ to print the letters on it.")

    all_files = {}
    for code, r in reports.items():
        for n, b in _files(r, settings[code]).items(): all_files[n] = b

    folder = _folder(mi)
    d1, d2, d3 = st.columns([1.5, 1.5, 3])
    d1.download_button(f"⬇ All {len(reports)} letters + counter lists (ZIP)", RL.zip_all(all_files),
                       f"Review Letters {mi_name(mi).replace(chr(39), '-')}.zip", "application/zip", type="primary", key="rl_zip")
    if d2.button(f"💾 Save to Reports/{folder.name}", key="rl_save"):
        folder.mkdir(parents=True, exist_ok=True)
        for n, b in all_files.items(): (folder / n).write_bytes(b)
        st.toast(f"Saved {len(all_files)} files to {folder}")

    # first time a complete month is seen: write its letters automatically with the default settings
    auto = os.environ.get("SALES_REVIEW_NO_AUTOSAVE") != "1"          # the smoke test turns this off
    if auto and mi == default_mi and complete and not folder.exists() and st.session_state.get("_rl_auto") != mi:
        st.session_state["_rl_auto"] = mi
        try:
            folder.mkdir(parents=True, exist_ok=True)
            for n, b in all_files.items(): (folder / n).write_bytes(b)
            d3.success(f"New month {mi_name(mi)} found: letters saved to Reports/{folder.name}.")
        except OSError as e:
            d3.warning(f"Could not save to {folder}: {e}")
    elif folder.exists():
        d3.caption(f"📁 Saved copy: Reports/{folder.name}")

    for tab, (code, r) in zip(st.tabs([r["stockist"] for r in reports.values()]), reports.items()):
        with tab:
            sev = RL.severity(r); lbl, col = SEV[sev]
            st.markdown(f"<div style='border-left:5px solid {col};padding:6px 12px;background:#F7F8FB;border-radius:6px'>"
                        f"<b>{lbl}</b> · Bazaar {RL.f1(r['bz_cy'])} KL vs target {RL.f1(r['bz_t'])} KL ({RL.fp(r['bz_ach'])}), "
                        f"{RL.fp(r['bz_g'], True)} vs LY · {r['month']} {RL.f1(r['m_cy'])} KL ({RL.fp(r['m_g'], True)} vs LY)</div>",
                        unsafe_allow_html=True)
            s = settings[code]; files = _files(r, s)
            b1, b2, _ = st.columns([1.4, 1.4, 3])
            (ln, lb), (cn, cb) = list(files.items())
            b1.download_button("⬇ Letter (Word)", lb, ln, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"rl_doc_{code}")
            b2.download_button("⬇ Lost / declining counters (Excel)", cb, cn, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"rl_xl_{code}")

            left, right = st.columns([1.15, 1])
            with left:
                st.markdown("##### What the letter says")
                F = RL.findings(r)
                for title, k in [("Overall", "overall"), ("Trend", "trend"), ("Counters", "counters"), ("Bulk buyers", "bulk"),
                                 ("Product mix", "mix"), ("SSF", "ssf"), ("Geography / returns", "geo"), ("Requirement", "need")]:
                    if F.get(k):
                        st.markdown(f"**{title}**\n" + "\n".join(f"- {x}" for x in F[k]))
                st.markdown("**Action points**\n" + "\n".join(f"{i}. {a}" for i, a in enumerate(RL.actions(r, s["days"], s["nc"]), 1)))
            with right:
                st.markdown("##### Mail")
                m = RL.mail(r, s["days"], sig, s["nc"])
                to, cc = con.get(_key(r["stockist"]), ("", ""))
                st.text_input("Subject", m["subject"], key=f"rl_subj_{code}_{mi}")
                st.code(m["body"], language=None, wrap_lines=True)
                q = urllib.parse.urlencode({"cc": cc, "subject": m["subject"], "body": m["body"]}, quote_via=urllib.parse.quote)
                st.link_button("✉️ Open in mail app", f"mailto:{urllib.parse.quote(to)}?{q}", disabled=False)
                if not to:
                    st.caption("No e-mail address for this SSR in Data/SSR Contacts.xlsx (To column) – fill it to pre-address the mail. "
                               "Attach the Word letter before sending.")
    st.caption("Bazaar = Bazaar + ServoXpert, ClearBlue excluded, volumes net of returns. Counters lost = billed in the same months last year, "
               "not billed this year. Severity: Strong if Bazaar is below 95% of target or below last year; Watch if the latest month is >10% below LY "
               "or lost counters exceed twice the new ones.")
