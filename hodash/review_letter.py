"""Monthly SSR performance-review letter (Bazaar focus): the numbers, the narrative, the Word letter and the mail.
analyse() works on any stockist / month in the data (cumulative Apr -> month vs LY like-for-like and target),
so a new month's Primary / Secondary files give next month's letters with no other change.
"""
from __future__ import annotations
import io, re, zipfile
from pathlib import Path
import numpy as np
import pandas as pd

from . import config as C
from .metrics import sel, growth, pct, target_for, monthly_target
from .periods import MONTHS, make_period, windows, fy_label, mi_of

# Counter groups that buy in bulk for members (fishermen co-op societies etc.) - tracked separately when material.
SOCIETY_PATTERN = r"machh|machim|machhu|fisher|sahakari|vkss|society|sanstha"
QUARTERS = {1: (1, 3), 2: (4, 6), 3: (7, 9), 4: (10, 12)}


def _span(fm_a: int, fm_b: int) -> str:
    return MONTHS[fm_a - 1] if fm_a == fm_b else f"{MONTHS[fm_a - 1]}–{MONTHS[fm_b - 1]}"


def _yy(fy: int, fm: int) -> str:
    y = fy - 1 if fm <= 9 else fy
    return f"{y % 100:02d}"


def _yspan(fy: int, fm: int) -> str:
    """Apr'26 / Apr–Sep'26 / Apr'25–Jan'26 (year shown on both ends once the span crosses into Jan)."""
    if fm == 1: return f"Apr'{_yy(fy, 1)}"
    if fm >= 10: return f"Apr'{_yy(fy, 1)}–{MONTHS[fm - 1]}'{_yy(fy, fm)}"
    return f"{_span(1, fm)}'{_yy(fy, fm)}"


def _per_counter(df):
    return df.groupby("customer_id").kl.sum()


# ------------------------------------------------------------------ numbers
def analyse(pri: pd.DataFrame, sec: pd.DataFrame, targets: pd.DataFrame, code: int, fy: int, fm: int,
            sec_last: pd.Timestamp, pri_last: pd.Timestamp) -> dict:
    p_all = pri[pri.stockist_code == code]
    s_all = sec[sec.stockist_code == code]
    name = (s_all.stockist.iloc[0] if len(s_all) else p_all.stockist.iloc[0] if len(p_all) else C.STOCKIST_ALIASES.get(code, str(code)))
    per = make_period("YTD (Apr to month)", fy, fm=fm)
    ws, wp = windows(per, sec_last), windows(per, pri_last)
    fms = per.fms
    r = dict(code=code, stockist=name, fy=fy, fm=fm, months=len(fms), partial=ws["cy"].partial, data_till=ws["cy"].end,
             span=_yspan(fy, fm), span_ly=_yspan(fy - 1, fm), month=f"{MONTHS[fm - 1]}'{_yy(fy, fm)}",
             month_ly=f"{MONTHS[fm - 1]}'{_yy(fy - 1, fm)}", fy_label=fy_label(fy), fy_ly=fy_label(fy - 1, True), fy_cy=fy_label(fy, True))

    # --- overall: primary + secondary by channel
    rows = []
    pc, pl = float(sel(p_all, wp["cy"]).kl.sum()), float(sel(p_all, wp["ly"]).kl.sum())
    rows.append(("Primary", pl, target_for(targets, [code], fy, fms, "Primary"), pc))
    scy, sly = sel(s_all, ws["cy"]), sel(s_all, ws["ly"])
    for ch in C.CHANNEL_ORDER:
        a, b = float(scy[scy.channel == ch].kl.sum()), float(sly[sly.channel == ch].kl.sum())
        t = target_for(targets, [code], fy, fms, "Secondary", ch)
        if a or b or (t == t and t): rows.append((f"Secondary – {ch}", b, t, a))
    tot_t = target_for(targets, [code], fy, fms, "Secondary")
    rows.append(("Secondary – Total", float(sly.kl.sum()), tot_t, float(scy.kl.sum())))
    r["overall"] = pd.DataFrame(rows, columns=["Line", "LY", "Target", "CY"])
    r["overall"]["Ach"] = [pct(c, t) if t == t else np.nan for c, t in zip(r["overall"].CY, r["overall"].Target)]
    r["overall"]["Growth"] = [growth(c, l) for c, l in zip(r["overall"].CY, r["overall"].LY)]
    r["sec_cy"], r["sec_ly"] = float(scy.kl.sum()), float(sly.kl.sum())

    oem = scy[scy.channel == "OEM"].groupby("customer").kl.sum().sort_values(ascending=False)
    r["oem_kl"] = float(oem.sum())
    r["oem_top"] = [(n, float(v)) for n, v in oem.head(2).items() if v > 0]
    r["oem_top_share"] = pct(sum(v for _, v in r["oem_top"]), r["oem_kl"])

    # --- Bazaar
    bz_all = s_all[s_all.channel == "Bazaar"]
    bc, bl = sel(bz_all, ws["cy"]), sel(bz_all, ws["ly"])
    r["bz_cy"], r["bz_ly"] = float(bc.kl.sum()), float(bl.kl.sum())
    r["bz_t"] = target_for(targets, [code], fy, fms, "Secondary", "Bazaar")
    r["bz_ach"], r["bz_g"] = pct(r["bz_cy"], r["bz_t"]), growth(r["bz_cy"], r["bz_ly"])
    r["bz_gap"] = r["bz_cy"] - r["bz_t"] if r["bz_t"] == r["bz_t"] else np.nan
    r["bz_share"], r["bz_share_ly"] = pct(r["bz_cy"], r["sec_cy"]), pct(r["bz_ly"], r["sec_ly"])

    mt = monthly_target(targets, fy, "Secondary", "Bazaar", [code]).reindex(range(1, 13))
    def by_month(df):
        return df.groupby("fm").kl.sum().reindex(range(1, 13)).fillna(0.0)
    m_ly, m_cy = by_month(bl), by_month(bc)
    r["bz_monthly"] = pd.DataFrame({"LY": m_ly, "Target": mt, "CY": m_cy}).loc[fms]
    qs = []
    for q, (a, b) in QUARTERS.items():
        if a > fm: break
        b_ = min(b, fm)
        qs.append(dict(q=f"Q{q}", label=_span(a, b_), ly=float(m_ly.loc[a:b_].sum()), cy=float(m_cy.loc[a:b_].sum()),
                       t=float(mt.loc[a:b_].sum()) if mt.loc[a:b_].notna().any() else np.nan, full=b_ == b))
    for q in qs: q["g"] = growth(q["cy"], q["ly"])
    r["quarters"] = qs
    r["m_cy"], r["m_ly"], r["m_t"] = float(m_cy[fm]), float(m_ly[fm]), float(mt[fm]) if mt[fm] == mt[fm] else np.nan
    r["m_g"], r["m_ach"] = growth(r["m_cy"], r["m_ly"]), pct(r["m_cy"], r["m_t"])

    # --- counters
    a, b = _per_counter(bl), _per_counter(bc)
    df = pd.concat([a.rename("ly"), b.rename("cy")], axis=1).fillna(0.0)
    names = bz_all.sort_values("date").groupby("customer_id")[["customer", "ssf_name", "district"]].last()
    df = df.join(names)
    lost, new = df[(df.ly > 0) & (df.cy <= 0)], df[(df.ly <= 0) & (df.cy > 0)]
    cont = df[(df.ly > 0) & (df.cy > 0)]
    dec, gro = cont[cont.cy < cont.ly * 0.9], cont[cont.cy > cont.ly * 1.1]
    r.update(cnt_cy=int((df.cy > 0).sum()), cnt_ly=int((df.ly > 0).sum()), lost_n=len(lost), lost_kl=float(lost.ly.sum()),
             new_n=len(new), new_kl=float(new.cy.sum()), dec_n=len(dec), dec_kl=float((dec.cy - dec.ly).sum()),
             gro_n=len(gro), gro_kl=float((gro.cy - gro.ly).sum()),
             lost_big_n=int((lost.ly >= 0.1).sum()), lost_big_kl=float(lost[lost.ly >= 0.1].ly.sum()))
    r["lost_top"] = lost.sort_values("ly", ascending=False).head(10)
    r["dec_top"] = dec.assign(chg=dec.cy - dec.ly).sort_values("chg").head(8)
    r["lost_list"] = lost.sort_values("ly", ascending=False)
    r["dec_list"] = dec.assign(chg=dec.cy - dec.ly).sort_values("chg")
    mc = lambda d, f: int(d[(d.fm == f) & (d.kl > 0)].customer_id.nunique())
    r["mcnt_cy"], r["mcnt_ly"] = mc(bc, fm), mc(bl, fm)

    # --- societies / bulk barrels
    soc_c, soc_l = bc[bc.customer.str.contains(SOCIETY_PATTERN, case=False, na=False)], bl[bl.customer.str.contains(SOCIETY_PATTERN, case=False, na=False)]
    r["soc_cy"], r["soc_ly"] = float(soc_c.kl.sum()), float(soc_l.kl.sum())
    r["soc_names"] = sorted(set(soc_l.customer.dropna()))[:8]
    drum = "210, 208,205,200 ltr"
    r["drum_cy"], r["drum_ly"] = float(bc[bc.pack_group == drum].kl.sum()), float(bl[bl.pack_group == drum].kl.sum())

    # --- product mix
    def cmp(dim, order=None):
        t = pd.DataFrame({"LY": bl.groupby(dim).kl.sum(), "CY": bc.groupby(dim).kl.sum()}).fillna(0.0)
        t["Change"] = t.CY - t.LY; t["Growth"] = [growth(c, l) for c, l in zip(t.CY, t.LY)]
        return t
    r["groups"] = cmp("product_group").sort_values("Change")
    r["cats"] = cmp("category")
    r["products"] = cmp("product").sort_values("Change")
    r["districts"] = cmp("district").sort_values("Change")

    # --- SSF
    ss = cmp("ssf_name")
    ss = ss[~ss.index.astype(str).str.strip().isin(["-", "", "nan"]) & ((ss.LY + ss.CY) >= 0.5)]
    ss["M_LY"] = bl[bl.fm == fm].groupby("ssf_name").kl.sum()
    ss["M_CY"] = bc[bc.fm == fm].groupby("ssf_name").kl.sum()
    ss["Lost"] = lost.groupby("ssf_name").size()
    r["ssf"] = ss.fillna({"M_LY": 0, "M_CY": 0, "Lost": 0}).sort_values("Growth", ascending=False)
    untag = bc[bc.ssf_name.astype(str).str.strip().isin(["-", "", "nan"]) & (bc.kl < 0)]
    r["returns_untagged"] = float(-untag.kl.sum())

    # --- balance-period requirement
    ann = monthly_target(targets, fy, "Secondary", "Bazaar", [code]).sum()
    r["bz_annual"] = float(ann) if ann else np.nan
    left = 12 - fm
    r["months_left"] = left
    r["need"] = r["bz_annual"] - r["bz_cy"] if r["bz_annual"] == r["bz_annual"] else np.nan
    r["need_pm"] = r["need"] / left if left and r["need"] == r["need"] else np.nan
    r["run_rate"] = r["bz_cy"] / len(fms)
    ly_rest = bz_all[(bz_all.fy == fy - 1) & (bz_all.fm > fm)]
    r["ly_rest"] = float(ly_rest.kl.sum())
    return r


# ------------------------------------------------------------------ narrative helpers
def f1(x, d=1):
    return "–" if x is None or x != x else f"{x:,.{d}f}"


def fp(x, sign=False):
    if x is None or x != x: return "–"
    if abs(x) >= 1000: return "n.m."
    d = 1 if abs(x) < 1 else 0
    return f"{x:+.{d}f}%" if sign else f"{x:.{d}f}%"


def severity(r: dict) -> str:
    """'strong' = Bazaar below target or degrowing, 'watch' = on target but slipping in the latest quarter/month, 'good' otherwise."""
    if (r["bz_ach"] == r["bz_ach"] and r["bz_ach"] < 95) or (r["bz_g"] == r["bz_g"] and r["bz_g"] < 0): return "strong"
    if (r["m_g"] == r["m_g"] and r["m_g"] < -10) or r["lost_n"] > r["new_n"] * 2: return "watch"
    return "good"


def _short(n: str, k=38) -> str:
    n = re.sub(r"\s+", " ", str(n)).strip()
    return n if len(n) <= k else n[:k - 1].rstrip() + "…"


def findings(r: dict) -> dict[str, list[str]]:
    """Plain-text bullets per section (used by both the letter and the mail)."""
    out: dict[str, list[str]] = {}
    sev = severity(r)
    # overall
    ov = []
    others = r["sec_cy"] - r["bz_cy"]
    if r["oem_kl"] > 0 and r["oem_top"]:
        names = " and ".join(f"{_short(n, 30)} – {f1(v)} KL" for n, v in r["oem_top"])
        ov.append(f"OEM volume is {f1(r['oem_kl'])} KL; the top {len(r['oem_top'])} OEM customer(s) ({names}) make up {fp(r['oem_top_share'])} of it.")
    if r["bz_share"] == r["bz_share"] and r["bz_share_ly"] == r["bz_share_ly"]:
        ov.append(f"Bazaar's share of your secondary sales is {fp(r['bz_share'])} against {fp(r['bz_share_ly'])} last year.")
    if r["bz_gap"] == r["bz_gap"] and r["bz_gap"] < 0:
        ov.append(f"Bazaar is {f1(-r['bz_gap'])} KL short of the {r['span']} target ({fp(r['bz_ach'])} achieved).")
    elif r["bz_gap"] == r["bz_gap"]:
        ov.append(f"Bazaar is {f1(r['bz_gap'])} KL ahead of the {r['span']} target ({fp(r['bz_ach'])} achieved).")
    out["overall"] = ov
    # trend
    tr = []
    qs = r["quarters"]
    if len(qs) >= 2:
        a, b = qs[-2], qs[-1]
        tr.append(f"{a['q']} {'grew' if a['g'] >= 0 else 'fell'} {fp(abs(a['g']))} vs LY; {b['q']}{'' if b['full'] else ' so far (' + b['label'] + ')'} "
                  f"{'grew' if b['g'] >= 0 else 'fell'} {fp(abs(b['g']))}.")
    g_ = r["m_g"]
    vs = ("against nil sale in" if g_ != g_ else "far above" if g_ >= 1000 else f"{fp(abs(g_))} {'below' if g_ < 0 else 'above'}")
    tr.append(f"{r['month']} sale of {f1(r['m_cy'])} KL is {vs} {r['month_ly']} ({f1(r['m_ly'])} KL)"
              + (f" and {fp(r['m_ach'])} of the month's target ({f1(r['m_t'])} KL)." if r["m_ach"] == r["m_ach"] else "."))
    out["trend"] = tr
    # counters
    ct = [f"Billed Bazaar counters: {r['cnt_cy']} against {r['cnt_ly']} last year ({r['cnt_cy'] - r['cnt_ly']:+d}, {fp(growth(r['cnt_cy'], r['cnt_ly']), True)}). "
          f"Counters billed in {r['month']}: {r['mcnt_cy']} against {r['mcnt_ly']} in {r['month_ly']}.",
          (f"{r['lost_n']} counters that bought in {r['span_ly']} have not been billed at all this year ({f1(r['lost_kl'])} KL of last year's volume). "
           if r["lost_n"] else f"Every counter billed in {r['span_ly']} has been billed again this year. ")
          + (f"Only {r['new_n']} new counters were added ({f1(r['new_kl'])} KL)." if r["new_n"] < max(r["lost_n"], 1) * 1.5
             else f"{r['new_n']} new counters were added ({f1(r['new_kl'])} KL)."),
          f"{r['dec_n']} continuing counters have reduced their offtake ({f1(r['dec_kl'])} KL); {r['gro_n']} have grown ({f1(r['gro_kl'], 1)} KL)."]
    if len(r["lost_top"]):
        ct.append("Major counters lost: " + ", ".join(f"{_short(n)} ({f1(v, 2)} KL)" for n, v in zip(r["lost_top"].customer.head(8), r["lost_top"].ly.head(8))) + ".")
    if len(r["dec_top"]):
        ct.append("Major counters with reduced offtake: " + ", ".join(f"{_short(n)} ({f1(v, 2)} KL)" for n, v in zip(r["dec_top"].customer.head(6), r["dec_top"].chg.head(6))) + ".")
    out["counters"] = ct
    # societies / drums
    sc = []
    if r["soc_ly"] >= 1 and r["soc_cy"] < r["soc_ly"] * 0.6:
        sc.append(f"Sales to co-operative / fishermen societies ({', '.join(_short(n, 28) for n in r['soc_names'][:5])}) fell from {f1(r['soc_ly'])} KL to {f1(r['soc_cy'])} KL.")
    if r["drum_ly"] >= 1 and r["drum_cy"] < r["drum_ly"] * 0.6:
        sc.append(f"Barrel (210 L) sales fell from {f1(r['drum_ly'])} KL to {f1(r['drum_cy'])} KL.")
    out["bulk"] = sc
    # mix
    mx = []
    g = r["groups"]
    for n, row in g[g.Change <= -0.5].head(4).iterrows():
        mx.append(f"{n}: {f1(row.LY)} → {f1(row.CY)} KL ({fp(row.Growth, True)}).")
    c = r["cats"]
    if "Mass" in c.index and "Aspirational" in c.index and c.loc["Aspirational", "Change"] < -1 and c.loc["Mass", "Change"] > 1:
        mx.append(f"The Aspirational range fell {f1(-c.loc['Aspirational', 'Change'])} KL while the Mass range rose {f1(c.loc['Mass', 'Change'])} KL. "
                  "Customers are being moved to cheaper grades rather than new volume being created.")
    pr = r["products"][r["products"].Change <= -0.5].head(4)
    if len(pr):
        mx.append("Biggest product drops: " + ", ".join(f"{_short(n, 32)} ({f1(v)} KL)" for n, v in pr.Change.items()) + ".")
    out["mix"] = mx
    # SSF
    sf = []
    for n, row in r["ssf"].iterrows():
        sf.append(f"{n}: {f1(row.LY)} → {f1(row.CY)} KL ({fp(row.Growth, True)}); {r['month']} {f1(row.M_CY)} KL vs {f1(row.M_LY)} KL; {int(row.Lost)} counters lost.")
    out["ssf"] = sf
    # geography / returns
    gr = []
    d = r["districts"]
    for n, row in d[(d.Change <= -1) & (d.LY >= 2) & (d.Growth < -5)].head(2).iterrows():
        gr.append(f"{n} district: {f1(row.LY)} → {f1(row.CY)} KL ({fp(row.Growth, True)}).")
    if r["returns_untagged"] >= 0.3:
        gr.append(f"{f1(r['returns_untagged'])} KL of Bazaar returns / credit notes are booked with no SSF tagged. Please explain these.")
    out["geo"] = gr
    # requirement
    rq = []
    if r["need_pm"] == r["need_pm"] and r["months_left"]:
        rq.append(f"To meet the annual Bazaar target of {f1(r['bz_annual'])} KL you need {f1(r['need'])} KL in the remaining {r['months_left']} months, "
                  f"i.e. {f1(r['need_pm'])} KL/month against the current run-rate of {f1(r['run_rate'])} KL/month "
                  f"({fp(growth(r['need_pm'], r['run_rate']), True)}). Last year's sale in the same months was {f1(r['ly_rest'])} KL.")
    out["need"] = rq
    out["severity"] = [sev]
    return out


def actions(r: dict, days: int = 5, new_counters: int | None = None) -> list[str]:
    nc = new_counters or max(20, int(round(max(r["cnt_ly"] - r["cnt_cy"], 0) / 2, -1)) or 20)
    if r["lost_n"]:
        big = f" (priority: the {r['lost_big_n']} counters with more than 100 L offtake last year)" if r["lost_big_n"] else ""
        a = [f"A counter-wise reactivation plan for the {r['lost_n']} lost counters{big} and the top {min(30, r['dec_n'])} declining counters, "
             "giving the responsible SSF and visit date for each."]
    else:
        a = [f"A counter-wise recovery plan for the top {min(30, r['dec_n'])} declining counters ({f1(r['dec_kl'])} KL lost on them), "
             "giving the responsible SSF and visit date for each."]
    if r["soc_ly"] >= 1 and r["soc_cy"] < r["soc_ly"] * 0.6:
        a.append("A plan to win back the co-operative / fishermen societies before the peak season, with joint visits with the undersigned this month.")
    a.append(f"A plan to add a minimum of {nc} new billing counters in the next quarter, with a focus on the weakest districts.")
    if r["need_pm"] == r["need_pm"]:
        a.append(f"SSF-wise, category-wise monthly Bazaar targets for the balance months, adding up to at least {f1(r['need_pm'])} KL/month.")
    weak = [n for n, row in r["ssf"].iterrows() if row.Growth == row.Growth and row.Growth < -5]
    if weak:
        a.append(f"A beat-productivity plan for {', '.join(weak)}: daily beat plan, minimum calls/day and weekly reporting to us.")
    a.append("A calendar of at least 3 Mechanic Meets and 2 Retailer Meets per month in the territory.")
    if r["returns_untagged"] >= 0.3:
        a.append(f"An explanation of the untagged Bazaar returns of {f1(r['returns_untagged'])} KL.")
    return a


# ------------------------------------------------------------------ Word letter
def _letterhead() -> Path | None:
    for p in sorted(C.DATA_DIR.glob("Letterhead*.docx")):
        if not p.name.startswith("~$"): return p
    return None


def letter_docx(r: dict, ref: str, date: str, address: str, prev_ref: str = "", days: int = 5, new_counters: int | None = None,
                signatory: str = "Amitesh Singh", designation: str = "Manager (Lubes – Reseller Sales), Mumbai Lubes Sales Area") -> bytes:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    lh = _letterhead()
    doc = Document(str(lh)) if lh else Document()
    if lh:                                         # keep the letterhead's header/footer, drop its body text
        body = doc.element.body
        for el in list(body):
            if el.tag.endswith("}sectPr"): continue
            body.remove(el)
    else:
        s = doc.sections[0]
        s.page_width, s.page_height = Cm(21), Cm(29.7)
        s.left_margin = s.right_margin = Cm(2); s.top_margin, s.bottom_margin = Cm(4.5), Cm(2)
    st_ = doc.styles["Normal"]; st_.font.name = "Arial"; st_.font.size = Pt(10.5); st_.paragraph_format.space_after = Pt(4)
    F, sev = findings(r), severity(r)

    def para(text="", bold=False, size=None, color=None, italic=False, style=None):
        p = doc.add_paragraph(style=style)
        if text:
            for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
                if not part: continue
                run = p.add_run(part); run.bold = bold or i % 2 == 1; run.italic = italic
                if size: run.font.size = Pt(size)
                if color: run.font.color.rgb = RGBColor.from_string(color)
        return p

    def head(text):
        p = para(text, bold=True, size=11, color="003087"); p.paragraph_format.space_before = Pt(8); return p

    def bullets(lines):
        for l in lines: para(l, style="List Bullet")

    def shade(cell, fill):
        pr = cell._tc.get_or_add_tcPr(); e = OxmlElement("w:shd")
        e.set(qn("w:val"), "clear"); e.set(qn("w:color"), "auto"); e.set(qn("w:fill"), fill); pr.append(e)

    def table(rows, bold_last=False):
        t = doc.add_table(rows=len(rows), cols=len(rows[0])); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                c = t.cell(i, j); p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(0)
                run = p.add_run(str(v)); run.font.size = Pt(9); run.bold = i == 0 or (bold_last and i == len(rows) - 1)
                if i == 0: shade(c, "D9E2F3")
        doc.add_paragraph()

    para(f"Date: {date}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    para(f"Ref.: {ref}")
    para("To,"); para(f"M/s. {r['stockist']} (SSR Code {r['code']})")
    if address: para(address)
    para(f"Subject: Review of Secondary Sales Performance – Bazaar Segment, {r['span']} ({r['fy_label']})", bold=True)
    para("Dear Channel Partner,"); para("Greetings from SERVO – World Class Lubricants!")
    intro = (f"This is further to our letter Ref. {prev_ref}. " if prev_ref else "")
    verdict = {"strong": "shows that the Bazaar position needs **immediate corrective action**.",
               "watch": "shows early signs of slippage in Bazaar that need your attention now.",
               "good": "shows healthy Bazaar performance. We appreciate your team's efforts."}[sev]
    para(intro + f"A review of your secondary sales for {r['span']} (compared with {r['span_ly']} and the target) {verdict}"
         + (f" Data is up to {r['data_till']:%d.%m.%Y}; last year is compared up to the same date." if r["partial"] else ""))

    head("1. Overall performance")
    o = r["overall"]
    rows = [[r["span"], f"{r['span_ly']} (KL)", "Target (KL)", f"{r['span']} (KL)", "Ach. vs Target", "Growth vs LY"]]
    for x in o.itertuples(): rows.append([x.Line, f1(x.LY), f1(x.Target), f1(x.CY), fp(x.Ach), fp(x.Growth, True)])
    table(rows, bold_last=True); bullets(F["overall"])

    head("2. Bazaar – month-wise trend")
    m = r["bz_monthly"]
    hdr = ["Bazaar (KL)"] + [MONTHS[f - 1] for f in m.index] + [q["q"] for q in r["quarters"]] + ["Total"]
    def line(lbl, s, key):
        return [lbl] + [f1(v) for v in s] + [f1(q[key]) for q in r["quarters"]] + [f1(s.sum() if s.notna().any() else np.nan)]
    table([hdr, line(r["fy_ly"], m.LY, "ly"), line(f"Target {r['fy_cy']}", m.Target, "t"), line(r["fy_cy"], m.CY, "cy")])
    bullets(F["trend"])

    head("3. Root causes identified from your secondary data")
    para("a) Counter coverage", bold=True); bullets(F["counters"])
    k = "b"
    if F["bulk"]:
        para(f"{k}) Bulk / institutional buyers", bold=True); bullets(F["bulk"]); k = chr(ord(k) + 1)
    if F["mix"]:
        para(f"{k}) Product mix", bold=True)
        g = r["groups"]; g = g[(g.LY >= 0.5) | (g.CY >= 0.5)].sort_values("Change")
        table([["Product group", r["span_ly"], r["span"], "Change", "Growth"]] +
              [[n, f1(x.LY), f1(x.CY), f"{x.Change:+.1f}", fp(x.Growth, True)] for n, x in g.iterrows()])
        bullets(F["mix"]); k = chr(ord(k) + 1)
    if len(r["ssf"]):
        para(f"{k}) Sales force (SSF) productivity", bold=True)
        table([["SSF", r["span_ly"], r["span"], "Growth", r["month_ly"], r["month"], "Counters lost"]] +
              [[n, f1(x.LY), f1(x.CY), fp(x.Growth, True), f1(x.M_LY), f1(x.M_CY), int(x.Lost)] for n, x in r["ssf"].iterrows()])
        k = chr(ord(k) + 1)
    if F["geo"]:
        para(f"{k}) Geography and returns", bold=True); bullets(F["geo"])

    if F["need"]:
        head("4. Balance-period requirement"); bullets(F["need"])
    head(f"5. You are hereby directed to submit, within {days} working days" if sev == "strong" else f"5. Please share, within {days} working days")
    for a in actions(r, days, new_counters): para(a, style="List Number")
    if r["oem_kl"] > 0 and r["bz_share"] == r["bz_share"] and r["bz_share"] < 50:
        para("OEM volumes are welcome, but they cannot replace a healthy retail network. They depend on a handful of customers "
             "and offer no counter presence or brand visibility in the market.")
    if sev == "strong":
        para("Continued Bazaar decline will be viewed seriously during the annual performance evaluation of the SSR.")
    para("We value our partnership and remain committed to supporting your team, including joint market visits, scheme support and BTL activities.")
    para("Kindly confirm receipt and the submission date of the action plan immediately." if sev != "good" else "Kindly keep up the momentum and share your plan for the coming month.")
    para("Happy SERVO Selling!"); para(""); para("Yours faithfully,"); para("For Indian Oil Corporation Limited"); para(""); para("")
    para(signatory, bold=True); para(designation)
    para(f"Data source: Secondary and Primary sales files and Target workbook ({r['span']} vs {r['span_ly']}). ClearBlue excluded, ServoXpert counted in Bazaar, "
         "Bazaar volume net of returns. A billed counter is a customer with positive sales in the period.", size=8, italic=True)
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


# ------------------------------------------------------------------ mail
def mail(r: dict, days: int = 5, signature: str = "", new_counters: int | None = None) -> dict[str, str]:
    F, sev = findings(r), severity(r)
    tone = {"strong": "Bazaar Sales Decline", "watch": "Bazaar Sales Slippage", "good": "Bazaar Sales Review"}[sev]
    subj = f"{tone} – {r['span']} – {'Immediate Action Plan Required' if sev == 'strong' else 'Action Plan'} | SSR {r['stockist']} ({r['code']})"
    L = ["Dear Sir,", "",
         f"Please find attached the review of your Bazaar secondary sales for {r['span']}.", "",
         f"- Bazaar: {f1(r['bz_cy'])} KL against a target of {f1(r['bz_t'])} KL ({fp(r['bz_ach'])}), {fp(r['bz_g'], True)} vs last year.",
         f"- Total secondary: {f1(r['sec_cy'])} KL ({fp(growth(r['sec_cy'], r['sec_ly']), True)} vs LY); Bazaar share {fp(r['bz_share'])} against {fp(r['bz_share_ly'])} last year."]
    L += [f"- {x}" for x in F["trend"]]
    L.append(f"- Billed counters: {r['cnt_cy']} against {r['cnt_ly']}; {r['lost_n']} counters lost, {r['new_n']} new counters added"
             + (f", {r['dec_n']} counters buying less ({f1(r['dec_kl'])} KL)." if r["dec_n"] else "."))
    L += [f"- {x}" for x in F["bulk"]]
    L += [f"- {x}" for x in F["mix"][:2]]
    L += [f"- {x}" for x in F["need"]]
    L += ["", f"You are directed to submit, within {days} working days:" if sev == "strong" else f"Please share, within {days} working days:"]
    L += [f"{i}. {a}" for i, a in enumerate(actions(r, days, new_counters), 1)]
    if sev == "strong": L += ["", "Continued Bazaar decline will be viewed seriously in the SSR's annual performance evaluation."]
    L += ["", "Please confirm receipt and your submission date by return mail.", "", "Regards,"]
    L += [s for s in signature.splitlines() if s.strip()] or ["Amitesh Singh", "Manager (Lubes – Reseller Sales), Mumbai Lubes Sales Area", "Indian Oil Corporation Limited"]
    return {"subject": subj, "body": "\n".join(L)}


def counter_lists_xlsx(r: dict) -> bytes:
    """Lost and declining counter lists (to attach / hand to the SSF)."""
    buf = io.BytesIO()
    cols = {"customer": "Customer", "ssf_name": "SSF", "district": "District", "ly": f"KL {r['span_ly']}", "cy": f"KL {r['span']}"}
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        r["lost_list"].reset_index().rename(columns={"customer_id": "Customer ID", **cols}).to_excel(w, sheet_name="Lost counters", index=False)
        r["dec_list"].reset_index().rename(columns={"customer_id": "Customer ID", "chg": "Change KL", **cols}).to_excel(w, sheet_name="Declining counters", index=False)
    return buf.getvalue()


def zip_all(files: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items(): z.writestr(n, b)
    return buf.getvalue()
