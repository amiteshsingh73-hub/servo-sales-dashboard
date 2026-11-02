"""HO LSE-interaction deck: fills the HO standard template slide by slide.

How it works
* The base deck is the HO template (blank) - or a deck you filled earlier, so manual entries carry over.
* Every slide is recognised by its TITLE, every column by its HEADING ("23-24", "Target 26-27 Prorata", "% Growth" ...),
  so a re-ordered or slightly changed template still lands the numbers in the right cells.
* Tables are resized to the real stockists (+ Total) - the template's SSR1..SSR4 / GSS1 rows are replaced.
* Cells that the sales files cannot provide (universe, plans, branding, action plans, ...) are left as they are in the
  base deck (blank in the template) for manual entry. % achievement / growth / balance / CAGR are worked out from the
  final row, so they also use targets typed in by hand.
* A slide that cannot be filled is left untouched and reported - the deck is always produced.
"""
from __future__ import annotations
import copy, datetime as dt, io, re
from pathlib import Path
import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt

from . import config as C, ho_tables as H, metrics as M, pptx_export as X
from .periods import fy_label, make_period, windows

NONE = None          # "leave this cell as it is"


# =============================================================== small helpers
def norm(s: str) -> str:
    s = str(s).lower().replace("–", "-").replace("—", "-").replace("’", "'")
    s = re.sub(r"\s*-\s*", "-", s)
    return re.sub(r"\s+", " ", s).strip()


def num(text):
    """Number in a cell ('1,095', '102%', '-', '15647 (est.)') -> float | None."""
    t = str(text).replace(",", "").strip()
    m = re.match(r"^-?\d+(\.\d+)?", t)
    return float(m.group(0)) if m else None


def fmt_v(v, d=None):
    if v is None or (isinstance(v, float) and not np.isfinite(v)): return "-"
    if d is not None: return f"{v:,.{d}f}".replace(",", "")
    a = abs(v)
    if a == 0: return "0"
    return f"{v:.0f}" if a >= 10 else f"{v:.1f}" if a >= 1 else f"{v:.2f}"


def fmt_p(v):
    if v is None or not np.isfinite(v) or abs(v) > 9999: return "-"
    return f"{round(v) + 0:.0f}%".replace("-0%", "0%")


def slide_title(slide) -> str:
    try:
        if slide.shapes.title is not None and slide.shapes.title.text.strip():
            return slide.shapes.title.text.strip()
    except Exception:
        pass
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            return sh.text_frame.text.strip()
    return ""


def main_table(slide):
    tbls = [sh for sh in slide.shapes if getattr(sh, "has_table", False) and sh.has_table]
    return max(tbls, key=lambda s: len(s.table.rows) * len(s.table.columns)).table if tbls else None


def set_text(cell, text):
    """Replace a cell's text but keep its font / colour / size (first run's formatting)."""
    if text is NONE: return
    tf = cell.text_frame
    p0 = tf.paragraphs[0]
    rpr = None
    if p0.runs: rpr = copy.deepcopy(p0.runs[0]._r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"))
    else:
        end = p0._p.find("{http://schemas.openxmlformats.org/drawingml/2006/main}endParaRPr")
        if end is not None:
            rpr = copy.deepcopy(end); rpr.tag = "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"
    for p in list(tf.paragraphs)[1:]: p._p.getparent().remove(p._p)
    for r in list(p0.runs): r._r.getparent().remove(r._r)
    run = p0.add_run(); run.text = str(text)
    if rpr is not None:
        old = run._r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
        if old is not None: run._r.remove(old)
        run._r.insert(0, rpr)


def header_texts(tbl, n_hdr: int) -> list[str]:
    """Combined heading text per column over the header rows (merged headings spread over their columns)."""
    ncol = len(tbl.columns)
    out = [""] * ncol
    for r in range(n_hdr):
        row = tbl.rows[r]
        j = 0
        while j < ncol:
            c = row.cells[j]
            if c.is_merge_origin:
                for k in range(j, min(ncol, j + c.span_width)): out[k] += " " + c.text.strip()
                j += c.span_width; continue
            if not c.is_spanned: out[j] += " " + c.text.strip()
            j += 1
    return [re.sub(r"\s+", " ", h).strip() for h in out]


def fit_rows(tbl, n_hdr: int, n_body: int):
    """Make the table have exactly n_body body rows (clone the first body row / drop extra rows; the last row keeps
    the template's 'Total' styling)."""
    trs = list(tbl._tbl.tr_lst)
    body = trs[n_hdr:]
    if not body: return
    proto, last = body[0], body[-1]
    while len(body) > n_body and len(body) > 1:
        victim = body[-2] if len(body) > 1 else body[-1]           # keep the styled last (total) row
        tbl._tbl.remove(victim); body.remove(victim)
    while len(body) < n_body:
        new = copy.deepcopy(proto); body[-1].addprevious(new); body.insert(len(body) - 1, new)


def old_rows(tbl, n_hdr):
    """Existing body rows keyed by their first-cell label (for keeping manual cells when the base is a filled deck)."""
    out = {}
    for r in list(tbl.rows)[n_hdr:]:
        cells = [c.text for c in r.cells]
        if cells and cells[0].strip(): out[_key(cells[0])] = cells
    return out


def _key(label) -> str:
    s = re.sub(r"[^a-z0-9]", "", str(label).lower())
    return s[:10]


def write_rows(tbl, n_hdr: int, rows: list[list], label_col=0, resize=True):
    """rows: list of cell values (str | None=keep). Keeps manual cells of the matching old row (same first-cell label)."""
    prev = old_rows(tbl, n_hdr)
    if resize: fit_rows(tbl, n_hdr, len(rows))
    body = list(tbl.rows)[n_hdr:]
    for r, vals in zip(body, rows):
        keep = prev.get(_key(vals[label_col])) if vals[label_col] else None
        for j, c in enumerate(r.cells):
            if c.is_spanned: continue
            v = vals[j] if j < len(vals) else NONE
            if v is NONE:
                v = keep[j] if keep is not None and j < len(keep) else ("" if resize else NONE)
            set_text(c, v)


# =============================================================== the numbers
class Calc:
    """Volumes for any slice of primary / secondary data, by financial year and for the selected period (CY / LY)."""

    def __init__(self, ctx: H.Ctx):
        self.ctx, self.fy, self.p = ctx, ctx.period.fy, ctx.period
        self.pri, self.sec, self.tg = ctx.pri, ctx.sec, ctx.targets
        self.wp, self.ws = ctx.wp, ctx.ws
        # a financial year is only reported when the files cover (nearly) all of it - a few stray invoices from an
        # older year must not overwrite the full-year figure typed in the deck
        full = lambda df: set(int(y) for y, n in df.groupby("fy").mi.nunique().items() if n >= 10)
        self.pri_years, self.sec_years = full(ctx.pri), full(ctx.sec)
        names = pd.concat([ctx.pri[["stockist_code", "stockist"]], ctx.sec[["stockist_code", "stockist"]]]).drop_duplicates("stockist_code")
        self.stockists = sorted(zip(names.stockist_code, names.stockist), key=lambda t: str(t[1]))
        self.fms = list(ctx.period.fms)
        self.months = max(len(self.fms), 1)

    # ---- volume helpers
    def fyv(self, df, y, col="kl", sec=False):
        if y not in (self.sec_years if sec else self.pri_years): return None
        return float(df.loc[df.fy == y, col].sum())

    def per(self, df, which="cy", sec=False, col="kl"):
        w = (self.ws if sec else self.wp)[which]
        return float(df.loc[(df.date >= w.start) & (df.date <= w.end), col].sum())

    def tgt(self, codes=None, fms=None, kind="Primary", ch=None):
        v = M.target_for(self.tg, codes if codes is not None else [c for c, _ in self.stockists], self.fy, fms or range(1, 13), kind, ch)
        return None if v != v else v

    def base(self, df, sec=False, target=None, prorata=None, col="kl"):
        """dict of the base values for one row; derived values are computed later from the final row."""
        d = {("fy", y): self.fyv(df, y, col, sec) for y in range(self.fy - 6, self.fy)}
        d.update(cy=self.per(df, "cy", sec, col), ly=self.per(df, "ly", sec, col), target=target, prorata=prorata)
        return d

    def soft(self, base: dict) -> dict:
        """Mark the targets of a row as apportioned estimates: a target already typed in the base deck is kept."""
        base["soft"] = True
        return base

    def apportion(self, part, whole, whole_target, whole_prorata, sec=False):
        """Target of a slice = parent's target x the slice's share of the parent's last full year."""
        a, b = self.fyv(part, self.fy - 1, sec=sec), self.fyv(whole, self.fy - 1, sec=sec)
        if not a or not b or whole_target is None: return None, None
        sh = a / b
        return whole_target * sh, (whole_prorata * sh if whole_prorata is not None else None)


# ---- column roles from headings
def year_of(h):
    m = re.findall(r"(?:20)?(\d{2})\s*-\s*(\d{2})", h)
    return 2000 + int(m[-1][1]) if m else None


def role(h: str, fy: int):
    h0 = norm(h)
    y = year_of(h0)
    if not h0 or "action" in h0: return None
    if "cagr" in h0: return "cagr"
    if "avg" in h0 and ("mth" in h0 or "month" in h0): return "avg"
    if "% tgt" in h0 or "% achievement" in h0 or "% target" in h0: return "ach"
    if "% growth" in h0: return "growth"
    if "% share" in h0 or "% of" in h0: return "share"
    if "balance" in h0: return "balance"
    if "pro rata" in h0 or "prorata" in h0: return "prorata"
    if "target" in h0:
        if y == fy and "apr" in h0: return "prorata"
        if y == fy or y is None and "bucket" not in h0: return "target"
        if "bucket" in h0: return "target"
        return None                                            # earlier years' targets are not in the files
    if "plan" in h0 or "annual" in h0:
        return "plan" if (y is None or y == fy) else None
    if ("ach" in h0 or "actual" in h0 or "sales" in h0) and y == fy: return "cy"
    if "apr" in h0 and y == fy - 1: return "ly"
    if y is not None and y < fy: return ("fy", y)
    return None


def row_cells(heads, roles, label_vals: dict, base: dict, fy: int, old: list | None, pct_digits=0, share=None):
    """Turn base values into cell texts (None = keep). Derived columns use the final (kept or computed) cell values."""
    out = []
    for j, h in enumerate(heads):
        if j in label_vals: out.append(label_vals[j]); continue
        r = roles[j]
        v = NONE
        if isinstance(r, tuple): v = base.get(r)
        elif r in ("cy", "ly"): v = base.get(r)
        elif r in ("target", "prorata"):
            typed = num(old[j]) if old is not None and j < len(old) else None
            v = NONE if (base.get("soft") and typed is not None) else base.get(r)
        elif r == "share": v = NONE if share is None else ("SHARE", share)
        out.append(v)

    def final(rl):
        for j, r in enumerate(roles):
            if r == rl:
                if out[j] is not NONE and not isinstance(out[j], str) and out[j] is not None: return out[j]
                t = out[j] if isinstance(out[j], str) else (old[j] if old is not None and j < len(old) else "")
                return num(t)
        return None

    tgt, pro, cy, ly = final("target"), final("prorata"), final("cy"), final("ly")
    t_ix = next((j for j, r in enumerate(roles) if r == "target"), None)
    tgt_from_data = t_ix is not None and out[t_ix] is not NONE and base.get("target") is not None
    yrs = sorted((r[1], j) for j, r in enumerate(roles) if isinstance(r, tuple))
    yv = [(y, out[j] if (out[j] is not NONE and not isinstance(out[j], str)) else num(old[j]) if old is not None and j < len(old) else None)
          for y, j in yrs]
    yv = [(y, v) for y, v in yv if v is not None]
    for j, r in enumerate(roles):
        if j in label_vals: continue
        v = out[j]
        if r == "ach": out[j] = fmt_p(cy / pro * 100) if cy is not None and pro else "-"
        elif r == "growth": out[j] = fmt_p((cy / ly - 1) * 100) if cy is not None and ly and abs(ly) >= 0.5 else "-"   # no % on a near-zero base
        elif r == "balance": out[j] = fmt_v(tgt - cy) if tgt is not None and cy is not None else "-"
        elif r == "plan": out[j] = fmt_v(tgt) if tgt_from_data and tgt is not None else NONE    # typed plans are kept
        elif r == "avg":
            last = base.get(("fy", fy - 1))
            out[j] = fmt_v(last / 12) if last is not None else "-"
        elif r == "cagr":
            pos = [(y, v) for y, v in yv if v and v > 0]
            out[j] = f"{((pos[-1][1] / pos[0][1]) ** (1 / (pos[-1][0] - pos[0][0])) - 1) * 100:.1f}" if len(pos) >= 2 and pos[-1][0] > pos[0][0] else "-"
        elif isinstance(v, tuple) and v and v[0] == "SHARE": out[j] = f"{v[1]:.1f}%" if v[1] == v[1] else "-"
        elif v is NONE: out[j] = NONE
        elif v is None: out[j] = "-"
        elif not isinstance(v, str): out[j] = fmt_v(v)
    return out


def fill_generic(tbl, n_hdr, calc: Calc, rows: list[tuple], fix_headers=None):
    """rows: [(label, base dict, share | None)]. Writes all rows with heading-driven columns."""
    heads = header_texts(tbl, n_hdr)
    if fix_headers: heads = fix_headers(heads)
    roles = [None] + [role(h, calc.fy) for h in heads[1:]]
    prev = old_rows(tbl, n_hdr)
    out = []
    for label, base, share in rows:
        old = prev.get(_key(label))
        out.append(row_cells(heads, roles, {0: label}, base, calc.fy, old, share=share))
    write_rows(tbl, n_hdr, out)


# =============================================================== slide fillers
def _stockist_rows(calc: Calc, df, sec=False, kind="Primary", ch=None, label_total="Total", target_from=None):
    rows, tot_t, tot_p = [], 0.0, 0.0
    for code, name in calc.stockists:
        d = df[df.stockist_code == code]
        if target_from == "apportion_primary":
            whole = calc.pri[calc.pri.stockist_code == code]
            t, p = calc.apportion(d, whole, calc.tgt([code]), calc.tgt([code], calc.fms))
        else:
            t, p = calc.tgt([code], None, kind, ch), calc.tgt([code], calc.fms, kind, ch)
        b_ = calc.base(d, sec, t, p)
        rows.append((name, calc.soft(b_) if target_from == "apportion_primary" else b_, None))
        tot_t = tot_t + t if t is not None and tot_t is not None else None
        tot_p = tot_p + p if p is not None and tot_p is not None else None
    bt = calc.base(df[df.stockist_code.isin([c for c, _ in calc.stockists])], sec, tot_t, tot_p)
    rows.append((label_total, calc.soft(bt) if target_from == "apportion_primary" else bt, None))
    return rows


def f_stockist_primary(slide, tbl, calc):
    fill_generic(tbl, 2, calc, _stockist_rows(calc, calc.pri))


def f_stockist_ta(slide, tbl, calc):
    fill_generic(tbl, 3, calc, _stockist_rows(calc, calc.pri))


def f_segment(slide, tbl, calc):
    pri, T, P = calc.pri, calc.tgt(), calc.tgt(None, calc.fms)
    rows = []
    for lab, seg in (("SERVO- AUTO", "Servo - Automotive"), ("SERVO-IND", "Servo - Industrial"), ("CO-BRANDED", "Co-branded (Genuine)")):
        d = pri[pri.hseg == seg]; t, p = calc.apportion(d, pri, T, P)
        rows.append((lab, calc.soft(calc.base(d, False, t, p)), None))
    rows.append(("TOTAL", calc.base(pri, False, T, P), None))
    fill_generic(tbl, 2, calc, rows)


def f_premium(slide, tbl, calc):
    prem = calc.pri[calc.pri.category == "Premium"]
    rows = _stockist_rows(calc, prem, target_from="apportion_primary")
    out = []
    for (label, base, _), code in zip(rows, [c for c, _ in calc.stockists] + [None]):
        whole = calc.pri if code is None else calc.pri[calc.pri.stockist_code == code]
        part = prem if code is None else prem[prem.stockist_code == code]
        a, b = calc.fyv(part, calc.fy - 1), calc.fyv(whole, calc.fy - 1)
        out.append((label if code is not None else rows[-1][0], base, (a / b * 100) if a is not None and b else float("nan")))
    fill_generic(tbl, 2, calc, out)


CAT_COLS = ["MASS", "Aspirational", "Premium", "Silver", "Industrial", "TOTAL"]


def _cat(df, which):
    c = df.category.astype(str)
    m = {"MASS": c.eq("Mass"), "Aspirational": c.eq("Aspirational"), "Premium": c.eq("Premium"), "Silver": c.eq("Silver")}
    m["Industrial"] = ~(m["MASS"] | m["Aspirational"] | m["Premium"] | m["Silver"])
    return {k: float(df.loc[v, "kl"].sum()) for k, v in m.items()} | {"TOTAL": float(df.kl.sum())}


def f_category(slide, tbl, calc):
    rows = []
    for code, name in calc.stockists + [(None, "Total")]:
        d = calc.pri if code is None else calc.pri[calc.pri.stockist_code == code]
        ly = _cat(d[d.fy == calc.fy - 1], None)
        w = calc.wp["cy"]; cy = _cat(d[(d.date >= w.start) & (d.date <= w.end)], None)
        vals = [name]
        vals += [fmt_v(ly[k]) for k in CAT_COLS] + [fmt_p(ly[k] / ly["TOTAL"] * 100) if ly["TOTAL"] else "-" for k in CAT_COLS]
        vals += [fmt_v(cy[k]) for k in CAT_COLS] + [fmt_p(cy[k] / cy["TOTAL"] * 100) if cy["TOTAL"] else "-" for k in CAT_COLS]
        rows.append(vals)
    write_rows(tbl, 2, rows)


def _is_hyper(df): return df["product"].astype(str).str.upper().str.contains("HYPER", regex=False)


def f_hyper(slide, tbl, calc):
    w, ws = calc.wp["cy"], calc.ws["cy"]
    rows = []
    for code, name in calc.stockists + [(None, "Total")]:
        p = calc.pri if code is None else calc.pri[calc.pri.stockist_code == code]
        s = calc.sec if code is None else calc.sec[calc.sec.stockist_code == code]
        p = p[_is_hyper(p) & (p.date >= w.start) & (p.date <= w.end)]
        s = s[_is_hyper(s) & (s.date >= ws.start) & (s.date <= ws.end)]
        ro, bz = s[s.channel == "RO"], s[s.channel == "Bazaar"]
        n = lambda d: d[d.kl > 0].customer_id.nunique()
        rows.append([name, fmt_v(float(p.kl.sum())), NONE, NONE, str(n(ro)), str(n(bz)), str(n(ro) + n(bz)),
                     fmt_v(float(ro.kl.sum())), fmt_v(float(bz.kl.sum())), fmt_v(float(ro.kl.sum() + bz.kl.sum())), NONE, NONE, NONE])
    write_rows(tbl, 2, rows)


GEN_BRANDS = [("Maruti", ["MARUTI"]), ("Hyundai", ["HYUNDAI"]), ("M&M", ["M&M", "MAHINDRA", "MNM"]), ("KIA", ["KIA"]),
              ("Honda", ["HONDA"]), ("TVS", ["TVS"]), ("ALT", ["ALT"]), ("MG Motors", ["MG MOTORS", "MG"])]


def f_genuine(slide, tbl, calc):
    pri, T, P = calc.pri, calc.tgt(), calc.tgt(None, calc.fms)
    gen = pri[pri.product_group == "Genuine Oil"]; spec = gen.spec.astype(str).str.upper().str.strip()
    rows = []

    def add(lab, d):
        t, p = calc.apportion(d, pri, T, P); rows.append((lab, calc.soft(calc.base(d, False, t, p)), None))
    add("Servo-Automotive", pri[pri.hseg == "Servo - Automotive"])
    add("Servo-Industrial", pri[pri.hseg == "Servo - Industrial"])
    add("Servo Total (A)", pri[pri.hseg != "Co-branded (Genuine)"])
    used = pd.Series(False, index=gen.index)
    for lab, keys in GEN_BRANDS:
        m = spec.isin(keys) & ~used; used |= m; add(lab, gen[m])
    add("Others", gen[~used])
    add("Gen. Oil Total (B)", gen)
    rows.append(("Gr Total (A+B)", calc.base(pri, False, T, P), None))

    def fix(heads):                       # the template repeats '2024-25'; the four year columns are the last four FYs
        yc = [j for j, h in enumerate(heads) if re.fullmatch(r"20\d{2}-\d{2}", h.strip())]
        for k, j in enumerate(yc): heads[j] = fy_label(calc.fy - len(yc) + k).replace("FY ", "")
        for k, j in enumerate(yc): set_text(tbl.cell(0, j), heads[j])
        return heads
    fill_generic(tbl, 1, calc, rows, fix_headers=fix)


def f_pgroup(slide, tbl, calc):
    pri, T, P = calc.pri, calc.tgt(), calc.tgt(None, calc.fms)
    rows = []

    def add(lab, groups):
        d = pri[pri.product_group.isin(groups)]; t, p = calc.apportion(d, pri, T, P)
        rows.append((lab, calc.soft(calc.base(d, False, t, p)), None))
    auto = ["4T Oil", "PCMO", "DEO", "Gear Oil", "Agri", "Grease", "Speciality"]
    ind = ["System 68", "Hydrashakti 68", "Other Industrial"]
    for lab, g in (("4T OIL", "4T Oil"), ("PCMO", "PCMO"), ("DEO", "DEO"), ("Gear Oil", "Gear Oil"), ("Agri product", "Agri"),
                   ("Grease", "Grease"), ("Speciality", "Speciality")): add(lab, [g])
    add("Autom. Oil Total", auto)
    for g in ind: add(g, [g])
    add("Total Industrial", ind)
    add("2T Oil", ["2T Oil"]); add("Genuine Oil", ["Genuine Oil"])
    rows.append(("Total channel", calc.base(pri, False, T, P), None))
    fill_generic(tbl, 1, calc, rows)


def f_pack(slide, tbl, calc):
    pri, T, P = calc.pri, calc.tgt(), calc.tgt(None, calc.fms)
    rows = []
    for g in C.PACK_GROUP_ORDER:
        d = pri[pri.pack_group == g]; t, p = calc.apportion(d, pri, T, P)
        rows.append((g, calc.base(d, False, t, p), None))
    rows.append(("Gr Total", calc.base(pri, False, T, P), None))
    fill_generic(tbl, 1, calc, rows)


def f_premium_plan(slide, tbl, calc):
    prem = calc.pri[calc.pri.category == "Premium"]
    fill_generic(tbl, 2, calc, _stockist_rows(calc, prem, target_from="apportion_primary"))


BUCKET_PACKS = ["> 6 ltr and < 15 ltr", "> 15 ltr to 25 ltr", "26 ltr"]      # buckets: 7 to 26 L (as in the HO deck)


def f_bucket(slide, tbl, calc):
    pri, T, P = calc.pri, calc.tgt(), calc.tgt(None, calc.fms)
    b = pri[pri.pack_group.isin(BUCKET_PACKS)]
    groups = [("DEO", ["DEO"]), ("GEAR OIL", ["Gear Oil"]), ("GREASE", ["Grease"]), ("HYDRAULIC (HYDRA+SY68)", ["Hydrashakti 68", "System 68"]),
              ("OTHER IND OIL", ["Other Industrial"]), ("AGRO OIL", ["Agri"])]
    used = []
    rows = []
    for lab, g in groups:
        d = b[b.product_group.isin(g)]; used += g
        whole = pri[pri.product_group.isin(g)]
        a, w = calc.fyv(d, calc.fy - 1), calc.fyv(whole, calc.fy - 1)
        t, p = calc.apportion(d, pri, T, P)
        rows.append((lab, calc.soft(calc.base(d, False, t, p)), a / w * 100 if a is not None and w else float("nan")))
    d = b[~b.product_group.isin(used)]; t, p = calc.apportion(d, pri, T, P)
    rows.append(("OTHER GRADE", calc.soft(calc.base(d, False, t, p)), float("nan")))
    t, p = calc.apportion(b, pri, T, P)
    a, w = calc.fyv(b, calc.fy - 1), calc.fyv(pri, calc.fy - 1)
    rows.append(("Total", calc.soft(calc.base(b, False, t, p)), a / w * 100 if a is not None and w else float("nan")))
    fill_generic(tbl, 2, calc, rows)


def f_sec_channel(slide, tbl, calc):
    rows = []
    for ch in C.CHANNEL_ORDER:
        d = calc.sec[calc.sec.channel == ch]
        rows.append((ch, calc.base(d, True, calc.tgt(None, None, "Secondary", ch), calc.tgt(None, calc.fms, "Secondary", ch)), None))
    rows.append(("Total", calc.base(calc.sec[calc.sec.channel.isin(C.CHANNEL_ORDER)], True,
                                    calc.tgt(None, None, "Secondary"), calc.tgt(None, calc.fms, "Secondary")), None))
    fill_generic(tbl, 2, calc, rows)


def _sec_by_stockist(ch):
    def f(slide, tbl, calc):
        d = calc.sec[calc.sec.channel == ch]
        fill_generic(tbl, 2, calc, _stockist_rows(calc, d, sec=True, kind="Secondary", ch=ch))
    return f


OEM_BRANDS = [("Maruti", ["MARUTI"]), ("Hyundai", ["HYUNDAI"]), ("Kia", ["KIA"]), ("MG Motors", ["MG"]), ("MNM- EM", ["MAHINDRA", "M&M", "MNM"])]


def _oem_slices(calc):
    oem = calc.sec[calc.sec.channel == "OEM"]
    sub = oem.cust_subgroup.astype(str).str.upper()
    used = pd.Series(False, index=oem.index)
    out = []
    for lab, keys in OEM_BRANDS:
        m = sub.str.contains("|".join(re.escape(k) for k in keys)) & ~used; used |= m
        out.append((lab, oem[m]))
    return oem, out, oem[~used]


def f_oem(slide, tbl, calc):
    oem, slices, rest = _oem_slices(calc)
    T, P = calc.tgt(None, None, "Secondary", "OEM"), calc.tgt(None, calc.fms, "Secondary", "OEM")
    rows = []
    for lab, d in slices + ([("Others", rest)] if len(rest) else []):
        t, p = calc.apportion(d, oem, T, P, sec=True); rows.append((lab, calc.soft(calc.base(d, True, t, p)), None))
    rows.append(("Total", calc.base(oem, True, T, P), None))
    fill_generic(tbl, 1, calc, rows)


def f_oem_pen(slide, tbl, calc):
    oem, slices, rest = _oem_slices(calc)
    T, P = calc.tgt(None, None, "Secondary", "OEM"), calc.tgt(None, calc.fms, "Secondary", "OEM")
    heads = header_texts(tbl, 1)
    prev = old_rows(tbl, 1)
    rows = []
    for lab, d in slices + [("Total", oem)]:
        t, p = (T, P) if lab == "Total" else calc.apportion(d, oem, T, P, sec=True)
        base = {"target": t, "prorata": p, "cy": calc.per(d, "cy", True), ("fy", calc.fy - 1): calc.fyv(d, calc.fy - 1, sec=True)}
        roles = [None]
        for h in heads[1:]:
            n = norm(h)
            roles.append(("fy", calc.fy - 1) if n.startswith("sales") else "cy" if n.startswith("actual") else
                         "plan" if n.startswith("plan") else role(h, calc.fy) if ("target" in n or "% tgt" in n) else None)
        rows.append(row_cells(heads, roles, {0: lab}, base, calc.fy, prev.get(_key(lab))))
    write_rows(tbl, 1, rows)


def f_summary(slide, tbl, calc):
    """HO slide 'TARGET VS ACHIEVEMENT' - 15 fixed parameter rows."""
    pri, sec, T, P = calc.pri, calc.sec, calc.tgt(), calc.tgt(None, calc.fms)
    heads = header_texts(tbl, 1)
    roles = [None, None, None] + [role(h, calc.fy) for h in heads[3:]]
    rows_old = [[c.text for c in r.cells] for r in list(tbl.rows)[1:]]

    def ch(c): return sec[sec.channel == c]

    def counters_pcs(y=None, which=None):
        bz = sec[sec.channel == "Bazaar"]
        if y is not None:
            if y not in calc.sec_years: return None
            d = bz[bz.fy == y]; n = _ctr(d)[_ctr(d).kl > 0].customer_id.nunique()
            return d.kl.sum() * 1000 / n / 12 if n else None
        w = calc.ws[which]; d = bz[(bz.date >= w.start) & (bz.date <= w.end)]
        n = _ctr(d)[_ctr(d).kl > 0].customer_id.nunique()
        return d.kl.sum() * 1000 / n / calc.months if n else None

    def new_counters(y=None, which=None):
        bz = _ctr(calc.ctx.sec_all[calc.ctx.sec_all.channel == "Bazaar"])
        first = bz[bz.kl > 0].groupby("customer_id").date.min()
        start = first.min() if len(first) else None
        if y is not None:
            ys, ye = pd.Timestamp(year=y - 1, month=4, day=1), pd.Timestamp(year=y, month=3, day=31)
            if start is None or ys < start + pd.DateOffset(months=11): return None       # needs history before that year
            return int(((first >= ys) & (first <= ye)).sum())
        w = calc.ws[which]
        return int(((first >= w.start) & (first <= w.end)).sum())

    def grp(g): return pri[pri.product_group == g]
    spec = {
        1: ("KL", lambda: calc.base(pri, False, T, P)),
        2: ("KL", lambda: calc.base(ch("RO"), True, calc.tgt(None, None, "Secondary", "RO"), calc.tgt(None, calc.fms, "Secondary", "RO"))),
        3: ("KL", lambda: calc.base(ch("Bazaar"), True, calc.tgt(None, None, "Secondary", "Bazaar"), calc.tgt(None, calc.fms, "Secondary", "Bazaar"))),
        4: ("KL", lambda: calc.base(ch("OEM"), True, calc.tgt(None, None, "Secondary", "OEM"), calc.tgt(None, calc.fms, "Secondary", "OEM"))),
        5: ("KL", lambda: calc.base(ch("Fleet"), True, calc.tgt(None, None, "Secondary", "Fleet"), calc.tgt(None, calc.fms, "Secondary", "Fleet"))),
        7: ("LTRS PER MONTH", lambda: {**{("fy", y): counters_pcs(y) for y in range(calc.fy - 6, calc.fy)}, "cy": counters_pcs(None, "cy"),
                                       "ly": counters_pcs(None, "ly")}),
        8: ("NO", lambda: {**{("fy", y): new_counters(y) for y in range(calc.fy - 6, calc.fy)}, "cy": new_counters(None, "cy"),
                           "ly": new_counters(None, "ly")}),
        9: ("KL", lambda: calc.soft(calc.base(pri[pri.category == "Premium"], False, *calc.apportion(pri[pri.category == "Premium"], pri, T, P)))),
        10: ("MT", lambda: calc.soft(calc.base(grp("Grease"), False, *calc.apportion(grp("Grease"), pri, T, P), col="mt"))),
        11: ("KL", lambda: calc.soft(calc.base(grp("4T Oil"), False, *calc.apportion(grp("4T Oil"), pri, T, P)))),
        12: ("KL", lambda: calc.soft(calc.base(grp("DEO"), False, *calc.apportion(grp("DEO"), pri, T, P)))),
        13: ("KL", lambda: calc.soft(calc.base(grp("PCMO"), False, *calc.apportion(grp("PCMO"), pri, T, P)))),
        14: ("KL", lambda: calc.soft(calc.base(grp("Gear Oil"), False, *calc.apportion(grp("Gear Oil"), pri, T, P)))),
        15: ("KL", lambda: calc.soft(calc.base(grp("Agri"), False, *calc.apportion(grp("Agri"), pri, T, P)))),
    }
    labels = {1: "PRIMARY SALES- CHANNEL", 14: "PRODUCT SEGMENT SALES- GEAR OIL"}
    out = []
    for i, old in enumerate(rows_old, 1):
        if i not in spec:
            out.append([NONE] * len(heads)); continue
        unit, fn = spec[i]
        try:
            base = fn()
        except Exception:
            out.append([NONE] * len(heads)); continue
        if i == 10 and not np.isfinite(calc.pri.mt.sum() if "mt" in calc.pri else np.nan):
            unit = "KL"; base = calc.base(grp("Grease"), False, *calc.apportion(grp("Grease"), pri, T, P))
        r = row_cells(heads, roles, {0: str(i), 1: labels.get(i, NONE), 2: unit}, base, calc.fy, old)
        if i in (7, 8):                 # integers / litres: no decimals
            r = [(f"{num(v):.0f}" if isinstance(v, str) and num(v) is not None and j >= 3 and "%" not in v else v) for j, v in enumerate(r)]
        out.append(r)
    write_rows(tbl, 1, out, resize=False)


def f_area(slide, tbl, calc):
    last = calc.fyv(calc.pri, calc.fy - 1)
    for r in tbl.rows:
        lab = norm(r.cells[0].text)
        if lab.startswith("total stockist"): set_text(r.cells[1], f"{len(calc.stockists):02d}")
        elif lab.startswith("average sales"):
            if last is not None: set_text(r.cells[1], f"{last / 12:.0f} KL ({fy_label(calc.fy - 1, True)})")


def f_industry(slide, tbl, calc):
    heads = [norm(h) for h in header_texts(tbl, 1)]
    for r in list(tbl.rows)[1:]:
        if norm(r.cells[0].text) not in ("iocl", "iocl (servo)", "iol"): continue
        for j, h in enumerate(heads):
            y = year_of(h)
            if h.startswith("sales in") and y in calc.pri_years:
                set_text(r.cells[j], fmt_v(calc.fyv(calc.pri, y, "mt")))
            elif h.startswith("ssr"):
                set_text(r.cells[j], str(len(calc.stockists)))
            elif h.startswith("sales/stockist"):
                v = calc.fyv(calc.pri, calc.fy - 1)
                if v is not None: set_text(r.cells[j], f"{v / max(len(calc.stockists), 1) / 12:.1f}")


def f_profile(slide, tbl, calc):
    heads = [norm(h) for h in header_texts(tbl, 1)]
    rows = []
    for code, name in calc.stockists:
        v = calc.fyv(calc.pri[calc.pri.stockist_code == code], calc.fy - 1)
        r = [name] + [NONE] * (len(heads) - 1)
        for j, h in enumerate(heads):
            if h.startswith("type"): r[j] = NONE
            if h.startswith("average sales"): r[j] = fmt_v(v / 12) if v is not None else "-"
        rows.append(r)
    write_rows(tbl, 1, rows)


def _bz(calc, code=None):
    d = calc.sec[calc.sec.channel == "Bazaar"]
    return d if code is None else d[d.stockist_code == code]


def _ctr(d):
    """Rows that count as Bazaar COUNTERS: ServoXpert workshops are left out of counter numbers (they have their own
    IG slide) but their volume stays in Bazaar sales - same convention as the HO deck."""
    return d[d.cust_group.astype(str).str.upper() != "SERVOXPERT"]


def _serviced(d, y):
    c = _ctr(d)
    return c[(c.fy == y) & (c.kl > 0)].customer_id.nunique()


def f_bz_ssr(slide, tbl, calc, district=False):
    n_hdr = 4
    prev = old_rows(tbl, n_hdr)
    off = 1 if district else 0
    rows = []
    for code, name in calc.stockists + [(None, "SO Total" if district else "Total")]:
        d = _bz(calc, code)
        lab0 = name
        old = prev.get(_key(name)) if not district else None
        if district:
            dist = d[d.fy == calc.fy - 1].groupby("district").kl.sum().sort_values(ascending=False)
            dlab = " + ".join(str(x) for x in dist.index[:2]) if code is not None else "SO Total"
            old = next((v for v in prev.values() if len(v) > 1 and _key(v[1]) == _key(name)), None) or prev.get(_key(dlab))
        u = [num(old[off + k]) if old is not None else None for k in (1, 2, 3)]
        r = [NONE] * len(tbl.columns)
        if district:
            r[0], r[1] = (dlab if code is not None else "SO Total"), (name if code is not None else "")
            if old is not None and old[0].strip() and code is not None: r[0] = NONE   # keep a district name typed by hand
        else:
            r[0] = name
        vals = []
        for y in (calc.fy - 2, calc.fy - 1):
            vals += [(_serviced(d, y) if y in calc.sec_years else None), (calc.fyv(d, y, sec=True))]
        for k, v in enumerate(vals):
            r[off + 4 + k] = (str(int(v)) if k % 2 == 0 else fmt_v(v)) if v is not None else NONE
        uni_c = (u[0] or 0) + (u[1] or 0) if (u[0] or u[1]) else None
        for k, (cnt, vol) in enumerate(((vals[0], vals[1]), (vals[2], vals[3]))):
            r[off + 8 + 2 * k] = fmt_p(cnt / uni_c * 100) if cnt is not None and uni_c else NONE
            r[off + 9 + 2 * k] = fmt_p(vol / u[2] * 100) if vol is not None and u[2] else NONE
        rows.append(r)
    if district:                                      # keep universe / plan columns of a filled deck row by row
        write_rows(tbl, n_hdr, rows, label_col=1)
    else:
        write_rows(tbl, n_hdr, rows)


def f_bz_dist(slide, tbl, calc): f_bz_ssr(slide, tbl, calc, district=True)


def _lubes_master_bazaar() -> pd.DataFrame:
    """Active Bazaar counters per stockist from the Lubes Customer Master (optional file)."""
    files = [f for f in C.DATA_DIR.glob("*") if re.search(r"customer master", f.name, re.I) and f.suffix.lower() in (".xlsx", ".xls")
             and not f.name.startswith(("~$", "."))]
    if not files: return pd.DataFrame()
    try:
        raw = pd.read_excel(max(files, key=lambda f: f.stat().st_mtime), header=None, dtype=str, engine="calamine")
        h = next(i for i in range(min(25, len(raw))) if "Customer Id" in raw.iloc[i].astype(str).str.strip().tolist())
        m = raw.iloc[h + 1:].copy(); m.columns = [str(c).strip() for c in raw.iloc[h]]
        m = m[m["Customer Group"].astype(str).str.upper().isin(["BAZAAR", "SERVOXPERT", "OTHERS"])]
        m = m[~m.get("Customer Status", pd.Series("", index=m.index)).astype(str).str.upper().isin(["CANCELLED"])]
        return pd.DataFrame({"cid": m["Customer Id"].astype(str).str.strip(), "code": pd.to_numeric(m["Stockist ID"], errors="coerce")})
    except Exception:
        return pd.DataFrame()


def f_bz_detail(slide, tbl, calc):
    heads = [norm(h) for h in header_texts(tbl, 2)]
    master = _lubes_master_bazaar()
    hist = _ctr(calc.ctx.sec_all[calc.ctx.sec_all.channel == "Bazaar"])
    first_all = hist[hist.kl > 0].groupby("customer_id").date.min()
    data_start = first_all.min() if len(first_all) else None
    rows = []
    y1 = calc.fy - 1
    for code, name in calc.stockists + [(None, "Total")]:
        d = _ctr(_bz(calc, code))
        r = [name] + [NONE] * (len(heads) - 1)
        s_prev = set(d[(d.fy == y1 - 1) & (d.kl > 0)].customer_id); s_cur = set(d[(d.fy == y1) & (d.kl > 0)].customer_id)
        first = d[d.kl > 0].groupby("customer_id").date.min()
        ys = pd.Timestamp(year=y1 - 1, month=4, day=1)
        for j, h in enumerate(heads):
            y = year_of(h)
            if "serviced" in h and y is not None and "target" not in h:
                r[j] = str(_serviced(d, y)) if y in calc.sec_years else NONE
            elif "inactive" in h:
                r[j] = str(len(s_prev - s_cur)) if (y1 - 1) in calc.sec_years else NONE
            elif "new counters" in h:
                r[j] = str(int(((first >= ys) & (first <= pd.Timestamp(year=y1, month=3, day=31))).sum())) \
                    if data_start is not None and ys >= data_start + pd.DateOffset(months=11) else NONE
            elif "nil selling" in h:
                if len(master):
                    ids = set(master[master.code == code].cid) if code is not None else set(master[master.code.isin([c for c, _ in calc.stockists])].cid)
                    r[j] = str(len(ids - s_cur))
            elif "net addition" in h:
                r[j] = str(len(s_cur) - len(s_prev)) if (y1 - 1) in calc.sec_years else NONE
        rows.append(r)
    write_rows(tbl, 2, rows)


def f_bz_pcs(slide, tbl, calc):
    heads = [norm(h) for h in header_texts(tbl, 2)]
    rows = []
    for code, name in calc.stockists + [(None, "TOTAL")]:
        d = _bz(calc, code)
        r = [name] + [NONE] * (len(heads) - 1)
        for j, h in enumerate(heads):
            y = year_of(h)
            if "target" in h or "lse area" in h or "action" in h: continue
            if y is not None and y < calc.fy:
                dd = d[d.fy == y]; n = _ctr(dd)[_ctr(dd).kl > 0].customer_id.nunique()
                r[j] = f"{dd.kl.sum() * 1000 / n / 12:.0f}" if n and y in calc.sec_years else NONE
            elif "ach" in h:
                w = calc.ws["cy"]; dd = d[(d.date >= w.start) & (d.date <= w.end)]
                n = _ctr(dd)[_ctr(dd).kl > 0].customer_id.nunique()
                r[j] = f"{dd.kl.sum() * 1000 / n / calc.months:.0f}" if n else "-"
        rows.append(r)
    write_rows(tbl, 2, rows)


def f_bz_track(slide, tbl, calc):
    heads = [norm(h) for h in header_texts(tbl, 2)]
    w = calc.ws["cy"]
    hist = _ctr(calc.ctx.sec_all[calc.ctx.sec_all.channel == "Bazaar"])
    first_all = hist[hist.kl > 0].groupby("customer_id").date.min()
    rows = []
    for code, name in calc.stockists + [(None, "Total")]:
        d = _ctr(_bz(calc, code))
        cy = d[(d.date >= w.start) & (d.date <= w.end)]
        a = set(d[(d.fy == calc.fy - 1) & (d.kl > 0)].customer_id)
        act = cy.groupby("customer_id").agg(kl=("kl", "sum"), inv=("invoice", "nunique"), sku=("sku_name", "nunique"))
        act = act[act.kl > 0]
        lit = act.kl * 1000
        newc = set(first_all[(first_all >= w.start) & (first_all <= w.end)].index) & set(act.index)
        prem_cy = cy[cy.category == "Premium"].groupby("customer_id").kl.sum()
        prem_ly = d[(d.fy == calc.fy - 1) & (d.category == "Premium")].groupby("customer_id").kl.sum()
        r = [name] + [NONE] * (len(heads) - 1)
        for j, h in enumerate(heads):
            v = None
            if "with servo" in h: v = len(a)
            elif "new counters added" in h: v = len(newc)
            elif h.startswith("total no"): v = len(a) + len(newc)
            elif "nil selling" in h: v = len(a - set(act.index))
            elif "nil premium" in h: v = int((~act.index.isin(prem_cy[prem_cy > 0].index)).sum())
            elif "premium counters" in h: v = int((prem_ly * 1000 > 50).sum())
            elif "(0-50l)" in h: v = int((lit <= 50).sum())
            elif "(50-150l)" in h: v = int(((lit > 50) & (lit <= 150)).sum())
            elif "(150-500l)" in h: v = int(((lit > 150) & (lit <= 500)).sum())
            elif "(>500l)" in h: v = int((lit > 500).sum())
            elif "(1 invoice)" in h: v = int((act.inv == 1).sum())
            elif "(2-3 invoice)" in h: v = int(act.inv.between(2, 3).sum())
            elif "(4-5 invoice)" in h: v = int((act.inv >= 4).sum())
            elif "(1 sku" in h: v = int((act.sku == 1).sum())
            elif "(2-3 sku" in h: v = int(act.sku.between(2, 3).sum())
            elif "(4-5 sku" in h: v = int(act.sku.between(4, 5).sum())
            elif "(6-10 sku" in h: v = int(act.sku.between(6, 10).sum())
            elif "(> 10sku" in h or "(>10sku" in h: v = int((act.sku > 10).sum())
            if v is not None: r[j] = str(v)
        rows.append(r)
    write_rows(tbl, 2, rows)


IG_ROWS = [("servoxpert- 2w", ["2/3W"]), ("servoxpert- pcmo", ["PASSENGER CAR"]), ("servoxpert- cv", ["CV", "COMMERCIAL VEHICLE"]),
           ("servoxpert- agri", ["AGRI", "TRACTOR"])]


def f_ig(slide, tbl, calc):
    heads = [norm(h) for h in header_texts(tbl, 2)]
    sx = calc.ctx.sec_all[calc.ctx.sec_all.cust_group.astype(str).str.upper() == "SERVOXPERT"]
    codes = [c for c, _ in calc.stockists]
    sx = sx[sx.stockist_code.isin(codes)]
    first = sx[sx.kl > 0].groupby("customer_id").date.min()
    sub = sx.drop_duplicates("customer_id", keep="last").set_index("customer_id").cust_subgroup.astype(str).str.upper()
    data_start = first.min() if len(first) else None
    w = calc.ws["cy"]

    def commissioned(ids, y=None):
        f = first.reindex(list(ids)).dropna()
        if y is None: return int(((f >= w.start) & (f <= w.end)).sum())
        ys, ye = pd.Timestamp(year=y - 1, month=4, day=1), pd.Timestamp(year=y, month=3, day=31)
        if data_start is None or ys < data_start: return None
        return int(((f >= ys) & (f <= ye)).sum())

    groups = {}
    used = set()
    for key, subs in IG_ROWS:
        ids = set(sub[sub.isin(subs)].index); used |= ids; groups[key] = ids
    groups["servoxpert- others"] = set(sub.index) - used
    groups["servoxpert- total"] = set(sub.index)
    for r in list(tbl.rows)[2:]:
        lab = norm(r.cells[0].text)
        key = next((k for k in groups if lab.startswith(k)), None)
        is_sales = lab.startswith("sales by iocl")
        if key is None and not is_sales: continue
        for j, h in enumerate(heads):
            y = year_of(h)
            if j == 0 or "plan" in h or "action" in h or "target" in h: continue
            if is_sales:
                if y is not None and y < calc.fy and y in calc.sec_years: set_text(r.cells[j], fmt_v(float(sx[sx.fy == y].kl.sum())))
                elif "apr" in h: set_text(r.cells[j], fmt_v(float(sx[(sx.date >= w.start) & (sx.date <= w.end)].kl.sum())))
                continue
            if y is not None and y < calc.fy:
                v = commissioned(groups[key], y)
                if v is not None: set_text(r.cells[j], str(v))
            elif "apr" in h:
                set_text(r.cells[j], str(commissioned(groups[key])))


def f_highlights(slide, tbl, calc):
    lines = H.highlights(calc.ctx)
    title = slide.shapes.title
    tid = title.shape_id if title is not None else None
    body = [sh for sh in slide.shapes if sh.has_text_frame and sh.shape_id != tid and "highlight" not in norm(sh.text_frame.text)]
    if body:
        set_text_frame(max(body, key=lambda sh: sh.width * sh.height).text_frame, lines)
    else:
        tb = X.textbox(slide, ["• " + l for l in lines], Inches(0.6), Inches(1.3), Inches(12.1), Inches(5.4), 16, "000000")
        for p in tb.text_frame.paragraphs: p.space_after = Pt(10)


def set_text_frame(tf, lines):
    p0 = tf.paragraphs[0]
    for p in list(tf.paragraphs)[1:]: p._p.getparent().remove(p._p)
    for r in list(p0.runs): r._r.getparent().remove(r._r)
    for i, l in enumerate(lines):
        p = p0 if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = ("• " if not l.startswith("•") else "") + l; r.font.size = Pt(16)
        p.space_after = Pt(10)


_MON = "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
_PERIOD_RX = re.compile(rf"\b({_MON})\s*[-–]\s*({_MON})(\s*['’]\s*\d{{2}})?\b(?!\s*['’]?\d{{2}}\s*-)")


def relabel_period(slide, ctx):
    """A filled deck may carry period labels typed for the previous review ('Actual 26-27 (Apr-Sep)', notes '... Apr-Sep'26 ...').
    On slides the dashboard filled, such month-range labels are rewritten to the period of the new numbers."""
    from .periods import mi_name
    a, b = mi_name(ctx.period.mi_start, False), mi_name(ctx.period.mi_end, False)
    new_y = (ctx.period.mi_end // 12) % 100
    texts = [c.text for sh in slide.shapes if getattr(sh, "has_table", False) and sh.has_table for r in sh.table.rows for c in r.cells] + \
            [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame]
    yrs = [int(re.sub(r"\D", "", m.group(3))) for t in texts for m in _PERIOD_RX.finditer(t) if m.group(3)]
    old_cy = max(yrs) if yrs else None                 # newest year on the slide = the old 'current' label; older ones are LY

    def fix(text):
        def one(m):
            if not m.group(3): return f"{a}-{b}"
            y = int(re.sub(r"\D", "", m.group(3)))
            return f"{a}-{b}’{(new_y - (old_cy - y)) % 100:02d}"
        return _PERIOD_RX.sub(one, text)

    def fix_tf(tf):
        for para in tf.paragraphs:
            for run in para.runs:
                if _PERIOD_RX.search(run.text): run.text = fix(run.text)
    for sh in slide.shapes:
        if getattr(sh, "has_table", False) and sh.has_table:
            for r in sh.table.rows:
                for c in r.cells:
                    if _PERIOD_RX.search(c.text): fix_tf(c.text_frame)
        elif sh.has_text_frame and _PERIOD_RX.search(sh.text_frame.text):
            fix_tf(sh.text_frame)


def f_cover(slide, tbl, calc):
    for sh in slide.shapes:
        if not sh.has_text_frame: continue
        for p in sh.text_frame.paragraphs:
            if norm(p.text).startswith("date") and (not p.text.split(":", 1)[-1].strip() or "_" in p.text) and p.runs:
                for r in list(p.runs)[1:]: r._r.getparent().remove(r._r)
                p.runs[0].text = f"Date: {dt.date.today():%d.%m.%Y}"


# =============================================================== slide catalogue (matched on the slide title)
SPEC = [
    ("cover", r"^lse interaction", "Cover (date filled)", f_cover),
    ("area", r"area map", "Area map (stockists, average sales)", f_area),
    ("industry", r"industry channel size", "Industry channel size (IOCL row, MT)", f_industry),
    ("highlights", r"major highlights", "Major highlights (auto text)", f_highlights),
    ("profile", r"stockist profile", "Stockist profile (names, average sales)", f_profile),
    ("summary", r"target vs achievement", "Target vs Achievement (15 parameters)", f_summary),
    ("segment", r"primary sales performance-product segment", "Primary – product segment", f_segment),
    ("stockist_ta", r"stockist wise target vs ach", "Primary – stockist target vs achievement by year", f_stockist_ta),
    ("stockist", r"primary sales performance-stockist wise", "Primary – stockist wise", f_stockist_primary),
    ("premium", r"primary sales-premium grade sales", "Primary – premium grade by stockist", f_premium),
    ("category", r"primary sales-segment wise", "Primary – category (Mass/Aspirational/Premium/…) by stockist", f_category),
    ("hyper", r"hyper series", "Hyper series placement & sales", f_hyper),
    ("genuine", r"servo & genuine oil", "SERVO & Genuine oil by brand", f_genuine),
    ("pgroup", r"product segment sales performance", "Product group (channel)", f_pgroup),
    ("pack", r"pack wise sales", "Pack wise (channel)", f_pack),
    ("premium_plan", r"gold/ ?premium grade sales plan", "Gold / premium grade plan", f_premium_plan),
    ("bucket", r"bucket grade", "Bucket grade (7–26 L packs)", f_bucket),
    ("sec_channel", r"secondary sales performance-channel", "Secondary – channel", f_sec_channel),
    ("sec_ro", r"secondary sales performance-ro channel", "Secondary – RO by stockist", _sec_by_stockist("RO")),
    ("sec_bazaar", r"secondary sales performance-bazaar channel", "Secondary – Bazaar by stockist", _sec_by_stockist("Bazaar")),
    ("ssr_channel", r"ssr wise channel wise", "SSR wise channel wise (your added slide)", None),
    ("bz_ssr", r"bazaar counter details & sales-ssr wise", "Bazaar counters & sales – SSR wise", f_bz_ssr),
    ("bz_dist", r"bazaar counter details & sales-district wise", "Bazaar counters & sales – district wise", f_bz_dist),
    ("bz_detail", r"^bazaar counter details$", "Bazaar counter details (serviced / new / inactive)", f_bz_detail),
    ("bz_pcs", r"bazaar counter sales per month", "Bazaar sales per counter per month", f_bz_pcs),
    ("bz_track", r"bazaar performance tracking", "Bazaar performance tracking (volume / invoice / SKU buckets)", f_bz_track),
    ("sec_oem", r"secondary sales performance-oem channel", "Secondary – OEM by brand", f_oem),
    ("oem_pen", r"oem segment wise penetration", "OEM penetration plan (sales / targets)", f_oem_pen),
    ("sec_fleet", r"secondary sales performance-fleet channel", "Secondary – Fleet by stockist", _sec_by_stockist("Fleet")),
    ("ig", r"independent garage", "Independent garage / ServoXpert", f_ig),
    ("marketing", r"marketing & promotional", "Marketing & promotion (manual)", None),
    ("inspection", r"inspection performance", "Inspection (manual)", None),
    ("development", r"developmental issues", "Developmental issues (manual)", None),
    ("visits", r"market visit", "Market & field visits (manual)", None),
    ("initiatives", r"new initiative", "New initiatives (manual)", None),
    ("concerns", r"area of concern", "Area of concern (manual)", None),
    ("thanks", r"thanks", "Thanks", None),
]
LABELS = {k: lab for k, _, lab, _ in SPEC}


def identify(slide) -> str | None:
    t = norm(slide_title(slide))
    for key, rx, _, _ in SPEC:
        if re.search(rx, t): return key
    if "thanks" in norm(" ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)): return "thanks"
    return None


def inventory(base_path) -> list[tuple[int, str | None, str]]:
    """[(slide no, key, title)] of the base deck."""
    prs = Presentation(str(base_path))
    return [(i, identify(s), slide_title(s)[:70]) for i, s in enumerate(prs.slides, 1)]


# =============================================================== extra slides
def _move_last(prs, new_index):
    """Move the slide added last to position new_index (0-based)."""
    ids = prs.slides._sldIdLst
    me = list(ids)[-1]
    ids.remove(me); ids.insert(max(0, min(new_index, len(ids))), me)


def add_ssr_channel(prs, calc, index):
    """'SSR Wise Channel Wise' (secondary by channel and stockist: annual / period target, base, sales, growth, % ach)."""
    rows = []
    for ch in ("Bazaar", "OEM", "Fleet", "RO"):
        rows.append({"Sl No": ch, "Stockist Name": "", "Annual Target": "", f"Target ({calc.ws['cy'].label})": "",
                     f"Base ({calc.ws['ly'].label})": "", f"Sales ({calc.ws['cy'].label})": "", "% Growth": "", "% Target Ach": ""})
        tot = [0.0, 0.0, 0.0, 0.0]
        for i, (code, name) in enumerate(calc.stockists, 1):
            d = calc.sec[(calc.sec.channel == ch) & (calc.sec.stockist_code == code)]
            a, p = calc.tgt([code], None, "Secondary", ch), calc.tgt([code], calc.fms, "Secondary", ch)
            b, s = calc.per(d, "ly", True), calc.per(d, "cy", True)
            for k, v in enumerate((a, p, b, s)): tot[k] += v or 0
            rows.append({"Sl No": str(i), "Stockist Name": name, "Annual Target": fmt_v(a), f"Target ({calc.ws['cy'].label})": fmt_v(p),
                         f"Base ({calc.ws['ly'].label})": fmt_v(b), f"Sales ({calc.ws['cy'].label})": fmt_v(s),
                         "% Growth": fmt_p((s / b - 1) * 100) if b and b >= 0.5 else "-", "% Target Ach": fmt_p(s / p * 100) if p else "-"})
        rows.append({"Sl No": "", "Stockist Name": "TOTAL", "Annual Target": fmt_v(tot[0]), f"Target ({calc.ws['cy'].label})": fmt_v(tot[1]),
                     f"Base ({calc.ws['ly'].label})": fmt_v(tot[2]), f"Sales ({calc.ws['cy'].label})": fmt_v(tot[3]),
                     "% Growth": fmt_p((tot[3] / tot[2] - 1) * 100) if tot[2] >= 0.5 else "-", "% Target Ach": fmt_p(tot[3] / tot[1] * 100) if tot[1] else "-"})
    s = X.new_slide(prs, "SSR Wise Channel Wise")
    X.add_table(s, pd.DataFrame(rows), "Secondary Sales - Channel Wise", top=Inches(0.95), unit="KL", max_h=Inches(6.3))
    _move_last(prs, index)


EXTRAS = {
    "trend": ("Monthly trend charts (Primary & Secondary vs LY)", X.s_trend),
    "channel_chart": ("Secondary channel chart + table", X.s_channel),
    "pri_sec": ("Primary vs Secondary by stockist", X.s_pri_sec),
    "pivot": ("Custom pivot (saved from the Pivot tab)", X.s_pivot),
}


# =============================================================== build
def build(ctx: H.Ctx, base_path, skip: set[str] | None = None, extras: list[str] | None = None,
          add_ssr_channel_slide=True) -> tuple[bytes, list[tuple[int, str, str]]]:
    """Returns (pptx bytes, report[(slide no in the base deck, label, status)])."""
    skip = skip or set()
    prs = Presentation(str(base_path))
    calc = Calc(ctx)
    report, drop = [], []
    own_ssr = None                                     # the user's own 'SSR Wise Channel Wise' slide, if the base has one
    for i, slide in enumerate(list(prs.slides)):
        key = identify(slide)
        label = LABELS.get(key, slide_title(slide)[:60] or f"Slide {i + 1}")
        if key in skip:
            drop.append(slide); report.append((i + 1, label, "left out")); continue
        if key == "ssr_channel":
            own_ssr = slide; continue
        fn = next((f for k, _, _, f in SPEC if k == key), None)
        if fn is None:
            report.append((i + 1, label, "kept as is - fill by hand" if key not in ("cover", "thanks") else "kept as is")); continue
        try:
            fn(slide, main_table(slide), calc)
            relabel_period(slide, ctx)
            report.append((i + 1, label, "filled"))
        except Exception as e:  # noqa: BLE001 - one slide never stops the deck
            report.append((i + 1, label, f"NOT filled ({type(e).__name__}: {str(e)[:80]}) - left as in the base deck"))
    def pos(key):
        return next((i for i, sl in enumerate(prs.slides) if identify(sl) == key), None)
    if "ssr_channel" not in skip and (add_ssr_channel_slide or own_ssr is not None):
        try:
            if own_ssr is not None:
                at = next(i for i, sl in enumerate(prs.slides) if sl.part is own_ssr.part)
                drop.append(own_ssr)                          # replaced by the rebuilt one (removed at the end)
            else:
                b = pos("sec_bazaar")
                at = (b + 1) if b is not None else len(prs.slides._sldIdLst)
            add_ssr_channel(prs, calc, at)
            report.append((at + 1, LABELS["ssr_channel"], "filled" if own_ssr is not None else "added"))
        except Exception as e:  # noqa: BLE001
            report.append((0, LABELS["ssr_channel"], f"NOT added ({type(e).__name__}: {str(e)[:80]})"))
    for k in extras or []:
        if k not in EXTRAS or (k == "pivot" and not ctx.pivot): continue
        n0 = len(prs.slides._sldIdLst)
        try:
            EXTRAS[k][1](prs, ctx)
            t = pos("thanks")
            if t is not None and len(prs.slides._sldIdLst) > n0: _move_last(prs, t)
            report.append((0, EXTRAS[k][0], "added"))
        except Exception as e:  # noqa: BLE001 - remove a half-made slide, carry on
            for sl in list(prs.slides)[n0:]: _delete_slide(prs, sl)
            report.append((0, EXTRAS[k][0], f"not added - no data for it in this period ({type(e).__name__})"))
    for sl in drop: _delete_slide(prs, sl)
    _renumber(prs)
    report.sort(key=lambda r: (r[0] == 0, r[0]))
    buf = io.BytesIO(); prs.save(buf)
    return buf.getvalue(), report


def _renumber(prs):
    """Give the remaining slides unique, sequential part names (slide1.xml ...) so the saved file never has two parts
    with the same name after slides were removed / added."""
    from pptx.opc.packuri import PackURI
    parts = [s.part for s in prs.slides]
    for i, part in enumerate(parts, 1): part.partname = PackURI(f"/ppt/slides/tmp_slide{i}.xml")
    for i, part in enumerate(parts, 1): part.partname = PackURI(f"/ppt/slides/slide{i}.xml")


def _delete_slide(prs, slide):
    ids = prs.slides._sldIdLst
    for sid in list(ids):
        if prs.part.related_part(sid.rId) is slide.part:
            prs.part.drop_rel(sid.rId); ids.remove(sid); return


def find_bases(folder_data: Path, folder_review: Path) -> list[Path]:
    """HO template first, then filled LSE decks (newest first) found in Data and the Sales Review folder."""
    tpl = sorted(folder_data.glob(C.HO_TEMPLATE_GLOB)) + sorted(folder_review.glob(C.HO_TEMPLATE_GLOB))
    filled = [f for d in (folder_data, folder_review) for f in d.glob("*.pptx")
              if not f.name.startswith(("~$", ".")) and re.search(r"lse interaction", f.name, re.I)]
    seen, out = set(), []
    for f in tpl[:1] + sorted(filled, key=lambda f: f.stat().st_mtime, reverse=True):
        if f.name not in seen: seen.add(f.name); out.append(f)
    return out


def deck_tables_xlsx(deck: bytes) -> bytes:
    """Every table of a built deck as an Excel sheet (same numbers as the slides)."""
    prs = Presentation(io.BytesIO(deck))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        used = set()
        for i, slide in enumerate(prs.slides, 1):
            for sh in slide.shapes:
                if not (getattr(sh, "has_table", False) and sh.has_table): continue
                rows = [[c.text.strip() for c in r.cells] for r in sh.table.rows]
                name = re.sub(r"[\[\]:*?/\\]", "", f"{i:02d} {slide_title(slide)}")[:31]
                while name in used: name = name[:29] + f"_{len(used)}"
                used.add(name)
                pd.DataFrame(rows).to_excel(xw, sheet_name=name, index=False, header=False)
        if not used: pd.DataFrame({"": ["No tables in this deck"]}).to_excel(xw, sheet_name="Info", index=False)
    return buf.getvalue()
