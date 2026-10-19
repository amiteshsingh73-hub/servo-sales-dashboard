"""Builds the PPT-ready tables + auto-written highlights, shared by the dashboard, Excel and PPTX export."""
from __future__ import annotations
from dataclasses import dataclass
import numpy as np
import pandas as pd
from . import config as C, metrics as M
from .periods import Period, windows, fy_label


@dataclass
class Ctx:
    """Everything a table needs: filtered data (already narrowed by sidebar), period, targets."""
    pri: pd.DataFrame
    sec: pd.DataFrame
    sec_all: pd.DataFrame          # unfiltered secondary history (needed for new/lost customers)
    targets: pd.DataFrame
    period: Period
    pri_last: pd.Timestamp
    sec_last: pd.Timestamp
    apportion: bool = True
    pivot: tuple | None = None      # (title, text DataFrame) saved from the Pivot tab

    @property
    def wp(self): return windows(self.period, self.pri_last)
    @property
    def ws(self): return windows(self.period, self.sec_last)
    @property
    def codes(self): return sorted(self.pri.stockist_code.unique()) if len(self.pri) else None


def table(ctx: Ctx, which: str, dim: str, **kw) -> pd.DataFrame:
    df, last = (ctx.pri, ctx.pri_last) if which == "pri" else (ctx.sec, ctx.sec_last)
    kw.setdefault("tkind", "Primary" if which == "pri" else "Secondary")
    return M.ho_table(df, dim, ctx.period, last, ctx.targets, apportion=ctx.apportion, **kw)


def segment_table(ctx):      return table(ctx, "pri", "hseg", order=["Servo - Automotive", "Servo - Industrial", "Co-branded (Genuine)"], rename_dim="Segment")
def stockist_table(ctx):     return table(ctx, "pri", "stockist", dim_is_stockist=True, rename_dim="SSR/GSS (Stockist)")
def product_group_table(ctx): return table(ctx, "pri", "product_group", order=C.PRODUCT_GROUP_ORDER, rename_dim="Product Group")
def pack_table(ctx):         return table(ctx, "pri", "pack_group", order=C.PACK_GROUP_ORDER, rename_dim="SKU / Pack Group")
def category_table(ctx):     return table(ctx, "pri", "category", order=["Premium", "Aspirational", "Silver", "Mass", "Special", "Industrial"], rename_dim="Product Category")
def sec_channel_table(ctx):  return table(ctx, "sec", "channel", order=C.CHANNEL_ORDER, rename_dim="Channel")
def sec_stockist_table(ctx): return table(ctx, "sec", "stockist", dim_is_stockist=True, rename_dim="SSR/GSS (Stockist)")


def premium_stockist_table(ctx):
    c = Ctx(**{**ctx.__dict__, "pri": ctx.pri[ctx.pri.category.eq("Premium")]})
    return table(c, "pri", "stockist", dim_is_stockist=False, rename_dim="SSR/GSS (Premium grade)")


def sec_channel_stockist_table(ctx, channel: str):
    c = Ctx(**{**ctx.__dict__, "sec": ctx.sec[ctx.sec.channel.eq(channel)]})
    return table(c, "sec", "stockist", dim_is_stockist=True, tchannel=channel, rename_dim=f"SSR/GSS ({channel})")


# ------------------------------------------------------------------ summary (HO slide 6 "Target vs Achievement")
def summary_table(ctx: Ctx) -> pd.DataFrame:
    p, fy = ctx.period, ctx.period.fy
    wp, ws = ctx.wp, ctx.ws
    years = [y for y in (fy - 3, fy - 2, fy - 1) if (ctx.pri.fy == y).any()]
    ycols = [fy_label(y, True) for y in years]
    cp, lyl = f"Actual {wp['cy'].label}", f"LY {wp['cy'].label.split('-')[0] if False else wp['ly'].label}"
    rows = []

    def add(name, unit, df, w, mult=1.0, target=False, ch=None):
        r = {"Parameter": name, "Unit": unit}
        for y, c in zip(years, ycols): r[c] = float(df[df.fy == y].kl.sum()) * mult
        a, b = M.kl(df, w["cy"]) * mult, M.kl(df, w["ly"]) * mult
        kind = "Secondary" if ch else "Primary"
        ta = M.target_for(ctx.targets, ctx.codes, fy, range(1, 13), kind, ch) if target else np.nan
        tp = M.target_for(ctx.targets, ctx.codes, fy, p.fms, kind, ch) if target else np.nan
        r.update({f"Target {fy_label(fy, True)}": ta, "Pro-rata target": tp, cp: a, lyl: b,
                  "% Achievement": M.pct(a, tp), "% Growth": M.growth(a, b)})
        rows.append(r)

    add("PRIMARY SALES - CHANNEL", "KL", ctx.pri, wp, target=True)
    for ch in C.CHANNEL_ORDER:
        add(f"{ch.upper()} SALES (secondary)", "KL", ctx.sec[ctx.sec.channel.eq(ch)], ws, target=True, ch=ch)
    add("GOLD / PREMIUM GRADE SALES", "KL", ctx.pri[ctx.pri.category.eq("Premium")], wp)
    add("GREASE SALES", "MT", ctx.pri[ctx.pri.product_group.eq("Grease")], wp, mult=1.0)
    for g, lbl in (("4T Oil", "4T OIL"), ("DEO", "DEO"), ("PCMO", "PCMO"), ("Gear Oil", "GEAR OIL"), ("Agri", "AGRI")):
        add(f"PRODUCT SEGMENT - {lbl}", "KL", ctx.pri[ctx.pri.product_group.eq(g)], wp)
    t = pd.DataFrame(rows)
    # per-counter productivity (secondary, bazaar)
    bz = ctx.sec[ctx.sec.channel.eq("Bazaar")]
    r = {"Parameter": "PER BAZAAR COUNTER SALES PER MONTH", "Unit": "Ltrs/month"}
    for y, c in zip(years, ycols):
        d = bz[(bz.fy == y) & (bz.kl > 0)]
        r[c] = d.kl.sum() * 1000 / d.customer_id.nunique() / 12 if len(d) else np.nan
    k = M.customer_kpis(ctx.sec, ctx.sec_all, ws, p)
    r[cp], r[lyl] = k["pcs"], k["pcs_ly"]; r["% Growth"] = M.growth(k["pcs"], k["pcs_ly"])
    t = pd.concat([t, pd.DataFrame([r])], ignore_index=True)
    if len(t):
        t.insert(0, "S.No.", range(1, len(t) + 1))
    return t


# ------------------------------------------------------------------ highlights text
def highlights(ctx: Ctx) -> list[str]:
    p, wp, ws = ctx.period, ctx.wp, ctx.ws
    out = []
    pc, pl = M.kl(ctx.pri, wp["cy"]), M.kl(ctx.pri, wp["ly"])
    tp = M.target_for(ctx.targets, ctx.codes, p.fy, p.fms)
    s = f"Primary channel sales of {pc:,.0f} KL in {wp['cy'].label}"
    if not np.isnan(tp): s += f" – {M.pct(pc, tp):.0f}% of pro-rata target ({tp:,.0f} KL)"
    if pl: s += f"{' and' if not np.isnan(tp) else ','} {M.growth(pc, pl):+.0f}% growth over {wp['ly'].label}"
    out.append(s + ".")
    sc, sl = M.kl(ctx.sec, ws["cy"]), M.kl(ctx.sec, ws["ly"])
    if sc:
        ch = M.compare(ctx.sec, "channel", ws, order=C.CHANNEL_ORDER)
        parts = [f"{i} {r.growth:+.0f}%" for i, r in ch.iterrows() if r.ly > 0 and r.cy > 0 and i in C.CHANNEL_ORDER[:3]]
        ts = M.target_for(ctx.targets, ctx.codes, p.fy, p.fms, "Secondary")
        tsx = f" – {M.pct(sc, ts):.0f}% of pro-rata target ({ts:,.0f} KL)" if not np.isnan(ts) else ""
        out.append(f"Secondary sales of {sc:,.0f} KL{tsx}" + (f", {M.growth(sc, sl):+.0f}% growth" if sl else "") + (f": {', '.join(parts)} over same period last year." if parts else "."))
        if pc: out.append(f"Primary to secondary ratio {pc / sc:.2f}x ({pc - sc:+,.0f} KL stock build at stockists vs secondary offtake).")
    pg = M.compare(ctx.pri, "category", wp)
    if "Premium" in pg.index and pg.loc["Premium", "ly"] > 0:
        r = pg.loc["Premium"]; out.append(f"Premium grade sales {r.growth:+.0f}% ({r.ly:,.0f} KL to {r.cy:,.0f} KL), {r.share:.1f}% of primary.")
    seg = M.compare(ctx.pri, "product_group", wp)
    seg = seg[((seg.ly > 5) | (seg.cy > 5)) & seg.growth.notna()]
    if len(seg) > 1:
        up, dn = seg.sort_values("growth", ascending=False).iloc[0], seg.sort_values("growth").iloc[0]
        out.append(f"Best segment: {seg.growth.idxmax()} ({up.growth:+.0f}%, {up.ly:,.0f}→{up.cy:,.0f} KL); weakest: {seg.growth.idxmin()} ({dn.growth:+.0f}%, {dn.ly:,.0f}→{dn.cy:,.0f} KL).")
    st = M.compare(ctx.pri, "stockist", wp)
    if not np.isnan(tp) and len(st):
        code = ctx.pri.drop_duplicates("stockist").set_index("stockist").stockist_code
        tps = M.target_by_stockist(ctx.targets, p.fy, p.fms)
        ach = {n: M.pct(r.cy, tps.get(code[n], np.nan)) for n, r in st.iterrows() if n in code}
        ok = [n for n, a in ach.items() if not np.isnan(a) and a >= 100]
        out.append(f"{len(ok)} of {len(ach)} stockists at or above pro-rata target" + (f" ({', '.join(ok)})." if ok else "."))
    k = M.customer_kpis(ctx.sec, ctx.sec_all, ws, p)
    if k["active"]:
        s = f"{k['active']:,} active counters" + (f" ({M.growth(k['active'], k['active_ly']):+.0f}% vs LY)" if k["active_ly"] else "")
        if k["new"] is not None: s += f"; {k['new']:,} new counters adding {k['new_kl']:,.0f} KL"
        s += f"; {k['lost']:,} counters from LY did not transact ({k['lost_kl']:,.0f} KL)."
        if k["pcs"] == k["pcs"]: s += f"  Bazaar per-counter sales {k['pcs']:.0f} L/month" + (f" vs {k['pcs_ly']:.0f} L LY." if k["pcs_ly"] == k["pcs_ly"] else ".")
        out.append(s)
    if wp["cy"].partial or ws["cy"].partial:
        out.append(f"Note: data available till {min(ctx.pri_last, ctx.sec_last):%d-%b-%Y}; comparison with LY is like-for-like to the same date.")
    return out
