"""SERVO IGNITE scheme, Oct-Dec'26 (circular HO/LUBES/RLS/2026-27/30, 02-Nov-2026) - calculations only, no screen code.

Incentive = 1.5% x weightage obtained x Q3 net primary invoice value (excl. GST and Hydrashakti).
Five parameters of 20% each; Hydrashakti volumes count for no parameter and earn nothing.
"""
from __future__ import annotations
import io, re
from dataclasses import dataclass, field
import numpy as np
import pandas as pd
from . import loader

CIRCULAR = "HO/LUBES/RLS/2026-27/30 dated 02-Nov-2026"
START, END = pd.Timestamp("2026-10-01"), pd.Timestamp("2026-12-31")
QUARTER_DAYS = (END - START).days + 1


@dataclass
class Rules:
    """Everything the circular fixes; the tab lets you change a value if HO clarifies it."""
    rate_pct: float = 1.5            # % of primary value at 100% weightage
    weight_pct: float = 20.0         # each parameter
    min_skus: int = 150              # 1. unique SKUs in primary
    sku_min_l: float = 20.0          # 1. litres a SKU needs in the quarter to count
    new_ssr: int = 50                # 2. new Bazaar counters, SSR
    new_gss: int = 15                # 2. new Bazaar counters, GSS
    new_min_l: float = 50.0          # 2. litres a new counter needs in the quarter
    hist_from: pd.Timestamp = pd.Timestamp("2025-04-01")     # 2. 'new' = not billed in this window
    hist_to: pd.Timestamp = pd.Timestamp("2026-09-30")
    min_ratio: float = 1.0           # 3. Bazaar / RO
    min_pcs: float = 90.0            # 4. litres per counter per month
    sku_per_counter: int = 4         # 5. unique SKUs per counter
    counter_sku_min_l: float = 20.0  # 5. litres a SKU needs at the counter to count
    sku_rule: str = "every"          # 5. 'every' = every counter must reach it, 'average' = average across counters


# ------------------------------------------------------------------ SKU = product code + pack size
def _pack(desc, raw) -> str | None:
    for txt in (desc, raw):
        s = re.sub(r"(\d+)/(\d+)\s*L", lambda m: f"{int(m.group(1)) / int(m.group(2)):g}L", str(txt).upper())   # 20 X 1/2 L
        u = loader._unit_litres(s)
        if u: return f"{u[0]:g}{'KG' if u[1] == 'KG' else 'L'}"
    return None


def sku_keys(df: pd.DataFrame) -> pd.Series:
    """'Product codes having same pack size are the same SKU': 20X0.9 L, 20x0.9 L promo and a 0.9 L promo can are one SKU."""
    if df.empty: return pd.Series(dtype=object, index=df.index)
    pairs = df[["pack_desc", "product_raw"]].astype(str).drop_duplicates()
    size = {(a, b): _pack(a, b) for a, b in pairs.itertuples(index=False)}
    packs = [size[(a, b)] for a, b in zip(df.pack_desc.astype(str), df.product_raw.astype(str))]
    return pd.Series([f"{c}|{p if p else 'pk' + str(k)}" for c, p, k in zip(df.pcode, packs, df.pk)], index=df.index)


def no_hydra(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    hy = df.product_group.astype(str).eq("Hydrashakti 68") | df["product"].astype(str).str.upper().str.contains("HYDRASHAKTI", regex=False) \
        | df.product_raw.astype(str).str.upper().str.contains("HYDRASHAKTI", regex=False)
    return df[~hy]


def _q(df: pd.DataFrame) -> pd.DataFrame:
    return df[(df.date >= START) & (df.date <= END)]


def _ctr(df: pd.DataFrame) -> pd.DataFrame:
    """Bazaar counters: ServoXpert workshops are not counted as counters (same convention as the HO deck)."""
    d = df[df.channel == "Bazaar"]
    return d[d.cust_group.astype(str).str.upper() != "SERVOXPERT"]


# ------------------------------------------------------------------ primary value
def value_lines(pri_q: pd.DataFrame, sec: pd.DataFrame, price_pct: float = 100.0) -> pd.DataFrame:
    """Value of each Q3 primary line excl. Hydrashakti. Uses an invoice-value column if the SAP export has one, otherwise
    litres x the secondary Rs/L of the same SKU (then product, then product group, then overall) x price_pct."""
    d = no_hydra(pri_q).copy()
    d["sku"] = sku_keys(d)
    d["litres"] = d.kl * 1000
    s = no_hydra(sec[(sec.date >= START - pd.DateOffset(months=12)) & (sec.date <= END) & (sec.kl > 0) & (sec.amount > 0)]).copy()
    overall = s.amount.sum() / (s.kl.sum() * 1000) if len(s) else np.nan
    s["sku"] = sku_keys(s)
    rate = lambda keys: (s.groupby(keys).amount.sum() / (s.groupby(keys).kl.sum() * 1000))
    est = d.sku.map(rate("sku")) if len(s) else pd.Series(np.nan, index=d.index)
    if len(s):
        est = est.fillna(d.pcode.map(rate("pcode"))).fillna(d.product_group.map(rate("product_group")))
    est = est.fillna(overall) * price_pct / 100
    d["est_value"] = d.litres * est
    sap = d["value"] if "value" in d else pd.Series(np.nan, index=d.index)
    d["line_value"] = sap.where(sap.notna(), d.est_value)
    d["from_sap"] = sap.notna()
    return d


# ------------------------------------------------------------------ evaluation
@dataclass
class Result:
    code: int
    name: str
    kind: str
    params: list = field(default_factory=list)
    met: int = 0
    weight: float = 0.0
    value_to_date: float = 0.0
    value_source: str = ""
    value_full: float = 0.0
    incentive_now: float = 0.0
    incentive_full: float = 0.0
    incentive_max: float = 0.0
    excl_hydra_kl: float = 0.0
    hydra_kl: float = 0.0
    sku_list: pd.DataFrame = None
    sku_gap: pd.DataFrame = None
    new_list: pd.DataFrame = None
    counter_list: pd.DataFrame = None


def evaluate(pri: pd.DataFrame, sec: pd.DataFrame, names: dict, rules: Rules, gss: set = frozenset(), actual_value: dict | None = None,
             price_pct: float = 100.0, last_p=None, last_s=None) -> list[Result]:
    actual_value = actual_value or {}
    last_p = last_p if last_p is not None else pri.date.max()
    last_s = last_s if last_s is not None else sec.date.max()
    days_p = int(np.clip((min(last_p, END) - START).days + 1, 0, QUARTER_DAYS))
    days_s = int(np.clip((min(last_s, END) - START).days + 1, 0, QUARTER_DAYS))
    months_s = days_s / QUARTER_DAYS * 3

    pri_q, sec_q = _q(pri), _q(sec)
    vals = value_lines(pri_q, sec, price_pct)
    sx = no_hydra(sec_q).copy()
    sx["sku"] = sku_keys(sx)
    ctr_q = _ctr(sx)

    # 5. SKUs per counter, across any SSR/GSS
    per_sku = ctr_q.assign(l=ctr_q.kl * 1000).groupby(["customer_id", "sku"]).l.sum()
    ok_sku = per_sku[per_sku >= rules.counter_sku_min_l]
    n_sku = ok_sku.groupby(level=0).size()
    all_sku = per_sku[per_sku > 0].groupby(level=0).size()

    # 2. billed in the look-back window, under any stockist
    billed = set(sec[(sec.date >= rules.hist_from) & (sec.date <= rules.hist_to) & (sec.kl > 0)].customer_id)

    # 12-month SKU history for the 'bill these next' list
    hist = no_hydra(pri[(pri.date >= START - pd.DateOffset(months=12)) & (pri.date < START)]).copy()
    hist["sku"] = sku_keys(hist)

    out = []
    for code in sorted(names, key=lambda c: str(names[c])):
        kind = "GSS" if code in gss else "SSR"
        r = Result(int(code), names[code], kind)

        # ---- 1. unique SKUs (primary)
        v = vals[vals.stockist_code == code]
        g = v.groupby("sku").agg(Product=("product", "first"), Pack=("pack_desc", "first"), Litres=("litres", "sum"))
        qual = g[g.Litres >= rules.sku_min_l]
        r.sku_list = g.assign(Counts=np.where(g.Litres >= rules.sku_min_l, "Yes", "No")).sort_values("Litres", ascending=False).reset_index(drop=True)
        h = hist[hist.stockist_code == code]
        hg = h.groupby("sku").agg(Product=("product", "first"), Pack=("pack_desc", "first"), l12=("kl", lambda x: x.sum() * 1000))
        hg["This quarter (L)"] = g.Litres.reindex(hg.index).fillna(0.0)
        hg = hg[(hg.l12 > 0) & (hg["This quarter (L)"] < rules.sku_min_l)]
        hg["Litres to add"] = rules.sku_min_l - hg["This quarter (L)"]
        r.sku_gap = hg.rename(columns={"l12": "Last 12 months (L)"}).sort_values("Last 12 months (L)", ascending=False).reset_index(drop=True)
        n1 = len(qual)
        r.params.append(_param("Unique SKUs (primary)", f"≥ {rules.min_skus} SKUs, each ≥ {rules.sku_min_l:g} L", n1, rules.min_skus, f"{n1} SKUs",
                               f"{max(rules.min_skus - n1, 0)} more SKUs" if n1 < rules.min_skus else ""))

        # ---- 2. new Bazaar counters
        c = ctr_q[ctr_q.stockist_code == code]
        lit = c.groupby("customer_id").kl.sum() * 1000
        fresh = lit[~lit.index.isin(billed) & (lit > 0)]
        need2 = rules.new_gss if kind == "GSS" else rules.new_ssr
        n2 = int((fresh >= rules.new_min_l).sum())
        info = c.sort_values("date").groupby("customer_id").agg(Counter=("customer", "last"), SSF=("ssf_name", "last"), District=("district", "last"))
        r.new_list = info.reindex(fresh.index).assign(**{"Litres": fresh, "Counts": np.where(fresh >= rules.new_min_l, "Yes", "No"),
                                                         "Litres to add": (rules.new_min_l - fresh).clip(lower=0)}).sort_values("Litres", ascending=False).reset_index()
        r.params.append(_param("New Bazaar counters", f"≥ {need2} new ({kind}), each ≥ {rules.new_min_l:g} L", n2, need2,
                               f"{n2} qualified ({len(fresh) - n2} more started)", f"{max(need2 - n2, 0)} more counters" if n2 < need2 else ""))

        # ---- 3. Bazaar / RO (Bazaar sales include ServoXpert, as in the HO deck)
        s = sx[sx.stockist_code == code]
        bz_l, ro_l = s[s.channel == "Bazaar"].kl.sum() * 1000, s[s.channel == "RO"].kl.sum() * 1000
        ratio = bz_l / ro_l if ro_l > 0 else (np.inf if bz_l > 0 else 0.0)
        r.params.append(_param("Bazaar ÷ RO sales", f"≥ {rules.min_ratio:g}", ratio, rules.min_ratio,
                               f"{'no RO sales' if ro_l <= 0 else f'{ratio:.2f}'}  ({bz_l / 1000:,.1f} KL ÷ {ro_l / 1000:,.1f} KL)",
                               f"{(rules.min_ratio * ro_l - bz_l) / 1000:,.1f} KL more Bazaar" if ratio < rules.min_ratio else ""))

        # ---- 4. PCS: Bazaar litres / (counters x 3); months elapsed while the quarter runs
        n_c = c[c.kl > 0].customer_id.nunique()
        pcs = bz_l / (n_c * months_s) if n_c and months_s else 0.0
        gap4 = rules.min_pcs * n_c * months_s - bz_l
        r.params.append(_param("Per-counter sales (PCS)", f"≥ {rules.min_pcs:g} L / counter / month", pcs, rules.min_pcs,
                               f"{pcs:,.0f} L  ({n_c:,} counters)", f"{gap4 / 1000:,.1f} KL more at these counters" if pcs < rules.min_pcs else ""))

        # ---- 5. SKUs per counter
        ids = lit[lit > 0].index
        counted = n_sku.reindex(ids).fillna(0).astype(int)
        share = (counted >= rules.sku_per_counter).mean() * 100 if len(ids) else 0.0
        avg = counted.mean() if len(ids) else 0.0
        if rules.sku_rule == "every":
            r.params.append(_param("SKUs per Bazaar counter", f"every counter ≥ {rules.sku_per_counter} SKUs, each ≥ {rules.counter_sku_min_l:g} L", share, 100.0,
                                   f"{int((counted >= rules.sku_per_counter).sum()):,} of {len(ids):,} counters ({share:.0f}%), average {avg:.1f}",
                                   f"{int((counted < rules.sku_per_counter).sum()):,} counters short" if share < 100 else ""))
        else:
            r.params.append(_param("SKUs per Bazaar counter", f"average ≥ {rules.sku_per_counter} SKUs, each ≥ {rules.counter_sku_min_l:g} L", avg, rules.sku_per_counter,
                                   f"average {avg:.1f}  ({share:.0f}% of counters at {rules.sku_per_counter}+)",
                                   f"{rules.sku_per_counter - avg:.1f} more SKUs per counter" if avg < rules.sku_per_counter else ""))
        r.counter_list = info.reindex(ids).assign(**{"Litres": lit.reindex(ids), "SKUs counted": counted,
                                                     "SKUs bought (any litres)": all_sku.reindex(ids).fillna(0).astype(int),
                                                     "SKUs to add": (rules.sku_per_counter - counted).clip(lower=0)}).sort_values("Litres", ascending=False).reset_index()

        # ---- value and incentive
        r.met = sum(p["met"] for p in r.params)
        r.weight = r.met * rules.weight_pct / 100
        typed = actual_value.get(code) or 0.0
        est_val = float(v.line_value.sum())
        r.value_to_date = typed if typed > 0 else est_val
        r.value_source = "typed (SAP)" if typed > 0 else ("SAP file" if len(v) and v.from_sap.all() else "estimate")
        r.value_full = r.value_to_date * QUARTER_DAYS / days_p if days_p else 0.0
        rate = rules.rate_pct / 100
        r.incentive_now = r.value_to_date * rate * r.weight
        r.incentive_full = r.value_full * rate * r.weight
        r.incentive_max = r.value_full * rate
        r.excl_hydra_kl = float(v.kl.sum())
        r.hydra_kl = float(pri_q[pri_q.stockist_code == code].kl.sum()) - r.excl_hydra_kl
        out.append(r)
    return out


def _param(name, criteria, actual, need, shown, gap) -> dict:
    return dict(Parameter=name, Criteria=criteria, actual=actual, need=need, Actual=shown, met=bool(actual >= need), Gap=gap)


def excel(results: list[Result], note: list[str]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame({"": note}).to_excel(xw, sheet_name="Read me", index=False, header=False)
        pd.DataFrame([{"Stockist": r.name, "Type": r.kind, "Parameters met": r.met, "Weightage": r.weight, "Value to date (Rs)": round(r.value_to_date),
                       "Value basis": r.value_source, "Value full quarter at run-rate (Rs)": round(r.value_full), "Incentive on value to date (Rs)": round(r.incentive_now),
                       "Incentive full quarter (Rs)": round(r.incentive_full), "Incentive if all 5 met (Rs)": round(r.incentive_max)} for r in results]
                     ).to_excel(xw, sheet_name="Summary", index=False)
        pd.DataFrame([{"Stockist": r.name, "Parameter": p["Parameter"], "Criteria": p["Criteria"], "Actual": p["Actual"], "Met": "Yes" if p["met"] else "No", "Gap": p["Gap"]}
                      for r in results for p in r.params]).to_excel(xw, sheet_name="Parameters", index=False)
        for r in results:
            tag = re.sub(r"[^A-Za-z]", "", r.name)[:12]
            for sh, df in (("SKUs", r.sku_list), ("SKUs to bill", r.sku_gap), ("New ctrs", r.new_list), ("Counter SKUs", r.counter_list)):
                if df is not None: df.round(1).to_excel(xw, sheet_name=f"{tag} {sh}"[:31], index=False)
    return buf.getvalue()
