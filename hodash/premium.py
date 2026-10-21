"""Secondary sales by customer and segment, with Premium-category sales CY vs LY and the premium grades billed.
Pure pandas (no Streamlit) so it can be tested and exported."""
from __future__ import annotations
import io
import numpy as np
import pandas as pd
from .metrics import growth, pct
from .periods import Window, month_start, month_end, mi_name

PREMIUM = "Premium"


def windows_for(mi_from: int, mi_to: int, last_date: pd.Timestamp) -> dict:
    """CY window for the chosen months and the same months a year earlier (cut at the same day if data stops mid-month)."""
    start, end = month_start(mi_from), month_end(mi_to)
    partial = pd.notna(last_date) and last_date < end
    if partial: end = max(last_date, start)
    lbl = mi_name(mi_from) if mi_from == mi_to else f"{mi_name(mi_from, False)}-{mi_name(mi_to)}"
    lbl_ly = mi_name(mi_from - 12) if mi_from == mi_to else f"{mi_name(mi_from - 12, False)}-{mi_name(mi_to - 12)}"
    ly_end = end - pd.DateOffset(years=1) if partial else month_end(mi_to - 12)
    return {"cy": Window(start, end, lbl, partial), "ly": Window(month_start(mi_from - 12), ly_end, lbl_ly, partial)}


def _sel(df, w): return df[(df.date >= w.start) & (df.date <= w.end)]


def _is_prem(df): return df.category.eq(PREMIUM)


def kpis(sec: pd.DataFrame, w: dict) -> dict:
    out = {}
    for k in ("cy", "ly"):
        d = _sel(sec, w[k]); p = d[_is_prem(d)]; pp = p[p.kl > 0]
        out[k] = dict(total=float(d.kl.sum()), prem=float(p.kl.sum()), share=pct(float(p.kl.sum()), float(d.kl.sum())),
                      grades=pp["product"].nunique(), skus=pp.sku_name.nunique(), prem_cust=pp.customer_id.nunique(),
                      cust=d[d.kl > 0].customer_id.nunique())
    return out


def _pair(d_cy, d_ly, keys):
    """Total and premium KL, CY and LY, by keys."""
    def agg(d):
        g = d.groupby(keys).agg(total=("kl", "sum"))
        p = d[_is_prem(d)]
        g["prem"] = p.groupby(keys).kl.sum()
        pp = p[p.kl > 0]
        g["grades"] = pp.groupby(keys)["product"].nunique()
        g["prem_cust"] = pp.groupby(keys).customer_id.nunique()
        return g
    a, b = agg(d_cy), agg(d_ly)
    t = a.join(b, how="outer", lsuffix="_cy", rsuffix="_ly").fillna(0.0)
    return t


def by_segment(sec, w) -> pd.DataFrame:
    t = _pair(_sel(sec, w["cy"]), _sel(sec, w["ly"]), "segment")
    if t.empty: return pd.DataFrame()
    t = t.sort_values("total_cy", ascending=False)
    out = pd.DataFrame({
        "Segment": t.index,
        f"Total CY (KL)": t.total_cy.values, "Total LY (KL)": t.total_ly.values,
        "Total % Growth": [growth(a, b) for a, b in zip(t.total_cy, t.total_ly)],
        "Premium CY (KL)": t.prem_cy.values, "Premium LY (KL)": t.prem_ly.values,
        "Premium % Growth": [growth(a, b) for a, b in zip(t.prem_cy, t.prem_ly)],
        "Premium share CY %": [pct(a, b) for a, b in zip(t.prem_cy, t.total_cy)],
        "Premium share LY %": [pct(a, b) for a, b in zip(t.prem_ly, t.total_ly)],
        "Premium grades billed CY": t.grades_cy.astype(int).values, "Premium grades billed LY": t.grades_ly.astype(int).values,
        "Customers buying premium CY": t.prem_cust_cy.astype(int).values})
    return _with_total(out, sec, w, "Segment")


def _with_total(out, sec, w, first):
    k = kpis(sec, w)
    tot = {first: "Total", "Total CY (KL)": k["cy"]["total"], "Total LY (KL)": k["ly"]["total"],
           "Total % Growth": growth(k["cy"]["total"], k["ly"]["total"]),
           "Premium CY (KL)": k["cy"]["prem"], "Premium LY (KL)": k["ly"]["prem"], "Premium % Growth": growth(k["cy"]["prem"], k["ly"]["prem"]),
           "Premium share CY %": k["cy"]["share"], "Premium share LY %": k["ly"]["share"],
           "Premium grades billed CY": k["cy"]["grades"], "Premium grades billed LY": k["ly"]["grades"],
           "Customers buying premium CY": k["cy"]["prem_cust"]}
    return pd.concat([out, pd.DataFrame([{c: tot.get(c, np.nan) for c in out.columns}])], ignore_index=True)


def by_customer(sec, w) -> pd.DataFrame:
    cy, ly = _sel(sec, w["cy"]), _sel(sec, w["ly"])
    t = _pair(cy, ly, "customer_id")
    if t.empty: return pd.DataFrame()
    t = t[(t.total_cy.abs() > 1e-9) | (t.total_ly.abs() > 1e-9)]
    info = pd.concat([ly, cy]).sort_values("date").groupby("customer_id").agg(
        Customer=("customer", "last"), Stockist=("stockist", "last"), Channel=("channel", "last"), District=("district", "last"))
    seg = cy.pivot_table(index="customer_id", columns="segment", values="kl", aggfunc="sum").fillna(0.0)
    seg.columns = [f"{c} CY (KL)" for c in seg.columns]
    seg = seg[seg.sum().sort_values(ascending=False).index]
    names = cy[_is_prem(cy) & (cy.kl > 0)].groupby("customer_id")["product"].agg(lambda x: ", ".join(sorted(set(x))))
    out = info.reindex(t.index).assign(**{
        "Total CY (KL)": t.total_cy, "Total LY (KL)": t.total_ly,
        "Total % Growth": [growth(a, b) for a, b in zip(t.total_cy, t.total_ly)],
        "Premium CY (KL)": t.prem_cy, "Premium LY (KL)": t.prem_ly,
        "Premium % Growth": [growth(a, b) for a, b in zip(t.prem_cy, t.prem_ly)],
        "Premium share CY %": [pct(a, b) for a, b in zip(t.prem_cy, t.total_cy)],
        "Premium grades billed CY": t.grades_cy.astype(int), "Premium grades billed LY": t.grades_ly.astype(int),
        "Premium grades (CY)": names.reindex(t.index).fillna("")})
    out = out.join(seg).fillna({c: 0.0 for c in seg.columns})
    out = out.sort_values("Total CY (KL)", ascending=False)
    out.index.name = "Customer Id"
    return out


def premium_grades(sec, w) -> pd.DataFrame:
    cy, ly = _sel(sec, w["cy"]), _sel(sec, w["ly"])
    cy, ly = cy[_is_prem(cy)], ly[_is_prem(ly)]
    a = cy.groupby(["product", "segment"]).agg(cy=("kl", "sum"), customers=("customer_id", "nunique"))
    b = ly.groupby(["product", "segment"]).agg(ly=("kl", "sum"))
    t = a.join(b, how="outer").fillna(0.0)
    if t.empty: return pd.DataFrame()
    t["% Growth"] = [growth(x, y) for x, y in zip(t.cy, t.ly)]
    t = t.sort_values("cy", ascending=False).reset_index()
    return t.rename(columns={"product": "Premium grade", "segment": "Segment", "cy": "CY (KL)", "ly": "LY (KL)", "customers": "Customers CY"})[
        ["Premium grade", "Segment", "CY (KL)", "LY (KL)", "% Growth", "Customers CY"]]


def excel(sheets: dict[str, pd.DataFrame], header: list[str]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame({"": header}).to_excel(xw, sheet_name="Filters", index=False, header=False)
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name[:31], index=df.index.name is not None)
            ws = xw.sheets[name[:31]]
            ws.freeze_panes = "B2"
            for col in ws.columns:
                hdr = str(col[0].value or "")
                ws.column_dimensions[col[0].column_letter].width = min(max(len(hdr), 10) + 2, 45)
                for c in col[1:]:
                    if isinstance(c.value, float): c.number_format = "0.0" if "%" in hdr else "#,##0.000"
    return buf.getvalue()
