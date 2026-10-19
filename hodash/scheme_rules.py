"""Rule book for the incentive calculator.

The rates/thresholds live in an Excel file ("Scheme Rules.xlsx", next to your circulars) so you can change them
when a new circular arrives - the calculator reads whatever is in the file for the selected period.
A default rule book is generated from the circulars (HO/LUBES/RLS/2026-27/01, 01C, 19B, 26 etc.) the first time.
Quarters without a circular yet are copied from the latest one and marked 'provisional'.
"""
from __future__ import annotations
import re
import numpy as np
import pandas as pd
from . import config as C

# period id -> (label, start, end, circulars, status)
PERIODS = {
    "FY27-Q1": ("Q1 FY26-27 (Apr-Jun'26)", "2026-04-01", "2026-06-30", "HO/LUBES/RLS/2026-27/01 (2-Apr-26), 09 (Jun'26 specials)", "verified"),
    "FY27-Q2": ("Q2 FY26-27 (Jul-Sep'26)", "2026-07-01", "2026-09-30", "01, 19B (Excellence), 26 (Blue Ocean)", "verified"),
    "FY27-Q3": ("Q3 FY26-27 (Oct-Dec'26)", "2026-10-01", "2026-12-31", "HO/LUBES/RLS/2026-27/01C (1-Oct-26)", "verified"),
    "FY27-Q4": ("Q4 FY26-27 (Jan-Mar'27)", "2027-01-01", "2027-03-31", "none yet - Q3 rules carried forward", "provisional"),
    "FY28-Q1": ("Q1 FY27-28 (Apr-Jun'27)", "2027-04-01", "2027-06-30", "none yet - Q3 rules carried forward", "provisional"),
    "FY28-Q2": ("Q2 FY27-28 (Jul-Sep'27)", "2027-07-01", "2027-09-30", "none yet - Q3 rules carried forward", "provisional"),
}

# key: (scheme, description, unit, H1-circular-01 value, Q3-circular-01C value)   None = not applicable / no condition
P = {
    # --- Early Bird: Rs/L on eligible primary volume by invoice day of month (paid in the SAP invoice, ZEBD)
    "eb_on": ("Early Bird", "Scheme active (1/0)", "flag", 1, 1),
    "eb1_from": ("Early Bird", "Slot 1 from day", "day", 1, 1), "eb1_to": ("Early Bird", "Slot 1 to day", "day", 15, 15), "eb1_rate": ("Early Bird", "Slot 1 rate", "Rs/L", 1.5, 1.5),
    "eb2_from": ("Early Bird", "Slot 2 from day", "day", 16, 16), "eb2_to": ("Early Bird", "Slot 2 to day", "day", 20, 20), "eb2_rate": ("Early Bird", "Slot 2 rate", "Rs/L", 1.0, 1.0),
    "eb3_from": ("Early Bird", "Slot 3 from day", "day", 21, 21), "eb3_to": ("Early Bird", "Slot 3 to day", "day", 25, 25), "eb3_rate": ("Early Bird", "Slot 3 rate", "Rs/L", 0.5, 0.5),
    "eb4_from": ("Early Bird", "Slot 4 from day", "day", 26, 26), "eb4_to": ("Early Bird", "Slot 4 to day", "day", 31, 31), "eb4_rate": ("Early Bird", "Slot 4 rate", "Rs/L", 0.0, 0.0),
    # --- MPI: monthly primary target achieved
    "mpi_on": ("MPI", "Scheme active (1/0)", "flag", 1, 1),
    "mpi_rate": ("MPI", "Rate on eligible primary volume of the month", "Rs/L", 1.0, 1.0),
    "mpi_hs_cap": ("MPI", "Hydrashakti 68 must be below this share of month primary (blank = no condition)", "share", None, 0.30),
    # --- QPI: quarterly primary target (+ secondary target)
    "qpi_on": ("QPI", "Scheme active (1/0)", "flag", 1, 1),
    "qpi_pri_rate": ("QPI", "Rate if quarterly PRIMARY target achieved (confirmed: Rs 1 + Rs 1 = Rs 2 when both met)", "Rs/L", 1.0, 1.0),
    "qpi_sec_rate": ("QPI", "Additional rate if quarterly SECONDARY target achieved (primary target also needed) - confirmed Rs 1 + Rs 1", "Rs/L", 1.0, 1.0),
    "qpi_hs_cap": ("QPI", "Hydrashakti 68 must be below this share of quarter primary (blank = none)", "share", None, 0.30),
    # --- Stockist categorisation (premium-grade upliftment of previous period)
    "cat_on": ("Categorisation", "Scheme active (1/0)", "flag", 1, 1),
    "cat_months": ("Categorisation", "Look-back period used (months)", "months", 6, 3),
    "cat_prem_share": ("Categorisation", "Premium stockist: premium-grade share of primary above", "share", 0.10, 0.10),
    "cat_prem_kl": ("Categorisation", "...or this many KL, whichever is LOWER", "KL", 120, 60),
    "cat_asp_share": ("Categorisation", "Aspirational stockist: premium-grade share above", "share", 0.05, 0.05),
    "cat_asp_kl": ("Categorisation", "...or this many KL, whichever is LOWER", "KL", 60, 30),
    "cat_prem_rate": ("Categorisation", "Premium stockist rate on Aspirational+Premium grade primary", "Rs/L", 2.0, 2.0),
    "cat_asp_rate": ("Categorisation", "Aspirational stockist rate on Aspirational+Premium grade primary", "Rs/L", 1.0, 1.0),
    # --- Special scheme on MASS grades (extra trade discount linked to premium-grade share; needs primary target)
    "mass_on": ("Mass trade discount", "Scheme active (1/0)", "flag", 0, 1),
    "mass_t1_share": ("Mass trade discount", "Tier 1: premium share above", "share", None, 0.12),
    "mass_t1_kl": ("Mass trade discount", "Tier 1: ...or KL, whichever LOWER", "KL", None, 40),
    "mass_t1_pct": ("Mass trade discount", "Tier 1 trade discount on MASS grades", "% of DLP", None, 0.05),
    "mass_t2_share": ("Mass trade discount", "Tier 2: premium share above", "share", None, 0.10),
    "mass_t2_kl": ("Mass trade discount", "Tier 2: ...or KL, whichever LOWER", "KL", None, 35),
    "mass_t2_pct": ("Mass trade discount", "Tier 2 trade discount", "% of DLP", None, 0.045),
    "mass_t3_share": ("Mass trade discount", "Tier 3: premium share above", "share", None, 0.08),
    "mass_t3_kl": ("Mass trade discount", "Tier 3: ...or KL, whichever LOWER", "KL", None, 25),
    "mass_t3_pct": ("Mass trade discount", "Tier 3 trade discount", "% of DLP", None, 0.035),
    "mass_base_pct": ("Mass trade discount", "Base trade discount already given on MASS (differential is paid)", "% of DLP", 0.025, 0.025),
    # --- Metro stockist incentive (value based)
    "metro_on": ("Metro", "Scheme active (1/0)", "flag", 1, 1),
    "metro_pct": ("Metro", "Incentive on Premium, Aspirational & Silver grades", "% of DLP", 0.005, 0.005),
    # --- SSF incentive (paid by stockist to SSF, reimbursed through SAP rebate quarterly)
    "ssf_on": ("SSF", "Scheme active (1/0)", "flag", 1, 1),
    "ssf_add_rate": ("SSF", "Rate on SSF bazaar volume if counter-addition target met", "Rs/L", 1.0, 1.0),
    "ssf_cover_rate": ("SSF", "Rate on SSF bazaar volume if >= coverage % of base counters billed", "Rs/L", 1.0, 1.0),
    "ssf_min_add": ("SSF", "Counter-addition target per SSF per quarter", "counters", 12, 12),
    "ssf_new_min_l": ("SSF", "A new counter counts only with this many litres in a month", "L", 50, 50),
    "ssf_cover_pct": ("SSF", "Coverage required of bazaar counters available at 31-Mar-26", "share", 0.80, 0.80),
    # --- Reseller Performance Incentive (Q3 onward; SDMS claim)
    "rpi_on": ("RPI", "Scheme active (1/0)", "flag", 0, 1),
    "rpi_cap_l": ("RPI", "Maximum litres per counter per quarter", "L", None, 1000),
    "rpi_new_since": ("RPI", "New counter = not billed in SDMS since this date (YYYY-MM-DD)", "date", None, "2026-01-01"),
    "rpi_n1_vol": ("RPI", "New counter slab 1: litres above", "L", None, 300), "rpi_n1_sku": ("RPI", "slab 1: premium SKUs", "SKUs", None, 4),
    "rpi_n1_prem": ("RPI", "slab 1: premium litres", "L", None, 80), "rpi_n1_rate": ("RPI", "slab 1 rate", "Rs/L", None, 3.5),
    "rpi_n2_vol": ("RPI", "New counter slab 2: litres above", "L", None, 200), "rpi_n2_sku": ("RPI", "slab 2: premium SKUs", "SKUs", None, 3),
    "rpi_n2_prem": ("RPI", "slab 2: premium litres", "L", None, 40), "rpi_n2_rate": ("RPI", "slab 2 rate", "Rs/L", None, 3.25),
    "rpi_n3_vol": ("RPI", "New counter slab 3: litres above", "L", None, 100), "rpi_n3_sku": ("RPI", "slab 3: premium SKUs", "SKUs", None, 2),
    "rpi_n3_prem": ("RPI", "slab 3: premium litres", "L", None, 20), "rpi_n3_rate": ("RPI", "slab 3 rate", "Rs/L", None, 3.0),
    "rpi_e_growth": ("RPI", "Existing counter: minimum growth vs LY same quarter", "share", None, 0.20),
    "rpi_e_sku": ("RPI", "Existing counter: premium SKUs", "SKUs", None, 3), "rpi_e_prem": ("RPI", "Existing counter: premium litres", "L", None, 40),
    "rpi_e_rate": ("RPI", "Existing counter rate", "Rs/L", None, 3.0),
    # --- Servo Festival Bonanza (IOC RO, Q3)
    "fb_on": ("Festival Bonanza", "Scheme active (1/0)", "flag", 0, 1),
    "fb_hi_growth": ("Festival Bonanza", "Higher slab: Apr-Dec growth above", "share", None, 0.15),
    "fb_hi_prem_l": ("Festival Bonanza", "Higher slab: min premium-grade litres in quarter", "L", None, 40),
    "fb_hi_rate": ("Festival Bonanza", "Higher slab rate", "Rs/L", None, 3.0),
    "fb_lo_growth": ("Festival Bonanza", "Lower slab: Apr-Dec growth above", "share", None, 0.10),
    "fb_lo_honda_l": ("Festival Bonanza", "Lower slab: min Honda 4T litres in quarter", "L", None, 40),
    "fb_lo_rate": ("Festival Bonanza", "Lower slab rate", "Rs/L", None, 2.0),
    # --- H1 only (circular 01): PPI / RGI / ServoXpert slabs
    "ppi_on": ("PPI", "Premium Performance Incentive active (1/0)", "flag", 1, 0),
    "ppi_min_l": ("PPI", "Bazaar counter premium-grade litres in quarter above", "L", 50, None),
    "ppi_rate": ("PPI", "Rate on premium-grade secondary volume", "Rs/L", 5.0, None),
    "rgi_on": ("RGI", "Reseller Growth Incentive active (1/0)", "flag", 1, 0),
    "rgi_growth": ("RGI", "Growth over LY same period above", "share", 0.10, None),
    "rgi_rate": ("RGI", "Rate (all grades except Hydrashakti 68)", "Rs/L", 1.0, None),
    "rgi_nil_base_l": ("RGI", "Base litres per quarter for a counter with nil base", "L", 120, None),
    "sx_on": ("ServoXpert", "ServoXpert slab scheme active (1/0)", "flag", 1, 0),
    "sx_s1_min": ("ServoXpert", "Slab 1 from litres", "L", 150, None), "sx_s2_min": ("ServoXpert", "Slab 2 above litres", "L", 300, None),
    "sx_s1_rate": ("ServoXpert", "Slab 1 rate", "Rs/L", 1.0, None), "sx_s2_rate": ("ServoXpert", "Slab 2 rate", "Rs/L", 2.0, None),
    "exc_on": ("Servo Excellence", "Primary target achievement scheme (Apr-Sep) active (1/0)", "flag", 1, 0),
    "exc_rate": ("Servo Excellence", "Rate on primary excl. Hydrashakti 68, hydraulic & OEM grades if Apr-Sep target met", "Rs/L", 3.0, None),
}
# Q2 FY27 differs from Q1 only by the Blue Ocean scheme (circular 26): same premium-share tiers on MASS grades
Q2_OVERRIDES = {"mass_on": 1, "mass_t1_share": 0.10, "mass_t1_kl": None, "mass_t1_pct": 0.05, "mass_t2_share": 0.08, "mass_t2_kl": None, "mass_t2_pct": 0.045,
                "mass_t3_share": 0.05, "mass_t3_kl": None, "mass_t3_pct": 0.035}
SOURCES = {"H1": "Circular 01 (2-Apr-26) / 19B / 26", "Q3": "Circular 01C (1-Oct-26)"}
COLS = ["Period", "Scheme", "Key", "Description", "Value", "Unit", "Source", "Status"]


def default_rules() -> pd.DataFrame:
    rows = []
    for pid, (_, _, _, circ, status) in PERIODS.items():
        fam = "H1" if pid in ("FY27-Q1", "FY27-Q2") else "Q3"
        for key, (scheme, desc, unit, h1, q3) in P.items():
            val = h1 if fam == "H1" else q3
            if pid == "FY27-Q2" and key in Q2_OVERRIDES: val = Q2_OVERRIDES[key]
            src = SOURCES[fam] if status == "verified" else "Carried forward from Q3 FY27 (01C) - update when circular is issued"
            rows.append({"Period": pid, "Scheme": scheme, "Key": key, "Description": desc, "Value": val, "Unit": unit, "Source": src, "Status": status})
    return pd.DataFrame(rows, columns=COLS)


NOTES = [
    "HOW THIS FILE WORKS",
    "One row per parameter per period. Edit the Value column (blank = no condition / not applicable) and save; the dashboard re-reads it automatically.",
    "When a new circular arrives: filter Period, update the Values, set Status to 'verified' and put the circular number in Source.",
    "Shares are decimals (0.10 = 10%). Flags: 1 = scheme active, 0 = inactive for that period.",
    "Periods Q4 FY27 and FY28 Q1/Q2 are PROVISIONAL copies of the Oct-Dec'26 circular (01C).",
    "NOT MODELLED (price/negotiated, claimed outside this calculator): Hydrashakti special net-price offers, OEM / Fleet / ServoXpert-garage customer-specific discounts, "
    "reseller product-discount annexures (SKU-wise, see the SDMS discount estimate), Ex-MI incentive slabs, Servo Exclusive Bonus (exclusive-counter flag not in sales data).",
]


def ensure_rulebook():
    """Creates the Excel rule book if it does not exist yet."""
    if C.RULES_FILE.exists(): return False
    C.RULES_FILE.parent.mkdir(parents=True, exist_ok=True)
    save_rules(default_rules())
    return True


def save_rules(df: pd.DataFrame, path=None):
    path = path or C.RULES_FILE
    per = pd.DataFrame([{"Period": k, "Label": v[0], "Start": v[1], "End": v[2], "Circulars": v[3], "Status": v[4]} for k, v in PERIODS.items()])
    with pd.ExcelWriter(path) as w:
        df[COLS].to_excel(w, sheet_name="Rules", index=False)
        per.to_excel(w, sheet_name="Periods", index=False)
        pd.DataFrame({"Read me": NOTES}).to_excel(w, sheet_name="Read me", index=False)


def load_rules() -> pd.DataFrame:
    ensure_rulebook()
    df = pd.read_excel(C.RULES_FILE, sheet_name="Rules")
    # forward-compat: add any parameter introduced later that the file does not have yet
    base = default_rules()
    have = set(zip(df.Period, df.Key))
    miss = base[[(p, k) not in have for p, k in zip(base.Period, base.Key)]]
    return pd.concat([df, miss], ignore_index=True) if len(miss) else df


class RuleSet:
    def __init__(self, df: pd.DataFrame, period: str):
        d = df[df.Period == period]
        self.period, self.rows = period, d
        self._v = {r.Key: r.Value for r in d.itertuples()}
        self.status = d.Status.iloc[0] if len(d) else "missing"

    def v(self, key, default=None):
        x = self._v.get(key, default)
        if x is None or (isinstance(x, float) and np.isnan(x)): return default
        return x

    def f(self, key, default=0.0):
        x = self.v(key, None)
        try:
            return float(x) if x is not None else default
        except (TypeError, ValueError):
            return default

    def on(self, flag_key) -> bool:
        return self.f(flag_key, 0) >= 1
