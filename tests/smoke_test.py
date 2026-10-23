"""Crash test: renders every tab of the dashboard headless under many period / filter combinations.
Run:  ~/.local/venvs/sales_review_ho/bin/python tests/smoke_test.py          (uses your real data folder, read-only)
Exit code 0 = no tab raised an error in any scenario.
"""
import os, sys, time
from pathlib import Path
from streamlit.testing.v1 import AppTest

APP = str(Path(__file__).resolve().parent.parent / "app.py")
sys.path.insert(0, str(Path(APP).parent))


def scenarios():
    yield "default", {}
    for kind in ["Month", "Quarter", "Half-year", "YTD (Apr to month)", "Full year", "Custom range"]:
        yield f"view={kind}", {"kind": kind}
    yield "oldest FY", {"fy": "oldest"}
    yield "future month (no data yet)", {"kind": "Month", "month": 12}
    yield "one stockist", {"sl_stockist": "first"}
    yield "channel=Fleet slice", {"sl_channel": ["Fleet"]}
    yield "rare district slice", {"sl_district": "rarest"}
    yield "rare product group", {"sl_product_group": "rarest"}
    yield "stockist + OEM", {"sl_stockist": "last", "sl_channel": ["OEM"]}
    yield "pack + category", {"sl_pack_group": "rarest", "sl_category": "rarest"}
    yield "premium tab: rare segment", {"pr_seg": "rarest"}
    yield "premium tab: HYPER series", {"pr_series": ["HYPER"]}
    yield "premium tab: one customer + stockist", {"pr_cust": "first", "pr_stk": "last"}


def apply(at, sc):
    sb = {w.key: w for w in at.sidebar.selectbox if w.key}
    ms = {w.key: w for w in at.multiselect if w.key}
    for k, v in sc.items():
        if k == "fy":
            sb["fy"].set_value(sb["fy"].options[-1] if False else sorted(sb["fy"].options)[0])
        elif k in sb:
            sb[k].set_value(v)
        elif k in ms:
            opts = ms[k].options
            if v == "first": v = [opts[0]]
            elif v == "last": v = [opts[-1]]
            elif v == "rarest": v = [opts[-1]]
            ms[k].set_value([o for o in v if o in opts] or opts[:1])
    return at


def main():
    bad = 0
    for name, sc in scenarios():
        t = time.time()
        at = AppTest.from_file(APP, default_timeout=600).run()
        if sc:
            at = apply(at, {k: v for k, v in sc.items() if k != "month"}).run()
            if "month" in sc:
                [w for w in at.sidebar.selectbox if w.key == "month"][0].set_value(sc["month"]); at.run()
        exc = [e.message for e in at.exception]
        errs = [e.value for e in at.error if "could not" in str(e.value).lower() or "problem" in str(e.value).lower()]
        caught = [m.value.split("</b>")[0].split("<b>")[-1] + ": " + m.value.split("(", 1)[-1][:120] for m in at.markdown if "class='oops'" in str(m.value)]
        errs += caught                      # sections that hit an error but were contained by ui.guard
        ok = not exc and not errs
        bad += not ok
        print(f"{'PASS' if ok else 'FAIL'}  {name:32s} {time.time() - t:5.1f}s", "" if ok else (exc + errs)[:3])
    print("\nALL SCENARIOS PASSED" if not bad else f"\n{bad} SCENARIO(S) FAILED")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
