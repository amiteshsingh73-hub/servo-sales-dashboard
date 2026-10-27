"""Widget fuzzer: renders the dashboard and then changes every filter / selector on every tab (first, last and a middle
option, multiselects with one / two picks, toggles flipped, text boxes filled), re-rendering all tabs after each change.
Any uncaught exception or any section that falls back to its error card is reported.

Runs on a temporary COPY of the Data folder (nothing in your real folders is written).
Run:  ~/.local/venvs/sales_review_ho/bin/python tests/fuzz_widgets.py            (~20-30 min)
      ~/.local/venvs/sales_review_ho/bin/python tests/fuzz_widgets.py --quick    (one option per widget)
"""
import os, shutil, sys, tempfile, time
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent.parent
SRC = Path(os.environ.get("SALES_DATA_DIR", APP_DIR.parent / "Data")).resolve()
QUICK = "--quick" in sys.argv

# never write into the real folders: copy Data and the scheme rule book to a temp place
TMP = Path(tempfile.mkdtemp()); (TMP / "Data").mkdir()
for f in SRC.iterdir():
    if f.is_file() and not f.name.startswith(("~$", ".")): shutil.copy2(f, TMP / "Data" / f.name)
os.environ["SALES_DATA_DIR"] = str(TMP / "Data")
os.environ["SALES_REVIEW_NO_AUTOSAVE"] = "1"
sch = Path(os.environ.get("SCHEME_DIR", APP_DIR.parent.parent / "Lube Scheme Circulars"))
(TMP / "Schemes").mkdir()
if (sch / "Scheme Rules.xlsx").exists(): shutil.copy2(sch / "Scheme Rules.xlsx", TMP / "Schemes" / "Scheme Rules.xlsx")
os.environ["SCHEME_DIR"] = str(TMP / "Schemes")
sys.path.insert(0, str(APP_DIR))

# ---- record the real option values of every choice widget as the app creates it
from streamlit.elements.widgets import selectbox as _sb, radio as _ra, multiselect as _ms, select_slider as _ss  # noqa: E402

REG: dict = {}


def _rec(kind, mixin, name):
    orig = getattr(mixin, name)

    def wrapper(self, label, options=(), *a, **kw):
        try:
            REG[(kind, kw.get("key") or str(label))] = list(options)
        except Exception:
            pass
        return orig(self, label, options, *a, **kw)
    setattr(mixin, name, wrapper)


_rec("selectbox", _sb.SelectboxMixin, "selectbox")
_rec("radio", _ra.RadioMixin, "radio")
_rec("multiselect", _ms.MultiSelectMixin, "multiselect")
_rec("select_slider", _ss.SelectSliderMixin, "select_slider")

from streamlit.testing.v1 import AppTest  # noqa: E402

SKIP_KEYS = {"autosave"}                                 # writes a file into exports/


def problems(at):
    out = [f"EXCEPTION {e.message[:250]}" for e in at.exception]
    out += ["SECTION ERROR " + str(m.value).split("<b>")[1].split("</b>")[0] + " - " + str(m.value).split("(", 1)[-1][:160]
            for m in at.markdown if "class='oops'" in str(m.value)]
    return out


def ident(w):
    return (w.type, w.key or str(w.label))


def all_widgets(at):
    ws = []
    for kind in ("selectbox", "radio", "multiselect", "select_slider", "toggle", "checkbox", "text_input", "number_input"):
        for w in getattr(at, kind):
            if (w.key or "") in SKIP_KEYS or str(getattr(w, "disabled", False)) == "True": continue
            ws.append(w)
    seen, uniq = set(), []
    for w in ws:
        if ident(w) not in seen: seen.add(ident(w)); uniq.append(w)
    return uniq


def find(at, idn):
    kind, k = idn
    for w in getattr(at, kind):
        if (w.key or str(w.label)) == k: return w
    return None


def choices(w):
    kind = w.type
    if kind in ("toggle", "checkbox"): return [not w.value]
    if kind == "text_input": return ["hyper", "zz-no-match"]
    if kind == "number_input": return [0, 10]
    opts = REG.get((kind, w.key or str(w.label)), [])
    if not opts: return []
    pick = [opts[0], opts[-1], opts[len(opts) // 2]]
    if kind == "multiselect": return [[opts[-1]], [opts[0], opts[len(opts) // 2]] if len(opts) > 1 else [opts[0]], []][: (1 if QUICK else 3)]
    if kind == "select_slider":
        rng = isinstance(w.value, (list, tuple))
        return ([(opts[0], opts[-1]), (opts[-1], opts[-1])] if rng else [opts[0], opts[-1]])[: (1 if QUICK else 2)]
    seen, out = set(), []
    for p in pick:
        if repr(p) not in seen: seen.add(repr(p)); out.append(p)
    return out[: (1 if QUICK else 3)]


def main():
    t0 = time.time()
    at = AppTest.from_file(str(APP_DIR / "app.py"), default_timeout=900).run()
    fails = [("initial render", p) for p in problems(at)]
    todo = [ident(w) for w in all_widgets(at)]
    print(f"{len(todo)} widgets found on first render ({time.time() - t0:.0f}s)")
    runs, done = 0, set()
    while todo:
        idn = todo.pop(0)
        if idn in done: continue
        done.add(idn)
        w = find(at, idn)
        if w is None: continue
        orig = w.value
        for v in choices(w):
            w = find(at, idn)
            if w is None: break
            try:
                w.set_value(v); at.run(); runs += 1
            except Exception as e:  # noqa: BLE001  harness could not drive this widget
                fails.append((f"{idn} = {v!r}", f"HARNESS {type(e).__name__}: {str(e)[:160]}")); continue
            for p in problems(at): fails.append((f"{idn} = {v!r}", p))
            todo += [ident(x) for x in all_widgets(at) if ident(x) not in done and ident(x) not in todo]   # widgets revealed by this choice
        w = find(at, idn)
        if w is not None:
            try: w.set_value(orig); at.run(); runs += 1
            except Exception: pass
        print(f"  {runs:4d} runs  {time.time() - t0:6.0f}s  {idn[0]:13s} {str(idn[1])[:50]}", flush=True)
    real = [f for f in fails if not f[1].startswith("HARNESS")]
    print(f"\n{runs} renders, {len(done)} widgets, {time.time() - t0:.0f}s")
    for where, p in fails: print(("FAIL " if not p.startswith("HARNESS") else "skip ") + where + "\n     " + p)
    print("\nNO CRASHES FOUND" if not real else f"\n{len(real)} PROBLEM(S) FOUND")
    shutil.rmtree(TMP, ignore_errors=True)
    return 1 if real else 0


if __name__ == "__main__":
    sys.exit(main())
