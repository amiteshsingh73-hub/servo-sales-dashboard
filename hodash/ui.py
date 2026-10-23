"""Look & feel: colours, CSS, KPI cards, section headers, chart template and the crash guard used by every tab."""
from __future__ import annotations
import contextlib, html, traceback
import numpy as np
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

NAVY, ORANGE, GREY, GREEN, RED, TEAL = "#0B1F66", "#E8710A", "#B7BECC", "#1A7F4B", "#C0392B", "#2A9D8F"
NAVY_LT = "#7C8DC4"
CAT = ["#0B1F66", "#E8710A", "#2A9D8F", "#8E6BBF", "#C0392B", "#D4A017", "#4E79A7", "#7F8C8D"]
FONT = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif"

# ------------------------------------------------------------------ chart template (applies to every plotly chart)
pio.templates["servo"] = go.layout.Template(layout=dict(
    font=dict(family=FONT, size=12, color="#2B3350"),
    title=dict(font=dict(size=15, color=NAVY), x=0, xanchor="left", y=.97),
    colorway=CAT, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(showgrid=False, linecolor="#D5DAE5", ticks="", automargin=True),
    yaxis=dict(gridcolor="rgba(120,130,150,.15)", zeroline=False, automargin=True),
    hoverlabel=dict(bgcolor="#FFFFFF", bordercolor="#D5DAE5", font=dict(family=FONT, color="#1B2440")),
    legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0, bgcolor="rgba(0,0,0,0)"),
    modebar=dict(orientation="v", bgcolor="rgba(255,255,255,.85)"),
    bargap=.25, margin=dict(l=8, r=36, t=78, b=8)))
pio.templates.default = "plotly+servo"

CSS = f"""<style>
:root {{ --navy:{NAVY}; --orange:{ORANGE}; --ink:#1B2440; --muted:#5B6475; --line:#E3E7EF; --card:#FFFFFF; }}
html, body, [class*="css"] {{ font-family:{FONT}; }}
.block-container {{ padding-top:2.4rem; padding-bottom:3rem; max-width:1560px; }}
[data-testid="stSidebar"] {{ background:linear-gradient(180deg,#0B1F66 0%,#13277A 100%); }}
[data-testid="stSidebar"] * {{ color:#E9EDFA; }}
[data-testid="stSidebar"] h3 {{ color:#FFB877 !important; font-size:.78rem; letter-spacing:.08em; text-transform:uppercase; margin:.6rem 0 .2rem; }}
[data-testid="stSidebar"] [data-baseweb="select"] > div, [data-testid="stSidebar"] input {{ background:#fff !important; color:#1B2440 !important; }}
[data-testid="stSidebar"] [data-baseweb="select"] * {{ color:#1B2440; }}
[data-testid="stSidebar"] [data-baseweb="tag"] {{ background:{ORANGE} !important; }}
[data-testid="stSidebar"] [data-baseweb="tag"] * {{ color:#fff !important; }}
[data-testid="stSidebar"] .stButton button {{ background:rgba(255,255,255,.1); border:1px solid rgba(255,255,255,.3); }}
[data-testid="stSidebar"] [data-testid="stExpander"] {{ background:rgba(255,255,255,.06); border-color:rgba(255,255,255,.18); }}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{ color:#C5CDE8 !important; }}

/* hero banner */
.hero {{ background:linear-gradient(115deg,#0B1F66 0%,#1C3596 62%,#2C4BC0 100%); color:#fff; border-radius:16px;
        padding:18px 24px; display:flex; justify-content:space-between; align-items:center; gap:18px; flex-wrap:wrap;
        box-shadow:0 8px 24px rgba(11,31,102,.18); border-bottom:4px solid {ORANGE}; margin-bottom:.6rem; }}
.hero .t {{ font-size:1.35rem; font-weight:800; letter-spacing:.01em; }}
.hero .p {{ font-size:.95rem; opacity:.9; margin-top:2px; }}
.hero .chips {{ display:flex; gap:8px; flex-wrap:wrap; }}
.hero .chip {{ background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.25); padding:4px 12px; border-radius:20px; font-size:.82rem; white-space:nowrap; }}

/* tabs as pills (works with old baseweb and new react-aria tab markup) */
.stTabs [role="tablist"] {{ gap:6px; flex-wrap:wrap; border-bottom:none !important; padding:4px 0 10px; }}
.stTabs [role="tab"] {{ background:#fff; border:1px solid var(--line) !important; border-radius:999px !important; padding:6px 16px !important; height:auto; font-weight:600; color:var(--muted); }}
.stTabs [role="tab"]:hover {{ border-color:{NAVY} !important; }}
.stTabs [role="tab"][aria-selected="true"] {{ background:{NAVY}; color:#fff !important; border-color:{NAVY} !important; box-shadow:0 3px 10px rgba(11,31,102,.25); }}
.stTabs [role="tab"][aria-selected="true"] p {{ color:#fff !important; }}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"], .stTabs [role="tablist"] > span {{ display:none; }}

/* KPI cards */
.kpi {{ background:var(--card); border:1px solid var(--line); border-radius:14px; padding:12px 16px 10px; height:100%;
        box-shadow:0 1px 3px rgba(16,24,40,.06); border-top:3px solid {NAVY}; position:relative; overflow:hidden; }}
.kpi.good {{ border-top-color:{GREEN}; }} .kpi.bad {{ border-top-color:{RED}; }} .kpi.warn {{ border-top-color:{ORANGE}; }}
.kpi .l {{ font-size:.74rem; color:var(--muted); font-weight:700; text-transform:uppercase; letter-spacing:.04em; }}
.kpi .v {{ font-size:clamp(1.3rem,2vw,1.85rem); font-weight:800; color:var(--ink); line-height:1.2; white-space:nowrap; margin-top:2px; }}
.kpi .v small {{ font-size:.55em; color:var(--muted); font-weight:600; margin-left:3px; }}
.kpi .s {{ font-size:.8rem; color:var(--muted); margin-top:2px; }}
.kpi svg.spark {{ display:block; width:100%; height:26px; margin-top:6px; opacity:.9; }}
.pill {{ display:inline-block; padding:1px 8px; border-radius:10px; font-weight:700; font-size:.78rem; margin-right:4px; }}
.pill.up {{ background:#E3F4EA; color:{GREEN}; }} .pill.dn {{ background:#FBE7E5; color:{RED}; }} .pill.nt {{ background:#EEF0F5; color:var(--muted); }}
.up {{ color:{GREEN}; font-weight:700; }} .dn {{ color:{RED}; font-weight:700; }}

/* progress (target tracker) */
.prog {{ background:var(--card); border:1px solid var(--line); border-radius:14px; padding:12px 16px; height:100%; box-shadow:0 1px 3px rgba(16,24,40,.06); }}
.prog .h {{ display:flex; justify-content:space-between; gap:8px; font-size:.8rem; font-weight:700; color:var(--muted); text-transform:uppercase; letter-spacing:.04em; }}
.prog .h b {{ color:var(--ink); font-size:1.05rem; text-transform:none; letter-spacing:0; }}
.prog .bar {{ height:10px; background:#EDF0F6; border-radius:6px; margin:8px 0 6px; overflow:hidden; position:relative; }}
.prog .bar i {{ display:block; height:100%; border-radius:6px; }}
.prog .bar em {{ position:absolute; left:100%; top:-3px; bottom:-3px; width:2px; background:var(--ink); opacity:.35; }}
.prog .f {{ font-size:.78rem; color:var(--muted); }}

/* section header */
.sec {{ display:flex; align-items:baseline; gap:10px; margin:1.3rem 0 .4rem; }}
.sec .bar {{ width:5px; height:20px; background:{ORANGE}; border-radius:3px; align-self:center; }}
.sec .t {{ font-size:1.12rem; font-weight:800; color:var(--ink); }}
.sec .d {{ font-size:.84rem; color:var(--muted); }}

/* highlight + note boxes */
.hl {{ background:linear-gradient(180deg,#FFF8F1,#FFFFFF); border:1px solid #F6D7B8; border-left:5px solid {ORANGE}; border-radius:12px; padding:12px 18px; }}
.hl ul {{ margin:0; padding-left:1.1rem; }} .hl li {{ margin:.25rem 0; color:var(--ink); }}
.slicechip {{ background:#E8F3EC; border:1px solid #9FD0AE; color:#17613A; border-radius:14px; padding:3px 12px; font-size:.84rem; display:inline-block; margin:2px 4px 2px 0; }}
.oops {{ background:#FFF6F5; border:1px solid #F3C3BD; border-radius:12px; padding:14px 18px; color:#7A271A; }}
[data-testid="stPlotlyChart"] {{ background:#fff; border:1px solid var(--line); border-radius:14px; padding:6px 8px 2px; box-shadow:0 1px 3px rgba(16,24,40,.06);
                                box-sizing:border-box; max-width:100%; overflow:hidden; }}
[data-testid="stVerticalBlockBorderWrapper"] {{ border-radius:14px; }}
div[data-testid="stDataFrame"] {{ border-radius:10px; overflow:hidden; }}
</style>"""


def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)


# ------------------------------------------------------------------ small HTML pieces
def _isnum(x) -> bool:
    try: return x is not None and not np.isnan(float(x))
    except (TypeError, ValueError): return False


def delta_pill(delta, suffix="%", label=""):
    if not _isnum(delta) or abs(delta) >= 999: return ""
    cls = "up" if delta >= 0 else "dn"
    return f"<span class='pill {cls}'>{'▲' if delta >= 0 else '▼'} {abs(delta):.1f}{suffix}</span>{label}"


def sparkline(values, w=200, h=26, color=NAVY) -> str:
    v = [float(x) for x in values if _isnum(x)]
    if len(v) < 2: return ""
    lo, hi = min(v), max(v); rng = (hi - lo) or 1
    pts = " ".join(f"{i * w / (len(v) - 1):.1f},{h - 2 - (x - lo) / rng * (h - 4):.1f}" for i, x in enumerate(v))
    return (f"<svg class='spark' viewBox='0 0 {w} {h}' preserveAspectRatio='none'><polyline fill='none' stroke='{color}' stroke-width='2' vector-effect='non-scaling-stroke' "
            f"stroke-linejoin='round' stroke-linecap='round' points='{pts}'/></svg>")


def kpi(label, value, sub="", delta=None, unit="", spark=None, tone=None) -> str:
    """KPI card. tone: None (auto from delta) | 'good' | 'bad' | 'warn' | 'plain'."""
    if tone is None: tone = "" if not _isnum(delta) else ("good" if delta >= 0 else "bad")
    elif tone == "plain": tone = ""
    sp = sparkline(spark) if spark is not None else ""
    u = f"<small>{html.escape(unit)}</small>" if unit else ""
    return (f"<div class='kpi {tone}'><div class='l'>{html.escape(str(label))}</div><div class='v'>{value}{u}</div>"
            f"<div class='s'>{delta_pill(delta)}{sub}</div>{sp}</div>")


def progress(label, actual, target, unit="KL", note="") -> str:
    if not _isnum(target) or target <= 0:
        return (f"<div class='prog'><div class='h'><span>{html.escape(label)}</span><b>{actual:,.0f} {unit}</b></div>"
                f"<div class='bar'><i style='width:0'></i></div><div class='f'>No target loaded {note}</div></div>")
    p = actual / target * 100
    col = GREEN if p >= 100 else (ORANGE if p >= 85 else RED)
    return (f"<div class='prog'><div class='h'><span>{html.escape(label)}</span><b style='color:{col}'>{p:.0f}%</b></div>"
            f"<div class='bar'><i style='width:{min(p, 100):.1f}%;background:{col}'></i></div>"
            f"<div class='f'>{actual:,.0f} of {target:,.0f} {unit} · {'surplus' if actual >= target else 'gap'} {abs(target - actual):,.0f} {unit} {note}</div></div>")


def cards(items, n=None):
    """Render a row of HTML cards (kpi()/progress() strings) in equal columns."""
    items = [i for i in items if i]
    if not items: return
    cols = st.columns(n or len(items))
    for c, h in zip(cols, items): c.markdown(h, unsafe_allow_html=True)


def section(title, desc=""):
    st.markdown(f"<div class='sec'><span class='bar'></span><span class='t'>{html.escape(title)}</span>"
                f"<span class='d'>{html.escape(desc)}</span></div>", unsafe_allow_html=True)


def hero(title, period, chips=()):
    ch = "".join(f"<span class='chip'>{html.escape(str(c))}</span>" for c in chips if c)
    st.markdown(f"<div class='hero'><div><div class='t'>{html.escape(title)}</div><div class='p'>{html.escape(period)}</div></div>"
                f"<div class='chips'>{ch}</div></div>", unsafe_allow_html=True)


def highlights(lines):
    if not lines: return
    st.markdown("<div class='hl'><ul>" + "".join(f"<li>{html.escape(l)}</li>" for l in lines) + "</ul></div>", unsafe_allow_html=True)


def layout(fig, h=360, legend=True, hover="x unified"):
    fig.update_layout(height=h, showlegend=legend, hovermode=hover)
    n_leg = sum(1 for tr in fig.data if tr.showlegend is not False)
    if any(isinstance(tr, go.Pie) for tr in fig.data) or (legend and n_leg > 4):
        # pies and many-series charts: legend down the right side so it never covers the title or the bars
        fig.update_layout(legend=dict(orientation="v", y=.5, x=1.02, yanchor="middle", traceorder="normal"), margin=dict(t=56, r=8))
    inside = [tr for tr in fig.data if isinstance(tr, go.Bar) and tr.textposition in ("inside", "auto")]
    for tr in inside: tr.update(textangle=0, insidetextanchor="middle")      # keep in-bar labels horizontal
    if inside and fig.layout.barmode == "stack":
        fig.update_layout(uniformtext=dict(minsize=9, mode="hide"))         # small slices: hide the label rather than squeeze it
    return fig


def wrap(label, width=14) -> str:
    """Break long axis labels over lines (plotly uses <br>)."""
    words, lines, cur = str(label).replace("/", "/ ").split(), [], ""
    for w in words:
        if cur and len(cur) + len(w) + 1 > width: lines.append(cur); cur = w
        else: cur = f"{cur} {w}".strip()
    if cur: lines.append(cur)
    return "<br>".join(lines[:3])


def growth_labels(fig, x, cy, ly, growth):
    """Growth % written above the taller of CY / LY so it never sits on the LY marker; adds head-room for it."""
    top = [max(a if a == a else 0, b if b == b else 0) for a, b in zip(cy, ly)]
    fig.add_scatter(x=list(x), y=top, mode="text", text=[f"{g:+.0f}%" if g == g and abs(g) < 999 else "" for g in growth],
                    textposition="top center", textfont=dict(size=11, color="#2B3350"), showlegend=False, hoverinfo="skip", cliponaxis=False)
    ymax = max(top + [0])
    fig.update_yaxes(range=[0, ymax * 1.2 if ymax > 0 else 1])
    return fig


def chart(fig, **kw):
    """Plotly chart (drawn as a card by the CSS)."""
    return st.plotly_chart(fig, width="stretch", config={"displaylogo": False}, **kw)


# ------------------------------------------------------------------ crash guard
@contextlib.contextmanager
def guard(where: str):
    """Any error inside a tab/section is shown as a small card; the rest of the dashboard keeps working.
    (Streamlit's own rerun/stop signals are BaseException and pass straight through.)"""
    try:
        yield
    except Exception as e:  # noqa: BLE001
        st.markdown(f"<div class='oops'><b>{html.escape(where)}</b> could not be shown for this selection "
                    f"({html.escape(type(e).__name__)}: {html.escape(str(e)[:200])}).<br>"
                    f"Try another period or filter - the other tabs are not affected.</div>", unsafe_allow_html=True)
        with st.expander("Technical details (for support)"):
            st.code(traceback.format_exc())
