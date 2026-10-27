"""Entry point of the Servo Sales Dashboard - a safety net around dashboard.py.

Streamlit runs this file; it runs dashboard.py. If anything unexpected goes wrong outside the per-tab guards, the
viewer gets a short message and a 'Try again' button instead of a broken page, and the details are written to
Sales Review/logs/dashboard_errors.log. Start:  ./run_dashboard.command   (or  streamlit run app.py)
"""
import runpy, traceback
from pathlib import Path

import streamlit as st

try:
    runpy.run_path(str(Path(__file__).with_name("dashboard.py")), run_name="__main__")
except Exception as e:  # noqa: BLE001  (Streamlit's rerun / stop signals are BaseException and pass through)
    tb = traceback.format_exc()
    try:
        from hodash.ui import log_error
        log_error("Whole page", tb)
    except Exception:
        pass
    st.markdown("<div class='oops' style='background:#FFF6F5;border:1px solid #F3C3BD;border-radius:12px;padding:14px 18px;color:#7A271A'>"
                f"<b>The dashboard</b> hit an unexpected problem ({type(e).__name__}: {str(e)[:200]}).<br>"
                "Your files are not affected. Click <b>Try again</b>; if it repeats, check that the sales files are closed in Excel "
                "and fully synced, or look at Sales Review/logs/dashboard_errors.log.</div>", unsafe_allow_html=True)
    if st.button("🔄 Try again", type="primary"):
        st.cache_data.clear(); st.rerun()
    with st.expander("Technical details (for support)"):
        st.code(tb)
