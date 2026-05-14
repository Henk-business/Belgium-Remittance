import streamlit as st

st.set_page_config(
    page_title="AR Suite · AB InBev",
    page_icon="🍺",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── AB InBev brand palette (2022 modern identity) ──────────────────────────
# Black  #0A0A0A  · Orange/Gold #F7951D  · White #FFFFFF  · Off-white #FAF8F5

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ── Global ───────────────────────────────────────────────────────────────── */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #FAF8F5;
}

/* ── Sidebar ──────────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: #0A0A0A !important;
    min-width: 230px;
    border-right: 1px solid rgba(247,149,29,0.15);
}

/* Orange top stripe on sidebar */
[data-testid="stSidebar"]::before {
    content: "";
    display: block;
    height: 3px;
    background: linear-gradient(90deg, #F7951D, #E07B00, #F7951D);
}

/* Sidebar text */
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] .stRadio label p,
[data-testid="stSidebar"] .stRadio label span {
    color: #E8E3DC !important;
    font-size: 13px;
}

/* Sidebar nav */
[data-testid="stSidebar"] .stRadio > div { gap: 2px; }
[data-testid="stSidebar"] .stRadio label {
    background: rgba(255,255,255,0.03);
    border-radius: 8px;
    padding: 9px 14px !important;
    margin: 1px 0;
    cursor: pointer;
    transition: all .15s ease;
    border: 1px solid transparent;
}
[data-testid="stSidebar"] .stRadio label:hover {
    background: rgba(255,199,44,0.12) !important;
    border-color: rgba(255,199,44,0.25) !important;
}
[data-testid="stSidebar"] .stRadio label:has(input:checked) {
    background: rgba(255,199,44,0.18) !important;
    border-color: #FFC72C !important;
}

[data-testid="stSidebarNav"] { display: none; }

/* ── Main layout ──────────────────────────────────────────────────────────── */
.block-container {
    padding-top: 0rem !important;
    padding-bottom: 2rem;
    max-width: 1200px;
}
/* Streamlit also injects padding via these selectors — zero them all */
section.main > div.block-container {
    padding-top: 0rem !important;
}
[data-testid="block-container"] {
    padding-top: 0rem !important;
}
#MainMenu, footer { visibility: hidden; }

/* ── Typography ───────────────────────────────────────────────────────────── */
h1, h2 { color: #0A0A0A !important; font-weight: 800 !important; letter-spacing: -0.02em !important; }
h3      { color: #1C1C1C !important; font-weight: 600 !important; }

.stMarkdown h2 {
    border-bottom: 2px solid #F7951D;
    padding-bottom: 6px;
    margin-bottom: 16px;
}
.stMarkdown h3 {
    color: #F7951D !important;
    font-size: 11px !important;
    font-weight: 700 !important;
    text-transform: uppercase;
    letter-spacing: 0.09em;
    margin-top: 20px;
}

/* ── Primary buttons ──────────────────────────────────────────────────────── */
.stButton > button {
    border-radius: 8px;
    font-weight: 600;
    font-size: 13px;
    letter-spacing: 0.01em;
    transition: all .15s ease;
    border: none;
}
.stButton > button[kind="primary"] {
    background: #0A0A0A !important;
    color: #F7951D !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.18);
}
.stButton > button[kind="primary"]:hover {
    background: #1C1C1C !important;
    box-shadow: 0 4px 16px rgba(0,0,0,0.28);
    transform: translateY(-1px);
}
.stButton > button[kind="secondary"] {
    background: white !important;
    color: #0A0A0A !important;
    border: 1.5px solid #D4CFC8 !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #F7951D !important;
    color: #F7951D !important;
}

/* ── File uploader ────────────────────────────────────────────────────────── */
[data-testid="stFileUploader"] {
    border: 2px dashed #D4CFC8;
    border-radius: 10px;
    background: white;
    transition: border-color .2s;
}
[data-testid="stFileUploader"]:hover { border-color: #F7951D; }

/* ── Tabs ─────────────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: transparent;
    border-bottom: 2px solid #E8E3DC;
    gap: 0;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0;
    padding: 8px 18px;
    font-size: 13px;
    font-weight: 500;
    color: #7A7065;
    border: none;
    background: transparent;
}
.stTabs [aria-selected="true"] {
    color: #0A0A0A !important;
    font-weight: 700 !important;
    border-bottom: 3px solid #F7951D !important;
    background: rgba(247,149,29,0.05) !important;
}

/* ── Expanders ────────────────────────────────────────────────────────────── */
[data-testid="stExpander"] {
    border: 1px solid #E8E3DC;
    border-radius: 10px;
    background: white;
    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
    margin-bottom: 8px;
}
[data-testid="stExpander"] summary {
    font-weight: 600;
    color: #0A0A0A;
    padding: 12px 16px;
}

/* ── Metrics ──────────────────────────────────────────────────────────────── */
[data-testid="stMetric"] {
    background: white;
    border-radius: 10px;
    padding: 14px 18px;
    border: 1px solid #E8E3DC;
    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
}
[data-testid="stMetricValue"] { color: #F7951D !important; font-weight: 800 !important; }

/* ── Form inputs ──────────────────────────────────────────────────────────── */
[data-testid="stSelectbox"] > div > div,
[data-testid="stTextInput"] > div > div > input,
[data-testid="stNumberInput"] > div > div > input,
[data-testid="stTextArea"] > div > div > textarea,
[data-testid="stDateInput"] > div > div > input {
    border-radius: 8px;
    border: 1.5px solid #D4CFC8;
    background: white;
}
[data-testid="stTextInput"] > div > div > input:focus {
    border-color: #F7951D;
    box-shadow: 0 0 0 2px rgba(247,149,29,0.15);
}

/* ── Download button ──────────────────────────────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: white !important;
    color: #0A0A0A !important;
    border: 2px solid #0A0A0A !important;
    border-radius: 8px;
    font-weight: 600;
    transition: all .15s;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #0A0A0A !important;
    color: #F7951D !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 14px rgba(0,0,0,0.2);
}

/* ── Checkbox ─────────────────────────────────────────────────────────────── */
[data-testid="stCheckbox"] span[aria-checked="true"] {
    background: #F7951D !important;
    border-color: #F7951D !important;
}

/* ── Dataframe ────────────────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid #E8E3DC;
}

/* ── Alerts ───────────────────────────────────────────────────────────────── */
[data-testid="stAlert"] { border-radius: 10px; border-left-width: 4px; }

/* ── Scrollbar ────────────────────────────────────────────────────────────── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: #FAF8F5; }
::-webkit-scrollbar-thumb { background: #D4CFC8; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #F7951D; }

/* ── Caption / small ──────────────────────────────────────────────────────── */
.stCaption, small { color: #7A7065 !important; font-size: 12px; }

/* ── HR ───────────────────────────────────────────────────────────────────── */
hr { border-color: #E8E3DC !important; }

/* ── Sidebar expander — force dark theme to match sidebar background ──────── */
[data-testid="stSidebar"] [data-testid="stExpander"] {
    background: rgba(255,255,255,0.04) !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
    border-radius: 10px !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary {
    color: #E8E3DC !important;
    font-size: 11px !important;
    padding: 8px 12px !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] summary svg {
    fill: #5A5550 !important;
}
[data-testid="stSidebar"] [data-testid="stExpander"] > div > div {
    background: transparent !important;
}
[data-testid="stSidebar"] .stButton > button {
    background: rgba(255,199,44,0.1) !important;
    color: #FFC72C !important;
    border: 1px solid rgba(255,199,44,0.25) !important;
    font-size: 12px !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,199,44,0.2) !important;
    border-color: #FFC72C !important;
    transform: none !important;
    box-shadow: none !important;
}
</style>
""", unsafe_allow_html=True)

APP_VERSION = "v130"

@st.cache_data(ttl=300, show_spinner=False)
def _check_github():
    """Check GitHub connectivity once per 5 minutes, not on every rerun."""
    try:
        from github_storage import github_configured, _repo, _headers
        import requests as _req
        if not github_configured():
            return False
        r = _req.get(f"https://api.github.com/repos/{_repo()}", headers=_headers(), timeout=4)
        return r.ok
    except Exception:
        return False

_gh_ok = _check_github()

_gh_dot   = "🟢" if _gh_ok else "🔴"
_gh_label = "GitHub connected" if _gh_ok else "GitHub offline"

# ── Sidebar branding ───────────────────────────────────────────────────────
st.sidebar.markdown(f"""
<div style='padding:22px 16px 18px; border-bottom:1px solid rgba(255,199,44,0.2); margin-bottom:12px;'>
    <div style='display:flex; align-items:center; gap:10px; margin-bottom:8px;'>
        <div style='background:#FFC72C;
                    width:40px; height:40px; border-radius:10px;
                    display:flex; align-items:center; justify-content:center;
                    font-size:22px; box-shadow:0 2px 12px rgba(255,199,44,0.35);
                    flex-shrink:0; line-height:1;'>🍺</div>
        <div>
            <div style='font-size:17px; font-weight:800; color:#FFFFFF;
                        letter-spacing:-0.02em; line-height:1.1;'>AR Suite</div>
            <div style='font-size:10px; color:#FFC72C; font-weight:600;
                        letter-spacing:0.1em; text-transform:uppercase; margin-top:1px;'>
                AB InBev · Belgium
            </div>
        </div>
    </div>
    <div style='font-size:11px; color:#5A5550; margin-top:4px;'>{_gh_dot} {_gh_label}</div>
    <div style='font-size:10px; color:#3A3530; margin-top:2px;'>{APP_VERSION}</div>
</div>
""", unsafe_allow_html=True)

PAGES = [
    "Home",
    "AR Calendar",
    "Remittance Reconciliation",
    "Account Splitter",
    "Customer Overview",
    "Bonus & Payout",
    "Help & FAQ",
]

# Handle redirect requests from page buttons (home page tool cards etc).
# Buttons cannot write to a radio's owned key mid-run, so they write to
# "_nav_to" instead. We read it here — BEFORE the radio renders — and apply.
if "_nav_to" in st.session_state:
    _dest = st.session_state.pop("_nav_to")
    if _dest in PAGES:
        st.session_state["active_page"] = _dest

if "active_page" not in st.session_state:
    st.session_state["active_page"] = "Home"
if st.session_state["active_page"] not in PAGES:
    st.session_state["active_page"] = "Home"

page = st.sidebar.radio(
    "Navigation",
    PAGES,
    key="active_page",
    label_visibility="collapsed",
)

# ── Persistent task widget in sidebar ─────────────────────────────────────
try:
    import datetime as _dt
    from calendar_data import CALENDAR, TYPE_COLORS as _TC

    _today    = _dt.date.today()
    _day      = _today.day
    _tasks    = CALENDAR.get(_day, [])

    # Find next day this month that has tasks
    _upcoming = []
    _next_day = None
    for _d in range(_day + 1, 32):
        _t = CALENDAR.get(_d, [])
        if _t:
            _next_day = _d
            _upcoming = _t
            break

    def _dot(typ):
        bg = _TC.get(typ, {"bg": "#888"})["bg"]
        return (f"<span style='display:inline-block;width:7px;height:7px;"
                f"border-radius:50%;background:{bg};flex-shrink:0;margin-top:4px;'></span>")

    def _pill(t):
        bg  = _TC.get(t["type"], {"bg":"#333","fg":"#FFC72C"})["bg"]
        fg  = _TC.get(t["type"], {"bg":"#333","fg":"#FFC72C"})["fg"]
        fmt = f" ({t['format']})" if t["format"] else ""
        return (f"<span style='background:{bg};color:{fg};font-size:9px;font-weight:700;"
                f"padding:1px 6px;border-radius:3px;letter-spacing:0.05em;white-space:nowrap;"
                f"border:1px solid rgba(255,255,255,0.08);'>"
                f"{t['type']}{fmt}</span>")

    # Month progress bar
    import calendar as _cal
    _, _dim = _cal.monthrange(_today.year, _today.month)
    _pct    = int((_day / _dim) * 100)

    month_section = (
        f"<div style='margin-bottom:14px;'>"
        f"<div style='font-size:10px;color:#9A9490;margin-bottom:5px;letter-spacing:0.04em;'>"
        f"{_today.strftime('%B')} &nbsp;·&nbsp; day {_day} of {_dim}</div>"
        f"<div style='background:rgba(255,255,255,0.1);border-radius:4px;height:4px;'>"
        f"<div style='background:#FFC72C;width:{_pct}%;height:4px;border-radius:4px;'></div>"
        f"</div></div>"
    )

    # Today tasks
    if _tasks:
        today_rows = "".join(
            f"<div style='display:flex;gap:8px;align-items:flex-start;"
            f"padding:6px 0;border-bottom:1px solid rgba(255,255,255,0.08);'>"
            f"{_dot(t['type'])}"
            f"<div style='flex:1;min-width:0;'>"
            f"<div style='font-size:12px;font-weight:600;color:#FFFFFF;line-height:1.3;"
            f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>{t['account']}</div>"
            f"<div style='margin-top:3px;'>{_pill(t)}</div>"
            f"</div></div>"
            for t in _tasks
        )
        expander_label = f"📋 Today · {len(_tasks)} task{'s' if len(_tasks)!=1 else ''}"
    else:
        today_rows     = (f"<div style='font-size:12px;color:#9A9490;"
                          f"padding:6px 0;'>Nothing scheduled today ✓</div>")
        expander_label = "📋 Today · clear"

    today_section = (
        f"<div style='font-size:9px;font-weight:700;color:#FFC72C;letter-spacing:0.1em;"
        f"text-transform:uppercase;margin-bottom:6px;'>"
        f"Today &nbsp;·&nbsp; {_today.strftime('%-d %b')}</div>"
        f"{today_rows}"
    )

    # Upcoming tasks
    if _upcoming and _next_day:
        _days_away  = _next_day - _day
        _up_heading = ("Tomorrow" if _days_away == 1
                       else f"In {_days_away} days") + f" · {_next_day} {_today.strftime('%b')}"
        up_rows = "".join(
            f"<div style='display:flex;gap:8px;align-items:flex-start;padding:5px 0;'>"
            f"{_dot(t['type'])}"
            f"<div style='flex:1;min-width:0;'>"
            f"<div style='font-size:12px;color:#C8C3BC;line-height:1.3;"
            f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>{t['account']}</div>"
            f"<div style='margin-top:3px;'>{_pill(t)}</div>"
            f"</div></div>"
            for t in _upcoming
        )
        upcoming_section = (
            f"<div style='font-size:9px;font-weight:700;color:#7A7065;letter-spacing:0.1em;"
            f"text-transform:uppercase;margin:12px 0 6px;'>{_up_heading}</div>"
            f"{up_rows}"
        )
    else:
        upcoming_section = (f"<div style='font-size:11px;color:#5A5550;margin-top:10px;'>"
                            f"No more tasks this month</div>")

    widget_html = (
        f"<div style='padding:2px 0 8px;'>"
        f"{month_section}{today_section}{upcoming_section}"
        f"</div>"
    )

    st.sidebar.markdown(
        f"<div style='font-size:11px;font-weight:600;color:#E8E3DC;"
        f"padding:10px 4px 4px;letter-spacing:0.02em;'>{expander_label}</div>",
        unsafe_allow_html=True
    )
    with st.sidebar.expander("", expanded=True):
        st.markdown(widget_html, unsafe_allow_html=True)
        if st.button("Open Calendar →", key="sb_cal_btn", use_container_width=True):
            st.session_state["_nav_to"] = "AR Calendar"
            st.rerun()

except Exception:
    pass

st.sidebar.markdown("""
<div style='position:fixed; bottom:0; left:0; width:230px;
            padding:12px 16px; background:#0A0A0A;
            border-top:1px solid rgba(255,199,44,0.1);'>
    <div style='font-size:10px; color:#3A3530; text-align:center; letter-spacing:0.05em;'>
        ACCOUNTS RECEIVABLE · BELGIUM
    </div>
</div>
""", unsafe_allow_html=True)

if page == "Home":
    import page_home; page_home.show()
elif page == "AR Calendar":
    import page_calendar; page_calendar.show()
elif page == "Remittance Reconciliation":
    import page_remittance; page_remittance.show()
elif page == "Account Splitter":
    import page_splitter; page_splitter.show()
elif page == "Customer Overview":
    import page_overview; page_overview.show()
elif page == "Bonus & Payout":
    import page_bonus; page_bonus.show()
elif page == "Help & FAQ":
    import page_faq; page_faq.show()
