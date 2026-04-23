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
    background: rgba(247,149,29,0.12) !important;
    border-color: rgba(247,149,29,0.2) !important;
}
[data-testid="stSidebar"] .stRadio label:has(input:checked) {
    background: rgba(247,149,29,0.18) !important;
    border-color: #F7951D !important;
}

[data-testid="stSidebarNav"] { display: none; }

/* ── Main layout ──────────────────────────────────────────────────────────── */
.block-container {
    padding-top: 1.2rem !important;
    padding-bottom: 2rem;
    max-width: 1200px;
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
</style>
""", unsafe_allow_html=True)

APP_VERSION = "v123"

_gh_ok = False
try:
    from github_storage import github_configured, _repo, _headers
    import requests as _req
    if github_configured():
        _r = _req.get(f"https://api.github.com/repos/{_repo()}", headers=_headers(), timeout=5)
        _gh_ok = _r.ok
except Exception:
    pass

_gh_dot   = "🟢" if _gh_ok else "🔴"
_gh_label = "GitHub connected" if _gh_ok else "GitHub offline"

# ── Sidebar branding ───────────────────────────────────────────────────────
st.sidebar.markdown(f"""
<div style='padding:22px 16px 18px; border-bottom:1px solid rgba(247,149,29,0.15); margin-bottom:12px;'>
    <div style='display:flex; align-items:center; gap:10px; margin-bottom:8px;'>
        <div style='background:linear-gradient(135deg,#F7951D,#E07B00);
                    width:38px; height:38px; border-radius:9px;
                    display:flex; align-items:center; justify-content:center;
                    font-size:20px; box-shadow:0 2px 10px rgba(247,149,29,0.4);
                    flex-shrink:0;'>⚡</div>
        <div>
            <div style='font-size:17px; font-weight:800; color:#FFFFFF;
                        letter-spacing:-0.02em; line-height:1.1;'>AR Suite</div>
            <div style='font-size:10px; color:#F7951D; font-weight:600;
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
    "🏠  Home",
    "🔍  Remittance Reconciliation",
    "📂  Account Splitter",
    "📊  Customer Overview",
    "🎁  Bonus & Payout",
    "❓  Help & FAQ",
]

if "active_page" in st.session_state and st.session_state["active_page"] in PAGES:
    default_idx = PAGES.index(st.session_state["active_page"])
else:
    default_idx = 0

page = st.sidebar.radio("Navigation", PAGES, index=default_idx, label_visibility="collapsed")
st.session_state["active_page"] = page

st.sidebar.markdown("""
<div style='position:fixed; bottom:0; left:0; width:230px;
            padding:12px 16px; background:#0A0A0A;
            border-top:1px solid rgba(247,149,29,0.12);'>
    <div style='font-size:10px; color:#3A3530; text-align:center; letter-spacing:0.05em;'>
        ACCOUNTS RECEIVABLE · BELGIUM
    </div>
</div>
""", unsafe_allow_html=True)

if page == "🏠  Home":
    import page_home; page_home.show()
elif page == "🔍  Remittance Reconciliation":
    import page_remittance; page_remittance.show()
elif page == "📂  Account Splitter":
    import page_splitter; page_splitter.show()
elif page == "📊  Customer Overview":
    import page_overview; page_overview.show()
elif page == "🎁  Bonus & Payout":
    import page_bonus; page_bonus.show()
elif page == "❓  Help & FAQ":
    import page_faq; page_faq.show()
