import streamlit as st

st.set_page_config(
    page_title="AR Suite · AB InBev",
    page_icon="🍺",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── AB InBev brand palette ──────────────────────────────────────────────────
# Primary: Deep Crimson, Gold, Near-Black, White
# Crimson  #C41230   Gold     #C8922A   Dark   #1A0A00   Warm grey #F5F0EB

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ── Global ───────────────────────────────────────────────────────────────── */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #FAF7F4;
}

/* ── Sidebar ──────────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1A0A00 0%, #2D1005 60%, #1A0A00 100%) !important;
    min-width: 230px;
    border-right: 1px solid rgba(200,146,42,0.2);
}

/* Sidebar logo area */
[data-testid="stSidebar"]::before {
    content: "";
    display: block;
    height: 4px;
    background: linear-gradient(90deg, #C41230, #C8922A, #C41230);
    margin-bottom: 0;
}

/* Sidebar text */
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] .stRadio label p,
[data-testid="stSidebar"] .stRadio label span {
    color: #F5EFE6 !important;
    font-size: 13px;
}

/* Sidebar nav items */
[data-testid="stSidebar"] .stRadio > div { gap: 3px; }
[data-testid="stSidebar"] .stRadio label {
    background: rgba(255,255,255,0.04);
    border-radius: 8px;
    padding: 9px 14px !important;
    margin: 1px 0;
    cursor: pointer;
    transition: all .18s ease;
    border: 1px solid transparent;
}
[data-testid="stSidebar"] .stRadio label:hover {
    background: rgba(196,18,48,0.18) !important;
    border-color: rgba(196,18,48,0.3) !important;
}
[data-testid="stSidebar"] .stRadio label[data-checked="true"],
[data-testid="stSidebar"] .stRadio label:has(input:checked) {
    background: rgba(196,18,48,0.25) !important;
    border-color: #C41230 !important;
}

/* Hide default nav */
[data-testid="stSidebarNav"] { display: none; }

/* ── Main layout ──────────────────────────────────────────────────────────── */
.block-container {
    padding-top: 1.8rem;
    padding-bottom: 2rem;
    max-width: 1200px;
}
#MainMenu, footer { visibility: hidden; }

/* ── Page headers (h1, h2) ────────────────────────────────────────────────── */
h1, h2 { color: #1A0A00 !important; font-weight: 700 !important; }
h3 { color: #2D1005 !important; font-weight: 600 !important; }

/* ── Streamlit headings via markdown ──────────────────────────────────────── */
.stMarkdown h2 { 
    border-bottom: 2px solid #C41230; 
    padding-bottom: 6px; 
    margin-bottom: 16px;
}
.stMarkdown h3 {
    color: #C41230 !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-top: 20px;
}

/* ── Buttons ──────────────────────────────────────────────────────────────── */
.stButton > button {
    border-radius: 8px;
    font-weight: 600;
    font-size: 13px;
    letter-spacing: 0.02em;
    transition: all .18s ease;
    border: none;
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #C41230, #9A0E25) !important;
    color: white !important;
    box-shadow: 0 2px 8px rgba(196,18,48,0.3);
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #D91535, #C41230) !important;
    box-shadow: 0 4px 14px rgba(196,18,48,0.4);
    transform: translateY(-1px);
}
.stButton > button[kind="secondary"] {
    background: white !important;
    color: #1A0A00 !important;
    border: 1px solid #D4C5B5 !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #C41230 !important;
    color: #C41230 !important;
}

/* ── File uploader ────────────────────────────────────────────────────────── */
[data-testid="stFileUploader"] {
    border: 2px dashed #D4C5B5;
    border-radius: 10px;
    background: white;
    transition: border-color .2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #C41230;
}

/* ── Tabs ─────────────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: transparent;
    border-bottom: 2px solid #E8DDD0;
    gap: 0;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0;
    padding: 8px 18px;
    font-size: 13px;
    font-weight: 500;
    color: #6B5744;
    border: none;
    background: transparent;
}
.stTabs [aria-selected="true"] {
    color: #C41230 !important;
    font-weight: 700 !important;
    border-bottom: 3px solid #C41230 !important;
    background: rgba(196,18,48,0.04) !important;
}

/* ── Expanders ────────────────────────────────────────────────────────────── */
[data-testid="stExpander"] {
    border: 1px solid #E8DDD0;
    border-radius: 10px;
    background: white;
    box-shadow: 0 1px 3px rgba(0,0,0,.05);
    margin-bottom: 8px;
}
[data-testid="stExpander"] summary {
    font-weight: 600;
    color: #1A0A00;
    padding: 12px 16px;
}

/* ── Info / success / warning / error boxes ───────────────────────────────── */
[data-testid="stAlert"] {
    border-radius: 10px;
    border-left-width: 4px;
}

/* ── Metrics ──────────────────────────────────────────────────────────────── */
[data-testid="stMetric"] {
    background: white;
    border-radius: 10px;
    padding: 12px 16px;
    border: 1px solid #E8DDD0;
    box-shadow: 0 1px 3px rgba(0,0,0,.04);
}
[data-testid="stMetricValue"] { color: #C41230 !important; font-weight: 700 !important; }

/* ── Selectbox / text inputs ──────────────────────────────────────────────── */
[data-testid="stSelectbox"] > div > div,
[data-testid="stTextInput"] > div > div > input,
[data-testid="stNumberInput"] > div > div > input,
[data-testid="stTextArea"] > div > div > textarea,
[data-testid="stDateInput"] > div > div > input {
    border-radius: 8px;
    border: 1px solid #D4C5B5;
    background: white;
}
[data-testid="stTextInput"] > div > div > input:focus {
    border-color: #C41230;
    box-shadow: 0 0 0 2px rgba(196,18,48,0.12);
}

/* ── Dataframe / table ────────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid #E8DDD0;
}

/* ── Download button ──────────────────────────────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: white !important;
    color: #C41230 !important;
    border: 2px solid #C41230 !important;
    border-radius: 8px;
    font-weight: 600;
    transition: all .18s;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #C41230 !important;
    color: white !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(196,18,48,0.3);
}

/* ── Checkbox ─────────────────────────────────────────────────────────────── */
[data-testid="stCheckbox"] span[aria-checked="true"] {
    background: #C41230 !important;
    border-color: #C41230 !important;
}

/* ── Spinner ──────────────────────────────────────────────────────────────── */
[data-testid="stSpinner"] > div {
    border-top-color: #C41230 !important;
}

/* ── Progress bar ─────────────────────────────────────────────────────────── */
[data-testid="stProgressBar"] > div {
    background: linear-gradient(90deg, #C41230, #C8922A) !important;
    border-radius: 4px;
}

/* ── Divider ──────────────────────────────────────────────────────────────── */
hr { border-color: #E8DDD0 !important; }

/* ── Caption / small text ─────────────────────────────────────────────────── */
.stCaption, small { color: #7A6555 !important; font-size: 12px; }

/* ── Radio buttons ────────────────────────────────────────────────────────── */
[data-testid="stRadio"] label[data-checked="true"] span {
    background: #C41230 !important;
    border-color: #C41230 !important;
}

/* ── Scrollbar ────────────────────────────────────────────────────────────── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #F5F0EB; }
::-webkit-scrollbar-thumb { background: #C8922A; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #C41230; }
</style>
""", unsafe_allow_html=True)

# Version
APP_VERSION = "v122"

# GitHub connection status
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
<div style='padding:24px 16px 20px; border-bottom:1px solid rgba(200,146,42,0.2); margin-bottom:14px;'>
    <div style='display:flex; align-items:center; gap:10px; margin-bottom:6px;'>
        <div style='background:linear-gradient(135deg,#C41230,#9A0E25);
                    width:36px; height:36px; border-radius:8px;
                    display:flex; align-items:center; justify-content:center;
                    font-size:20px; box-shadow:0 2px 8px rgba(196,18,48,0.4);'>🍺</div>
        <div>
            <div style='font-size:17px; font-weight:700; color:#F5EFE6; letter-spacing:0.02em;'>
                AR Suite
            </div>
            <div style='font-size:10px; color:#C8922A; font-weight:600; letter-spacing:0.08em; text-transform:uppercase;'>
                AB InBev · {APP_VERSION}
            </div>
        </div>
    </div>
    <div style='font-size:11px; color:#A08060; margin-top:8px;'>{_gh_dot} {_gh_label}</div>
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

page = st.sidebar.radio(
    "Navigation",
    PAGES,
    index=default_idx,
    label_visibility="collapsed",
)

st.session_state["active_page"] = page

# ── Bottom sidebar footer ──────────────────────────────────────────────────
st.sidebar.markdown("""
<div style='position:fixed; bottom:0; left:0; width:230px;
            padding:14px 16px; background:#1A0A00;
            border-top:1px solid rgba(200,146,42,0.2);'>
    <div style='font-size:10px; color:#7A5030; text-align:center; letter-spacing:0.04em;'>
        ACCOUNTS RECEIVABLE · BELGIUM
    </div>
</div>
""", unsafe_allow_html=True)

if page == "🏠  Home":
    import page_home
    page_home.show()
elif page == "🔍  Remittance Reconciliation":
    import page_remittance
    page_remittance.show()
elif page == "📂  Account Splitter":
    import page_splitter
    page_splitter.show()
elif page == "📊  Customer Overview":
    import page_overview
    page_overview.show()
elif page == "🎁  Bonus & Payout":
    import page_bonus
    page_bonus.show()
elif page == "❓  Help & FAQ":
    import page_faq
    page_faq.show()
