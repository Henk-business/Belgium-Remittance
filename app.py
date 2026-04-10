import streamlit as st

st.set_page_config(
    page_title="AR Suite",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

[data-testid="stSidebar"] {
    background: #0f172a !important;
    min-width: 220px;
}

/* White text for all sidebar content */
[data-testid="stSidebar"] .stRadio label p,
[data-testid="stSidebar"] .stRadio label span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span {
    color: #f1f5f9 !important;
}

[data-testid="stSidebar"] .stRadio > div { gap: 4px; }
[data-testid="stSidebar"] .stRadio label {
    background: rgba(255,255,255,0.04);
    border-radius: 8px;
    padding: 8px 12px !important;
    margin: 2px 0;
    cursor: pointer;
    transition: background .15s;
}
[data-testid="stSidebar"] .stRadio label:hover {
    background: rgba(255,255,255,0.12) !important;
}

[data-testid="stSidebarNav"] { display: none; }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
#MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.sidebar.markdown("""
<div style='padding:20px 12px 20px; border-bottom:1px solid #1e293b; margin-bottom:12px;'>
    <div style='font-size:20px; font-weight:700; color:#f1f5f9;'>💼 AR Suite</div>
    <div style='font-size:11px; color:#475569; margin-top:4px;'>Accounts Receivable Tools</div>
</div>
""", unsafe_allow_html=True)

PAGES = [
    "🏠  Home",
    "🔍  Remittance Reconciliation",
    "📂  Account Splitter",
    "📊  Customer Overview",
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
