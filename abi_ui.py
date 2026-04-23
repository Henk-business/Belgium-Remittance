"""
AB InBev UI helper — shared branded components for all AR Suite pages.
Import and call page_header() at the top of each show() function.
"""
import streamlit as st


def page_header(title: str, subtitle: str, icon: str = ""):
    """Render a consistent branded page header."""
    st.markdown(f"""
    <div style="display:flex; align-items:center; gap:14px; 
                padding-bottom:16px; border-bottom:2px solid #E8DDD0; margin-bottom:24px;">
        <div style="background:linear-gradient(135deg,#C41230,#9A0E25);
                    width:44px; height:44px; border-radius:10px;
                    display:flex; align-items:center; justify-content:center;
                    font-size:22px; box-shadow:0 2px 8px rgba(196,18,48,0.3); flex-shrink:0;">
            {icon}
        </div>
        <div>
            <div style="font-size:22px; font-weight:800; color:#1A0A00; line-height:1.2;">{title}</div>
            <div style="font-size:13px; color:#7A6555; margin-top:2px;">{subtitle}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def section_header(number: str, label: str):
    """Render a numbered section label (replaces ### 1 · Upload)."""
    st.markdown(f"""
    <div style="display:flex; align-items:center; gap:10px; margin:22px 0 12px;">
        <div style="background:#C41230; color:white; font-size:11px; font-weight:800;
                    width:22px; height:22px; border-radius:50%;
                    display:flex; align-items:center; justify-content:center;
                    flex-shrink:0;">{number}</div>
        <div style="font-size:13px; font-weight:700; color:#1A0A00;
                    text-transform:uppercase; letter-spacing:0.06em;">{label}</div>
        <div style="flex:1; height:1px; background:#E8DDD0;"></div>
    </div>
    """, unsafe_allow_html=True)


def result_banner(text: str, kind: str = "success"):
    """Branded result / status banner."""
    colours = {
        "success": ("#F0FDF4", "#166534", "#BBF7D0", "✓"),
        "info":    ("#FFF7ED", "#9A3412", "#FED7AA", "ℹ"),
        "warning": ("#FFFBEB", "#92400E", "#FDE68A", "⚠"),
    }
    bg, fg, border, icon = colours.get(kind, colours["info"])
    st.markdown(f"""
    <div style="background:{bg}; border:1px solid {border}; border-left:4px solid {fg};
                border-radius:8px; padding:12px 16px; margin:12px 0;
                display:flex; align-items:center; gap:10px;">
        <span style="font-size:16px;">{icon}</span>
        <span style="font-size:13px; color:{fg}; font-weight:500;">{text}</span>
    </div>
    """, unsafe_allow_html=True)
