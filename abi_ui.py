"""
AB InBev UI helper — shared branded components for all AR Suite pages.
Colours: Black #0A0A0A · Orange/Gold #F7951D · White #FFFFFF · Dark grey #1C1C1C
"""
import streamlit as st


def page_header(title: str, subtitle: str, icon: str = ""):
    """Render a consistent branded page header — flush with top, never clipped."""
    st.markdown(f"""
    <div style="display:flex; align-items:center; gap:14px;
                padding:18px 0 16px 0;
                border-bottom:2px solid #F0EDE8;
                margin-bottom:22px;">
        <div style="background:linear-gradient(135deg,#F7951D,#E07B00);
                    width:46px; height:46px; border-radius:10px;
                    display:flex; align-items:center; justify-content:center;
                    font-size:22px; box-shadow:0 3px 10px rgba(247,149,29,0.35);
                    flex-shrink:0;">
            {icon}
        </div>
        <div>
            <div style="font-size:22px; font-weight:800; color:#0A0A0A;
                        line-height:1.2; letter-spacing:-0.02em;">{title}</div>
            <div style="font-size:13px; color:#7A7065; margin-top:3px;">{subtitle}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def section_header(number: str, label: str):
    """Numbered section divider."""
    st.markdown(f"""
    <div style="display:flex; align-items:center; gap:10px; margin:24px 0 12px;">
        <div style="background:#0A0A0A; color:#F7951D; font-size:11px; font-weight:800;
                    width:22px; height:22px; border-radius:50%;
                    display:flex; align-items:center; justify-content:center;
                    flex-shrink:0; font-family:monospace;">{number}</div>
        <div style="font-size:12px; font-weight:700; color:#0A0A0A;
                    text-transform:uppercase; letter-spacing:0.08em;">{label}</div>
        <div style="flex:1; height:1px; background:#E8E3DC;"></div>
    </div>
    """, unsafe_allow_html=True)
