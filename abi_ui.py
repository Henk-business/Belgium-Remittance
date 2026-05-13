"""
AB InBev UI helper — shared branded components for all AR Suite pages.
Colours: Black #0A0A0A · Yellow-Orange #FFC72C · White #FFFFFF
"""
import streamlit as st
import datetime


def page_header(title: str, subtitle: str, icon: str = ""):
    """Render a consistent branded page header."""
    st.markdown(
        f"<div style='display:flex;align-items:center;gap:14px;"
        f"padding:20px 0 16px 0;border-bottom:2px solid #F0EDE8;margin-bottom:22px;'>"
        f"<div style='background:#FFC72C;width:46px;height:46px;border-radius:10px;"
        f"display:flex;align-items:center;justify-content:center;font-size:22px;"
        f"box-shadow:0 3px 10px rgba(255,199,44,0.35);flex-shrink:0;'>{icon}</div>"
        f"<div>"
        f"<div style='font-size:22px;font-weight:800;color:#0A0A0A;"
        f"line-height:1.2;letter-spacing:-0.02em;'>{title}</div>"
        f"<div style='font-size:13px;color:#7A7065;margin-top:3px;'>{subtitle}</div>"
        f"</div></div>",
        unsafe_allow_html=True
    )


def section_header(number: str, label: str):
    """Numbered section divider."""
    st.markdown(
        f"<div style='display:flex;align-items:center;gap:10px;margin:24px 0 12px;'>"
        f"<div style='background:#0A0A0A;color:#FFC72C;font-size:11px;font-weight:800;"
        f"width:22px;height:22px;border-radius:50%;display:flex;align-items:center;"
        f"justify-content:center;flex-shrink:0;font-family:monospace;'>{number}</div>"
        f"<div style='font-size:12px;font-weight:700;color:#0A0A0A;"
        f"text-transform:uppercase;letter-spacing:0.08em;'>{label}</div>"
        f"<div style='flex:1;height:1px;background:#E8E3DC;'></div>"
        f"</div>",
        unsafe_allow_html=True
    )


def today_bar():
    """
    Compact today's-tasks bar shown on every tool page.
    Renders nothing if nothing is scheduled today.
    """
    try:
        from calendar_data import CALENDAR, TYPE_COLORS
    except ImportError:
        return

    today = datetime.date.today()
    tasks = CALENDAR.get(today.day, [])
    if not tasks:
        return

    pills = ""
    for t in tasks:
        bg  = TYPE_COLORS.get(t["type"], {"bg": "#E8E3DC", "fg": "#0A0A0A"})["bg"]
        fg  = TYPE_COLORS.get(t["type"], {"bg": "#E8E3DC", "fg": "#0A0A0A"})["fg"]
        fmt = f" ({t['format']})" if t["format"] else ""
        pills += (
            f"<span style='background:{bg};color:{fg};font-size:10px;"
            f"font-weight:700;padding:2px 8px;border-radius:4px;"
            f"letter-spacing:0.05em;white-space:nowrap;margin-right:4px;'>"
            f"{t['type']}{fmt}: {t['account']}</span>"
        )

    st.markdown(
        f"<div style='background:#0A0A0A;border-radius:10px;padding:10px 16px;"
        f"margin-bottom:16px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;"
        f"border:1px solid rgba(255,199,44,0.2);'>"
        f"<span style='font-size:11px;font-weight:700;color:#FFC72C;"
        f"white-space:nowrap;letter-spacing:0.06em;text-transform:uppercase;flex-shrink:0;'>"
        f"📅 Today</span>"
        f"<span style='color:rgba(255,255,255,0.3);flex-shrink:0;'>|</span>"
        f"<div style='display:flex;flex-wrap:wrap;gap:4px;align-items:center;'>{pills}</div>"
        f"</div>",
        unsafe_allow_html=True
    )
