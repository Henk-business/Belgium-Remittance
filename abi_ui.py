

def today_bar():
    """
    Compact 'what's due today' bar — call at the top of every page's show().
    Shows a slim strip with today's tasks if any, links to the Calendar page.
    """
    import datetime
    try:
        from calendar_data import CALENDAR, TYPE_COLORS
    except ImportError:
        return

    today = datetime.date.today()
    tasks = CALENDAR.get(today.day, [])
    if not tasks:
        return  # nothing today — don't show the bar at all

    pills = ""
    for t in tasks:
        c   = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
        fmt = f" ({t['format']})" if t["format"] else ""
        pills += (
            f"<span style='background:{c['bg']};color:{c['fg']};font-size:10px;"
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
        f"<span style='display:flex;flex-wrap:wrap;gap:4px;align-items:center;'>{pills}</span>"
        f"</div>",
        unsafe_allow_html=True
    )
