"""
AR Calendar — monthly task schedule for the Belgium AR team.
Shows what needs to be done on any given day of the month.
"""
import streamlit as st
import datetime
import calendar as cal_lib
from calendar_data import CALENDAR, TYPE_COLORS


def show():
    from abi_ui import page_header
    page_header(
        "AR Calendar",
        "Monthly schedule — direct debits, overviews, retours and meetings.",
        "📅"
    )

    today = datetime.date.today()
    day   = today.day

    # ── Today's tasks banner ───────────────────────────────────────────────
    today_tasks = CALENDAR.get(day, [])

    if today_tasks:
        task_pills = ""
        for t in today_tasks:
            c = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
            label = f"{t['type']}"
            if t["format"]: label += f" ({t['format']})"
            note = f" — {t['note']}" if t["note"] else ""
            task_pills += f"""
            <div style="display:flex; align-items:center; gap:10px; 
                        padding:10px 16px; background:white;
                        border:1px solid #E8E3DC; border-radius:10px;
                        border-left:4px solid {c['bg']}; margin-bottom:6px;">
                <span style="background:{c['bg']}; color:{c['fg']};
                             font-size:10px; font-weight:700; padding:2px 8px;
                             border-radius:4px; letter-spacing:0.06em; flex-shrink:0;">
                    {t['type']}
                </span>
                <span style="font-weight:700; color:#0A0A0A; font-size:14px;">
                    {t['account']}
                </span>
                <span style="font-size:12px; color:#7A7065;">{note}</span>
            </div>"""

        st.markdown(f"""
        <div style="background:#0A0A0A; border-radius:14px; padding:20px 24px;
                    margin-bottom:24px; border:1px solid rgba(255,199,44,0.2);">
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px;">
                <div style="background:#FFC72C; width:36px; height:36px; border-radius:8px;
                            display:flex; align-items:center; justify-content:center;
                            font-size:18px; flex-shrink:0;">📋</div>
                <div>
                    <div style="font-size:16px; font-weight:800; color:#FFFFFF;">
                        Today — {today.strftime('%A %-d %B %Y')}
                    </div>
                    <div style="font-size:12px; color:#FFC72C; margin-top:1px;">
                        {len(today_tasks)} task{'s' if len(today_tasks)!=1 else ''} scheduled
                    </div>
                </div>
            </div>
            {task_pills}
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div style="background:#F5F3F0; border-radius:14px; padding:20px 24px;
                    margin-bottom:24px; border:1px solid #E8E3DC; text-align:center;">
            <div style="font-size:28px; margin-bottom:8px;">✅</div>
            <div style="font-weight:700; color:#0A0A0A; font-size:16px;">
                Nothing scheduled today — {today.strftime('%A %-d %B')}
            </div>
            <div style="font-size:13px; color:#7A7065; margin-top:4px;">
                Enjoy the breathing room.
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── Month selector ─────────────────────────────────────────────────────
    col_m, col_y, _ = st.columns([1, 1, 3])
    with col_m:
        months = ["January","February","March","April","May","June",
                  "July","August","September","October","November","December"]
        sel_month = st.selectbox("Month", months, index=today.month - 1, key="cal_month")
    with col_y:
        sel_year = st.number_input("Year", value=today.year, min_value=2024,
                                   max_value=2030, step=1, key="cal_year")

    month_num = months.index(sel_month) + 1
    _, days_in_month = cal_lib.monthrange(sel_year, month_num)
    first_weekday, _ = cal_lib.monthrange(sel_year, month_num)  # 0=Mon

    # Legend
    st.markdown("""
    <div style="display:flex; gap:10px; flex-wrap:wrap; margin:8px 0 16px;">
    """ + "".join(f"""
        <div style="display:flex; align-items:center; gap:5px;">
            <div style="width:10px; height:10px; border-radius:3px;
                        background:{v['bg']}; border:1px solid #D4CFC8;"></div>
            <span style="font-size:11px; color:#5A5550;">{v['label']}</span>
        </div>""" for v in TYPE_COLORS.values()) + """
    </div>""", unsafe_allow_html=True)

    # ── Full month calendar grid ───────────────────────────────────────────
    day_names = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]

    # Header row
    header_html = "".join(
        f'<div style="font-size:11px; font-weight:700; color:#7A7065; '
        f'text-align:center; padding:6px 0; letter-spacing:0.06em;">{d}</div>'
        for d in day_names
    )

    # Build calendar cells
    cells = []
    # Empty cells before first day
    for _ in range(first_weekday):
        cells.append('<div></div>')

    for d in range(1, days_in_month + 1):
        tasks = CALENDAR.get(d, [])
        is_today = (d == today.day and month_num == today.month and sel_year == today.year)
        is_weekend = (first_weekday + d - 1) % 7 >= 5

        bg      = "#F5F3F0" if is_weekend else "white"
        border  = "2px solid #FFC72C" if is_today else "1px solid #E8E3DC"
        day_fg  = "#FFC72C" if is_today else ("#BABABA" if is_weekend else "#0A0A0A")
        day_bg  = "#0A0A0A" if is_today else "transparent"
        day_br  = "50%" if is_today else "0"

        task_html = ""
        for t in tasks[:4]:  # max 4 shown in cell
            c = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
            fmt = f" ({t['format']})" if t["format"] else ""
            task_html += f"""
            <div style="background:{c['bg']}; color:{c['fg']};
                        font-size:9px; font-weight:600; padding:2px 5px;
                        border-radius:3px; margin-top:2px; line-height:1.3;
                        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                {t['type']}{fmt}: {t['account']}
            </div>"""
        if len(tasks) > 4:
            task_html += f'<div style="font-size:9px; color:#7A7065; margin-top:2px;">+{len(tasks)-4} more</div>'

        cells.append(f"""
        <div style="background:{bg}; border:{border}; border-radius:8px;
                    padding:6px; min-height:80px; position:relative;">
            <div style="display:inline-flex; width:22px; height:22px;
                        background:{day_bg}; border-radius:{day_br};
                        align-items:center; justify-content:center; margin-bottom:2px;">
                <span style="font-size:12px; font-weight:700; color:{day_fg};">{d}</span>
            </div>
            {task_html}
        </div>""")

    # Pad end of grid to complete last week
    while len(cells) % 7 != 0:
        cells.append('<div></div>')

    grid_html = f"""
    <div style="display:grid; grid-template-columns:repeat(7,1fr); gap:4px; margin-bottom:24px;">
        {header_html}
        {"".join(cells)}
    </div>"""

    st.markdown(grid_html, unsafe_allow_html=True)

    # ── Day detail picker ──────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**🔍 Look up a specific day**")
    pick_day = st.number_input(
        f"Day in {sel_month}", min_value=1, max_value=days_in_month,
        value=today.day if month_num == today.month else 1, key="cal_pick"
    )
    pick_tasks = CALENDAR.get(pick_day, [])
    picked_date = datetime.date(sel_year, month_num, pick_day)

    if pick_tasks:
        st.markdown(f"**{picked_date.strftime('%A %-d %B')} — {len(pick_tasks)} task(s):**")
        for t in pick_tasks:
            c = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
            fmt = f" · {t['format']} format" if t["format"] else ""
            note = f" · {t['note']}" if t["note"] else ""
            st.markdown(f"""
            <div style="display:flex; align-items:center; gap:10px; padding:10px 14px;
                        background:white; border:1px solid #E8E3DC; border-radius:8px;
                        border-left:4px solid {c['bg']}; margin-bottom:5px;">
                <span style="background:{c['bg']}; color:{c['fg']}; font-size:10px;
                             font-weight:700; padding:2px 8px; border-radius:4px;
                             letter-spacing:0.06em; flex-shrink:0;">{t['type']}</span>
                <span style="font-weight:700; color:#0A0A0A;">{t['account']}</span>
                <span style="font-size:12px; color:#7A7065;">{fmt}{note}</span>
            </div>""", unsafe_allow_html=True)
    else:
        st.info(f"Nothing scheduled on {picked_date.strftime('%-d %B')}.")
