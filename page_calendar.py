"""AR Calendar — monthly task schedule for the Belgium AR team."""
import streamlit as st
import datetime
import calendar as cal_lib
from calendar_data import CALENDAR, TYPE_COLORS


def show():
    from abi_ui import page_header
    page_header("AR Calendar",
                "Monthly schedule — direct debits, overviews, UAC and meetings.",
                "📅")

    today     = datetime.date.today()
    day       = today.day

    # ── Today's tasks banner ───────────────────────────────────────────────
    today_tasks = CALENDAR.get(day, [])

    if today_tasks:
        pills = ""
        for t in today_tasks:
            c    = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
            fmt  = f" ({t['format']})" if t["format"] else ""
            note = f"<span style='font-size:11px;color:rgba(255,255,255,0.5);margin-left:4px;'>{t['note']}</span>" if t["note"] else ""
            bg_col = c["bg"]
            fg_col = c["fg"]
            pills += (
                f"<div style='display:inline-flex;align-items:center;gap:6px;"
                f"background:rgba(255,255,255,0.07);border:1px solid rgba(255,255,255,0.12);"
                f"border-left:3px solid {bg_col};border-radius:8px;"
                f"padding:8px 14px;margin:3px 0;'>"
                f"<span style='background:{bg_col};color:{fg_col};font-size:10px;"
                f"font-weight:700;padding:1px 7px;border-radius:3px;letter-spacing:0.06em;"
                f"white-space:nowrap;'>{t['type']}{fmt}</span>"
                f"<span style='font-weight:600;color:#fff;font-size:13px;'>{t['account']}</span>"
                f"{note}</div>\n"
            )

        html = (
            f"<div style='background:#0A0A0A;border-radius:14px;padding:20px 24px;"
            f"margin-bottom:24px;border:1px solid rgba(255,199,44,0.2);'>"
            f"<div style='font-size:11px;font-weight:700;color:#FFC72C;letter-spacing:0.08em;"
            f"text-transform:uppercase;margin-bottom:10px;'>"
            f"📋 Today — {today.strftime('%A %-d %B %Y')} &nbsp;·&nbsp; "
            f"{len(today_tasks)} task{'s' if len(today_tasks)!=1 else ''}</div>"
            f"<div style='display:flex;flex-direction:column;gap:0px;'>{pills}</div>"
            f"</div>"
        )
        st.markdown(html, unsafe_allow_html=True)
    else:
        st.markdown(
            f"<div style='background:#F5F3F0;border-radius:14px;padding:20px 24px;"
            f"margin-bottom:24px;border:1px solid #E8E3DC;text-align:center;'>"
            f"<div style='font-size:24px;margin-bottom:6px;'>✅</div>"
            f"<div style='font-weight:700;color:#0A0A0A;font-size:15px;'>"
            f"Nothing scheduled today — {today.strftime('%A %-d %B')}</div>"
            f"<div style='font-size:13px;color:#7A7065;margin-top:4px;'>Enjoy the breathing room.</div>"
            f"</div>",
            unsafe_allow_html=True
        )

    # ── Month selector ─────────────────────────────────────────────────────
    col_m, col_y, _ = st.columns([1, 1, 3])
    with col_m:
        months = ["January","February","March","April","May","June",
                  "July","August","September","October","November","December"]
        sel_month = st.selectbox("Month", months, index=today.month-1, key="cal_month")
    with col_y:
        sel_year = st.number_input("Year", value=today.year, min_value=2024,
                                   max_value=2030, step=1, key="cal_year")

    month_num = months.index(sel_month) + 1
    _, days_in_month = cal_lib.monthrange(sel_year, month_num)
    first_weekday, _ = cal_lib.monthrange(sel_year, month_num)

    # ── Legend (build as single string — avoids Streamlit nested-div sanitisation) ──
    legend_items = "".join(
        f"<span style='display:inline-flex;align-items:center;gap:5px;margin-right:12px;'>"
        f"<span style='display:inline-block;width:10px;height:10px;border-radius:3px;"
        f"background:{v['bg']};border:1px solid #D4CFC8;'></span>"
        f"<span style='font-size:11px;color:#5A5550;'>{v['label']}</span></span>"
        for v in TYPE_COLORS.values()
    )
    st.markdown(
        f"<div style='margin:8px 0 16px;line-height:2;'>{legend_items}</div>",
        unsafe_allow_html=True
    )

    # ── Calendar grid ──────────────────────────────────────────────────────
    day_names = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
    header = "".join(
        f"<div style='font-size:11px;font-weight:700;color:#7A7065;"
        f"text-align:center;padding:6px 0;letter-spacing:0.06em;'>{d}</div>"
        for d in day_names
    )

    cells = ["<div></div>"] * first_weekday

    for d in range(1, days_in_month + 1):
        tasks    = CALENDAR.get(d, [])
        is_today = (d == today.day and month_num == today.month and sel_year == today.year)
        is_wknd  = (first_weekday + d - 1) % 7 >= 5

        bg     = "#F5F3F0" if is_wknd else "white"
        border = "2px solid #FFC72C" if is_today else "1px solid #E8E3DC"
        num_bg = "#0A0A0A" if is_today else "transparent"
        num_fg = "#FFC72C" if is_today else ("#BABABA" if is_wknd else "#0A0A0A")
        num_br = "50%" if is_today else "3px"

        task_chips = ""
        for t in tasks[:4]:
            c   = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
            fmt = f" ({t['format']})" if t["format"] else ""
            task_chips += (
                f"<span style='display:block;background:{c['bg']};color:{c['fg']};"
                f"font-size:9px;font-weight:600;padding:2px 5px;border-radius:3px;"
                f"margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>"
                f"{t['type']}{fmt}: {t['account']}</span>"
            )
        if len(tasks) > 4:
            task_chips += (
                f"<span style='font-size:9px;color:#7A7065;margin-top:2px;display:block;'>"
                f"+{len(tasks)-4} more</span>"
            )

        cells.append(
            f"<div style='background:{bg};border:{border};border-radius:8px;"
            f"padding:6px;min-height:80px;'>"
            f"<span style='display:inline-flex;width:22px;height:22px;background:{num_bg};"
            f"border-radius:{num_br};align-items:center;justify-content:center;"
            f"font-size:12px;font-weight:700;color:{num_fg};margin-bottom:2px;'>{d}</span>"
            f"{task_chips}</div>"
        )

    while len(cells) % 7 != 0:
        cells.append("<div></div>")

    grid = (
        f"<div style='display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin-bottom:24px;'>"
        f"{header}{''.join(cells)}</div>"
    )
    st.markdown(grid, unsafe_allow_html=True)

    # ── Day detail picker ──────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**🔍 Look up a specific day**")
    pick_day = st.number_input(
        f"Day in {sel_month}", min_value=1, max_value=days_in_month,
        value=min(today.day, days_in_month) if month_num == today.month else 1,
        key="cal_pick"
    )
    pick_tasks   = CALENDAR.get(pick_day, [])
    picked_date  = datetime.date(sel_year, month_num, pick_day)

    if pick_tasks:
        st.markdown(f"**{picked_date.strftime('%A %-d %B')} — {len(pick_tasks)} task(s):**")
        for t in pick_tasks:
            c    = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
            fmt  = f" · {t['format']} format" if t["format"] else ""
            note = f" · {t['note']}" if t["note"] else ""
            st.markdown(
                f"<div style='display:flex;align-items:center;gap:10px;padding:10px 14px;"
                f"background:white;border:1px solid #E8E3DC;border-radius:8px;"
                f"border-left:4px solid {c['bg']};margin-bottom:5px;'>"
                f"<span style='background:{c['bg']};color:{c['fg']};font-size:10px;"
                f"font-weight:700;padding:2px 8px;border-radius:4px;"
                f"letter-spacing:0.06em;white-space:nowrap;'>{t['type']}</span>"
                f"<span style='font-weight:700;color:#0A0A0A;'>{t['account']}</span>"
                f"<span style='font-size:12px;color:#7A7065;'>{fmt}{note}</span>"
                f"</div>",
                unsafe_allow_html=True
            )
    else:
        st.info(f"Nothing scheduled on {picked_date.strftime('%-d %B')}.")
