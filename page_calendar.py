"""
AR Calendar — multi-calendar, editable monthly task schedule.
Calendars are stored in Streamlit session state (and GitHub if configured).
"""
import streamlit as st
import datetime
import calendar as cal_lib
import json


# ── Calendar storage helpers ───────────────────────────────────────────────

def _default_calendars():
    """Seed the built-in Wholesale Scope calendar from calendar_data.py."""
    try:
        from calendar_data import CALENDAR
        return {"Wholesale Scope": CALENDAR}
    except ImportError:
        return {"Wholesale Scope": {}}


def _load_calendars():
    """Load calendars from session state, seeding defaults if needed."""
    if "ar_calendars" not in st.session_state:
        st.session_state["ar_calendars"] = _default_calendars()
    return st.session_state["ar_calendars"]


def _save_calendars(cals):
    st.session_state["ar_calendars"] = cals


def _get_active(cals):
    key = st.session_state.get("ar_active_calendar")
    if key not in cals:
        key = list(cals.keys())[0]
        st.session_state["ar_active_calendar"] = key
    return key


# ── Type colours ──────────────────────────────────────────────────────────
TYPE_COLORS = {
    "DD":       {"bg": "#FFC72C", "fg": "#0A0A0A", "label": "Direct Debit"},
    "Overview": {"bg": "#0A0A0A", "fg": "#FFC72C", "label": "Overview"},
    "UAC":      {"bg": "#C41230", "fg": "#FFFFFF",  "label": "UAC"},
    "Meeting":  {"bg": "#2E75B6", "fg": "#FFFFFF",  "label": "Meeting"},
}
TASK_TYPES  = list(TYPE_COLORS.keys())
FORMAT_OPTS = ["", "DD", "Manual"]


# ── Main show() ───────────────────────────────────────────────────────────
def show():
    from abi_ui import page_header
    page_header("AR Calendar",
                "Monthly schedules — direct debits, overviews, UAC and meetings.",
                "📅")

    today = datetime.date.today()

    cals        = _load_calendars()
    active_name = _get_active(cals)

    # ── Top bar: calendar selector + actions ───────────────────────────────
    sel_c, new_c, del_c, edit_c = st.columns([3, 2, 1, 1])

    with sel_c:
        chosen = st.selectbox(
            "Calendar",
            list(cals.keys()),
            index=list(cals.keys()).index(active_name),
            key="cal_selector",
            label_visibility="collapsed",
        )
        if chosen != active_name:
            st.session_state["ar_active_calendar"] = chosen
            active_name = chosen

    with new_c:
        with st.popover("➕ New calendar"):
            new_name = st.text_input("Calendar name", key="cal_new_name",
                                     placeholder="e.g. Retail Scope")
            if st.button("Create", key="cal_create_btn", type="primary"):
                name = new_name.strip()
                if name and name not in cals:
                    cals[name] = {}
                    st.session_state["ar_active_calendar"] = name
                    _save_calendars(cals)
                    st.rerun()
                elif name in cals:
                    st.error("Name already exists.")

    with del_c:
        if len(cals) > 1:
            with st.popover("🗑"):
                st.caption(f"Delete **{active_name}**?")
                if st.button("Confirm delete", key="cal_del_confirm", type="primary"):
                    del cals[active_name]
                    st.session_state["ar_active_calendar"] = list(cals.keys())[0]
                    _save_calendars(cals)
                    st.rerun()

    with edit_c:
        edit_mode = st.toggle("✏️", key="cal_edit_mode",
                              help="Edit tasks", value=False)

    active_cal = cals[active_name]

    # ── Rename calendar ────────────────────────────────────────────────────
    if edit_mode:
        with st.expander("✏️ Rename this calendar"):
            rename_val = st.text_input("New name", value=active_name, key="cal_rename_inp")
            if st.button("Save name", key="cal_rename_btn"):
                new_n = rename_val.strip()
                if new_n and new_n != active_name:
                    if new_n in cals:
                        st.error("Name already in use.")
                    else:
                        cals[new_n] = cals.pop(active_name)
                        st.session_state["ar_active_calendar"] = new_n
                        _save_calendars(cals)
                        st.rerun()

    # ── Today's tasks banner ───────────────────────────────────────────────
    today_tasks = active_cal.get(today.day, [])

    if today_tasks:
        pills = ""
        for t in today_tasks:
            c      = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
            fmt    = f" ({t['format']})" if t.get("format") else ""
            note   = (f"<span style='font-size:11px;color:rgba(255,255,255,0.5);"
                      f"margin-left:4px;'>{t['note']}</span>") if t.get("note") else ""
            bg_col = c["bg"]; fg_col = c["fg"]
            pills += (
                f"<div style='display:inline-flex;align-items:center;gap:6px;"
                f"background:rgba(255,255,255,0.07);border:1px solid rgba(255,255,255,0.12);"
                f"border-left:3px solid {bg_col};border-radius:8px;padding:8px 14px;margin:3px 0;'>"
                f"<span style='background:{bg_col};color:{fg_col};font-size:10px;"
                f"font-weight:700;padding:1px 7px;border-radius:3px;letter-spacing:0.06em;"
                f"white-space:nowrap;'>{t['type']}{fmt}</span>"
                f"<span style='font-weight:600;color:#fff;font-size:13px;'>{t['account']}</span>"
                f"{note}</div>\n"
            )
        st.markdown(
            f"<div style='background:#0A0A0A;border-radius:14px;padding:20px 24px;"
            f"margin-bottom:24px;border:1px solid rgba(255,199,44,0.2);'>"
            f"<div style='font-size:11px;font-weight:700;color:#FFC72C;letter-spacing:0.08em;"
            f"text-transform:uppercase;margin-bottom:10px;'>"
            f"📋 Today — {today.strftime('%A %-d %B %Y')} &nbsp;·&nbsp; "
            f"{len(today_tasks)} task{'s' if len(today_tasks)!=1 else ''} · {active_name}</div>"
            f"<div style='display:flex;flex-direction:column;gap:0;'>{pills}</div></div>",
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f"<div style='background:#F5F3F0;border-radius:14px;padding:16px 24px;"
            f"margin-bottom:24px;border:1px solid #E8E3DC;text-align:center;'>"
            f"<div style='font-weight:700;color:#0A0A0A;font-size:15px;'>"
            f"✅ Nothing scheduled today — {today.strftime('%A %-d %B')} · {active_name}</div>"
            f"</div>",
            unsafe_allow_html=True
        )

    # ── Month selector ─────────────────────────────────────────────────────
    months = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]
    col_m, col_y, _ = st.columns([1, 1, 3])
    with col_m:
        sel_month = st.selectbox("Month", months, index=today.month-1, key="cal_month")
    with col_y:
        sel_year  = st.number_input("Year", value=today.year,
                                    min_value=2024, max_value=2030, step=1, key="cal_year")

    month_num = months.index(sel_month) + 1
    _, days_in_month = cal_lib.monthrange(sel_year, month_num)
    first_weekday, _ = cal_lib.monthrange(sel_year, month_num)

    # ── Legend ─────────────────────────────────────────────────────────────
    legend = "".join(
        f"<span style='display:inline-flex;align-items:center;gap:5px;margin-right:12px;'>"
        f"<span style='display:inline-block;width:10px;height:10px;border-radius:3px;"
        f"background:{v['bg']};border:1px solid #D4CFC8;'></span>"
        f"<span style='font-size:11px;color:#5A5550;'>{v['label']}</span></span>"
        for v in TYPE_COLORS.values()
    )
    st.markdown(f"<div style='margin:8px 0 16px;line-height:2;'>{legend}</div>",
                unsafe_allow_html=True)

    # ── Calendar grid ──────────────────────────────────────────────────────
    day_names = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
    header = "".join(
        f"<div style='font-size:11px;font-weight:700;color:#7A7065;"
        f"text-align:center;padding:6px 0;letter-spacing:0.06em;'>{d}</div>"
        for d in day_names
    )
    cells = ["<div></div>"] * first_weekday

    for d in range(1, days_in_month + 1):
        tasks    = active_cal.get(d, [])
        is_today = (d == today.day and month_num == today.month and sel_year == today.year)
        is_wknd  = (first_weekday + d - 1) % 7 >= 5
        bg     = "#F5F3F0" if is_wknd else "white"
        border = "2px solid #FFC72C" if is_today else "1px solid #E8E3DC"
        num_bg = "#0A0A0A" if is_today else "transparent"
        num_fg = "#FFC72C" if is_today else ("#BABABA" if is_wknd else "#0A0A0A")
        num_br = "50%" if is_today else "3px"
        chips  = ""
        for t in tasks[:4]:
            c   = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
            fmt = f" ({t['format']})" if t.get("format") else ""
            chips += (
                f"<span style='display:block;background:{c['bg']};color:{c['fg']};"
                f"font-size:9px;font-weight:600;padding:2px 5px;border-radius:3px;"
                f"margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>"
                f"{t['type']}{fmt}: {t['account']}</span>"
            )
        if len(tasks) > 4:
            chips += f"<span style='font-size:9px;color:#7A7065;margin-top:2px;display:block;'>+{len(tasks)-4} more</span>"
        cells.append(
            f"<div style='background:{bg};border:{border};border-radius:8px;padding:6px;min-height:80px;'>"
            f"<span style='display:inline-flex;width:22px;height:22px;background:{num_bg};"
            f"border-radius:{num_br};align-items:center;justify-content:center;"
            f"font-size:12px;font-weight:700;color:{num_fg};margin-bottom:2px;'>{d}</span>"
            f"{chips}</div>"
        )

    while len(cells) % 7 != 0:
        cells.append("<div></div>")

    st.markdown(
        f"<div style='display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin-bottom:24px;'>"
        f"{header}{''.join(cells)}</div>",
        unsafe_allow_html=True
    )

    # ── Day detail picker ──────────────────────────────────────────────────
    st.markdown("---")
    pick_col, _ = st.columns([2, 3])
    with pick_col:
        pick_day = st.number_input(
            f"🔍 Look up a specific day in {sel_month}",
            min_value=1, max_value=days_in_month,
            value=min(today.day, days_in_month) if month_num == today.month else 1,
            key="cal_pick"
        )

    pick_tasks  = active_cal.get(pick_day, [])
    picked_date = datetime.date(sel_year, month_num, pick_day)

    if pick_tasks:
        st.markdown(f"**{picked_date.strftime('%A %-d %B')} — {len(pick_tasks)} task(s):**")
        for t in pick_tasks:
            c    = TYPE_COLORS.get(t["type"], {"bg":"#E8E3DC","fg":"#0A0A0A"})
            fmt  = f" · {t.get('format','')} format" if t.get("format") else ""
            note = f" · {t.get('note','')}" if t.get("note") else ""
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
    st.markdown("---")
    import pandas as pd

    if edit_mode:
        st.markdown("**✏️ Edit tasks — click any cell to change, add rows at the bottom, select row + delete to remove**")
        st.caption("Hit **💾 Save changes** when done.")
    else:
        st.markdown("**📋 All scheduled tasks** — toggle ✏️ to edit")

    # Flatten calendar → dataframe
    rows = []
    for d in sorted(active_cal.keys()):
        for t in active_cal[d]:
            rows.append({
                "Day":     int(d),
                "Type":    t.get("type", ""),
                "Account": t.get("account", ""),
                "Format":  t.get("format", ""),
                "Note":    t.get("note", ""),
            })
    df_cal = pd.DataFrame(rows, columns=["Day","Type","Account","Format","Note"]) \
             if rows else pd.DataFrame(columns=["Day","Type","Account","Format","Note"])

    if edit_mode:
        edited = st.data_editor(
            df_cal,
            use_container_width=True,
            num_rows="dynamic",
            key="cal_table_editor",
            column_config={
                "Day": st.column_config.NumberColumn(
                    "Day", help="Day of month (1–31)",
                    min_value=1, max_value=31, step=1, required=True, width="small",
                ),
                "Type": st.column_config.SelectboxColumn(
                    "Type", options=TASK_TYPES, required=True, width="small",
                ),
                "Account": st.column_config.TextColumn(
                    "Account", required=True, width="medium",
                ),
                "Format": st.column_config.SelectboxColumn(
                    "Format", options=FORMAT_OPTS, width="small",
                ),
                "Note": st.column_config.TextColumn("Note", width="large"),
            },
            hide_index=True,
        )

        if st.button("💾 Save changes", type="primary", key="cal_save_btn"):
            new_cal = {}
            for _, row in edited.iterrows():
                try:
                    d = int(row["Day"])
                except (ValueError, TypeError):
                    continue
                if not (1 <= d <= 31):
                    continue
                acc = str(row.get("Account", "") or "").strip()
                if not acc:
                    continue
                task = {
                    "type":    str(row.get("Type", "Overview")),
                    "account": acc.upper(),
                    "format":  str(row.get("Format", "") or ""),
                    "note":    str(row.get("Note", "") or ""),
                }
                new_cal.setdefault(d, []).append(task)
            active_cal.clear()
            active_cal.update(new_cal)
            _save_calendars(cals)
            st.success("✅ Calendar saved.")
            st.rerun()
    else:
        st.dataframe(
            df_cal,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Day":     st.column_config.NumberColumn("Day",    width="small"),
                "Type":    st.column_config.TextColumn("Type",     width="small"),
                "Account": st.column_config.TextColumn("Account",  width="medium"),
                "Format":  st.column_config.TextColumn("Format",   width="small"),
                "Note":    st.column_config.TextColumn("Note",     width="large"),
            },
        )
