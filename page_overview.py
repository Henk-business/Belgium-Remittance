import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import traceback

from overview_engine import prepare_df, build_overview
from common import get_email, mailto_link, LANG_LABELS, detect_customer_name


def show():
    st.markdown("## 📊 Customer Overview")
    st.caption(
        "Generate a year-by-year breakdown for a customer. "
        "Each year shows all clearing groups with their transactions, "
        "sorted newest to oldest by net due date."
    )

    # ── UPLOAD ────────────────────────────────────────────────────────────────
    st.markdown("### 1 · Upload SAP export")
    uploaded = st.file_uploader(
        "SAP Export — FBL5N full history (.xlsx)",
        type=["xlsx","xls"],
        label_visibility="collapsed",
        key="ov_file",
    )

    if not uploaded:
        st.info(
            "Export the full transaction history from SAP (FBL5N) for this customer. "
            "Include all years you need. The tool uses the SAP grouping structure directly."
        )
        return

    # ── PARSE FILE — cache in session to survive reruns ───────────────────────
    file_key = f"ov_df_{uploaded.name}_{uploaded.size}"
    if file_key not in st.session_state:
        try:
            raw_bytes = uploaded.read()
            df, amt_col = prepare_df(BytesIO(raw_bytes))
            st.session_state[file_key]          = df
            st.session_state[file_key+"_amt"]   = amt_col
            st.session_state[file_key+"_bytes"] = raw_bytes
        except Exception as e:
            st.error(f"Could not read file: {e}")
            with st.expander("Detail"):
                st.code(traceback.format_exc())
            return

    df      = st.session_state[file_key]
    amt_col = st.session_state[file_key+"_amt"]

    # Detect accounts (from real rows only)
    acc_col = next(
        (c for c in df.columns
         if c.lower() in ("account","customer","debitor","debiteurnummer",
                          "konto","klant","debiteur")), None
    )
    accounts = []
    if acc_col:
        accounts = sorted(set(
            str(a).strip().split(".")[0]
            for a in df[acc_col].dropna()
            if str(a).strip() not in ("","nan","None")
        ))

    # Detect year range from net due date (preferred) or doc date
    net_due_col = next(
        (c for c in df.columns if "net due" in c.lower()
         or "vervaldatum" in c.lower()), None
    )
    date_col_for_range = net_due_col or next(
        (c for c in df.columns if "document date" in c.lower()
         or "belegdatum" in c.lower()), None
    )
    if date_col_for_range and date_col_for_range in df.columns:
        valid = df[date_col_for_range].dropna()
        yr_min = int(valid.min().year) if len(valid) else datetime.date.today().year - 5
        yr_max = int(valid.max().year) if len(valid) else datetime.date.today().year
    else:
        yr_min = datetime.date.today().year - 5
        yr_max = datetime.date.today().year

    n_real = df[acc_col].notna().sum() if acc_col else len(df)
    st.success(
        f"✓ File loaded — {n_real:,} transactions  ·  "
        f"{len(accounts)} account(s)  ·  "
        f"Detected range: {yr_min}–{yr_max}"
    )

    # ── SETTINGS ──────────────────────────────────────────────────────────────
    st.markdown("### 2 · Settings")

    # Mode: single year (current overview) vs multi-year
    mode_col, _ = st.columns([2,3])
    with mode_col:
        ov_mode = st.radio(
            "Overview type",
            ["📋 Current overview (single period)", "📅 Multi-year overview"],
            key="ov_mode_radio",
            horizontal=True,
        )
    single_mode = ov_mode.startswith("📋")

    # ── Row 1: always-visible settings ───────────────────────────────────────
    r1a, r1b, r1c = st.columns(3)
    with r1a:
        lang = st.selectbox(
            "Language", ["en","nl","fr"],
            format_func=lambda x: {"en":"🇬🇧 English","nl":"🇳🇱 Dutch","fr":"🇫🇷 French"}[x],
            key="ov_lang_w",
        )
    with r1b:
        _auto_cname = detect_customer_name(df_raw) if "df_raw" in dir() else ""
        customer_name = st.text_input(
            "Customer name", key="ov_cname_w",
            placeholder="e.g. ACME Corp",
            value=st.session_state.get("ov_cname_detected", _auto_cname),
        )
    with r1c:
        if len(accounts) > 1:
            account_filter = st.selectbox(
                "Account", ["All accounts"] + accounts,
                key="ov_acc_w",
            )
        else:
            account_filter = accounts[0] if accounts else "All accounts"
            st.text_input("Account", value=account_filter,
                          key="ov_acc_disp_w", disabled=True)

    # ── Row 2: mode-specific settings ─────────────────────────────────────────
    r2a, r2b, r2c, r2d = st.columns(4)
    with r2a:
        # Current only: reference date
        ref_date = st.date_input(
            "Reference date",
            value=datetime.date.today(),
            key="ov_refdate_w",
            disabled=not single_mode,
            help="Only used in Current overview mode",
        )
    with r2b:
        # Current only: remove not due
        remove_not_due = st.checkbox(
            "Remove invoices not yet due",
            value=True, key="ov_remove_nd",
            disabled=not single_mode,
        )
    with r2c:
        # Multi-year only: from year
        year_from = st.number_input(
            "From year", min_value=2000, max_value=2099,
            value=yr_min, step=1, key="ov_from_input",
            disabled=single_mode,
            help="Only used in Multi-year mode",
        )
        if single_mode:
            year_from = yr_max
    with r2d:
        # Multi-year only: to year
        year_to = st.number_input(
            "To year", min_value=2000, max_value=2099,
            value=yr_max, step=1, key="ov_to_input",
            disabled=single_mode,
            help="Only used in Multi-year mode",
        )
        if single_mode:
            year_to = yr_max

    # Current mode: month range selector
    _month_names = ['January','February','March','April','May','June',
                    'July','August','September','October','November','December']
    if single_mode:
        _today = datetime.date.today()
        _cur_month = _today.month
        mt0, mt1, mt2, _ = st.columns([1, 1, 1, 1])
        with mt0:
            use_month_filter = st.checkbox(
                "Filter by month range", value=False, key="ov_use_months",
            )
        with mt1:
            month_from = st.selectbox(
                "From month", options=list(range(1,13)),
                format_func=lambda x: _month_names[x-1],
                index=0, key="ov_month_from",
                disabled=not use_month_filter,
            )
        with mt2:
            month_to = st.selectbox(
                "To month", options=list(range(1,13)),
                format_func=lambda x: _month_names[x-1],
                index=_cur_month-1, key="ov_month_to",
                disabled=not use_month_filter,
            )
        if not use_month_filter:
            month_from, month_to = 1, 12
    else:
        month_from, month_to = 1, 12

    year_from = int(year_from)
    year_to   = int(year_to)

    if year_from > year_to:
        st.error("'From year' must be before or equal to 'To year'.")
        return

    n_years = year_to - year_from + 1
    if single_mode:
        st.caption(f"Current overview — all transactions in {yr_max}.")
    else:
        st.caption(f"Will generate {n_years} year section(s) on one sheet.")

    # ── GENERATE ──────────────────────────────────────────────────────────────
    st.markdown("### 3 · Generate")
    gen_col, _ = st.columns([1, 2])
    with gen_col:
        generate = st.button(
            "▶  Generate Overview",
            use_container_width=True,
            type="primary",
            key="ov_run",
        )

    if generate:
        work_df = df.copy()

        # Filter to selected account — must keep blank separator rows intact
        if account_filter != "All accounts" and acc_col:
            # Mark which real rows belong to this account
            acc_str = work_df[acc_col].astype(str).str.strip().str.split(".").str[0]
            is_real  = acc_str.isin([str(account_filter)])
            is_blank = acc_str.isin(["", "nan", "None"]) | work_df[acc_col].isna()

            # Walk through rows: keep a blank row only if it immediately follows
            # real rows that belonged to this account
            keep = []
            last_real_kept = False
            for idx in work_df.index:
                if is_real[idx]:
                    keep.append(idx)
                    last_real_kept = True
                elif is_blank[idx]:
                    if last_real_kept:
                        keep.append(idx)
                    last_real_kept = False
                else:
                    last_real_kept = False

            work_df = work_df.loc[keep].copy()

        if work_df[acc_col].notna().sum() == 0 if acc_col else len(work_df) == 0:
            st.error("No data found for the selected account.")
            return

        with st.spinner(f"Building {n_years}-year overview…"):
            try:
                result = build_overview(
                    work_df, amt_col,
                    year_from, year_to,
                    customer_name=customer_name.strip(),
                    account_id=(account_filter if account_filter != "All accounts" else ""),
                    lang=lang,
                    reference_date=ref_date if single_mode else None,
                    remove_not_due=remove_not_due if single_mode else False,
                    month_from=month_from if single_mode else 1,
                    month_to=month_to if single_mode else 12,
                )
                st.session_state["ov_result"]      = result
                st.session_state["ov_acc"]         = account_filter
                st.session_state["ov_from"]        = year_from
                st.session_state["ov_to"]          = year_to
                st.session_state["ov_single_mode"] = single_mode
                st.session_state["ov_cname"]   = customer_name.strip()
                st.session_state["ov_nrows"]   = n_real
                st.session_state["ov_ready"]   = True
                st.session_state["ov_refdate_s"] = ref_date
                st.session_state["ov_lang"]    = lang
            except Exception as e:
                st.error(f"Error: {e}")
                with st.expander("Detail"):
                    st.code(traceback.format_exc())

    if not st.session_state.get("ov_ready") or "ov_result" not in st.session_state:
        return

    # ── RESULT ────────────────────────────────────────────────────────────────
    result   = st.session_state["ov_result"]
    acc_lbl  = st.session_state["ov_acc"]
    from_yr  = st.session_state["ov_from"]
    to_yr    = st.session_state["ov_to"]
    cname    = st.session_state["ov_cname"]

    st.markdown("---")
    st.success(f"Done — {to_yr - from_yr + 1} year section(s)")

    parts = []
    if cname: parts.append(cname.replace(" ","_")[:20])
    if acc_lbl != "All accounts": parts.append(str(acc_lbl))
    parts.append(f"{from_yr}-{to_yr}")
    filename = "Overview_" + "_".join(parts) + ".xlsx"

    st.download_button(
        "⬇  Download Overview Excel",
        data=result.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="ov_dl",
    )

    # ── EMAIL DRAFT ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📧 Email draft")

    e1, e2, e3, e4 = st.columns(4)
    with e1:
        email_lang = st.selectbox(
            "Email language",
            ["en", "nl", "fr"],
            format_func=lambda x: {"en": "🇬🇧 English", "nl": "🇳🇱 Dutch", "fr": "🇫🇷 French"}[x],
            key="ov_email_lang",
        )
    with e2:
        sender  = st.text_input("Your name",  key="ov_sender",  placeholder="Your Name",
                               value=st.session_state.get("_persist_sender",""))
    with e3:
        company = st.text_input("Company",    key="ov_company", placeholder="Your Company",
                               value=st.session_state.get("_persist_company",""))
    with e4:
        to_email = st.text_input("Customer email", key="ov_to_email", placeholder="customer@example.com")

    # Persist sender/company across pages
    if sender: st.session_state["_persist_sender"]  = sender
    if company: st.session_state["_persist_company"] = company

    subject, body = get_email(
        "overview", email_lang,
        customer_name=cname or f"Account {acc_lbl}",
        account_id=acc_lbl if acc_lbl != "All accounts" else "",
        sender_name=sender or "[Your Name]",
        company_name=company or "[Your Company]",
    )

    st.text_input("Subject", value=subject, key="ov_email_subj")
    st.text_area("Body", value=body, height=200, key="ov_email_body")

    st.caption(
        "📎 After downloading the Excel above, attach it manually to the email. "
        "Click the button below to open a pre-filled draft in your email client."
    )

    if to_email:
        mailto = mailto_link(to_email, subject, body)
        st.markdown(
            f'''<a href="{mailto}" style="display:block;text-align:center;
            padding:10px 20px;background:#2E75B6;color:white;border-radius:8px;
            text-decoration:none;font-weight:bold;margin-top:8px;">
            ✉  Open in email client</a>''',
            unsafe_allow_html=True,
        )
    else:
        st.info("Enter the customer email above to enable the open-in-email button.")
