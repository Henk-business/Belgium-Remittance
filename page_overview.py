import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import traceback

from overview_engine import prepare_df, build_overview


def show():
    st.markdown("## 📊 Customer Overview")
    st.caption(
        "Generate a year-by-year breakdown for a customer. "
        "Each year appears as its own section showing all transactions "
        "that occurred during that year, with a subtotal per year and a grand total."
    )

    # ── UPLOAD ────────────────────────────────────────────────────────────────
    st.markdown("### 1 · Upload SAP export")
    st.markdown("**SAP Export** — FBL5N full history for this customer (.xlsx)")
    uploaded = st.file_uploader(
        "SAP file", type=["xlsx", "xls"],
        label_visibility="collapsed", key="ov_file",
    )

    if not uploaded:
        st.info(
            "Export from SAP (FBL5N) the full transaction history for the customer, "
            "covering all years you need. The tool groups transactions by document date "
            "into yearly sections automatically."
        )
        return

    # ── PARSE ─────────────────────────────────────────────────────────────────
    try:
        file_bytes = uploaded.read()
        df, amt_col = prepare_df(BytesIO(file_bytes))
    except Exception as e:
        st.error(f"Could not read file: {e}")
        with st.expander("Detail"):
            st.code(traceback.format_exc())
        return

    # Detect accounts
    acc_col = next(
        (c for c in df.columns
         if c.lower() in ("account","customer","debitor","debiteurnummer",
                          "konto","klant","debiteur")), None
    )
    accounts = []
    if acc_col:
        accounts = sorted([
            str(a).strip().split(".")[0]
            for a in df[acc_col].dropna().unique()
            if str(a).strip() not in ("","nan")
        ])

    # Detect year range from document dates
    doc_date_col = next(
        (c for c in df.columns
         if "document date" in c.lower()
         or "belegdatum" in c.lower()
         or "boekingsdatum" in c.lower()), None
    )
    if doc_date_col and doc_date_col in df.columns:
        valid = df[doc_date_col].dropna()
        yr_min = int(valid.min().year) if len(valid) else datetime.date.today().year - 5
        yr_max = int(valid.max().year) if len(valid) else datetime.date.today().year
    else:
        yr_min = datetime.date.today().year - 5
        yr_max = datetime.date.today().year

    st.success(
        f"File loaded — {len(df):,} rows  ·  "
        f"{len(accounts)} account(s)  ·  "
        f"Dates: {yr_min}–{yr_max}"
    )

    # ── SETTINGS ──────────────────────────────────────────────────────────────
    st.markdown("### 2 · Settings")

    s1, s2, s3, s4 = st.columns(4)
    with s1:
        year_from = st.number_input(
            "From year", min_value=2000, max_value=2099,
            value=yr_min, step=1, key="ov_from_input",
        )
    with s2:
        year_to = st.number_input(
            "To year", min_value=2000, max_value=2099,
            value=yr_max, step=1, key="ov_to_input",
        )
    with s3:
        customer_name = st.text_input(
            "Customer name", key="ov_cname_w",
            placeholder="e.g. ACME Corp",
        )
    with s4:
        if len(accounts) > 1:
            account_filter = st.selectbox(
                "Account",
                ["All accounts"] + accounts,
                key="ov_acc_w",
            )
        else:
            account_filter = accounts[0] if accounts else "All accounts"
            st.text_input("Account", value=account_filter,
                          key="ov_acc_disp_w", disabled=True)

    if int(year_from) > int(year_to):
        st.error("'From year' must be before or equal to 'To year'.")
        return

    n_years = int(year_to) - int(year_from) + 1
    st.caption(
        f"Will generate {n_years} year section(s) on one sheet — "
        "all transactions grouped by document date."
    )

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

        # Filter to selected account
        if account_filter != "All accounts" and acc_col:
            work_df = work_df[
                work_df[acc_col].astype(str).str.strip().str.split(".").str[0]
                == str(account_filter)
            ].copy()

        if len(work_df) == 0:
            st.error("No data found for the selected account.")
            return

        with st.spinner(f"Building {n_years}-year overview…"):
            try:
                result = build_overview(
                    work_df, amt_col,
                    int(year_from), int(year_to),
                    customer_name=customer_name.strip(),
                    account_id=(account_filter
                                if account_filter != "All accounts" else ""),
                )
                st.session_state["ov_result"]  = result
                st.session_state["ov_acc"]     = account_filter
                st.session_state["ov_from"]    = int(year_from)
                st.session_state["ov_to"]      = int(year_to)
                st.session_state["ov_cname"]   = customer_name.strip()
                st.session_state["ov_nrows"]   = len(work_df)
                st.session_state["ov_ready"]   = True
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
    nrows    = st.session_state.get("ov_nrows", 0)

    st.markdown("---")
    st.success(
        f"Done — {to_yr - from_yr + 1} year sections  ·  {nrows:,} transactions total"
    )

    # Build filename
    parts = []
    if cname:
        parts.append(cname.replace(" ", "_")[:20])
    if acc_lbl != "All accounts":
        parts.append(str(acc_lbl))
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
