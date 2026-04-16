import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import traceback

from reconcile_engine import (run_reconciliation, build_recon_report, build_statement,
                               find_amount_combinations, build_amount_match_report)
from common import get_email, LANG_LABELS, mailto_link


def show():
    st.markdown("## 🔍 Remittance Reconciliation")
    st.caption("Upload a SAP export and a client remittance. SAP is the source of truth.")

    tab1, tab2 = st.tabs(["📄 Remittance matching", "💰 Amount-only matching"])

    with tab1:
        _show_remittance()

    with tab2:
        _show_amount_match()


def _show_remittance():

    # ── FILES ────────────────────────────────────────────────────────────────
    st.markdown("### 1 · Upload files")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**SAP Export** — FBL5N or ALV open items (.xlsx)")
        sap_file = st.file_uploader("SAP", type=["xlsx", "xls"],
                                    label_visibility="collapsed", key="rem_sap")
    with col2:
        st.markdown("**Client Remittance** — payment advice from customer (.xlsx)")
        rem_file = st.file_uploader("Remittance", type=["xlsx", "xls"],
                                    label_visibility="collapsed", key="rem_rem")

    # ── DETAILS ──────────────────────────────────────────────────────────────
    st.markdown("### 2 · Payment details")
    d1, d2, d3 = st.columns(3)
    with d1:
        cname = st.text_input("Customer name", key="rem_cname", placeholder="e.g. Acme Corp")
    with d2:
        pmt = st.number_input("Payment amount (€)", min_value=0.0, value=0.0,
                              step=0.01, format="%.2f", key="rem_pmt")
    with d3:
        pmt_date = st.date_input("Payment date", value=None, key="rem_date")

    # ── RUN ──────────────────────────────────────────────────────────────────
    st.markdown("### 3 · Run")
    run_col, _ = st.columns([1, 2])
    with run_col:
        run = st.button("▶  Run Reconciliation", use_container_width=True,
                        type="primary", key="rem_run")

    if run:
        if not sap_file or not rem_file:
            st.error("Please upload both files.")
        else:
            with st.spinner("Matching against SAP…"):
                try:
                    results = run_reconciliation(
                        BytesIO(sap_file.read()),
                        BytesIO(rem_file.read()),
                        float(pmt) if pmt and pmt > 0 else None,
                        cname.strip() or "Customer",
                    )
                    st.session_state["rem_results"] = results
                except Exception as e:
                    st.error(f"Error: {e}")
                    with st.expander("Detail"):
                        st.code(traceback.format_exc())

    if "rem_results" not in st.session_state:
        return

    results = st.session_state["rem_results"]
    mi  = results["matched_inv"]
    mc  = results["matched_cred"]
    ac  = results["already_cleared"]
    nf  = results["not_found"]
    mfr = results["missing"]

    # ── METRICS ──────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Results")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("✓ Invoices Matched",  len(mi),  f"€{results['t_inv']:,.0f}")
    m2.metric("✓ Credits Matched",   len(mc),  f"€{results['t_cred']:,.0f}")
    m3.metric("⚠ Already Cleared",   len(ac),  "Potential doubles")
    m4.metric("✗ Not Found in SAP",  len(nf))
    m5.metric("SAP Open Only",       len(mfr), f"€{results['t_missing']:,.0f}")

    if ac:
        with st.expander(f"⚠️  Already Cleared — {len(ac)} potential double payments", expanded=True):
            st.warning("These are on the remittance but already cleared in SAP. Check before processing.")
            st.dataframe(pd.DataFrame([{
                "SAP Reference": i["sap_ref"],
                "Cleared Date":  str(i.get("cleared_date", ""))[:10],
                "Clearing Doc":  str(i.get("cleared_by", "")),
            } for i in ac]), use_container_width=True, hide_index=True)

    if nf:
        with st.expander(f"✗  Not Found in SAP — {len(nf)} items"):
            st.dataframe(pd.DataFrame([{
                "Value from Remittance": i["sap_ref"],
                "Context": i.get("context", ""),
            } for i in nf]), use_container_width=True, hide_index=True)

    if mi:
        with st.expander(f"✓  Matched Invoices — {len(mi)} items"):
            st.dataframe(pd.DataFrame([{
                "SAP Ref":    i["sap_ref"],
                "Due Date":   str(i.get("sap_due_date", ""))[:10],
                "Amount (€)": i.get("sap_amount"),
            } for i in mi]), use_container_width=True, hide_index=True)

    if mc:
        with st.expander(f"✓  Matched Credit Notes — {len(mc)} items"):
            st.dataframe(pd.DataFrame([{
                "SAP Ref":    i["sap_ref"],
                "Amount (€)": i.get("sap_amount"),
            } for i in mc]), use_container_width=True, hide_index=True)

    # ── DOWNLOADS ─────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Downloads")

    today     = datetime.date.today()
    safe_name = (results["customer_name"] or "Customer").replace(" ", "_")[:25]
    dl1, dl2  = st.columns(2)

    with dl1:
        st.markdown("**Full reconciliation report**")
        st.caption("All details — matched, doubles, not found, SAP open items")
        recon_bytes = build_recon_report(results)
        st.download_button(
            "⬇  Reconciliation Report",
            data=recon_bytes.getvalue(),
            file_name=f"Recon_{safe_name}_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="dl_recon",
        )

    with dl2:
        st.markdown("**Customer statement — what you still owe**")
        st.caption("Clean statement to send to the customer")
        stmt_bytes = build_statement(results, today)
        st.download_button(
            "⬇  Customer Statement",
            data=stmt_bytes.getvalue(),
            file_name=f"Statement_{safe_name}_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="dl_stmt",
        )

    # ── EMAIL ─────────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📧 Generate follow-up email")

    e1, e2, e3 = st.columns(3)
    with e1:
        lang = st.selectbox("Language", list(LANG_LABELS.keys()),
                            format_func=lambda x: LANG_LABELS[x], key="rem_lang")
    with e2:
        sender  = st.text_input("Your name",  key="rem_sender",  placeholder="Your Name")
    with e3:
        company = st.text_input("Company",    key="rem_company", placeholder="Your Company")

    to_email = st.text_input("Customer email address (optional — for mailto link)",
                             key="rem_to", placeholder="customer@example.com")

    pmt_display = (
        f"\u20ac{results['payment_amount']:,.2f}"
        if results["payment_amount"]
        else "your recent payment"
    )

    subject, body = get_email(
        "remittance", lang,
        customer_name=results["customer_name"] or "Sir/Madam",
        payment_amount=pmt_display,
        sender_name=sender  or "[Your Name]",
        company_name=company or "[Your Company]",
    )

    with st.expander("📋  Email draft", expanded=True):
        st.text_input("Subject", value=subject, key="rem_email_subj")
        st.text_area("Body", value=body, height=260, key="rem_email_body")

    if to_email:
        mailto = mailto_link(to_email, subject, body)
        st.markdown(
            f'<a href="{mailto}" style="display:inline-block;background:linear-gradient(135deg,#0f2942,#1d4ed8);'
            f'color:white;font-weight:600;padding:12px 28px;border-radius:8px;'
            f'text-decoration:none;font-size:14px;margin-top:8px;">📧 Open in Email Client</a>'
            f'<div style="font-size:11px;color:#94a3b8;margin-top:6px;">'
            f'Opens your default email app pre-filled. Attach the statement Excel before sending.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.caption("Enter the customer email above to get a one-click mailto link.")


def _show_amount_match():
    st.markdown("### 💰 Amount-only matching")
    st.caption(
        "Customer paid without sending a remittance? Enter the payment amount and upload "
        "the SAP open items — the system will find which invoice combinations add up to the payment."
    )

    st.markdown("### 1 · Upload SAP export")
    sap_file = st.file_uploader(
        "SAP export (FBL5N open items)", type=["xlsx","xls"],
        label_visibility="collapsed", key="amt_sap"
    )

    st.markdown("### 2 · Payment details")
    a1, a2, a3 = st.columns(3)
    with a1:
        cname = st.text_input("Customer name", key="amt_cname_w", placeholder="e.g. Acme Corp")
    with a2:
        pmt_amt = st.number_input(
            "Payment amount (€)", min_value=0.01, value=1000.0,
            step=0.01, format="%.2f", key="amt_pmt_w"
        )
    with a3:
        tolerance = st.number_input(
            "Tolerance (€)", min_value=0.0, value=0.05,
            step=0.01, format="%.2f", key="amt_tol",
            help="Maximum allowed difference between invoice total and payment amount"
        )

    run_col, _ = st.columns([1, 2])
    with run_col:
        run = st.button("🔍  Find matching invoices", use_container_width=True,
                        type="primary", key="amt_run")

    if run:
        if not sap_file:
            st.error("Please upload the SAP export.")
        else:
            with st.spinner("Searching for matching invoice combinations…"):
                try:
                    from io import BytesIO
                    matches, sap_df = find_amount_combinations(
                        BytesIO(sap_file.read()),
                        float(pmt_amt),
                        tolerance=float(tolerance),
                    )
                    st.session_state["amt_matches"]  = matches
                    st.session_state["amt_pmt_s"]      = float(pmt_amt)
                    st.session_state["amt_cname_s"]    = cname.strip()
                    st.session_state["amt_n_open"]   = len(sap_df[sap_df["is_open"]])
                except Exception as e:
                    st.error(f"Error: {e}")
                    import traceback
                    with st.expander("Detail"):
                        st.code(traceback.format_exc())

    if "amt_matches" not in st.session_state:
        return

    matches  = st.session_state["amt_matches"]
    pmt_amt  = st.session_state["amt_pmt_s"]
    cname    = st.session_state["amt_cname_s"]
    n_open   = st.session_state.get("amt_n_open", 0)

    st.markdown("---")
    if not matches:
        st.warning(
            f"No invoice combinations found that sum to €{pmt_amt:,.2f} "
            f"(within ±€{tolerance:.2f}). Try increasing the tolerance or check "
            f"whether the payment includes a credit note offset."
        )
        return

    st.success(f"Found {len(matches)} possible combination(s) from {n_open} open invoices.")

    for i, match in enumerate(matches, 1):
        conf_color = "🟢" if match["diff"] < 0.01 else "🟡"
        with st.expander(
            f"{conf_color} Option {i} — {match.get('label', str(match['n']) + ' item(s)')}  ·  "
            f"Total €{match['total']:,.2f}  ·  {match['confidence']}",
            expanded=(i == 1)
        ):
            rows = []
            for inv in match["invoices"]:
                rows.append({
                    "Document №":   inv.get("doc_number_str",""),
                    "Assignment":   inv.get("ref",""),
                    "Doc Date":     inv.get("doc_date",""),
                    "Net Due":      inv.get("net_due",""),
                    "Type":         inv.get("doc_type",""),
                    "Amount (€)":   inv.get("amount",0),
                })
            import pandas as pd
            st.dataframe(
                pd.DataFrame(rows),
                use_container_width=True, hide_index=True
            )
            st.caption(
                f"Sum of selected invoices: **€{match['total']:,.2f}**  ·  "
                f"Payment amount: **€{pmt_amt:,.2f}**  ·  "
                f"Difference: **€{match['diff']:,.2f}**"
            )
            if match.get('excluded_credits'):
                excl = match['excluded_credits']
                st.info(
                    f"ℹ️ {len(excl)} credit note(s) totalling "
                    f"**€{abs(sum(excl)):,.2f}** were excluded — these were likely "
                    f"added to the account after the customer submitted their payment. "
                    f"Credits excluded: {', '.join(f'€{abs(a):,.2f}' for a in excl)}"
                )

    # Download Excel with all options
    st.markdown("---")
    if st.button("📥  Download all options as Excel", key="amt_dl_btn"):
        try:
            report = build_amount_match_report(
                matches, pmt_amt,
                customer_name=cname,
                today=datetime.date.today(),
            )
            st.session_state["amt_report"] = report
        except Exception as e:
            st.error(f"Error building report: {e}")

    if "amt_report" in st.session_state:
        safe = (cname or "Customer").replace(" ","_")[:20]
        st.download_button(
            "⬇  Download match report",
            data=st.session_state["amt_report"].getvalue(),
            file_name=f"AmountMatch_{safe}_{pmt_amt:.0f}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="amt_dl",
        )
