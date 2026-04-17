import streamlit as st


def show():
    st.markdown("""
    <style>
    .tool-card {
        background: white; border: 1px solid #e2e8f0; border-radius: 16px;
        padding: 28px; box-shadow: 0 1px 4px rgba(0,0,0,.06);
    }
    .tool-card .icon { font-size: 32px; margin-bottom: 14px; }
    .tool-card h3 { font-size: 16px; font-weight: 700; margin: 0 0 8px; color: #0f172a; }
    .tool-card p  { font-size: 13px; color: #64748b; line-height: 1.6; margin: 0 0 12px; }
    .feat { font-size: 12px; color: #475569; padding: 3px 0; }
    .feat::before { content: "→  "; color: #1d4ed8; font-weight: 600; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="background:linear-gradient(135deg,#0f172a,#1e3a5f,#1d4ed8);
                border-radius:16px; padding:40px; color:white; margin-bottom:28px;">
        <div style="font-size:32px;font-weight:700;margin-bottom:8px;">💼 AR Suite</div>
        <div style="font-size:15px;opacity:.75;max-width:520px;line-height:1.6;">
            A unified toolkit for your Accounts Receivable team.
            Match client remittances, split multi-account exports,
            generate customer statements, and draft emails — all in one place.
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="tool-card">
            <div class="icon">🔍</div>
            <h3>Remittance Reconciliation</h3>
            <p>Upload a SAP export and a client remittance. Automatically matches
               invoices, flags doubles, and shows what the customer still owes.</p>
            <div class="feat">SAP is the source of truth — ignores client sign conventions</div>
            <div class="feat">Matches by Assignment, Document Number, or substring</div>
            <div class="feat">Flags already-cleared items (potential doubles)</div>
            <div class="feat">Generates a customer statement (Section A/B/C format)</div>
            <div class="feat">Draft follow-up email in English, Dutch, or French</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Remittance Tool →", use_container_width=True, key="btn_rem"):
            st.session_state["active_page"] = "🔍  Remittance Reconciliation"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="tool-card">
            <div class="icon">📂</div>
            <h3>Account Splitter</h3>
            <p>Upload a SAP extract containing multiple customer accounts.
               Splits into separate sheets, removes invoices not yet due,
               and applies custom templates per customer.</p>
            <div class="feat">One sheet per customer in a single workbook</div>
            <div class="feat">Auto-removes invoices not yet due</div>
            <div class="feat">Custom Excel templates per customer account</div>
            <div class="feat">Summary tab with totals per account</div>
            <div class="feat">Draft payment reminder email per account (EN/NL/FR)</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Account Splitter →", use_container_width=True, key="btn_spl"):
            st.session_state["active_page"] = "📂  Account Splitter"
            st.rerun()

    st.write("")
    col3, col4 = st.columns(2)
    with col4:
        st.markdown("""
        <div class="tool-card">
            <div class="icon">🎁</div>
            <h3>Bonus & Payout</h3>
            <p>Match your SAP customers against a bonus partner file,
               and check that all X payouts are clean with no blocks.</p>
            <div class="feat">Highlights matching, missing, and extra accounts</div>
            <div class="feat">Adds SAP-only accounts into the bonus file</div>
            <div class="feat">Flags X payouts with B or U payment blocks</div>
            <div class="feat">Shows open invoices on or before a chosen date</div>
            <div class="feat">Alerts on any B-blocked items in the export</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Bonus & Payout →", use_container_width=True, key="btn_bonus"):
            st.session_state["active_page"] = "🎁  Bonus & Payout"
            st.rerun()

    st.write("")
    col5, col6 = st.columns(2)
    with col5:
        st.markdown("""
        <div class="tool-card">
            <div class="icon">❓</div>
            <h3>Help & FAQ</h3>
            <p>Not sure where to start or how something works? The FAQ covers
               all three tools in plain language — what they do, when to use
               them, and how to get the best results.</p>
            <div class="feat">Step-by-step guide for each tool</div>
            <div class="feat">Explains templates, account groups, and chunking</div>
            <div class="feat">Grouping logic and year assignment explained</div>
            <div class="feat">Troubleshooting common issues</div>
            <div class="feat">General questions about storage and SAP exports</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Help & FAQ →", use_container_width=True, key="btn_faq"):
            st.session_state["active_page"] = "❓  Help & FAQ"
            st.rerun()

    with col6:
        st.write("")

    with col3:
        st.markdown("""
        <div class="tool-card">
            <div class="icon">📊</div>
            <h3>Customer Overview</h3>
            <p>Generate a year-by-year breakdown for a customer.
               Each year shows all transactions, split by G/L account,
               with carry-over payments linked to their original invoices.</p>
            <div class="feat">All transactions grouped by year</div>
            <div class="feat">G/L split per year (Beer vs Rent)</div>
            <div class="feat">Invoices paid next year shown together with payment</div>
            <div class="feat">Document types translated (Invoice, Credit note, Payment…)</div>
            <div class="feat">English, Dutch and French output</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Customer Overview →", use_container_width=True, key="btn_ov"):
            st.session_state["active_page"] = "📊  Customer Overview"
            st.rerun()
