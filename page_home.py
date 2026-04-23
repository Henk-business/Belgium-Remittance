import streamlit as st


def show():
    # ── Hero banner ────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    .abi-hero {
        background: linear-gradient(135deg, #1A0A00 0%, #3D1408 50%, #1A0A00 100%);
        border-radius: 16px;
        padding: 44px 40px;
        color: white;
        margin-bottom: 28px;
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(200,146,42,0.3);
    }
    .abi-hero::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
        background: linear-gradient(90deg, #C41230, #C8922A, #C41230);
    }
    .abi-hero::after {
        content: "🍺";
        position: absolute;
        right: 40px;
        top: 50%;
        transform: translateY(-50%);
        font-size: 80px;
        opacity: 0.12;
    }
    .abi-badge {
        display: inline-block;
        background: rgba(200,146,42,0.2);
        color: #C8922A;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        padding: 4px 12px;
        border-radius: 20px;
        border: 1px solid rgba(200,146,42,0.4);
        margin-bottom: 14px;
    }
    .abi-hero h1 {
        font-size: 30px !important;
        font-weight: 800 !important;
        color: white !important;
        margin: 0 0 10px !important;
        line-height: 1.2;
    }
    .abi-hero p {
        font-size: 15px;
        color: rgba(255,255,255,0.7);
        max-width: 520px;
        line-height: 1.7;
        margin: 0;
    }
    .tool-card {
        background: white;
        border: 1px solid #E8DDD0;
        border-radius: 14px;
        padding: 26px;
        box-shadow: 0 2px 8px rgba(26,10,0,0.06);
        height: 100%;
        transition: all .2s ease;
        position: relative;
        overflow: hidden;
    }
    .tool-card::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
        background: linear-gradient(90deg, #C41230, #C8922A);
        opacity: 0;
        transition: opacity .2s;
    }
    .tool-card:hover { box-shadow: 0 8px 24px rgba(196,18,48,0.12); transform: translateY(-2px); }
    .tool-card:hover::before { opacity: 1; }
    .tool-icon {
        width: 48px; height: 48px;
        background: linear-gradient(135deg, rgba(196,18,48,0.1), rgba(200,146,42,0.1));
        border-radius: 12px;
        display: flex; align-items: center; justify-content: center;
        font-size: 24px;
        margin-bottom: 14px;
    }
    .tool-card h3 {
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #1A0A00 !important;
        margin: 0 0 8px !important;
        text-transform: none !important;
        letter-spacing: 0 !important;
    }
    .tool-card p {
        font-size: 13px;
        color: #6B5744;
        line-height: 1.65;
        margin: 0 0 14px;
    }
    .feat {
        font-size: 12px;
        color: #4A3828;
        padding: 3px 0;
        display: flex;
        align-items: flex-start;
        gap: 6px;
    }
    .feat::before {
        content: "→";
        color: #C41230;
        font-weight: 700;
        flex-shrink: 0;
    }
    .stat-strip {
        background: linear-gradient(135deg, #C41230, #9A0E25);
        border-radius: 12px;
        padding: 20px 28px;
        color: white;
        display: flex;
        gap: 40px;
        margin-bottom: 28px;
        align-items: center;
    }
    .stat-item { text-align: center; }
    .stat-num { font-size: 24px; font-weight: 800; color: #FFE082; }
    .stat-lbl { font-size: 11px; opacity: 0.8; text-transform: uppercase; letter-spacing: 0.06em; }
    </style>

    <div class="abi-hero">
        <div class="abi-badge">AB InBev · Belgium AR Team</div>
        <h1>Accounts Receivable Suite</h1>
        <p>A unified toolkit built for the Belgian AR team. Match remittances,
           split customer exports, generate multi-year overviews, and manage bonus
           payouts — all in one place, all in your language.</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Tool cards ─────────────────────────────────────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-icon">🔍</div>
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
        if st.button("Open Remittance Tool →", use_container_width=True,
                     key="btn_rem", type="primary"):
            st.session_state["active_page"] = "🔍  Remittance Reconciliation"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-icon">📂</div>
            <h3>Account Splitter</h3>
            <p>Upload a SAP extract containing multiple customer accounts.
               Splits into separate sheets, removes invoices not yet due,
               and applies custom templates per customer.</p>
            <div class="feat">One sheet per customer in a single workbook</div>
            <div class="feat">Custom templates with NEGO-style POC grouping</div>
            <div class="feat">Chunked payment batches for large accounts</div>
            <div class="feat">EN / NL / FR document type translation</div>
            <div class="feat">Payment reminder emails per account</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Account Splitter →", use_container_width=True,
                     key="btn_spl", type="primary"):
            st.session_state["active_page"] = "📂  Account Splitter"
            st.rerun()

    st.write("")
    col3, col4 = st.columns(2)

    with col3:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-icon">📊</div>
            <h3>Customer Overview</h3>
            <p>Generate a year-by-year or current-period breakdown for any customer.
               Groups transactions by clearing document, shows reconciliation status,
               and exports a branded Excel overview.</p>
            <div class="feat">Multi-year history grouped by clearing document</div>
            <div class="feat">Current open items with overdue detection</div>
            <div class="feat">Year totals showing reconciliation balance</div>
            <div class="feat">EN / NL / FR with translated descriptions</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Customer Overview →", use_container_width=True,
                     key="btn_ov", type="primary"):
            st.session_state["active_page"] = "📊  Customer Overview"
            st.rerun()

    with col4:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-icon">🎁</div>
            <h3>Bonus & Payout Tools</h3>
            <p>Two tools in one: compare SAP accounts against a bonus file to find
               mismatches and missing accounts, or scan for X-payout entries and
               flag open invoices for the payout & block checker.</p>
            <div class="feat">Customer matching — SAP vs bonus file</div>
            <div class="feat">Highlights missing accounts and discrepancies</div>
            <div class="feat">X-payout scanner with block flag detection</div>
            <div class="feat">Exports colour-coded Excel report</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Bonus Tools →", use_container_width=True,
                     key="btn_bon", type="primary"):
            st.session_state["active_page"] = "🎁  Bonus & Payout"
            st.rerun()

    # ── Footer strip ────────────────────────────────────────────────────────
    st.write("")
    st.markdown("""
    <div style="background:#F5F0EB; border-radius:12px; padding:16px 24px;
                border:1px solid #E8DDD0; display:flex; align-items:center;
                justify-content:space-between; gap:20px; margin-top:8px;">
        <div style="font-size:12px; color:#7A6555;">
            💡 Need help? Check the
            <strong style="color:#C41230;">Help & FAQ</strong> page for step-by-step guides
            on each tool.
        </div>
        <div style="font-size:11px; color:#A08060; white-space:nowrap;">
            AB InBev · Belgium · AR Suite
        </div>
    </div>
    """, unsafe_allow_html=True)
