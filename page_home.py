import streamlit as st


def show():
    st.markdown("""
    <style>
    .abi-hero {
        background: #0A0A0A;
        border-radius: 16px;
        padding: 44px 44px 44px 44px;
        color: white;
        margin-bottom: 28px;
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(247,149,29,0.2);
    }
    .abi-hero::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
        background: linear-gradient(90deg, #FFC72C, #E09A00, #FFC72C);
    }
    /* Abstract geometric pattern */
    .abi-hero::after {
        content: "";
        position: absolute;
        right: -30px; top: -30px;
        width: 280px; height: 280px;
        border-radius: 50%;
        background: radial-gradient(circle, rgba(247,149,29,0.08) 0%, transparent 70%);
        pointer-events: none;
    }
    .abi-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(247,149,29,0.15);
        color: #FFC72C;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        padding: 5px 14px;
        border-radius: 20px;
        border: 1px solid rgba(247,149,29,0.3);
        margin-bottom: 16px;
    }
    .abi-hero-title {
        font-size: 32px;
        font-weight: 800;
        color: #FFFFFF;
        margin: 0 0 10px;
        line-height: 1.15;
        letter-spacing: -0.03em;
    }
    .abi-hero-title span { color: #FFC72C; }
    .abi-hero-sub {
        font-size: 15px;
        color: rgba(255,255,255,0.55);
        max-width: 500px;
        line-height: 1.7;
        margin: 0;
    }
    /* Tool cards */
    .tool-card {
        background: white;
        border: 1px solid #E8E3DC;
        border-radius: 14px;
        padding: 28px 26px 22px;
        height: 100%;
        transition: all .2s ease;
        position: relative;
        overflow: hidden;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }
    .tool-card::after {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
        background: linear-gradient(90deg, #FFC72C, #E09A00);
        opacity: 0;
        transition: opacity .2s;
    }
    .tool-card:hover {
        box-shadow: 0 8px 28px rgba(0,0,0,0.1);
        transform: translateY(-2px);
        border-color: #D4CFC8;
    }
    .tool-card:hover::after { opacity: 1; }
    /* Icon box */
    .tool-icon-wrap {
        width: 52px; height: 52px;
        background: #0A0A0A;
        border-radius: 12px;
        display: flex; align-items: center; justify-content: center;
        margin-bottom: 16px;
        box-shadow: 0 3px 10px rgba(0,0,0,0.15);
    }
    .tool-icon-wrap svg { width: 26px; height: 26px; }
    .tool-name {
        font-size: 16px;
        font-weight: 700;
        color: #0A0A0A;
        margin: 0 0 8px;
        letter-spacing: -0.01em;
    }
    .tool-desc {
        font-size: 13px;
        color: #7A7065;
        line-height: 1.65;
        margin: 0 0 16px;
    }
    .feat {
        font-size: 12px;
        color: #3A3530;
        padding: 3px 0;
        display: flex;
        align-items: flex-start;
        gap: 7px;
    }
    .feat-dot {
        width: 5px; height: 5px;
        background: #FFC72C;
        border-radius: 50%;
        flex-shrink: 0;
        margin-top: 5px;
    }
    </style>

    <!-- Hero -->
    <div class="abi-hero">
        <div class="abi-badge">⚡ AB InBev · Belgium · AR Team</div>
        <div class="abi-hero-title">Accounts <span>Receivable</span> Suite</div>
        <p class="abi-hero-sub">
            A unified internal toolkit for the Belgian AR team.
            Match remittances, split customer exports, generate overviews,
            and manage bonus payouts — in EN, NL or FR.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # ── Tool cards ─────────────────────────────────────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-icon-wrap">
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M9 5H7C5.89543 5 5 5.89543 5 7V19C5 20.1046 5.89543 21 7 21H17C18.1046 21 19 20.1046 19 19V7C19 5.89543 18.1046 5 17 5H15" stroke="#FFC72C" stroke-width="2" stroke-linecap="round"/>
                    <path d="M9 5C9 3.89543 9.89543 3 11 3H13C14.1046 3 15 3.89543 15 5C15 6.10457 14.1046 7 13 7H11C9.89543 7 9 6.10457 9 5Z" stroke="#FFC72C" stroke-width="2"/>
                    <path d="M9 12L11 14L15 10" stroke="#FFC72C" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                </svg>
            </div>
            <div class="tool-name">Remittance Reconciliation</div>
            <p class="tool-desc">Upload a SAP export and a client remittance. Automatically matches
               invoices, flags doubles, and shows what the customer still owes.</p>
            <div class="feat"><div class="feat-dot"></div>SAP is the source of truth — ignores client sign conventions</div>
            <div class="feat"><div class="feat-dot"></div>Matches by Assignment, Document Number, or substring</div>
            <div class="feat"><div class="feat-dot"></div>Flags already-cleared items (potential doubles)</div>
            <div class="feat"><div class="feat-dot"></div>Draft follow-up email in EN / NL / FR</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Remittance Tool →", use_container_width=True, key="btn_rem", type="primary"):
            st.session_state["active_page"] = "Remittance Reconciliation"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-icon-wrap">
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M3 6C3 4.89543 3.89543 4 5 4H9.58579C9.851 4 10.1054 4.10536 10.2929 4.29289L11 5H19C20.1046 5 21 5.89543 21 7V18C21 19.1046 20.1046 20 19 20H5C3.89543 20 3 19.1046 3 18V6Z" stroke="#FFC72C" stroke-width="2" stroke-linejoin="round"/>
                    <path d="M9 13H15M12 10V16" stroke="#FFC72C" stroke-width="2" stroke-linecap="round"/>
                </svg>
            </div>
            <div class="tool-name">Account Splitter</div>
            <p class="tool-desc">Upload a multi-account SAP extract. Splits into one sheet per customer,
               removes invoices not yet due, and applies custom templates per customer.</p>
            <div class="feat"><div class="feat-dot"></div>One sheet per customer in a single workbook</div>
            <div class="feat"><div class="feat-dot"></div>NEGO-style POC grouping &amp; chunked payment batches</div>
            <div class="feat"><div class="feat-dot"></div>EN / NL / FR document type translation</div>
            <div class="feat"><div class="feat-dot"></div>Payment reminder emails per account</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Account Splitter →", use_container_width=True, key="btn_spl", type="primary"):
            st.session_state["active_page"] = "Account Splitter"
            st.rerun()

    st.write("")
    col3, col4 = st.columns(2)

    with col3:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-icon-wrap">
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <rect x="3" y="3" width="18" height="18" rx="2" stroke="#FFC72C" stroke-width="2"/>
                    <path d="M3 9H21" stroke="#FFC72C" stroke-width="2"/>
                    <path d="M9 9V21" stroke="#FFC72C" stroke-width="2"/>
                    <path d="M13 13H17M13 17H15" stroke="#FFC72C" stroke-width="2" stroke-linecap="round"/>
                </svg>
            </div>
            <div class="tool-name">Customer Overview</div>
            <p class="tool-desc">Year-by-year or current-period breakdown for any customer.
               Groups by clearing document, shows reconciliation status per year,
               and exports a branded Excel overview.</p>
            <div class="feat"><div class="feat-dot"></div>Multi-year history or current open items only</div>
            <div class="feat"><div class="feat-dot"></div>Year totals showing full reconciliation balance</div>
            <div class="feat"><div class="feat-dot"></div>EN / NL / FR with translated descriptions</div>
            <div class="feat"><div class="feat-dot"></div>Works with single-account or multi-year SAP exports</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Customer Overview →", use_container_width=True, key="btn_ov", type="primary"):
            st.session_state["active_page"] = "Customer Overview"
            st.rerun()

    with col4:
        st.markdown("""
        <div class="tool-card">
            <div class="tool-icon-wrap">
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M12 2L15.09 8.26L22 9.27L17 14.14L18.18 21.02L12 17.77L5.82 21.02L7 14.14L2 9.27L8.91 8.26L12 2Z" stroke="#FFC72C" stroke-width="2" stroke-linejoin="round"/>
                    <path d="M12 7V12L14.5 14.5" stroke="#FFC72C" stroke-width="1.5" stroke-linecap="round"/>
                </svg>
            </div>
            <div class="tool-name">Bonus & Payout Tools</div>
            <p class="tool-desc">Two tools in one: compare SAP accounts against a bonus file to find
               mismatches, or scan for X-payout entries and flag open invoices
               for the payout &amp; block checker.</p>
            <div class="feat"><div class="feat-dot"></div>SAP vs bonus file — highlights missing accounts</div>
            <div class="feat"><div class="feat-dot"></div>X-payout scanner with B/U block flag detection</div>
            <div class="feat"><div class="feat-dot"></div>Exports colour-coded Excel report</div>
            <div class="feat"><div class="feat-dot"></div>Date-based open invoice detection</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Open Bonus Tools →", use_container_width=True, key="btn_bon", type="primary"):
            st.session_state["active_page"] = "Bonus & Payout"
            st.rerun()

    # ── Help strip ─────────────────────────────────────────────────────────
    st.write("")
    st.markdown("""
    <div style="background:#0A0A0A; border-radius:12px; padding:16px 24px;
                display:flex; align-items:center; justify-content:space-between;
                gap:20px; margin-top:4px;">
        <div style="display:flex; align-items:center; gap:10px;">
            <div style="background:rgba(247,149,29,0.15); border-radius:8px;
                        width:32px; height:32px; display:flex; align-items:center;
                        justify-content:center; font-size:15px; flex-shrink:0;">❓</div>
            <div style="font-size:13px; color:rgba(255,255,255,0.7);">
                New here? The <strong style="color:#FFC72C;">Help &amp; FAQ</strong>
                page has step-by-step guides for every tool.
            </div>
        </div>
        <div style="font-size:11px; color:#3A3530; white-space:nowrap; flex-shrink:0;">
            AB InBev · Belgium AR
        </div>
    </div>
    """, unsafe_allow_html=True)
