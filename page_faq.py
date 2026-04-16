import streamlit as st


def show():
    st.markdown("## ❓ Help & FAQ")
    st.markdown("Everything you need to know about the AR Suite tools.")

    with st.expander("🔍  Remittance Reconciliation — what does it do?", expanded=False):
        st.markdown("""
**What it does**

Two tools in one page for handling customer payments.

---

**Tab 1 — Remittance matching**

Matches a customer's payment advice against your SAP open items and flags discrepancies.

*How to use it:*
1. Export open items from SAP (FBL5N) for that customer.
2. Upload the SAP export and the customer's remittance file.
3. Enter customer name, payment amount, and payment date.
4. Click **Run Reconciliation**.

*Output:* Colour-coded Excel (matched / unmatched / partial), plus a draft email in EN/NL/FR.

---

**Tab 2 — Amount-only matching**

Customer paid without sending a remittance? Enter the payment amount and the system finds
which combinations of open invoices add up to it.

*How to use it:*
1. Export open items from SAP (FBL5N).
2. Upload the export and enter the payment amount.
3. Set a tolerance (default €0.05).
4. Click **Find matching invoices**.

*How matching works:* Tries exact single invoice → two-invoice pairs → greedy subset-sum
across all open invoices. Returns up to 5 options ranked by closeness, with 🟢 exact or
🟡 within-tolerance confidence. Downloadable as Excel with one sheet per option.
        """)

    with st.expander("📂  Account Splitter — what does it do?", expanded=False):
        st.markdown("""
**What it does**

Splits a SAP FBL5N export into one sheet per customer, removes SAP-internal columns,
optionally removes invoices not yet due, and applies custom layouts where configured.

**How to use it**

1. Export open items from SAP (can include multiple accounts).
2. Upload, set reference date, tick/untick "Remove invoices not yet due".
3. Click **Split**.
4. Use **Download all accounts** for a quick combined file, or use individual buttons below
   for custom layouts.

**Document language selector**

Translates the Document Type column — RV → Invoice/Facture/Factuur, RV− → Credit note,
ZP → Payment, RS− → Bonus, AB → Clearing, etc. Set independently of the email language.

**Colour convention:** Positive (invoices) = red. Negative (credits/payments) = green.

---

**Custom templates**

Upload a customer's Excel as a template — the system reads column order, widths, and header
style and reproduces it automatically.
*Setup:* Customer templates → Upload a new customer template.

**Account groups**

Combine two or more accounts into one download.
- **Multi-sheet:** one sheet per account (needs template on primary account).
- **Flat merge:** all rows in one sheet, sorted by net due date. No template needed.
*Setup:* Customer templates → Account groups → enter label + account numbers → Save.

**Special layouts**

- **POC grouped** (NEGOBOISSONS): rows grouped by 29xxxxx Reference Key 3 with subtotals.
  Triggered automatically when template has POC structure.
- **Chunked** (30111788): invoices split into €40k batches.
  *Setup:* Customer templates → Chunking rules.
        """)

    with st.expander("📊  Customer Overview — what does it do?", expanded=False):
        st.markdown("""
**What it does**

Generates a formatted account overview in two modes:

- **📋 Current overview** — open items as of a reference date, not-yet-due invoices removed.
- **📅 Multi-year overview** — transactions grouped by year for multi-year history.

**Settings per mode**

| Control | Current | Multi-year |
|---|---|---|
| Reference date | ✅ Active | 🔘 Greyed |
| Remove not yet due | ✅ Active | 🔘 Greyed |
| From / To year | 🔘 Greyed | ✅ Active |

**How to use it**

1. Export full transaction history from SAP (FBL5N) including all required years.
2. Upload, select mode, set dates/years, choose language.
3. Click **Generate Overview** and download the Excel.

**Grouping logic**

Transactions grouped by SAP Clearing Document. Year assigned by the oldest net due date
of invoice rows — AB/ZP/DZ clearing rows are ignored so old clearings don't drag current
invoices into past years.

**Columns included**

Account, Assignment, Document Number, Reference Key 3, Document Date, Net due date,
Description, Amount, Payment Method, G/L Account, Arrears. Clearing date and Clearing
Document are excluded from all outputs.

**G/L split:** Beer (2400000) and Rent (2530009) shown in separate sub-sections per year.

**Document descriptions:** RV+ = Invoice, RV− = Credit note, ZP/DZ = Payment,
RS− = Bonus, RS+ = Re-invoice, AB = Clearing / Ajustement comptable (FR),
RV− = Note de crédit (FR).

**Colour:** Positive = red, Negative = green. Total rows = white on dark blue.

**Email draft**

Appears below the download button. Set name, company, customer email, language.
Click "Open in email client" to open a pre-filled draft — attach the Excel manually.
        """)

    with st.expander("🗂  Customer templates & groups — how do they work?", expanded=False):
        st.markdown("""
**Templates**

Upload a customer's preferred Excel to reproduce their layout automatically on every split.
The system auto-detects column order, widths, header style, and POC structure.

*Upload:* Customer templates → Upload a new customer template → enter account number → upload.

*POC detection:* Template with 29xxxxx in column A → automatically POC-grouped layout.

*Plain table detection:* SAP-style column headers detected regardless of merged title rows above.
Title rows (account number, date, line count) are updated automatically with fresh values.

---

**Chunking rules**

Split invoices into payment batches: Customer templates → Chunking rules → enter account +
batch size → Save. Stored in GitHub, applied automatically on every split.

---

**Account groups**

*Setup:* Customer templates → Account groups → label + account numbers (comma-separated)
→ tick Flat merge if needed → Save group.

- **Multi-sheet:** separate sheet per account. Template required for primary account.
- **Flat merge:** all rows combined in one sheet. No template needed. For accounts like
  30351345 + 30104410 that should appear as one combined list.
        """)

    with st.expander("⚙️  General questions", expanded=False):
        st.markdown("""
**Are my files stored anywhere?**

Uploaded SAP files are processed in memory only and never stored. Templates, rules, and
account groups are saved to your private GitHub repository.

**Why does the app sometimes show a connection error?**

Occasional GitHub API timeouts. Handled gracefully — the page still works, just without
cached templates at that moment. Refreshing usually resolves it.

**Why does the app ask me to sign in?**

Set the app to Public in Streamlit Cloud dashboard → Share → Public. Anyone with the URL
can then open it without an account.

**What SAP export format do I need?**

Standard FBL5N .xlsx export. Include all columns — the app strips internal ones
automatically. For Customer Overview, export full transaction history (not just open items).

**Red = invoices, Green = credits — is that right?**

Yes. Positive amounts (invoices, money owed to you) = red. Negative amounts (credit notes,
payments received) = green. This is the Belgian AR convention.

**Something looks wrong**

Common causes: wrong reference date · SAP export pre-filtered · template on wrong account
number · account group primary account has no template uploaded.
        """)
