import streamlit as st


def show():
    st.markdown("## ❓ Help & FAQ")
    st.markdown("Everything you need to know about the AR Suite tools.")

    # ── TOOL 1 ────────────────────────────────────────────────────────────────
    with st.expander("🔍  Remittance Reconciliation — what does it do?", expanded=False):
        st.markdown("""
**What it does**

When a customer sends a payment with a remittance advice (a list of which invoices they're paying),
this tool matches their list against your SAP open items and flags any discrepancies.

**When to use it**

Use it when a customer sends you a payment file or remittance email and you want to quickly check
which invoices are covered, which are missing, and what the outstanding balance is after the payment.

**How to use it**

1. Export the open items for that customer from SAP (FBL5N).
2. Upload the SAP export and the customer's remittance file.
3. The tool matches by document number / assignment and produces a reconciliation report.

**Output**

- A colour-coded Excel showing matched, unmatched, and partially-matched items.
- A draft email you can send to the customer summarising the reconciliation.
        """)

    # ── TOOL 2 ────────────────────────────────────────────────────────────────
    with st.expander("📂  Account Splitter — what does it do?", expanded=False):
        st.markdown("""
**What it does**

Takes a SAP FBL5N export containing multiple customer accounts and splits it into
one sheet per customer in a single workbook. Removes internal SAP columns, optionally
removes invoices not yet due, and applies custom layouts where configured.

**When to use it**

Use it when you need to send individual account statements to multiple customers at once.
Export all accounts in one go from SAP, upload once, and download a clean file per customer.

**How to use it**

1. Export open items from SAP (FBL5N) — can include multiple accounts.
2. Set the reference date (invoices due after this date are removed if the option is ticked).
3. Click **Split**. Download either the combined file or individual account files.

**Custom templates**

You can upload a customer's own Excel layout as a template. When that account appears in a
split, the output will match their preferred format automatically.

**Account groups**

Two accounts that belong to the same customer (e.g. a North and South entity) can be grouped
so they download as one combined file. Two modes are available:
- **Separate sheets** — one sheet per account + a summary (like North & South Beverages).
- **Flat merge** — all rows combined into one single sheet.

**Special layouts**

- **POC grouped** (e.g. NEGOBOISSONS): rows grouped by Reference Key 3 (29xxxxx) with subtotals.
- **Chunked** (e.g. 30111788): invoices split into €40k batches for payment processing.

**Colour convention**

Positive amounts (invoices) = **red**. Negative amounts (credits/payments) = **green**.
        """)

    # ── TOOL 3 ────────────────────────────────────────────────────────────────
    with st.expander("📊  Customer Overview — what does it do?", expanded=False):
        st.markdown("""
**What it does**

Generates a formatted overview of a customer's account history. Two modes:

- **Current overview** — all open items as of a chosen reference date, with invoices not yet
  due removed. Ideal for sending a regular monthly statement.
- **Multi-year overview** — transactions grouped by year, showing what happened in each
  calendar year. Useful when a customer requests a history spanning several years.

**When to use it**

- A customer asks "what do we currently owe you?" → use **Current overview**.
- A customer asks "can you send us everything from 2020 to today?" → use **Multi-year overview**.

**How to use it**

1. Export the full transaction history from SAP (FBL5N) for the customer — include all
   years you need.
2. Upload the file, select the mode, set dates if needed, choose language.
3. Click **Generate Overview** and download the Excel.

**Grouping logic**

Transactions are grouped by SAP Clearing Document. Each group (invoice + credit notes +
payment) appears together and nets to zero when fully cleared. The year a group belongs to
is determined by the oldest net due date of the invoice rows in that group — AB/clearing
rows are ignored so one old clearing entry doesn't drag a group of current invoices into a
past year.

**G/L split**

If a customer has both Beer (2400000) and Rent (2530009) transactions, they appear in
separate sub-sections within each year with their own subtotals.

**Languages**

The Excel output and the email draft can each be set independently to English, Dutch, or French.

**Document type descriptions**

The Document Type column is replaced with a human-readable description:
RV+ = Invoice, RV− = Credit note, ZP/DZ = Payment, RS− = Bonus,
RS+ = Re-invoice, AB = Clearing, Payment Method X = Payout to customer.
        """)

    # ── TEMPLATES ─────────────────────────────────────────────────────────────
    with st.expander("🗂  Customer templates — how do they work?", expanded=False):
        st.markdown("""
**What they are**

Templates let you save a customer's preferred Excel layout so the splitter
reproduces it automatically every time that account appears in a split.

**How to upload**

In the Account Splitter → **Customer templates** section → expand **Upload a new customer template**,
enter the account number, and upload the customer's Excel file.

**What gets detected automatically**

- Column order and which columns to show
- Row heights and column widths
- Header style (plain table vs branded layout with merged title rows)
- POC grouping structure (NEGOBOISSONS-style)

**Chunking rules**

For accounts that need invoices split into payment batches (e.g. every €40,000),
use the **Chunking rules** section to set the batch size. The rule is stored permanently
in GitHub and applied automatically.

**Account groups**

Use the **Account groups** section to combine two or more accounts into one download.
Set the label, enter the account numbers comma-separated, and choose flat or multi-sheet.
        """)

    # ── GENERAL ───────────────────────────────────────────────────────────────
    with st.expander("⚙️  General questions", expanded=False):
        st.markdown("""
**Are my files stored anywhere?**

Uploaded SAP files are processed in memory and never written to disk or stored.
Templates and rules are saved to your private GitHub repository (configured in Streamlit secrets).

**Why does the app ask me to sign in?**

The app is hosted on Streamlit Cloud. If you see a login prompt, check that the app is set to
Public in the Streamlit Cloud dashboard (Share → Public). You can share the URL directly with
colleagues — no account needed to view a public app.

**Can I automate the SAP export?**

Not directly from this app — SAP requires a manual FBL5N export. However, SAP can be configured
to email scheduled exports automatically, which could then be forwarded and uploaded here.

**What SAP export format do I need?**

Standard FBL5N open items export as .xlsx. Include all columns — the app strips the ones
it doesn't need automatically. For the Customer Overview, export the full history (not just
open items) to get all years.

**Something looks wrong in the output**

Common causes:
- Wrong reference date selected (invoices that should be removed are still showing).
- The SAP export was filtered before uploading (e.g. only one document type exported).
- A template uploaded for the wrong account number.

If the issue persists, use the thumbs down button on any response to send feedback.
        """)
