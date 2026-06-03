import streamlit as st


def show():
    from abi_ui import page_header, today_bar
    page_header("Help & FAQ",
                "Step-by-step guides and answers for every AR Suite tool.",
                "❓")
    today_bar()

    # ── Remittance ─────────────────────────────────────────────────────────
    with st.expander("🔍  Remittance Reconciliation", expanded=False):
        st.markdown("""
**What it does**

Three tabs for handling customer payments — remittance matching, amount-only matching,
and invoice/credit note matching.

---

**Tab 1 — Remittance matching**

Matches a customer's payment advice against your SAP open items line by line.

1. Export open items from SAP FBL5N for that customer → `.xlsx`
2. Upload the SAP export and the customer's remittance file (Excel **or PDF**)
3. Enter customer name, payment amount, and payment date
4. Click **Run Reconciliation**

Output: colour-coded Excel (matched / unmatched / already cleared / partial), a discrepancy
summary, and a draft follow-up email in EN / NL / FR that you can open directly in your mail client.

**Matching logic:** tries Assignment number first, then Document Number, then Reference Key 3,
then substring. SAP is the source of truth — the tool ignores sign conventions in the remittance.
Already-cleared items are flagged separately as potential doubles.

**PDF remittances:** the tool extracts tables from the PDF automatically (using pdfplumber),
falling back to raw text extraction if no tables are found.

---

**Tab 2 — Amount-only matching**

Customer paid without sending a remittance? Enter the amount and the tool finds which
combination of invoices and credit notes adds up to it.

1. Upload SAP export and enter the payment amount
2. Set a tolerance (default €0.05)
3. Click **Find matching invoices**

**Matching strategy — oldest invoices first:** the tool always tries to match starting
from the oldest due-date invoice, then progressively adds newer invoices until the
target amount is reached. Credits are included automatically when they help reduce
an overshoot. Returns up to 5 options ranked by closeness — 🟢 exact or 🟡 within tolerance.
Downloadable as Excel with one sheet per option.

---

**Tab 3 — Invoice / Credit matching**

Upload any SAP open-items export and the system automatically pairs each invoice with
the credit notes that best offset it.

- **Oldest invoices matched first** — the oldest due-date invoice is matched before newer ones
- **Exact matches** (net = €0.00) found first, then closest within **€100 difference**
- **Each credit note used once only** — a credit already used for one invoice match cannot
  appear in another match
- After all possible matches, the output shows unmatched invoices (no credits left within €100)
  and remaining unmatched credits

Output: 4-sheet Excel — Summary · Matches (invoice + credit detail) · Unmatched Invoices · Unmatched Credits.
        """)

    # ── Account Splitter ───────────────────────────────────────────────────
    with st.expander("📂  Account Splitter", expanded=False):
        st.markdown("""
**What it does**

Splits a SAP FBL5N export (single or multiple accounts) into one sheet per customer.
Strips internal SAP columns, translates document type codes, and applies custom layouts
where configured.

**Basic usage**

1. Export open items from SAP (can include multiple accounts)
2. Upload the file
3. Set the global **reference date** (e.g. last day of the month)
4. Tick **Remove invoices not yet due** if you want to exclude future-dated invoices
5. Click **Split**
6. Select document language (EN / NL / FR)
7. Download all accounts (standard layout) or individual custom downloads below

---

**Full column translation per language**

All column headers translate automatically based on the selected language:

| Column | EN | NL | FR |
|---|---|---|---|
| Account | Account | Rekening | Compte |
| Assignment | Assignment | Toewijzing | Affectation |
| Document Number | Document Number | Documentnummer | Numéro de document |
| Net due date | Net due date | Netto vervaldatum | Date d'échéance nette |
| Document Type | Description | Omschrijving | Description |
| Amount | Amount in local currency | Bedrag in lokale valuta | Montant en devise locale |
| Arrears | Arrears after net due date | Achterstand na vervaldatum | Arriérés après échéance |
| G/L Account | G/L Account | Grootboekrekening | Compte général |

Document type values also translate per language:
Invoice / Factuur / Facture · Credit note / Creditnota / Note de crédit · Payment / Betaling / Paiement · Clearing / Verrekening / Ajustement · Bonus / Bonus / Bonus.

---

**Auto-detected customer language**

The system automatically looks up each account's preferred language from the customer
master file (392 accounts). When you split:
- The summary table shows each account's name and a language flag (🇳🇱 🇫🇷 🇬🇧)
- Each account's sheet is generated in that customer's own language
- Accounts without a language record default to English

---

**Per-account reference dates**

When "Remove invoices not yet due" is ticked, an **⚙️ Per-account date overrides** expander
appears. Each account defaults to the global date but can be set individually.

---

**Results summary table**

After splitting, a table shows each account with Name, language flag, reference date,
row count, and total. Accounts with a date override show ⚙ next to their date.
Excluded accounts show greyed out with a ↩ restore button.

---

**Account exclusion**

Click ✕ next to any account in the summary table to exclude it from all downloads.
Click ↩ to restore. Useful when one account in a multi-account export should not
be sent out this cycle.

---

**Custom templates**

Upload a customer's preferred Excel format — the tool reproduces their column order,
widths, and header style automatically on every split.

*Setup:* Customer templates → Upload a new customer template → enter account number → upload.
Requires admin login to save or delete templates.

Special layouts detected automatically:
- **POC grouped** (NEGOBOISSONS): rows grouped by 29xxxxx Reference Key 3 with subtotals
- **Chunked** (30111788): invoices split into €40k payment batches

---

**Account groups**

Combine two or more accounts into one download.
- **Multi-sheet:** one sheet per account in a single workbook (needs template on primary account)
- **Flat merge:** all rows combined in one sheet, sorted by net due date

*Setup:* Customer templates → Account groups → label + comma-separated account numbers → Save group.

---

**Admin access**

Saving, replacing, or deleting templates and rules requires an admin password. Enter it in
the **🔐 Admin login** expander inside Customer templates. Once unlocked, the session stays
open until you close the browser tab.
        """)

    # ── Customer Overview ──────────────────────────────────────────────────
    with st.expander("📊  Customer Overview", expanded=False):
        st.markdown("""
**What it does**

Generates a branded Excel account overview in two modes for any customer.

---

**Mode 1 — Current Overview**

Shows only open (outstanding) items as of a chosen reference date.

- **Auto-detects structure:** if the export has 3+ blank separator rows in the middle
  it is treated as a clearing-group history export and only the current open section
  is shown. 1–2 mid-file blank rows = GL subtotals only (Beer/Rent split), treated as
  flat open-items. 0 blank rows = all rows are current open.
- **Remove not yet due:** hides invoices with a net due date after the reference date
- **Remove current overdues:** hides already-overdue items
- Works with single-account exports *and* full multi-year exports — extracts just the
  current open section automatically
- **Zero-net clearing groups removed** from current open: balanced AB+ZP pairs that
  net to €0 are filtered out (they are already settled and do not belong in the open section).
  Standalone items with no Reference Key 3 (e.g. standalone AB adjustments) are always kept.

*Output format:* title row → subtitle row → column headers → data rows → Net Balance total.

---

**Mode 2 — Multi-year Overview**

Shows the full clearing history grouped by year, newest year first.

- **Current Open Items** section at the top
- Each year shows: group count · invoice total · credits total · net balance
- **Year bucketing — majority invoice year:** a group's year is determined by the
  majority of its invoice net-due-date years (e.g. 8 invoices due 2026, 1 credit due 2025
  → bucketed into 2026). Ties go to the most recent year. Falls back to credit dates
  if no invoices; then to clearing row dates as a last resort.
- Yellow zero-subtotal rows show payment grouping boundaries
- Year totals in mid-blue; grand Net Balance in dark navy at the bottom
- Arrears column shows the **original SAP export values** (not recalculated)

**AB clearing annotations:** when an open AB (clearing adjustment) is visible in the
Current Open Items section, the tool automatically adds a soft blue annotation row
beneath it explaining what the AB originated from:

> ↳ Clearing origin: Invoice 13080 15/02/2026 €51,300.69  |  Payment 2000045123 28/02/2026 €-51,100.00

This tells the customer exactly what over/under-payment caused the adjustment, so they
are not confused by an AB entry standing alone on the account.

---

**Auto-detected customer name & language**

When a single account is detected in the export, the customer's name and preferred
language are auto-filled from the customer master (392 accounts). Language can be
overridden manually. For multi-account exports, selecting a specific account from the
dropdown shows a hint if that account's language differs from the current selection.

---

**G/L split (both modes)**

When a customer has both Beer (G/L 2400000) and Rent (2530009) invoices, the overview
automatically groups them into separate sections with individual subtotals and a combined
Net Balance.

| Language | Beer | Rent |
|---|---|---|
| EN | Beer | Rent |
| NL | Bier | Huur |
| FR | Bière | Loyer |

---

**Columns included**

Account · Assignment · Document Number · Reference Key 3 · Document Date · Net due date ·
Description · Amount · Payment Method · G/L Account · Arrears after net due date.

Excluded: Case ID · Status · Dunning Block · Clearing date · Clearing Document · and all
other internal SAP columns.

---

**Document descriptions** (language-aware)

RV+ = Invoice · RV− = Credit note · ZP/DZ = Payment · RS− = Bonus · RS+ = Re-invoice ·
AB = Clearing

**Colours:** Positive amounts (invoices) = red · Negative (credits/payments) = green ·
Total rows = white on dark navy.
        """)

    # ── AR Calendar ────────────────────────────────────────────────────────
    with st.expander("📅  AR Calendar", expanded=False):
        st.markdown("""
**What it does**

A monthly task scheduler showing direct debits, overviews, UAC runs, and meetings by
day of month. Repeats every month. Multiple calendars can be created for different scopes
(e.g. Wholesale, Retail).

---

**Weekend shift**

If a scheduled task falls on a Saturday it automatically moves to **Friday** (day −1).
If it falls on a Sunday it moves to **Monday** (day +1). The chip on the calendar shows a
small ⇒ indicator with the original day number (e.g. `DD: SABIKO ⇒28`) so it is always
clear that the task was shifted. The edit table always shows the original scheduled days
so you can see and edit the source schedule without the display shift confusing things.

---

**Today banner**

When tasks are scheduled for today, a black banner appears at the top of the page showing
each task with a colour-coded pill. Also visible in the **sidebar** on every page.

Sidebar widget also shows:
- Month progress bar
- **Next up** — the next day this month with tasks and how many days away it is
- **Open Calendar →** button

---

**Calendar selector**

Dropdown at the top of the calendar page. Switch between multiple calendars instantly.

**➕ New calendar** — click the button, type a name, hit Create.
**🗑 Delete calendar** — only available when 2+ calendars exist.
**Rename** — flip on edit mode, use the rename expander at the top.

---

**Viewing the calendar**

A 7-column month grid shows all scheduled tasks as colour-coded chips inside each day cell.
Today is highlighted with a gold border. Weekends are greyed out.

**Colour coding:**
- 🟡 Yellow = Direct Debit (DD)
- ⚫ Black/gold = Overview
- 🔴 Red = UAC
- 🔵 Blue = Meeting

---

**Editing tasks**

Toggle **✏️ edit mode** on. The full task list appears as an interactive spreadsheet:

- **Click any cell** to edit inline
- **Add a row** — use the + button, fill in Day (1–31), Type, Account, optionally Format and Note
- **Delete a row** — tick the checkbox then click the bin icon
- **Move a task** — change the Day number
- Hit **💾 Save changes** to apply

| Type | Colour | Used for |
|---|---|---|
| DD | Gold | Direct debit run |
| Overview | Black/gold | Generate and send customer overview |
| UAC | Red | UAC PRIK & TIK processing |
| Meeting | Blue | WHS Dutch meetings, etc. |
        """)

    # ── Templates & groups ─────────────────────────────────────────────────
    with st.expander("🗂  Customer templates & groups", expanded=False):
        st.markdown("""
**Templates**

Upload a customer's preferred Excel layout to reproduce it automatically on every split.
The system auto-detects column order, widths, header style, and POC structure.

*Upload:* Customer templates → Upload a new customer template → enter account number → upload.
Requires admin login.

- **POC detection:** template with 29xxxxx Reference Key in column A → automatically
  POC-grouped layout (NEGOBOISSONS style)
- **Plain table detection:** SAP-style column headers detected regardless of merged title rows above
- **Chunked layout:** enter account + batch size in Chunking rules → invoices split into batches

---

**Account groups**

*Setup:* Customer templates → Account groups → label + account numbers (comma-separated)
→ tick Flat merge if needed → Save group. Requires admin login.

- **Multi-sheet:** one sheet per account. Template required on primary account.
- **Flat merge:** all rows combined, sorted by net due date.

---

**Admin access**

A password set in Streamlit secrets protects all write/delete actions.
Enter the password once per session in the 🔐 Admin login expander.
        """)

    # ── Bonus & Payout ─────────────────────────────────────────────────────
    with st.expander("🎁  Bonus & Payout Tools", expanded=False):
        st.markdown("""
**Tab 1 — Customer matching**

Compare your SAP customer list against a bonus/partner file. Shows which accounts match,
which are in the bonus file but missing from SAP, and which SAP accounts are absent.

1. Upload SAP export and the bonus/partner file (reads ALL sheets automatically)
2. Confirm which column holds the account number in each
3. Click **Run matching**

Output: annotated bonus file (green = match, orange = not in SAP, yellow = added from SAP),
a Summary sheet, and a list of missing accounts.

---

**Tab 2 — Payout & block checker**

Scans a SAP export to verify payout eligibility.

1. Upload the SAP export
2. Set the invoice cutoff date (default: 21st of the month)
3. Click **Run check**

The output Excel contains **four sheets:**

- **X Payouts — OK (RS bonus lines only):** clean payout-eligible accounts showing
  **only the RS bonus credit rows** — no RV invoices, no ZP payments, no clearing docs.
  This is a pure list of what will be paid out: RS debits marked X, no blockers.
  Positive RS rows (debit re-invoices) are excluded.
- **X Payouts BLOCKED:** RS rows with B or U payment block
- **B-Blocked Items:** any row with a B block anywhere in the export
- **Open Invoices by cutoff:** invoices open on or before the cutoff date

**Clerk filter:** only the 82 allowed clerk codes are included in the payout check
(C1–C7, B1–B9, N1–N2, P1–P6, F1–F5, I1–I6, A1–A8, D1–D6, H1–H7, Z1–Z9, 70–75, 90, W1–W8, X1–X2).
        """)

    # ── General ────────────────────────────────────────────────────────────
    with st.expander("⚙️  General questions", expanded=False):
        st.markdown("""
**Are my files stored anywhere?**

Uploaded SAP exports are processed entirely in memory and never stored anywhere.
Templates, chunking rules, and account groups are saved to your private GitHub repository.
Nothing goes to a third-party server.

**Is the app secure for confidential data?**

That depends on deployment. If running on **Streamlit Community Cloud** (streamlit.io),
data passes through Streamlit's AWS servers — acceptable for internal tooling but check
with your IT team for sensitive financial data. For maximum security, deploy on **Azure App
Service** or an on-premise server so data never leaves the AB InBev network.

**What SAP export format do I need?**

Standard FBL5N `.xlsx` export. Include all columns — the app strips internal ones
automatically. For multi-year Customer Overview, export the full transaction history
(not just open items). Single-account and multi-account exports both work for all tools.

**Why do my arrears numbers in the multi-year overview differ from the export?**

They shouldn't — the multi-year overview preserves the exact arrears values from the SAP
export. The current overview tool recalculates arrears against your chosen reference date,
but the multi-year tool never overwrites them.

**Why does the total in the splitter differ from the raw export total?**

Usually because "Remove invoices not yet due" is ticked. Also check per-account date
overrides — if an account has a specific override date set, only invoices due on or
before that date are included. Historical overdue rows (net due date in a prior year)
are always included regardless.

**Red = invoices, Green = credits — is that right?**

Yes. Positive amounts (invoices) = red. Negative amounts (credit notes, payments) = green.
This is the Belgian AR convention used across all outputs.

**GitHub status shows 🔴 offline**

The app checks GitHub connectivity once every 5 minutes. A 🔴 means either GitHub is
unreachable or your token/repo in Streamlit secrets is not configured. Templates fall back
to session-only storage until connectivity is restored.

**Customer language lookup — which accounts are included?**

392 accounts are in the customer master file (combined from two source files). Each account
has a preferred language (N = Dutch, F = French, E = English, I = English fallback).
The lookup works with or without leading zeros — both 30124101 and 0030124101 resolve correctly.
Accounts not in the master default to English.

**Quality-of-life features**

- **Sidebar task widget** — always shows today's tasks with weekend shift applied
- **Freeze panes** on all generated Excels so headers stay fixed when scrolling
- **Consistent filenames** — all downloads include the account number and date
- **Language persists** within a session once set in any tool
- **Sender name & company persist** across pages once entered in any email draft
- **GitHub status** — 🟢 connected or 🔴 offline in sidebar (checked every 5 minutes)
        """)
