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

Two tabs for handling customer payments — remittance matching and amount-only matching.

---

**Tab 1 — Remittance matching**

Matches a customer's payment advice against your SAP open items line by line.

1. Export open items from SAP FBL5N for that customer → `.xlsx`
2. Upload the SAP export and the customer's remittance file
3. Enter customer name, payment amount, and payment date
4. Click **Run Reconciliation**

Output: colour-coded Excel (matched / unmatched / already cleared / partial), a discrepancy
summary, and a draft follow-up email in EN / NL / FR that you can open directly in your mail client.

**Matching logic:** tries Assignment number first, then Document Number, then Reference Key 3,
then substring. SAP is the source of truth — the tool ignores sign conventions in the remittance.
Already-cleared items are flagged separately as potential doubles.

---

**Tab 2 — Amount-only matching**

Customer paid without sending a remittance? Enter the amount and the tool finds which
combination of invoices adds up to it.

1. Upload SAP export and enter the payment amount
2. Set a tolerance (default €0.05)
3. Click **Find matching invoices**

Tries exact single invoice → two-invoice pairs → greedy subset-sum across all open invoices.
Returns up to 5 options ranked by closeness — 🟢 exact match or 🟡 within tolerance.
Downloadable as Excel with one sheet per option.
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
6. Select document language (EN / NL / FR) — translates RV → Invoice/Factuur/Facture,
   ZP/DZ → Payment, RS− → Bonus, AB → Clearing, etc.
7. Download all accounts (standard layout) or individual custom downloads below

---

**Per-account reference dates**

When "Remove invoices not yet due" is ticked and there are multiple accounts, an
**⚙️ Per-account date overrides** expander appears. Each account defaults to the global
date but can be set individually — useful when different customers have different month-end
cut-offs (e.g. account A = 25th, account B = 31st).

---

**Results summary table**

After splitting, a table shows each account with its row count and total. Accounts with
a date override show ⚙ next to their date.

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

- Automatically detects the current-open section from the SAP export (rows before the
  first DZ/ZP payment row, or all rows if the export is already filtered to open items)
- **Remove not yet due:** hides invoices with a net due date after the reference date
  (default = off, so all open items show)
- **Remove current overdues:** hides already-overdue items, leaving only future-dated invoices
- **Filter by month range:** narrow results to a specific month window within the year
- Works with single-account exports *and* full multi-year exports — it extracts just
  the current open section automatically

*Output format:* title row → subtitle row → column headers → data rows → Net Balance total.

---

**Mode 2 — Multi-year Overview**

Shows the full clearing history grouped by year, newest year first.

- **Current Open Items** section at the top (can be removed with "Remove current overdues")
- Each year shows: group count · invoice total · credits total · net balance
- Within each year, SAP clearing groups are preserved with yellow zero-subtotal rows showing
  payment grouping boundaries
- Year totals in mid-blue; grand Net Balance in dark navy at the bottom
- Arrears column shows the **original SAP export values** (not recalculated) so they match
  the source exactly

---

**G/L split (both modes)**

When a customer has both Beer (G/L 2400000) and Rent (2530009) invoices, the current
overview automatically groups them into separate sections with individual subtotals and
a combined Net Balance. Labels translate per language:

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

**Today banner**

When tasks are scheduled for today, a black banner appears at the top of the page showing
each task with a colour-coded pill. Also visible in the **sidebar** on every page so you
never have to leave your current tool to check.

Sidebar widget also shows:
- Month progress bar (how far through the month you are)
- **Next up** — the next day this month with tasks and how many days away it is
- **Open Calendar →** button

---

**Calendar selector**

Dropdown at the top of the calendar page. Switch between multiple calendars instantly.

**➕ New calendar** — click the button, type a name, hit Create. Starts empty.

**🗑 Delete calendar** — only available when 2+ calendars exist. Requires a confirmation click.

**Rename** — flip on edit mode, use the rename expander at the top.

---

**Viewing the calendar**

A 7-column month grid shows all scheduled tasks as colour-coded chips inside each day cell.
Today is highlighted with a gold border. Weekends are greyed out. Use the Month / Year
selectors to navigate.

**Colour coding:**
- 🟡 Yellow = Direct Debit (DD)
- ⚫ Black/gold = Overview
- 🔴 Red = UAC
- 🔵 Blue = Meeting

---

**Editing tasks**

Toggle **✏️ edit mode** on (top-right of the page). The full task list appears as an
interactive spreadsheet table below the calendar:

- **Click any cell** to edit it inline — Type and Format are dropdowns, Account and Note
  are free text
- **Add a row** — use the + button at the bottom of the table, fill in Day (1–31),
  Type, Account, and optionally Format and Note
- **Delete a row** — tick the checkbox on the left of the row, then click the bin icon
  that appears in the table header
- **Move a task** to a different day — just change the Day number in the Day column
- Hit **💾 Save changes** to apply — the visual grid updates immediately

The table shows all tasks across all 31 days in one place so you can see and edit
everything without navigating day by day.

---

**Task types**

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
- **Plain table detection:** SAP-style column headers detected regardless of merged title
  rows above. Title rows are refreshed with current values on each split.
- **Chunked layout:** enter account + batch size in Chunking rules → invoices split into
  payment batches of that size.

---

**Account groups**

*Setup:* Customer templates → Account groups → label + account numbers (comma-separated)
→ tick Flat merge if needed → Save group. Requires admin login.

- **Multi-sheet:** one sheet per account. Template required on primary account.
- **Flat merge:** all rows combined, sorted by net due date. No template needed.
  Useful for linked accounts like 30351345 + 30104410.

---

**Admin access**

A password set in Streamlit secrets protects all write/delete actions (save template,
replace template, delete template, save rule, delete rule, save group, delete group).
Enter the password once per session in the 🔐 Admin login expander.
        """)

    # ── Bonus & Payout ─────────────────────────────────────────────────────
    with st.expander("🎁  Bonus & Payout Tools", expanded=False):
        st.markdown("""
**Tab 1 — Customer matching**

Compare your SAP customer list against a bonus/partner file. Shows which accounts match,
which are in the bonus file but missing from SAP, and which SAP accounts are absent from
the bonus file.

1. Upload SAP export and the bonus/partner file
2. Confirm which column holds the account number in each
3. Click **Run matching**

Output: annotated bonus file (green = match, orange = not in SAP, yellow = added from SAP),
a Summary sheet with counts, and a list of missing accounts.

---

**Tab 2 — Payout & block checker**

Scans a SAP export to verify:
- All **Payment Method X** (payout-to-customer) rows have no B or U payment block
- No **B-blocked** items anywhere in the export
- No **open invoices** on or before a chosen cutoff date (default: 21st of the month)

1. Upload the SAP export
2. Set the invoice cutoff date
3. Click **Run check**

Output: dashboard with counts and colour-coded alerts, plus downloadable Excel with four
sheets — X payouts OK · X payouts blocked · B-blocked items · open invoices by cutoff.
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
Service** or an on-premise server so data never leaves the AB InBev network. Azure AD
authentication can also be added so only AB InBev employees can access the app.

**What SAP export format do I need?**

Standard FBL5N `.xlsx` export. Include all columns — the app strips internal ones
automatically. For multi-year Customer Overview, export the full transaction history
(not just open items). Single-account and multi-account exports both work for all tools.

**Why do my arrears numbers in the multi-year overview differ from the export?**

They shouldn't — as of v143, the multi-year overview preserves the exact arrears values
from the SAP export. The current overview tool recalculates arrears against your chosen
reference date (so they reflect that specific date), but the multi-year tool never overwrites them.

**Why does the total in the splitter differ from the raw export total?**

Usually because "Remove invoices not yet due" is ticked. Any invoice with a net due date
after the reference date is excluded. Also check per-account date overrides — if an
account has a specific override date set, only invoices due on or before that date are included.
Historical overdue rows (net due date in a prior year) are always included regardless.

**Red = invoices, Green = credits — is that right?**

Yes. Positive amounts (invoices, money owed to you) = red. Negative amounts (credit notes,
payments received) = green. This is the Belgian AR convention used across all outputs.

**The app bounces back to Home when I click a tool in the sidebar**

This was a known issue fixed in v132. If you're still seeing it, make sure you've replaced
`app.py` and `page_home.py` with the v132+ versions.

**GitHub status shows 🔴 offline**

The app checks GitHub connectivity once every 5 minutes. A 🔴 means either GitHub is
unreachable or your token/repo in Streamlit secrets isn't configured. Templates fall back
to session-only storage until connectivity is restored.

**Quality-of-life features**

- **Sidebar task widget** — always shows today's scheduled tasks and what's coming next,
  from whichever calendar is currently selected. Collapses cleanly when not needed.
- **Version** shown in sidebar (e.g. v145)
- **GitHub status** — 🟢 connected or 🔴 offline in sidebar (checked every 5 minutes, not on every click)
- **Freeze panes** on all generated Excels so headers stay fixed when scrolling
- **Consistent filenames** — all downloads include the account number and date
- **Language persists** within a session once set in any tool
- **Sender name & company persist** across pages once entered in any email draft
        """)
