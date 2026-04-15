"""
Customer yearly overview engine.

GROUPING: Uses SAP blank-row separators (rows with no Account) to identify groups.
YEAR RULE: A group belongs to the year of the oldest positive RV invoice in it.
           If no positive RV exists, uses the oldest document date in the group.
SORT:      Within each year: groups sorted newest→oldest by the oldest doc date.
           Within each group: rows in original SAP order (preserved as-is).
G/L SPLIT: If multiple G/L accounts in a year, split into sub-sections with subtotals.
COLUMNS:   Same strip rules as Account Splitter PLUS keep Clearing date, Clearing Document,
           Payment Method, G/L Account. Document Type replaced with description.
LANGUAGE:  EN / NL / FR.
COLOURS:   Positive = RED (invoices), Negative = GREEN (credits/payments).
"""

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime
import warnings
warnings.filterwarnings("ignore")

# ── COLOURS ───────────────────────────────────────────────────────────────────
DK_BLUE  = "1F3864"
MD_BLUE  = "2E75B6"
LT_BLUE  = "BDD7EE"
YELLOW   = "FFEE09"
WHITE    = "FFFFFF"
GREY     = "F2F2F2"
POS_FG   = "C00000"   # red   – invoices / positive
NEG_FG   = "375623"   # green – credits / payments / negative
BLACK_FG = "000000"

# ── STRIP COLUMNS (same as splitter, but KEEP the overview-specific ones) ─────
STRIP_COLS = {
    "Reason code","Clerk Abbreviation","Cleared/open items symbol",
    "Case ID","Status","Dunning Block","Disputed item","Payment Block",
    "Net due date symbol","Text","Dunning Level","Last Dunned",
    "Reversed with","Document Header Text","User Name","Special G/L ind.",
    "Billing Document","Reference Key 1",
    # KEPT: Payment Method, G/L Account, Clearing date, Clearing Document
}

GL_LABELS = {
    "en": {"2400000": "Beer",  "2530009": "Rent"},
    "nl": {"2400000": "Bier",  "2530009": "Huur"},
    "fr": {"2400000": "Bière", "2530009": "Loyer"},
}

T = {
    "en": {
        "title_suffix":    "Customer Overview",
        "subtitle":        "Grouped by clearing document  ·  Year = oldest invoice in group  ·  Positive = invoices (red)  ·  Negative = credits / payments (green)",
        "year_banner":     "{yr}  ·  {n} groups  ·  Invoices: €{inv}  ·  Credits: €{cred}  ·  Net: €{net}",
        "year_total":      "{yr} — Total",
        "grand_total":     "Grand Total {a}–{b}",
        "net_balance":     "Net Balance",
        "no_transactions": "No transactions in {yr}",
        "gl_subtotal":     "{lbl} — Subtotal",
        "group_subtotal":  "Subtotal",
        "gl_other":        "Other",
        "desc_col":        "Description",
        "doc_types": {
            "RV+": "Invoice",          "RV-": "Credit note",
            "ZP":  "Payment",          "DZ":  "Payment",
            "RS+": "Re-invoice",       "RS-": "Bonus",
            "AB":  "Clearing",         "X_PAY": "Payout to customer",
        },
    },
    "nl": {
        "title_suffix":    "Klantoverzicht",
        "subtitle":        "Gegroepeerd per verrekeningsdocument  ·  Jaar = oudste factuur in groep  ·  Positief = facturen (rood)  ·  Negatief = creditnota's / betalingen (groen)",
        "year_banner":     "{yr}  ·  {n} groepen  ·  Facturen: €{inv}  ·  Creditnota's: €{cred}  ·  Netto: €{net}",
        "year_total":      "{yr} — Totaal",
        "grand_total":     "Eindtotaal {a}–{b}",
        "net_balance":     "Nettosaldo",
        "no_transactions": "Geen transacties in {yr}",
        "gl_subtotal":     "{lbl} — Subtotaal",
        "group_subtotal":  "Subtotaal",
        "gl_other":        "Overig",
        "desc_col":        "Omschrijving",
        "doc_types": {
            "RV+": "Factuur",          "RV-": "Creditnota",
            "ZP":  "Betaling",         "DZ":  "Betaling",
            "RS+": "Refactuur",        "RS-": "Bonus",
            "AB":  "Verrekening",      "X_PAY": "Uitbetaling aan klant",
        },
    },
    "fr": {
        "title_suffix":    "Aperçu client",
        "subtitle":        "Groupé par document de compensation  ·  Année = facture la plus ancienne  ·  Positif = factures (rouge)  ·  Négatif = avoirs / paiements (vert)",
        "year_banner":     "{yr}  ·  {n} groupes  ·  Factures: €{inv}  ·  Avoirs: €{cred}  ·  Net: €{net}",
        "year_total":      "{yr} — Total",
        "grand_total":     "Total général {a}–{b}",
        "net_balance":     "Solde net",
        "no_transactions": "Aucune transaction en {yr}",
        "gl_subtotal":     "{lbl} — Sous-total",
        "group_subtotal":  "Sous-total",
        "gl_other":        "Autre",
        "desc_col":        "Description",
        "doc_types": {
            "RV+": "Facture",          "RV-": "Avoir",
            "ZP":  "Paiement",         "DZ":  "Paiement",
            "RS+": "Re-facturation",   "RS-": "Bonus",
            "AB":  "Compensation",     "X_PAY": "Paiement au client",
        },
    },
}


def _t(lang, key, **kw):
    val = T.get(lang, T["en"]).get(key, T["en"].get(key, key))
    try:    return val.format(**kw) if kw else val
    except: return val


def _desc(doc_type, amount, pay_method, lang):
    dt  = str(doc_type  or "").strip().upper()
    pm  = str(pay_method or "").strip().upper()
    amt = float(amount) if str(amount) not in ("", "nan") else 0.0
    d   = T.get(lang, T["en"])["doc_types"]
    if pm == "X":          return d["X_PAY"]
    if dt == "RV":         return d["RV+"] if amt >= 0 else d["RV-"]
    if dt in ("ZP","DZ"):  return d["ZP"]
    if dt == "RS":         return d["RS+"] if amt >= 0 else d["RS-"]
    if dt == "AB":         return d["AB"]
    return d.get(dt, "")


def _gl_lbl(gl, lang):
    g = str(gl or "").strip().split(".")[0]
    return GL_LABELS.get(lang, GL_LABELS["en"]).get(g,
           _t(lang, "gl_other") if g else "")


# ── EXCEL HELPERS ─────────────────────────────────────────────────────────────
def _fill(rgb):   return PatternFill("solid", fgColor=rgb)
def _font(bold=False, color=BLACK_FG, size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)
def _align(ha="left"):
    return Alignment(horizontal=ha, vertical="center")
def _thin():
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)

def _w(ws, row, col, val=None, bold=False, bg=WHITE, fg=BLACK_FG,
       size=10, ha="left", fmt=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font = _font(bold=bold, color=fg, size=size)
    c.fill = _fill(bg); c.alignment = _align(ha); c.border = _thin()
    if fmt: c.number_format = fmt
    return c

def _mw(ws, row, c1, c2, val=None, bold=False, bg=WHITE,
        fg=BLACK_FG, size=10, ha="center"):
    ws.merge_cells(f"{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}")
    c = ws.cell(row=row, column=c1, value=val)
    c.font = _font(bold=bold, color=fg, size=size)
    c.fill = _fill(bg); c.alignment = _align(ha); c.border = _thin()
    for col in range(c1+1, c2+1):
        ws.cell(row=row, column=col).fill   = _fill(bg)
        ws.cell(row=row, column=col).border = _thin()
    return c


# ── PARSE ─────────────────────────────────────────────────────────────────────

def prepare_df(file_obj):
    """Read raw SAP export, keeping ALL rows (including blank subtotal rows)."""
    df = pd.read_excel(file_obj, sheet_name=0, header=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        if any(kw in col.lower() for kw in ["date", "datum"]):
            # Skip "Arrears after net due date" — it contains numbers (days), not dates
            if "arrears" in col.lower():
                continue
            df[col] = pd.to_datetime(df[col], errors="coerce")
    amt_col = next((c for c in df.columns if "amount" in c.lower()
                    or "bedrag" in c.lower() or "betrag" in c.lower()), None)
    if amt_col:
        df[amt_col] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
    return df, amt_col


def _parse_groups(df, amt_col):
    """
    Split df into groups using blank Account rows as separators.
    Each group is a list of row dicts (real SAP rows only, no blank rows).
    Returns list of groups, each group = list of row Series.
    """
    acc_col = next((c for c in df.columns
                    if c.lower() in ("account","customer","debitor","konto")), None)
    groups, current = [], []
    for _, row in df.iterrows():
        acc = str(row.get(acc_col, "") or "").strip() if acc_col else ""
        if acc and acc not in ("nan", "None", ""):
            current.append(row)
        else:
            if current:
                groups.append(current)
                current = []
    if current:
        groups.append(current)
    return groups


def _group_year(group, doc_date_col, doc_type_col, amt_col):
    """
    Year of a group = year of the oldest net due date among NON-clearing rows.
    AB/Clearing rows are excluded so one old clearing doesn't drag current
    invoices into a past year. Fallback: all rows, then doc date.
    """
    net_due_col = next((c for c in group[0].index
                        if "net due" in c.lower()
                        or "vervaldatum" in c.lower()), None) if group else None

    SKIP_TYPES = {"AB", "ZP", "DZ"}   # clearing/payment types to skip for year assignment

    due_dates_inv, due_dates_all = [], []
    doc_dates_inv, doc_dates_all = [], []
    for row in group:
        dt = str(row.get(doc_type_col, "") or "").strip().upper() if doc_type_col else ""
        is_clearing = dt in SKIP_TYPES
        if net_due_col:
            nd = row.get(net_due_col)
            if nd is not None and pd.notna(nd):
                due_dates_all.append(nd)
                if not is_clearing:
                    due_dates_inv.append(nd)
        if doc_date_col:
            dd = row.get(doc_date_col)
            if dd is not None and pd.notna(dd):
                doc_dates_all.append(dd)
                if not is_clearing:
                    doc_dates_inv.append(dd)

    # Prefer invoice dates; fall back to all dates
    dates = (due_dates_inv or due_dates_all or
             doc_dates_inv or doc_dates_all)
    if not dates:
        return None
    oldest = min(dates)
    return oldest.year if hasattr(oldest, "year") else int(str(oldest)[:4])


# ── BUILD ─────────────────────────────────────────────────────────────────────


def build_overview(df: pd.DataFrame, amt_col: str,
                   year_from: int, year_to: int,
                   customer_name:  str  = "",
                   account_id:     str  = "",
                   lang:           str  = "en",
                   reference_date       = None,
                   remove_not_due: bool = False) -> BytesIO:

    years = list(range(year_from, year_to + 1))

    # Apply "remove not yet due" filter based on reference date
    import datetime as _dt
    if reference_date is None:
        reference_date = _dt.date.today()
    ref_ts    = pd.Timestamp(reference_date)
    today_str = ref_ts.strftime("%d/%m/%Y")

    if remove_not_due:
        net_due_col_filt = next((c for c in df.columns
                                 if "net due" in c.lower()
                                 or "vervaldatum" in c.lower()), None)
        if net_due_col_filt and net_due_col_filt in df.columns:
            due = pd.to_datetime(df[net_due_col_filt], errors="coerce")
            df  = df[due.isna() | (due <= ref_ts)].copy()

    # Key column names
    doc_date_col = next((c for c in df.columns if "document date" in c.lower()
                         or "belegdatum" in c.lower()), None)
    doc_type_col = next((c for c in df.columns if "document type" in c.lower()
                         or "belegtyp"   in c.lower()), None)
    pay_meth_col = next((c for c in df.columns if "payment method" in c.lower()), None)
    gl_col       = next((c for c in df.columns if "g/l account" in c.lower()
                         or "sachkonto" in c.lower()), None)

    # Display columns: strip unwanted, replace Document Type with __DESC__
    kept = [c for c in df.columns if c not in STRIP_COLS]
    display_cols = []
    for c in kept:
        display_cols.append("__DESC__" if c == doc_type_col else c)
    ncols = len(display_cols)

    # Which kept columns are dates
    date_cols = {c for c in kept
                 if any(kw in c.lower() for kw in ["date","datum"])
                 or (c in df.columns and pd.api.types.is_datetime64_any_dtype(df[c]))}
    amt_ci = (display_cols.index(amt_col) + 1) if amt_col and amt_col in display_cols else None
    desc_name = _t(lang, "desc_col")

    # Parse all groups
    all_groups = _parse_groups(df, amt_col)

    if year_from == year_to:
        # Single mode: show ALL groups flat — no year filtering at all
        by_year = {year_to: list(all_groups)}
        years   = [year_to]
    else:
        # Assign year to each group
        by_year = {yr: [] for yr in years}
        for grp in all_groups:
            yr = _group_year(grp, doc_date_col, doc_type_col, amt_col)
            if yr is not None and year_from <= yr <= year_to:
                by_year[yr].append(grp)

    # Sort groups within each year: newest net due date first
    net_due_col_sort = next((c for c in df.columns
                             if "net due" in c.lower()
                             or "vervaldatum" in c.lower()), None)
    def grp_sort_key(grp, _nd=net_due_col_sort, _dd=doc_date_col):
        SKIP = {"AB","ZP","DZ"}
        due_inv, due_all = [], []
        for row in grp:
            dt = str(row.get(doc_type_col,"") or "").strip().upper() if doc_type_col else ""
            nd = row.get(_nd) if _nd else None
            if nd is not None and pd.notna(nd):
                due_all.append(nd)
                if dt not in SKIP: due_inv.append(nd)
        dates = due_inv or due_all
        return min(dates) if dates else pd.Timestamp.min
    for yr in years:
        by_year[yr].sort(key=grp_sort_key, reverse=True)  # newest first

    # ── Workbook ──────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overview"

    # Column widths
    for ci, col in enumerate(display_cols, 1):
        if col == "__DESC__":
            ws.column_dimensions[get_column_letter(ci)].width = 26
        elif col in date_cols:
            ws.column_dimensions[get_column_letter(ci)].width = 13
        elif amt_col and col == amt_col:
            ws.column_dimensions[get_column_letter(ci)].width = 18
        else:
            samp = df[col].astype(str) if col in df.columns else pd.Series([""])
            w    = max(len(col), samp.str.len().max() if len(samp) else 0)
            ws.column_dimensions[get_column_letter(ci)].width = min(max(w + 2, 9), 28)

    # ── Title rows ────────────────────────────────────────────────────────────
    r = 1
    if year_from == year_to:
        title = "  ·  ".join(filter(None, [
            customer_name,
            f"Account {account_id}" if account_id else "",
            _t(lang, "title_suffix"),
        ]))
    else:
        title = "  ·  ".join(filter(None, [
            customer_name,
            f"Account {account_id}" if account_id else "",
            f"{year_from}–{year_to}  {_t(lang,'title_suffix')}",
        ]))
    _mw(ws, r, 1, ncols, title, bold=True, bg=DK_BLUE, fg=WHITE, size=14)
    ws.row_dimensions[r].height = 34; r += 1
    sub_extra = f"  ·  {today_str}" + ("  ·  Invoices not yet due removed" if remove_not_due else "")
    _mw(ws, r, 1, ncols, _t(lang, "subtitle") + sub_extra, bg=MD_BLUE, fg=WHITE, size=9)
    ws.row_dimensions[r].height = 16; r += 2

    # ── Helpers ───────────────────────────────────────────────────────────────
    def col_headers(bg):
        nonlocal r
        for ci, col in enumerate(display_cols, 1):
            hdr = desc_name if col == "__DESC__" else col
            c   = ws.cell(row=r, column=ci, value=hdr)
            c.font = _font(bold=True, color=WHITE, size=9)
            c.fill = _fill(bg); c.alignment = _align("center"); c.border = _thin()
        ws.row_dimensions[r].height = 15; r += 1

    def data_row(row_data, bg):
        nonlocal r
        for ci, col in enumerate(display_cols, 1):
            is_amt  = (ci == amt_ci)
            is_date = (col in date_cols)
            is_desc = (col == "__DESC__")

            if is_desc:
                val = _desc(
                    row_data.get(doc_type_col, "") if doc_type_col else "",
                    row_data.get(amt_col, 0)       if amt_col else 0,
                    row_data.get(pay_meth_col, "")  if pay_meth_col else "",
                    lang,
                )
                fg = BLACK_FG
            elif col not in row_data.index:
                val, fg = "", BLACK_FG
            elif is_amt:
                val = float(row_data[col]) if pd.notna(row_data[col]) else 0.0
                fg  = POS_FG if val >= 0 else NEG_FG
            elif is_date:
                try:
                    val = (pd.Timestamp(row_data[col]).to_pydatetime()
                           if pd.notna(row_data[col]) else "")
                except Exception:
                    val = str(row_data[col]) if pd.notna(row_data[col]) else ""
                fg = BLACK_FG
            elif pd.isna(row_data[col]):
                val, fg = "", BLACK_FG
            elif gl_col and col == gl_col:
                raw = str(row_data[col]).strip().split(".")[0]
                lbl = _gl_lbl(raw, lang)
                val, fg = (f"{raw} — {lbl}" if lbl else raw), BLACK_FG
            elif isinstance(row_data[col], float) and row_data[col] == int(row_data[col]):
                val, fg = int(row_data[col]), BLACK_FG
            else:
                val, fg = row_data[col], BLACK_FG

            cell = ws.cell(row=r, column=ci, value=val)
            cell.font      = _font(color=fg, size=9)
            cell.fill      = _fill(bg)
            cell.alignment = _align("right" if is_amt else "left")
            cell.border    = _thin()
            if is_amt:
                cell.number_format = "#,##0.00"
            elif is_date and isinstance(val, datetime.datetime):
                cell.number_format = "DD/MM/YYYY"
        ws.row_dimensions[r].height = 13; r += 1

    def subtotal_row(label, total, bg, fg_label=WHITE, size=9):
        nonlocal r
        for ci in range(1, ncols + 1):
            if ci == 1:
                _w(ws, r, ci, label, bold=True, bg=bg, fg=fg_label, size=size)
            elif ci == amt_ci:
                c2 = ws.cell(row=r, column=ci, value=total)
                c2.font = _font(bold=True, color=WHITE, size=size)
                c2.fill = _fill(bg); c2.alignment = _align("right")
                c2.number_format = "#,##0.00"; c2.border = _thin()
            else:
                ws.cell(row=r, column=ci).fill   = _fill(bg)
                ws.cell(row=r, column=ci).border = _thin()
        ws.row_dimensions[r].height = 14; r += 1

    # ── Year sections ─────────────────────────────────────────────────────────
    grand_total = 0.0

    for yr in years:
        yr_groups = by_year[yr]
        all_rows  = [row for g in yr_groups for row in g]
        if all_rows:
            amts    = [float(rw.get(amt_col, 0) or 0) for rw in all_rows if amt_col]
            yr_inv  = sum(a for a in amts if a > 0)
            yr_cred = sum(a for a in amts if a < 0)
            yr_net  = sum(amts)
        else:
            yr_inv = yr_cred = yr_net = 0.0
        grand_total += yr_net

        # Year banner — skip in single mode (title row already has customer info)
        if year_from != year_to:
            _mw(ws, r, 1, ncols,
                _t(lang, "year_banner", yr=yr, n=len(yr_groups),
                   inv=f"{yr_inv:,.2f}", cred=f"{yr_cred:,.2f}", net=f"{yr_net:,.2f}"),
                bold=True, bg=DK_BLUE, fg=WHITE, size=11)
            ws.row_dimensions[r].height = 22; r += 1

        if not yr_groups:
            _mw(ws, r, 1, ncols, _t(lang, "no_transactions", yr=yr),
                bg=GREY, fg=BLACK_FG, size=9)
            ws.row_dimensions[r].height = 16; r += 2; continue

        # Determine G/L sub-sections for this year
        if gl_col and gl_col in df.columns:
            gl_raw    = [str(rw.get(gl_col,"") or "").strip().split(".")[0]
                         for g in yr_groups for rw in g]
            known     = [g for g in ["2400000","2530009"] if g in gl_raw]
            others    = sorted({g for g in gl_raw
                                if g not in known and g not in ("","nan","None")})
            has_blank = any(g in ("","nan","None") for g in gl_raw)
            gl_sections = known + others + ([""] if has_blank else [])
            multi_gl    = len([g for g in gl_sections if g]) > 1
        else:
            gl_sections, multi_gl = [None], False

        for gl_key in gl_sections:
            # Filter groups to this GL section
            if gl_key is None:
                sec_groups = yr_groups
            else:
                sec_groups = []
                for grp in yr_groups:
                    if gl_key == "":
                        filtered = [rw for rw in grp
                                    if gl_col and str(rw.get(gl_col,"") or "").strip()
                                    .split(".")[0] in ("","nan","None")]
                    else:
                        filtered = [rw for rw in grp
                                    if gl_col and str(rw.get(gl_col,"") or "").strip()
                                    .split(".")[0] == gl_key]
                    if filtered:
                        sec_groups.append(filtered)

            if not sec_groups:
                continue

            sec_amts  = [float(rw.get(amt_col,0) or 0)
                         for g in sec_groups for rw in g if amt_col]
            sec_total = sum(sec_amts)

            # G/L sub-header
            if multi_gl:
                if gl_key in ("","nan","None",""):
                    gl_title = _t(lang, "gl_other")
                else:
                    lbl      = _gl_lbl(gl_key, lang)
                    gl_title = f"{gl_key} — {lbl}" if lbl else gl_key
                _mw(ws, r, 1, ncols, f"  {gl_title}",
                    bold=True, bg=LT_BLUE, fg=DK_BLUE, size=9, ha="left")
                ws.row_dimensions[r].height = 14; r += 1

            # Column headers (once per GL section)
            col_headers(MD_BLUE)

            # Each group
            for grp in sec_groups:
                grp_total = sum(float(rw.get(amt_col,0) or 0)
                                for rw in grp if amt_col)
                for ri, row_data in enumerate(grp):
                    data_row(row_data, GREY if ri % 2 == 0 else WHITE)

                # Group subtotal
                subtotal_row(_t(lang, "group_subtotal"), grp_total, MD_BLUE)

                # Blank gap between groups
                ws.row_dimensions[r].height = 5; r += 1

            # No G/L subtotals — groups already have their own subtotals

        # Year total
        subtotal_row(_t(lang, "year_total", yr="" if year_from == year_to else yr).strip(" —"),
                     yr_net, DK_BLUE, size=10)

        # Gap between years
        ws.row_dimensions[r].height = 10; r += 1
        ws.row_dimensions[r].height = 10; r += 1

    # ── Grand total ───────────────────────────────────────────────────────────
    _mw(ws, r, 1, ncols,
        _t(lang, "grand_total", a=year_from, b=year_to),
        bold=True, bg=DK_BLUE, fg=WHITE, size=12)
    ws.row_dimensions[r].height = 28; r += 1

    for ci in range(1, ncols + 1):
        if ci == 1:
            _w(ws, r, ci, _t(lang, "net_balance"),
               bold=True, bg=DK_BLUE, fg=WHITE, size=11)
        elif ci == amt_ci:
            c2 = ws.cell(row=r, column=ci, value=grand_total)
            c2.font = _font(bold=True, color=WHITE, size=13)
            c2.fill = _fill(DK_BLUE); c2.alignment = _align("right")
            c2.number_format = "#,##0.00"; c2.border = _thin()
        else:
            ws.cell(row=r, column=ci).fill   = _fill(DK_BLUE)
            ws.cell(row=r, column=ci).border = _thin()
    ws.row_dimensions[r].height = 26
    ws.freeze_panes = "A4"

    out = BytesIO()
    wb.save(out); out.seek(0)
    return out
