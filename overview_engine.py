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
DK_BLUE  = "FF1F3864"
MD_BLUE  = "FF2E75B6"
LT_BLUE  = "FFBDD7EE"
YELLOW   = "FFEE09"
WHITE    = "FFFFFFFF"
GREY     = "FFF2F2F2"
POS_FG   = "FFC00000"   # red   – invoices / positive
NEG_FG   = "FF375623"   # green – credits / payments / negative
BLACK_FG = "FF000000"

# ── STRIP COLUMNS (same as splitter, but KEEP the overview-specific ones) ─────
STRIP_COLS = {
    "Reason code","Clerk Abbreviation","Cleared/open items symbol",
    "Case ID","Status","Dunning Block","Disputed item","Payment Block",
    "Net due date symbol","Text","Dunning Level","Last Dunned",
    "Reversed with","Document Header Text","User Name","Special G/L ind.",
    "Billing Document","Reference Key 1",
    "Clearing date","Clearing Document",
    # KEPT: Payment Method, G/L Account
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
        "year_banner":     "{yr}  ·  {n} groupes  ·  Factures: €{inv}  ·  Notes de crédit: €{cred}  ·  Net: €{net}",
        "year_total":      "{yr} — Total",
        "grand_total":     "Total général {a}–{b}",
        "net_balance":     "Solde net",
        "no_transactions": "Aucune transaction en {yr}",
        "gl_subtotal":     "{lbl} — Sous-total",
        "group_subtotal":  "Sous-total",
        "gl_other":        "Autre",
        "desc_col":        "Description",
        "doc_types": {
            "RV+": "Facture",          "RV-": "Note de crédit",
            "ZP":  "Paiement",         "DZ":  "Paiement",
            "RS+": "Re-facturation",   "RS-": "Bonus",
            "AB":  "Ajustement comptable",     "X_PAY": "Virement au client",
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


def _recalc_arrears(df: pd.DataFrame, reference_date) -> pd.DataFrame:
    """Recalculate arrears based on reference date: (ref - net_due_date).days"""
    if reference_date is None:
        return df
    ndd_col = next((c for c in df.columns if "net due" in c.lower()), None)
    arr_col = next((c for c in df.columns if "arrears" in c.lower()), None)
    if not ndd_col or not arr_col:
        return df
    df = df.copy()
    ref_ts = pd.Timestamp(reference_date)
    due    = pd.to_datetime(df[ndd_col], errors="coerce")
    mask   = due.notna()
    days   = (ref_ts - due[mask]).dt.days
    if str(df[arr_col].dtype) == 'string':
        df.loc[mask, arr_col] = days.astype(str)
    else:
        try:
            df[arr_col] = pd.to_numeric(df[arr_col], errors='coerce')
            df.loc[mask, arr_col] = days
        except Exception:
            df.loc[mask, arr_col] = days.astype(str)
    return df
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


def build_current_overview(df: pd.DataFrame, amt_col: str,
                            reference_date=None,
                            remove_not_due: bool = False,
                            remove_overdues: bool = False,
                            month_from: int = 1,
                            month_to: int = 12) -> BytesIO:
    """
    Current overview — flat format matching the NL export style:
    - Alternating white / light-blue rows
    - Yellow row when a clearing group nets to zero
    - Sorted newest net due date first
    - No blank rows between groups
    """
    import datetime as _dt
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ref_ts = pd.Timestamp(reference_date) if reference_date else pd.Timestamp.now()
    ndd_col = next((c for c in df.columns if "net due" in c.lower()), None)
    arr_col = next((c for c in df.columns if "arrears" in c.lower()), None)
    doc_type_col = next((c for c in df.columns if "document type" in c.lower()), None)
    acc_col = next((c for c in df.columns if c.lower() in ("account","konto","debitor")), None)

    # ── Filters ───────────────────────────────────────────────────────────────
    if remove_not_due and ndd_col:
        due = pd.to_datetime(df[ndd_col], errors="coerce")
        df  = df[due.isna() | (due <= ref_ts)].copy()

    if remove_overdues and ndd_col and arr_col:
        # Remove rows where arrears > 0 (already past due date)
        arr = pd.to_numeric(df[arr_col], errors="coerce").fillna(0)
        df  = df[arr <= 0].copy()

    if (month_from != 1 or month_to != 12) and ndd_col:
        due2 = pd.to_datetime(df[ndd_col], errors="coerce")
        df   = df[due2.isna() | ((due2.dt.month >= month_from) & (due2.dt.month <= month_to))].copy()

    if ndd_col:
        df[ndd_col] = pd.to_datetime(df[ndd_col], errors="coerce")
    if amt_col:
        df[amt_col] = pd.to_numeric(df[amt_col], errors="coerce")
    if arr_col:
        df[arr_col] = pd.to_numeric(df[arr_col], errors="coerce")

    # Recalculate arrears based on reference date
    if reference_date:
        df = _recalc_arrears(df, reference_date)

    # ── Split into clearing-doc groups ────────────────────────────────────────
    groups = []
    current = []
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

    # ── Sort groups newest net due first ──────────────────────────────────────
    SKIP = {"AB", "ZP", "DZ"}
    def _oldest_due(grp):
        dates = []
        for row in grp:
            dt = str(row.get(doc_type_col, "") or "").strip().upper() if doc_type_col else ""
            nd = row.get(ndd_col) if ndd_col else None
            if nd is not None and pd.notna(nd) and dt not in SKIP:
                dates.append(pd.Timestamp(nd))
        return min(dates) if dates else pd.Timestamp.min

    groups.sort(key=lambda g: -int(_oldest_due(g).timestamp()))

    # ── Colours ───────────────────────────────────────────────────────────────
    HDR_FILL  = "FF1F3864"   # dark blue header
    ROW_WHITE = "FFFFFFFF"
    ROW_BLUE  = "FFEBF3FB"   # light blue alternate
    ROW_YELL  = "FFFFFF00"   # yellow subtotal
    COL_POS   = "FFC00000"   # red = positive amount
    COL_NEG   = "FF375623"   # green = negative amount
    COL_WHT   = "FFFFFFFF"
    COL_BLK   = "FF000000"

    def _fill(rgb): return PatternFill("solid", fgColor=rgb)
    def _font(bold=False, color=COL_BLK, size=9):
        return Font(name="Arial", bold=bold, color=color, size=size)
    def _thin():
        s = Side(style="thin", color="DDDDDD")
        return Border(left=s, right=s, top=s, bottom=s)
    def _aln(h="left"):
        return Alignment(horizontal=h, vertical="center", wrap_text=False)

    # ── Columns to display ────────────────────────────────────────────────────
    STRIP = {
        "Reason code","Clerk Abbreviation","Cleared/open items symbol",
        "Disputed item","Payment Block","Net due date symbol",
        "Text","Clearing date","Clearing Document","Dunning Level",
        "Last Dunned","Reversed with","Document Header Text","User Name",
        "Special G/L ind.","Billing Document","Reference Key 1",
        "doc_number_str","ref","sap_class","is_open","header_text",
        "clearing_date","clearing_doc","text",
    }
    display_cols = [c for c in df.columns if c not in STRIP]
    ncols = len(display_cols)

    col_widths = {
        "Account":10,"Assignment":14,"Document Number":18,
        "Reference Key 3":14,"Document Date":13,"Net due date":13,
        "Document Type":13,"Amount in local currency":20,
        "Arrears after net due date":24,"Payment Method":13,
        "G/L Account":18,"Case ID":10,"Status":10,
        "Dunning Block":13,"Disputed item":13,
    }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overview"

    for ci, col in enumerate(display_cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = col_widths.get(col, max(len(col)+2,12))

    # Header row
    for ci, h in enumerate(display_cols, 1):
        cell = ws.cell(1, ci, value=h)
        cell.font = _font(bold=True, color=COL_WHT, size=9)
        cell.fill = _fill(HDR_FILL)
        cell.alignment = _aln("center")
        cell.border = _thin()
    ws.row_dimensions[1].height = 15
    ws.freeze_panes = "A2"

    r = 2
    row_idx = 0  # for alternating colour

    for grp in groups:
        grp_total = sum(
            float(row.get(amt_col, 0) or 0)
            for row in grp if amt_col and pd.notna(row.get(amt_col))
        )

        for row in grp:
            bg = ROW_WHITE if row_idx % 2 == 0 else ROW_BLUE
            for ci, col in enumerate(display_cols, 1):
                val = row.get(col, "")
                if isinstance(val, pd.Timestamp):
                    val = val.to_pydatetime()
                elif not isinstance(val, (str, int, float, _dt.datetime, type(None))):
                    val = str(val)
                elif isinstance(val, float):
                    if val != val:  # NaN
                        val = None
                    elif val == int(val):
                        val = int(val)
                is_amt = amt_col and col == amt_col
                cell = ws.cell(r, ci, value=val if val != "" else None)
                if is_amt and isinstance(val, (int, float)) and val is not None:
                    cell.font = _font(color=COL_POS if val > 0 else (COL_NEG if val < 0 else COL_BLK))
                    cell.number_format = "#,##0.00"
                    cell.alignment = _aln("right")
                elif isinstance(val, _dt.datetime):
                    cell.font = _font()
                    cell.number_format = "DD/MM/YYYY"
                    cell.alignment = _aln("left")
                else:
                    cell.font = _font()
                    cell.alignment = _aln("left")
                cell.fill = _fill(bg)
                cell.border = _thin()
            ws.row_dimensions[r].height = 13
            r += 1
            row_idx += 1

        # Yellow subtotal row when group nets to zero
        if abs(grp_total) < 0.02:
            for ci in range(1, ncols + 1):
                cell = ws.cell(r, ci)
                cell.fill = _fill(ROW_YELL)
                cell.border = _thin()
                if ci == (display_cols.index(amt_col) + 1 if amt_col in display_cols else 8):
                    cell.value = 0
                    cell.font = _font(bold=True)
                    cell.number_format = "#,##0.00"
                    cell.alignment = _aln("right")
            ws.row_dimensions[r].height = 8
            r += 1
            row_idx += 1

    # ── Grand total row ───────────────────────────────────────────────────────
    grand_total = sum(
        float(row.get(amt_col, 0) or 0)
        for grp in groups for row in grp
        if amt_col and pd.notna(row.get(amt_col))
    )
    amt_ci = (display_cols.index(amt_col) + 1) if amt_col and amt_col in display_cols else None
    for ci in range(1, ncols + 1):
        cell = ws.cell(r, ci)
        cell.fill = _fill(HDR_FILL)
        cell.border = _thin()
        if ci == 1:
            cell.value = "Net Balance"
            cell.font = _font(bold=True, color=COL_WHT, size=10)
            cell.alignment = _aln("left")
        elif ci == amt_ci:
            cell.value = grand_total
            cell.font = _font(bold=True, color=COL_WHT, size=11)
            cell.number_format = "#,##0.00"
            cell.alignment = _aln("right")
    ws.row_dimensions[r].height = 20

    out = BytesIO()
    wb.save(out); out.seek(0)
    return out



def build_overview(df: pd.DataFrame, amt_col: str,
                   year_from: int, year_to: int,
                   customer_name: str = "",
                   account_id:    str = "",
                   lang:          str = "en",
                   remove_overdues: bool = False) -> BytesIO:
    """
    Multi-year overview — preserves SAP blank-row group separators exactly.

    Structure:
      1. CURRENT OPEN ITEMS  — groups whose rows all precede the first
         historical clearing block (index boundary detected automatically).
         Plain white/light-blue alternating rows (no special background).
         Current-open total row (mid-blue).
      2. One section per calendar year, newest year first.
         Each year: dark-blue banner → mid-blue column headers → data rows
         → yellow zero-subtotal rows (payment grouping boundaries) →
         mid-blue year total.
      3. Net Balance grand total (dark blue).

    Colours match build_current_overview exactly:
      - Dark navy  #1F3864 — year banners, grand total
      - Mid blue   #2E75B6 — column headers, year totals, open total
      - White      #FFFFFF / Light blue #EBF3FB — alternating data rows
      - Yellow     #FFFF00 — zero-netting group separator rows
      - Red        #C00000 — positive amounts (invoices)
      - Green      #375623 — negative amounts (credits / payments)
    """
    import datetime as _dt
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    HDR_FILL  = "FF1F3864"   # dark blue  — year banner + grand total
    BAND_FILL = "FF2E75B6"   # mid blue   — column headers + year/open totals
    ROW_WHITE = "FFFFFFFF"   # plain white — data rows (incl. current open)
    ROW_BLUE  = "FFEBF3FB"   # light blue  — alternating data rows
    ROW_YELL  = "FFFFFF00"   # yellow      — zero-netting group separator
    COL_POS   = "FFC00000"   # red         — positive amounts
    COL_NEG   = "FF375623"   # green       — negative amounts
    COL_WHT   = "FFFFFFFF"
    COL_BLK   = "FF000000"

    def _fill(rgb): return PatternFill("solid", fgColor=rgb)
    def _font(bold=False, color=COL_BLK, size=9):
        return Font(name="Arial", bold=bold, color=color, size=size)
    def _thin():
        s = Side(style="thin", color="DDDDDD")
        return Border(left=s, right=s, top=s, bottom=s)
    def _aln(h="left"):
        return Alignment(horizontal=h, vertical="center", wrap_text=False)

    STRIP = {
        "Reason code","Clerk Abbreviation","Cleared/open items symbol",
        "Disputed item","Payment Block","Net due date symbol",
        "Text","Clearing date","Clearing Document","Dunning Level",
        "Last Dunned","Reversed with","Document Header Text","User Name",
        "Special G/L ind.","Billing Document","Reference Key 1",
        "doc_number_str","ref","sap_class","is_open","header_text",
        "clearing_date","clearing_doc","text",
    }
    display_cols = [c for c in df.columns if c not in STRIP and c is not None and str(c) != "None"]
    ncols  = len(display_cols)
    col_widths = {
        "Account":10,"Assignment":14,"Document Number":18,
        "Reference Key 3":14,"Document Date":13,"Net due date":13,
        "Document Type":13,"Amount in local currency":20,
        "Arrears after net due date":24,"Payment Method":13,
        "G/L Account":18,"Case ID":10,"Status":10,
        "Dunning Block":13,"Disputed item":13,
    }
    amt_ci = (display_cols.index(amt_col)+1) if amt_col and amt_col in display_cols else None

    ndd_col      = next((c for c in df.columns if "net due"       in c.lower()), None)
    arr_col      = next((c for c in df.columns if "arrears"       in c.lower()), None)
    doc_type_col = next((c for c in df.columns if "document type" in c.lower()), None)
    doc_date_col = next((c for c in df.columns if c.lower() == "document date"), None)
    acc_col      = next((c for c in df.columns if c.lower() in ("account","konto","debitor")), None)

    if ndd_col:      df[ndd_col]      = pd.to_datetime(df[ndd_col],      errors="coerce")
    if doc_date_col: df[doc_date_col] = pd.to_datetime(df[doc_date_col], errors="coerce")
    if amt_col:      df[amt_col]      = pd.to_numeric( df[amt_col],      errors="coerce")
    if arr_col:      df[arr_col]      = pd.to_numeric( df[arr_col],      errors="coerce")

    if remove_overdues and arr_col:
        df = df[df[arr_col].fillna(0) <= 0].copy()

    # Recalculate arrears to today
    df = _recalc_arrears(df, _dt.date.today())

    # ── Parse SAP groups using blank Account rows as separators ──────────────
    # Each group = list of (original_index, row) tuples
    groups_raw, cur = [], []
    for idx, row in df.iterrows():
        acc = str(row.get(acc_col, "") or "").strip() if acc_col else ""
        if acc and acc not in ("nan", "None", ""):
            cur.append((idx, row))
        else:
            if cur:
                groups_raw.append(cur)
                cur = []
    if cur:
        groups_raw.append(cur)

    # ── Split: current-open groups vs historical ──────────────────────────────
    # SAP exports place current open items first, followed by cleared/historical
    # groups each terminated by a DZ or ZP payment row.
    # Detection strategy: a group is "historical" (cleared) when it contains at
    # least one DZ or ZP row (actual payment documents). AB rows appear in BOTH
    # open and historical sections so cannot be used as the discriminator.
    # We scan from the end of groups_raw backwards: the first group (from the
    # front) that has NO DZ/ZP row and whose dataframe indices all fall before
    # the first DZ/ZP row in the entire file is the open-items block.
    # Simpler and robust: find the minimum dataframe index of any DZ or ZP row.
    # Every group whose rows all sit before that index is current-open.
    PAY_TYPES = {"DZ", "ZP"}

    # Find the first (lowest) dataframe index that has a DZ or ZP doc type
    first_pay_index = None
    if doc_type_col:
        for idx, row in df.iterrows():
            dt = str(row.get(doc_type_col, "") or "").strip().upper()
            if dt in PAY_TYPES:
                first_pay_index = idx
                break

    if first_pay_index is None:
        # No payment rows at all — everything is current open
        current_open_groups = [[r for _, r in grp] for grp in groups_raw]
        historical_groups   = []
    else:
        # Groups whose maximum row index falls before the first payment row
        # are current-open; everything else is historical
        current_open_groups = []
        historical_groups   = []
        for grp in groups_raw:
            max_idx = max(i for i, _ in grp)
            if max_idx < first_pay_index:
                current_open_groups.append([r for _, r in grp])
            else:
                historical_groups.append([r for _, r in grp])

    # ── Bucket historical groups by year ──────────────────────────────────────
    SKIP_TYPES = {"AB", "ZP", "DZ"}

    def _group_year_local(grp):
        dates = []
        for row in grp:
            dt = str(row.get(doc_type_col, "") or "").strip().upper()
            if dt in SKIP_TYPES:
                continue
            for c in [ndd_col, doc_date_col]:
                v = row.get(c) if c else None
                if v is not None and pd.notna(v):
                    dates.append(pd.Timestamp(v))
        if not dates:
            return None
        return min(dates).year

    year_groups: dict = {}
    for grp in historical_groups:
        yr = _group_year_local(grp)
        if yr is not None and year_from <= yr <= year_to:
            year_groups.setdefault(yr, []).append(grp)

    # ── Helper: write a single data row ──────────────────────────────────────
    def _write_data_row(ws, r, row, bg, amt_ci):
        for ci, col in enumerate(display_cols, 1):
            val = row.get(col, "")
            if isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            elif not isinstance(val, (str, int, float, _dt.datetime, type(None))):
                val = str(val)
            elif isinstance(val, float):
                if val != val: val = None
                elif val == int(val): val = int(val)
            is_amt = amt_col and col == amt_col
            cell = ws.cell(r, ci, value=val if val != "" else None)
            if is_amt and isinstance(val, (int, float)) and val is not None:
                cell.font = _font(color=COL_POS if val > 0 else (COL_NEG if val < 0 else COL_BLK))
                cell.number_format = "#,##0.00"
                cell.alignment = _aln("right")
            elif isinstance(val, _dt.datetime):
                cell.font = _font()
                cell.number_format = "DD/MM/YYYY"
                cell.alignment = _aln("left")
            else:
                cell.font = _font()
                cell.alignment = _aln("left")
            cell.fill   = _fill(bg)
            cell.border = _thin()
        ws.row_dimensions[r].height = 13

    def _write_col_headers(ws, r):
        for ci, h in enumerate(display_cols, 1):
            cell = ws.cell(r, ci, value=h)
            cell.font      = _font(bold=True, color=COL_WHT, size=9)
            cell.fill      = _fill(BAND_FILL)
            cell.alignment = _aln("center")
            cell.border    = _thin()
        ws.row_dimensions[r].height = 15

    def _write_zero_subtotal(ws, r, amt_ci):
        for ci in range(1, ncols+1):
            cell = ws.cell(r, ci)
            cell.fill   = _fill(ROW_YELL)
            cell.border = _thin()
            if ci == amt_ci:
                cell.value         = 0
                cell.font          = _font(bold=True)
                cell.number_format = "#,##0.00"
                cell.alignment     = _aln("right")
        ws.row_dimensions[r].height = 8

    def _write_band_total(ws, r, label, total, amt_ci):
        for ci in range(1, ncols+1):
            cell = ws.cell(r, ci)
            cell.fill   = _fill(BAND_FILL)
            cell.border = _thin()
        ws.cell(r, 1).value     = label
        ws.cell(r, 1).font      = _font(bold=True, color=COL_WHT, size=10)
        ws.cell(r, 1).alignment = _aln("left")
        if amt_ci:
            c = ws.cell(r, amt_ci)
            c.value         = total
            c.font          = _font(bold=True, color=COL_WHT, size=10)
            c.number_format = "#,##0.00"
            c.alignment     = _aln("right")
        ws.row_dimensions[r].height = 18

    # ── Build workbook ────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overview"
    for ci, col in enumerate(display_cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = col_widths.get(col, max(len(str(col))+2, 12))

    r = 1
    row_idx = 0
    grand_total = 0.0

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 1 — CURRENT OPEN ITEMS
    # ══════════════════════════════════════════════════════════════════════════
    if current_open_groups:
        # Dark-blue banner
        for ci in range(1, ncols+1):
            cell = ws.cell(r, ci)
            cell.fill   = _fill(HDR_FILL)
            cell.border = _thin()
        _open_labels = {"en": "Current Open Items",
                        "nl": "Huidige Openstaande Posten",
                        "fr": "Postes Ouverts Actuels"}
        lbl = _open_labels.get(lang, "Current Open Items")
        ws.cell(r, 1).value     = lbl
        ws.cell(r, 1).font      = _font(bold=True, color=COL_WHT, size=12)
        ws.cell(r, 1).alignment = _aln("left")
        ws.row_dimensions[r].height = 22
        r += 1

        # Column headers
        _write_col_headers(ws, r); r += 1

        open_total = 0.0
        for grp in current_open_groups:
            grp_total = sum(
                float(row.get(amt_col, 0) or 0)
                for row in grp if amt_col and pd.notna(row.get(amt_col))
            )
            open_total += grp_total
            for row in grp:
                bg = ROW_WHITE if row_idx % 2 == 0 else ROW_BLUE
                _write_data_row(ws, r, row, bg, amt_ci)
                r += 1; row_idx += 1
            if abs(grp_total) < 0.02:
                _write_zero_subtotal(ws, r, amt_ci)
                r += 1; row_idx += 1

        grand_total += open_total
        _open_total_lbl = {"en": "Current Open — Total",
                           "nl": "Huidige Openstaand — Totaal",
                           "fr": "Postes Ouverts — Total"}
        _write_band_total(ws, r, _open_total_lbl.get(lang, "Current Open — Total"), open_total, amt_ci)
        r += 1
        # Gap
        ws.row_dimensions[r].height = 8; r += 1

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 2+ — ONE SECTION PER YEAR (newest first)
    # ══════════════════════════════════════════════════════════════════════════
    for yr in sorted(year_groups.keys(), reverse=True):
        groups = year_groups[yr]
        if not groups:
            continue

        yr_total = sum(
            float(row.get(amt_col, 0) or 0)
            for grp in groups for row in grp
            if amt_col and pd.notna(row.get(amt_col))
        )
        grand_total += yr_total

        # Dark-blue year banner
        for ci in range(1, ncols+1):
            cell = ws.cell(r, ci)
            cell.fill   = _fill(HDR_FILL)
            cell.border = _thin()
        ws.cell(r, 1).value     = str(yr)
        ws.cell(r, 1).font      = _font(bold=True, color=COL_WHT, size=12)
        ws.cell(r, 1).alignment = _aln("left")
        ws.row_dimensions[r].height = 22
        r += 1

        # Mid-blue column headers
        _write_col_headers(ws, r); r += 1

        # Data rows + yellow zero-subtotals
        row_idx = 0
        for grp in groups:
            grp_total = sum(
                float(row.get(amt_col, 0) or 0)
                for row in grp if amt_col and pd.notna(row.get(amt_col))
            )
            for row in grp:
                bg = ROW_WHITE if row_idx % 2 == 0 else ROW_BLUE
                _write_data_row(ws, r, row, bg, amt_ci)
                r += 1; row_idx += 1
            # Yellow separator row for every group that nets to zero
            if abs(grp_total) < 0.02:
                _write_zero_subtotal(ws, r, amt_ci)
                r += 1; row_idx += 1

        # Mid-blue year total
        yr_label = _t(lang, "year_total", yr=yr)
        _write_band_total(ws, r, yr_label, yr_total, amt_ci)
        r += 1

        # Small gap before next year
        ws.row_dimensions[r].height = 8; r += 1

    # ══════════════════════════════════════════════════════════════════════════
    # GRAND TOTAL — dark blue
    # ══════════════════════════════════════════════════════════════════════════
    for ci in range(1, ncols+1):
        cell = ws.cell(r, ci)
        cell.fill   = _fill(HDR_FILL)
        cell.border = _thin()
    ws.cell(r, 1).value     = _t(lang, "net_balance")
    ws.cell(r, 1).font      = _font(bold=True, color=COL_WHT, size=11)
    ws.cell(r, 1).alignment = _aln("left")
    if amt_ci:
        ws.cell(r, amt_ci).value         = grand_total
        ws.cell(r, amt_ci).font          = _font(bold=True, color=COL_WHT, size=12)
        ws.cell(r, amt_ci).number_format = "#,##0.00"
        ws.cell(r, amt_ci).alignment     = _aln("right")
    ws.row_dimensions[r].height = 22
    ws.freeze_panes = "A2"

    out = BytesIO()
    wb.save(out); out.seek(0)
    return out
