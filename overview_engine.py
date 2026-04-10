"""
Customer yearly overview engine.

- Each year section contains ALL relevant rows: activity in that year PLUS
  payments from year+1 that are linked to that year's invoices (via Assignment).
- Within each year: rows are split by G/L Account (Beer / Rent / Other),
  each sub-group gets its own subtotal.
- Document Type column REPLACED with human-readable description.
- Payment Method and G/L Account columns kept.
- EN / NL / FR language support throughout.
- Positive = RED (invoices), Negative = GREEN (credits/payments).
"""

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime
import warnings
warnings.filterwarnings("ignore")

DK_BLUE  = "1F3864"
MD_BLUE  = "2E75B6"
LT_BLUE  = "D6E4F0"
WHITE    = "FFFFFF"
GREY     = "F2F2F2"
POS_FG   = "C00000"
NEG_FG   = "375623"
BLACK_FG = "000000"

STRIP_COLS = {
    "Reason code","Clerk Abbreviation","Cleared/open items symbol",
    "Case ID","Status","Dunning Block","Disputed item","Payment Block",
    "Net due date symbol","Text","Clearing date","Clearing Document",
    "Dunning Level","Last Dunned","Reversed with","Document Header Text",
    "User Name","Special G/L ind.","Billing Document","Reference Key 1",
    # G/L Account and Payment Method intentionally NOT stripped
}

GL_LABELS = {
    "en": {"2400000":"Beer",  "2530009":"Rent"},
    "nl": {"2400000":"Bier",  "2530009":"Huur"},
    "fr": {"2400000":"Bière", "2530009":"Loyer"},
}

T = {
    "en": {
        "title_suffix":     "Customer Overview",
        "subtitle":         "All transactions by document date  ·  Positive = invoices (red)  ·  Negative = credits / payments (green)",
        "year_banner":      "{yr}  ·  {n} transactions  ·  Invoices: €{inv}  ·  Credits: €{cred}  ·  Net: €{net}",
        "year_total":       "{yr} — Total",
        "grand_total":      "Grand Total {from_yr}–{to_yr}",
        "net_balance":      "Net Balance",
        "no_transactions":  "No transactions in {yr}",
        "gl_other":         "Other",
        "gl_subtotal":      "{lbl} — Subtotal",
        "transactions":     "transactions",
        "doc_types": {
            "RV+":"Invoice", "RV-":"Credit note",
            "ZP":"Payment",  "DZ":"Payment",
            "RS+":"Re-invoice (bonus correction)", "RS-":"Bonus",
            "AB":"Clearing", "X_PAY":"Payout to customer",
        },
    },
    "nl": {
        "title_suffix":     "Klantoverzicht",
        "subtitle":         "Alle transacties op boekingsdatum  ·  Positief = facturen (rood)  ·  Negatief = creditnota's / betalingen (groen)",
        "year_banner":      "{yr}  ·  {n} transacties  ·  Facturen: €{inv}  ·  Creditnota's: €{cred}  ·  Netto: €{net}",
        "year_total":       "{yr} — Totaal",
        "grand_total":      "Eindtotaal {from_yr}–{to_yr}",
        "net_balance":      "Nettosaldo",
        "no_transactions":  "Geen transacties in {yr}",
        "gl_other":         "Overig",
        "gl_subtotal":      "{lbl} — Subtotaal",
        "transactions":     "transacties",
        "doc_types": {
            "RV+":"Factuur", "RV-":"Creditnota",
            "ZP":"Betaling",  "DZ":"Betaling",
            "RS+":"Refactuur (bonuscorrectie)", "RS-":"Bonus",
            "AB":"Verrekening", "X_PAY":"Uitbetaling aan klant",
        },
    },
    "fr": {
        "title_suffix":     "Aperçu client",
        "subtitle":         "Toutes les transactions par date comptable  ·  Positif = factures (rouge)  ·  Négatif = avoirs / paiements (vert)",
        "year_banner":      "{yr}  ·  {n} transactions  ·  Factures: €{inv}  ·  Avoirs: €{cred}  ·  Net: €{net}",
        "year_total":       "{yr} — Total",
        "grand_total":      "Total général {from_yr}–{to_yr}",
        "net_balance":      "Solde net",
        "no_transactions":  "Aucune transaction en {yr}",
        "gl_other":         "Autre",
        "gl_subtotal":      "{lbl} — Sous-total",
        "transactions":     "transactions",
        "doc_types": {
            "RV+":"Facture", "RV-":"Avoir",
            "ZP":"Paiement",  "DZ":"Paiement",
            "RS+":"Re-facturation (correction bonus)", "RS-":"Bonus",
            "AB":"Compensation", "X_PAY":"Paiement au client",
        },
    },
}


def _t(lang, key, **kw):
    val = T.get(lang, T["en"]).get(key, T["en"].get(key, key))
    try:
        return val.format(**kw) if kw else val
    except Exception:
        return val


def _doc_desc(doc_type, amount, pay_method, lang):
    dt  = str(doc_type or "").strip().upper()
    amt = float(amount) if str(amount) not in ("","nan") else 0.0
    pm  = str(pay_method or "").strip().upper()
    d   = T.get(lang, T["en"])["doc_types"]
    if pm == "X":
        return d["X_PAY"]
    if dt == "RV":
        return d["RV+"] if amt >= 0 else d["RV-"]
    if dt in ("ZP","DZ"):
        return d["ZP"]
    if dt == "RS":
        return d["RS+"] if amt >= 0 else d["RS-"]
    if dt == "AB":
        return d["AB"]
    return d.get(dt, "")


def _gl_label(gl_val, lang):
    g = str(gl_val or "").strip().split(".")[0]
    return GL_LABELS.get(lang, GL_LABELS["en"]).get(g, _t(lang,"gl_other") if g else "")


def _fill(rgb): return PatternFill("solid", fgColor=rgb)
def _font(bold=False, color=BLACK_FG, size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)
def _align(ha="left"):
    return Alignment(horizontal=ha, vertical="center")
def _thin():
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)

def _w(ws, row, col, val=None, bold=False, fill=WHITE, fg=BLACK_FG,
       size=10, ha="left", fmt=None):
    cell = ws.cell(row=row, column=col, value=val)
    cell.font      = _font(bold=bold, color=fg, size=size)
    cell.fill      = _fill(fill)
    cell.alignment = _align(ha)
    cell.border    = _thin()
    if fmt: cell.number_format = fmt
    return cell

def _merge(ws, row, c1, c2, val=None, bold=False, fill=WHITE,
           fg=BLACK_FG, size=10, ha="center"):
    ws.merge_cells(f"{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}")
    cell = ws.cell(row=row, column=c1, value=val)
    cell.font      = _font(bold=bold, color=fg, size=size)
    cell.fill      = _fill(fill)
    cell.alignment = _align(ha)
    for c in range(c1+1, c2+1):
        ws.cell(row=row, column=c).fill   = _fill(fill)
        ws.cell(row=row, column=c).border = _thin()
    cell.border = _thin()
    return cell


def prepare_df(file_obj):
    df = pd.read_excel(file_obj, sheet_name=0, header=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.drop(columns=[c for c in df.columns if c in STRIP_COLS])

    for col in df.columns:
        if any(kw in col.lower() for kw in ["date","datum"]):
            df[col] = pd.to_datetime(df[col], errors="coerce")

    amt_col = next((c for c in df.columns
                    if "amount" in c.lower() or "bedrag" in c.lower()
                    or "betrag" in c.lower()), None)
    if amt_col:
        df[amt_col] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)

    doc_col = next((c for c in df.columns
                    if c.lower() in ("document number","belegnummer")), None)
    if doc_col:
        df = df[df[doc_col].notna() &
                ~df[doc_col].astype(str).str.strip().isin(["","nan","0","0.0"])].copy()

    return df.reset_index(drop=True), amt_col


def build_overview(df: pd.DataFrame, amt_col: str,
                   year_from: int, year_to: int,
                   customer_name: str = "",
                   account_id: str    = "",
                   lang: str          = "en") -> BytesIO:

    years = list(range(year_from, year_to + 1))

    doc_date_col = next((c for c in df.columns
                         if "document date" in c.lower()
                         or "belegdatum" in c.lower()
                         or "boekingsdatum" in c.lower()), None)
    doc_num_col  = next((c for c in df.columns
                         if c.lower() in ("document number","belegnummer")), None)
    doc_type_col = next((c for c in df.columns
                         if "document type" in c.lower()
                         or "belegtyp" in c.lower()
                         or "boekingssoort" in c.lower()), None)
    pay_meth_col = next((c for c in df.columns
                         if "payment method" in c.lower()), None)
    gl_col       = next((c for c in df.columns
                         if "g/l account" in c.lower()
                         or "sachkonto" in c.lower()), None)
    assign_col   = next((c for c in df.columns
                         if "assignment" in c.lower()
                         or "zuordnung" in c.lower()), None)

    # Build display columns — REPLACE Document Type with description
    base_cols = [c for c in df.columns
                 if c not in STRIP_COLS and not c.startswith("_")]
    display_cols = []
    for c in base_cols:
        if c == doc_type_col:
            display_cols.append("_DESC_")   # placeholder replaced at write time
        else:
            display_cols.append(c)

    # Column header names shown in Excel
    def col_header(c, lang):
        if c == "_DESC_":
            return _t(lang, "title_suffix").split()[0]  # unused, handled specially
        return c

    ncols = len(display_cols)

    date_col_names = {c for c in base_cols
                      if any(kw in c.lower() for kw in ["date","datum"])
                      or (c in df.columns
                          and pd.api.types.is_datetime64_any_dtype(df[c]))}
    # amt_ci: position of amount col in display_cols (1-based)
    amt_ci = (display_cols.index(amt_col)+1) if amt_col and amt_col in display_cols else None

    # ── Workbook ──────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overview"

    for ci, col in enumerate(display_cols, 1):
        if col == "_DESC_":
            ws.column_dimensions[get_column_letter(ci)].width = 30
        elif col in date_col_names:
            ws.column_dimensions[get_column_letter(ci)].width = 13
        elif amt_col and col == amt_col:
            ws.column_dimensions[get_column_letter(ci)].width = 18
        else:
            sample = df[col].astype(str) if col in df.columns else pd.Series([""])
            w = max(len(col), sample.str.len().max() if len(sample) else 0)
            ws.column_dimensions[get_column_letter(ci)].width = min(max(w+2,9),30)

    # ── Title ─────────────────────────────────────────────────────────────────
    r = 1
    title = "  ·  ".join(filter(None,[
        customer_name,
        f"Account {account_id}" if account_id else "",
        f"{year_from}–{year_to}  {_t(lang,'title_suffix')}",
    ]))
    _merge(ws,r,1,ncols,title,bold=True,fill=DK_BLUE,fg=WHITE,size=14)
    ws.row_dimensions[r].height = 36
    r += 1
    _merge(ws,r,1,ncols,_t(lang,"subtitle"),fill=MD_BLUE,fg=WHITE,size=9)
    ws.row_dimensions[r].height = 16
    r += 2

    # ── Helper: write rows with G/L sub-grouping ───────────────────────────
    def write_gl_sections(rows_df, r):
        """Write rows grouped by G/L Account. Returns next row."""
        if len(rows_df) == 0:
            return r

        if gl_col and gl_col in rows_df.columns:
            gl_vals  = rows_df[gl_col].astype(str).str.strip().str.split(".").str[0]
            known    = [g for g in ["2400000","2530009"] if g in gl_vals.unique()]
            others   = sorted(g for g in gl_vals.unique()
                               if g not in known and g not in ("","nan","None"))
            no_gl    = [""] if gl_vals.isin(["","nan","None"]).any() else []
            groups   = known + others + no_gl
            multi    = len([g for g in groups if g != ""]) > 1
        else:
            groups, multi = [None], False

        for gl_grp in groups:
            if gl_grp is None:
                grp_df = rows_df
                gl_lbl = None
            elif gl_grp == "":
                grp_df = rows_df[gl_vals.isin(["","nan","None"])]
                gl_lbl = _t(lang,"gl_other") if multi else None
            else:
                grp_df = rows_df[gl_vals == gl_grp]
                raw_lbl = _gl_label(gl_grp, lang)
                gl_lbl  = f"{gl_grp} — {raw_lbl}" if raw_lbl else gl_grp

            if len(grp_df) == 0:
                continue

            # G/L sub-header
            if multi and gl_lbl:
                gl_total = grp_df[amt_col].sum() if amt_col and amt_col in grp_df.columns else 0
                sub_lbl  = f"  {gl_lbl}"
                _merge(ws,r,1,ncols,sub_lbl,bold=True,fill=LT_BLUE,fg=DK_BLUE,size=9,ha="left")
                ws.row_dimensions[r].height = 15
                r += 1

            # Column headers
            for ci, col in enumerate(display_cols, 1):
                hdr = _t(lang,"title_suffix").split()[0] if col == "_DESC_" else col
                # Use "Description" / "Omschrijving" / "Description" as header
                if col == "_DESC_":
                    hdr = {"en":"Description","nl":"Omschrijving","fr":"Description"}.get(lang,"Description")
                cell = ws.cell(row=r, column=ci, value=hdr)
                cell.font      = _font(bold=True,color=WHITE,size=9)
                cell.fill      = _fill(MD_BLUE)
                cell.alignment = _align("center")
                cell.border    = _thin()
            ws.row_dimensions[r].height = 15
            r += 1

            # Data rows
            sort_c = doc_date_col if doc_date_col and doc_date_col in grp_df.columns else None
            sdf    = grp_df.sort_values(sort_c) if sort_c else grp_df

            for ri, (_, row_data) in enumerate(sdf.iterrows()):
                row_fill = GREY if ri%2==0 else WHITE
                for ci, col in enumerate(display_cols, 1):
                    is_amt  = (ci == amt_ci)
                    is_date = (col in date_col_names)
                    is_desc = (col == "_DESC_")

                    if is_desc:
                        cell_val = _doc_desc(
                            row_data.get(doc_type_col,"") if doc_type_col else "",
                            row_data.get(amt_col,0)       if amt_col else 0,
                            row_data.get(pay_meth_col,"") if pay_meth_col else "",
                            lang,
                        )
                        fg = BLACK_FG
                    elif col not in df.columns:
                        cell_val, fg = "", BLACK_FG
                    elif is_amt:
                        cell_val = float(row_data[col]) if pd.notna(row_data[col]) else 0.0
                        fg = POS_FG if cell_val >= 0 else NEG_FG
                    elif is_date:
                        try:
                            cell_val = (pd.Timestamp(row_data[col]).to_pydatetime()
                                        if pd.notna(row_data[col]) else "")
                        except Exception:
                            cell_val = str(row_data[col]) if pd.notna(row_data[col]) else ""
                        fg = BLACK_FG
                    elif pd.isna(row_data[col]):
                        cell_val, fg = "", BLACK_FG
                    elif isinstance(row_data[col],float) and row_data[col]==int(row_data[col]):
                        cell_val, fg = int(row_data[col]), BLACK_FG
                    else:
                        cell_val, fg = row_data[col], BLACK_FG

                    cell = ws.cell(row=r, column=ci, value=cell_val)
                    cell.font      = _font(color=fg,size=9)
                    cell.fill      = _fill(row_fill)
                    cell.alignment = _align("right" if is_amt else "left")
                    cell.border    = _thin()
                    if is_amt:
                        cell.number_format = "#,##0.00"
                    elif is_date and isinstance(cell_val,datetime.datetime):
                        cell.number_format = "DD/MM/YYYY"
                ws.row_dimensions[r].height = 13
                r += 1

            # G/L subtotal (only if multiple groups)
            if multi and gl_lbl:
                gl_total = grp_df[amt_col].sum() if amt_col and amt_col in grp_df.columns else 0
                lbl_text = _t(lang,"gl_subtotal",lbl=gl_lbl)
                for ci in range(1, ncols+1):
                    if ci == 1:
                        _w(ws,r,ci,lbl_text,bold=True,fill=LT_BLUE,fg=DK_BLUE,size=9)
                    elif ci == amt_ci:
                        fg2 = POS_FG if gl_total>=0 else NEG_FG
                        c2  = ws.cell(row=r,column=ci,value=gl_total)
                        c2.font=_font(bold=True,color=fg2,size=9); c2.fill=_fill(LT_BLUE)
                        c2.alignment=_align("right"); c2.number_format="#,##0.00"
                        c2.border=_thin()
                    else:
                        ws.cell(row=r,column=ci).fill=_fill(LT_BLUE)
                        ws.cell(row=r,column=ci).border=_thin()
                ws.row_dimensions[r].height = 14
                r += 1

        return r

    # ── Pre-compute cross-year pairs ─────────────────────────────────────────
    # Invoice in year Y paid in year Y+1 → pull BOTH into year Y+1 section.
    # This prevents invoice showing in Y and payment in Y+1 as a duplicate.
    pulled_indices = set()   # original invoice row indices pulled out of their year
    pulled_into    = {}      # yr -> [extra DataFrames to prepend]

    if doc_date_col and doc_num_col and assign_col and doc_type_col:
        for yi in range(len(years)-1):
            yr_pull  = years[yi]
            next_yr  = years[yi+1]
            s  = pd.Timestamp(yr_pull,1,1);  e  = pd.Timestamp(yr_pull,12,31,23,59,59)
            ns = pd.Timestamp(next_yr,1,1);  ne = pd.Timestamp(next_yr,12,31,23,59,59)

            yr_rows  = df[df[doc_date_col].notna() & (df[doc_date_col]>=s)  & (df[doc_date_col]<=e)]
            nxt_rows = df[df[doc_date_col].notna() & (df[doc_date_col]>=ns) & (df[doc_date_col]<=ne)]

            pay_types = {"ZP","DZ","AB"}
            nxt_pays  = nxt_rows[nxt_rows[doc_type_col].astype(str).str.strip().str.upper().isin(pay_types)]
            if len(nxt_pays) == 0:
                continue

            yr_doc_nums   = set(yr_rows[doc_num_col].astype(str).str.strip().tolist())
            linked_pays   = nxt_pays[nxt_pays[assign_col].astype(str).str.strip().isin(yr_doc_nums)]
            if len(linked_pays) == 0:
                continue

            matched_nums  = set(linked_pays[assign_col].astype(str).str.strip().tolist())
            orig_invoices = yr_rows[yr_rows[doc_num_col].astype(str).str.strip().isin(matched_nums)]

            for idx in orig_invoices.index:
                pulled_indices.add(idx)
            # Also exclude the linked payments from base_df of next_yr
            # since we add them explicitly via pulled_into
            for idx in linked_pays.index:
                pulled_indices.add(idx)
            pulled_into.setdefault(next_yr, [])
            pulled_into[next_yr].append(orig_invoices)
            pulled_into[next_yr].append(linked_pays)

    # ── Year sections ─────────────────────────────────────────────────────────
    grand_total = 0.0

    for yr in years:
        if doc_date_col and doc_date_col in df.columns:
            s = pd.Timestamp(yr,1,1); e = pd.Timestamp(yr,12,31,23,59,59)
            base_df = df[df[doc_date_col].notna() &
                         (df[doc_date_col]>=s) & (df[doc_date_col]<=e)].copy()
        else:
            base_df = df.copy()

        # Remove invoices pulled into a later year
        base_df = base_df[~base_df.index.isin(pulled_indices)].copy()

        # Add invoices+payments pulled INTO this year from previous year
        extras = pulled_into.get(yr, [])
        combined_df = pd.concat([base_df]+extras, ignore_index=True) if extras else base_df

        yr_total  = combined_df[amt_col].sum() if amt_col and amt_col in combined_df.columns else 0
        yr_inv    = combined_df[combined_df[amt_col]>0][amt_col].sum() if amt_col and amt_col in combined_df.columns else 0
        yr_cred   = combined_df[combined_df[amt_col]<0][amt_col].sum() if amt_col and amt_col in combined_df.columns else 0
        grand_total += yr_total

        # ── Year banner ───────────────────────────────────────────────────────
        banner = _t(lang,"year_banner",
                    yr=yr, n=len(combined_df),
                    inv=f"{yr_inv:,.2f}", cred=f"{yr_cred:,.2f}",
                    net=f"{yr_total:,.2f}")
        _merge(ws,r,1,ncols,banner,bold=True,fill=DK_BLUE,fg=WHITE,size=11)
        ws.row_dimensions[r].height = 22
        r += 1

        if len(combined_df) == 0:
            _merge(ws,r,1,ncols,_t(lang,"no_transactions",yr=yr),
                   fill=GREY,fg=BLACK_FG,size=9)
            ws.row_dimensions[r].height = 16
            r += 2
            continue

        # ── Write rows with G/L grouping ──────────────────────────────────────
        r = write_gl_sections(combined_df, r)

        # ── Year total row ────────────────────────────────────────────────────
        for ci in range(1, ncols+1):
            if ci == 1:
                _w(ws,r,ci,_t(lang,"year_total",yr=yr),
                   bold=True,fill=DK_BLUE,fg=WHITE,size=10)
            elif ci == amt_ci:
                fg = POS_FG if yr_total>=0 else NEG_FG
                cell = ws.cell(row=r,column=ci,value=yr_total)
                cell.font=_font(bold=True,color=fg,size=10); cell.fill=_fill(DK_BLUE)
                cell.alignment=_align("right"); cell.number_format="#,##0.00"
                cell.border=_thin()
            else:
                ws.cell(row=r,column=ci).fill=_fill(DK_BLUE)
                ws.cell(row=r,column=ci).border=_thin()
        ws.row_dimensions[r].height = 18
        r += 1

        # Gap
        ws.row_dimensions[r].height = 10; r += 1
        ws.row_dimensions[r].height = 10; r += 1

    # ── Grand total ───────────────────────────────────────────────────────────
    _merge(ws,r,1,ncols,_t(lang,"grand_total",from_yr=year_from,to_yr=year_to),
           bold=True,fill=DK_BLUE,fg=WHITE,size=12)
    ws.row_dimensions[r].height = 28; r += 1

    gt_fg = POS_FG if grand_total>=0 else NEG_FG
    for ci in range(1, ncols+1):
        if ci == 1:
            _w(ws,r,ci,_t(lang,"net_balance"),bold=True,fill=DK_BLUE,fg=WHITE,size=11)
        elif ci == amt_ci:
            cell = ws.cell(row=r,column=ci,value=grand_total)
            cell.font=_font(bold=True,color=gt_fg,size=13); cell.fill=_fill(DK_BLUE)
            cell.alignment=_align("right"); cell.number_format="#,##0.00"; cell.border=_thin()
        else:
            ws.cell(row=r,column=ci).fill=_fill(DK_BLUE)
            ws.cell(row=r,column=ci).border=_thin()
    ws.row_dimensions[r].height = 26

    ws.freeze_panes = "A4"
    out = BytesIO(); wb.save(out); out.seek(0)
    return out
