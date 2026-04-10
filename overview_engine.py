"""
Customer yearly overview engine.

Layout: ONE sheet, each year as a section with header + data rows + subtotal.
Data:   All transactions where Document Date falls within that year.
Input:  Same SAP FBL5N format as the Account Splitter (same column stripping).
Colours: Positive (invoices) = RED, Negative (credits) = GREEN.
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
LT_BLUE  = "D6E4F0"
WHITE    = "FFFFFF"
GREY     = "F2F2F2"
POS_FG   = "C00000"   # red  = invoices / positive
NEG_FG   = "375623"   # green = credits / negative
BLACK_FG = "000000"

# Same strip list as Account Splitter
STRIP_COLS = {
    "Reason code", "Clerk Abbreviation", "Cleared/open items symbol",
    "Case ID", "Status", "Dunning Block", "Disputed item", "Payment Block",
    "Payment Method", "Net due date symbol", "G/L Account", "Text",
    "Clearing date", "Clearing Document", "Dunning Level", "Last Dunned",
    "Reversed with", "Document Header Text", "User Name", "Special G/L ind.",
    "Billing Document", "Reference Key 1",
}


def _fill(rgb): return PatternFill("solid", fgColor=rgb)
def _font(bold=False, color=BLACK_FG, size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)
def _align(ha="left"):
    return Alignment(horizontal=ha, vertical="center")
def _thin():
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)


def _w(ws, row, col, val=None, bold=False, fill=WHITE, fg=BLACK_FG,
       size=10, ha="left", fmt=None, border=True):
    cell = ws.cell(row=row, column=col, value=val)
    cell.font      = _font(bold=bold, color=fg, size=size)
    cell.fill      = _fill(fill)
    cell.alignment = _align(ha)
    if fmt:   cell.number_format = fmt
    if border: cell.border = _thin()
    return cell


def _merge(ws, row, c1, c2, val=None, bold=False, fill=WHITE,
           fg=BLACK_FG, size=10, ha="center"):
    ws.merge_cells(
        f"{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}"
    )
    cell = ws.cell(row=row, column=c1, value=val)
    cell.font      = _font(bold=bold, color=fg, size=size)
    cell.fill      = _fill(fill)
    cell.alignment = _align(ha)
    for c in range(c1 + 1, c2 + 1):
        ws.cell(row=row, column=c).fill = _fill(fill)
    return cell


# ── PREPARE DATA ──────────────────────────────────────────────────────────────

def prepare_df(file_obj):
    """Load SAP export, strip unwanted columns, parse dates + amounts."""
    df = pd.read_excel(file_obj, sheet_name=0, header=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    # Strip unwanted columns
    df = df.drop(columns=[c for c in df.columns if c in STRIP_COLS])

    # Parse date columns
    for col in df.columns:
        if any(kw in col.lower() for kw in ["date", "datum"]):
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Parse amount
    amt_col = next(
        (c for c in df.columns
         if "amount" in c.lower() or "bedrag" in c.lower()
         or "betrag" in c.lower()), None
    )
    if amt_col:
        df[amt_col] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)

    # Remove SAP subtotal rows (no Document Number)
    doc_col = next(
        (c for c in df.columns
         if c.lower() in ("document number", "belegnummer")), None
    )
    if doc_col:
        df = df[
            df[doc_col].notna() &
            ~df[doc_col].astype(str).str.strip().isin(["", "nan", "0", "0.0"])
        ].copy()

    return df.reset_index(drop=True), amt_col


# ── MAIN BUILDER ──────────────────────────────────────────────────────────────

def build_overview(df: pd.DataFrame, amt_col: str,
                   year_from: int, year_to: int,
                   customer_name: str = "",
                   account_id: str    = "") -> BytesIO:
    """
    One-sheet overview: each year is a section (year banner → column headers
    → data rows → subtotal row → blank gap).
    Grand total at the very end.
    """
    years = list(range(year_from, year_to + 1))

    doc_date_col = next(
        (c for c in df.columns
         if "document date" in c.lower()
         or "belegdatum" in c.lower()
         or "boekingsdatum" in c.lower()), None
    )
    if not doc_date_col:
        doc_date_col = next(
            (c for c in df.columns
             if any(kw in c.lower() for kw in ["date","datum"])
             and pd.api.types.is_datetime64_any_dtype(df[c])), None
        )

    # Columns to display (same strip rules, drop SAP internals)
    display_cols = [c for c in df.columns if c not in STRIP_COLS
                    and not c.startswith("_")]

    # Identify date and amount column positions (1-based in display_cols)
    date_col_names = {
        c for c in display_cols
        if any(kw in c.lower() for kw in ["date","datum"])
        or (c in df.columns and pd.api.types.is_datetime64_any_dtype(df[c]))
    }
    amt_ci = (display_cols.index(amt_col) + 1) if amt_col and amt_col in display_cols else None
    ncols  = len(display_cols)

    # ── Create workbook ───────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Overview"

    # Auto column widths — set once based on data
    for ci, col in enumerate(display_cols, 1):
        sample = df[col].astype(str) if col in df.columns else pd.Series([""])
        max_len = max(len(col), sample.str.len().max() if len(sample) else 0)
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len + 2, 9), 36)

    # ── Main title ────────────────────────────────────────────────────────────
    r = 1
    title = "  ·  ".join(filter(None, [
        customer_name,
        f"Account {account_id}" if account_id else "",
        f"Overview {year_from}–{year_to}",
    ]))
    _merge(ws, r, 1, ncols, title,
           bold=True, fill=DK_BLUE, fg=WHITE, size=14)
    ws.row_dimensions[r].height = 36
    r += 1
    _merge(ws, r, 1, ncols,
           "All transactions by document date  ·  Positive = invoices (red)  ·  Negative = credits (green)",
           bold=False, fill=MD_BLUE, fg=WHITE, size=9)
    ws.row_dimensions[r].height = 16
    r += 2   # one blank row

    # ── Year sections ─────────────────────────────────────────────────────────
    grand_total     = 0.0
    grand_invoices  = 0.0
    grand_credits   = 0.0

    for yr in years:
        # Filter to this year by document date
        if doc_date_col and doc_date_col in df.columns:
            start = pd.Timestamp(yr, 1, 1)
            end   = pd.Timestamp(yr, 12, 31, 23, 59, 59)
            yr_df = df[
                df[doc_date_col].notna() &
                (df[doc_date_col] >= start) &
                (df[doc_date_col] <= end)
            ].copy()
        else:
            yr_df = df.copy()

        yr_total   = yr_df[amt_col].sum() if amt_col and amt_col in yr_df.columns else 0.0
        yr_inv     = yr_df[yr_df[amt_col] > 0][amt_col].sum() if amt_col and amt_col in yr_df.columns else 0.0
        yr_cred    = yr_df[yr_df[amt_col] < 0][amt_col].sum() if amt_col and amt_col in yr_df.columns else 0.0
        grand_total    += yr_total
        grand_invoices += yr_inv
        grand_credits  += yr_cred

        # ── Year banner ───────────────────────────────────────────────────────
        yr_fg = POS_FG if yr_total >= 0 else NEG_FG
        banner = (
            f"{yr}   ·   {len(yr_df)} transactions   ·   "
            f"Invoices: €{yr_inv:,.2f}   Credits: €{yr_cred:,.2f}   "
            f"Net: €{yr_total:,.2f}"
        )
        _merge(ws, r, 1, ncols, banner,
               bold=True, fill=DK_BLUE, fg=WHITE, size=11)
        ws.row_dimensions[r].height = 22
        r += 1

        if len(yr_df) == 0:
            _merge(ws, r, 1, ncols, "No transactions in this year",
                   bold=False, fill=GREY, fg=BLACK_FG, size=9)
            ws.row_dimensions[r].height = 16
            r += 2
            continue

        # ── Column headers ────────────────────────────────────────────────────
        for ci, col in enumerate(display_cols, 1):
            cell = ws.cell(row=r, column=ci, value=col)
            cell.font      = _font(bold=True, color=WHITE, size=9)
            cell.fill      = _fill(MD_BLUE)
            cell.alignment = _align("center")
            cell.border    = _thin()
        ws.row_dimensions[r].height = 15
        r += 1

        # ── Data rows ─────────────────────────────────────────────────────────
        yr_df_sorted = (
            yr_df.sort_values(doc_date_col)
            if doc_date_col and doc_date_col in yr_df.columns
            else yr_df
        )

        for ri, (_, row_data) in enumerate(yr_df_sorted.iterrows()):
            row_fill = GREY if ri % 2 == 0 else WHITE
            for ci, col in enumerate(display_cols, 1):
                val     = row_data.get(col, "")
                is_amt  = (ci == amt_ci)
                is_date = (col in date_col_names)

                if is_amt:
                    cell_val = float(val) if pd.notna(val) else 0.0
                    fg = POS_FG if cell_val >= 0 else NEG_FG
                elif is_date:
                    try:
                        cell_val = (pd.Timestamp(val).to_pydatetime()
                                    if pd.notna(val) else "")
                    except Exception:
                        cell_val = str(val) if pd.notna(val) else ""
                    fg = BLACK_FG
                elif pd.isna(val):
                    cell_val = ""
                    fg = BLACK_FG
                elif isinstance(val, float) and val == int(val):
                    cell_val = int(val)
                    fg = BLACK_FG
                else:
                    cell_val = val
                    fg = BLACK_FG

                cell = ws.cell(row=r, column=ci, value=cell_val)
                cell.font      = _font(color=fg, size=9)
                cell.fill      = _fill(row_fill)
                cell.alignment = _align("right" if is_amt else "left")
                cell.border    = _thin()
                if is_amt:
                    cell.number_format = "#,##0.00"
                elif is_date and isinstance(cell_val, datetime.datetime):
                    cell.number_format = "DD/MM/YYYY"

            ws.row_dimensions[r].height = 13
            r += 1

        # ── Subtotal row ──────────────────────────────────────────────────────
        for ci in range(1, ncols + 1):
            if ci == 1:
                _w(ws, r, ci, f"{yr} TOTAL",
                   bold=True, fill=DK_BLUE, fg=WHITE, size=10)
            elif ci == amt_ci:
                fg = POS_FG if yr_total >= 0 else NEG_FG
                cell = ws.cell(row=r, column=ci, value=yr_total)
                cell.font          = _font(bold=True, color=fg, size=10)
                cell.fill          = _fill(DK_BLUE)
                cell.alignment     = _align("right")
                cell.number_format = "#,##0.00"
                cell.border        = _thin()
            else:
                ws.cell(row=r, column=ci).fill   = _fill(DK_BLUE)
                ws.cell(row=r, column=ci).border = _thin()
        ws.row_dimensions[r].height = 18
        r += 1

        # ── 2 blank rows between years ────────────────────────────────────────
        ws.row_dimensions[r].height = 8
        r += 1
        ws.row_dimensions[r].height = 8
        r += 1

    # ── Grand total ───────────────────────────────────────────────────────────
    gt_fg    = POS_FG if grand_total >= 0 else NEG_FG
    gt_label = (
        f"GRAND TOTAL {year_from}–{year_to}   ·   "
        f"Invoices: €{grand_invoices:,.2f}   "
        f"Credits: €{grand_credits:,.2f}   "
        f"Net: €{grand_total:,.2f}"
    )
    _merge(ws, r, 1, ncols, gt_label,
           bold=True, fill=DK_BLUE, fg=WHITE, size=12)
    ws.row_dimensions[r].height = 28
    r += 1

    # Grand total amount cell (standalone, large)
    for ci in range(1, ncols + 1):
        if ci == 1:
            _w(ws, r, ci, "NET BALANCE",
               bold=True, fill=DK_BLUE, fg=WHITE, size=11)
        elif ci == amt_ci:
            cell = ws.cell(row=r, column=ci, value=grand_total)
            cell.font          = _font(bold=True, color=gt_fg, size=13)
            cell.fill          = _fill(DK_BLUE)
            cell.alignment     = _align("right")
            cell.number_format = "#,##0.00"
        else:
            ws.cell(row=r, column=ci).fill = _fill(DK_BLUE)
    ws.row_dimensions[r].height = 24

    ws.freeze_panes = "A4"

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out
