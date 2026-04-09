"""
POC-grouped Excel builder for accounts like NEGOBOISSONS.

Structure per group:
  - Pink header row: [POC name | col headers] with POC number in col A
  - First data row: pink + bold (same fill as header)
  - Remaining data rows: plain white/grey alternating
  - Subtotal row: amount only in col I, all other cols empty
  - 2 blank rows separator

Customer block at top:
  - Row 1: account number (col C)
  - Row 2: customer name (col C)
  - Row 3: address line 1 (col C)
  - Row 4: address line 2 (col C)

Grand total: "Grand total" label in col K, value in col M, positioned at row 8.

POC grouping: rows are grouped by Reference Key 3 value.
Only POC numbers starting with "29" are treated as groups.
Rows with no matching POC (e.g. blank ref key) are placed in a catch-all group at the end.
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
PINK_FILL  = "FDB9FD"   # header rows
WHITE_FILL = "FFFFFF"
GREY_FILL  = "F2F2F2"
GREEN_FG   = "166534"
RED_FG     = "B91C1C"
BLACK_FG   = "000000"


def _fill(rgb):
    return PatternFill("solid", fgColor=rgb)


def _font(bold=False, color=BLACK_FG, size=9, italic=False):
    return Font(name="Calibri", bold=bold, color=color, size=size, italic=italic)


def _align(ha="left"):
    return Alignment(horizontal=ha, vertical="center")


def _thin_border():
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)


def _write(ws, row, col, value, bold=False, fill_rgb=WHITE_FILL,
           fg=BLACK_FG, size=9, ha="left", fmt=None, italic=False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = _font(bold=bold, color=fg, size=size, italic=italic)
    cell.fill      = _fill(fill_rgb)
    cell.alignment = _align(ha)
    if fmt:
        cell.number_format = fmt
    return cell


# ── POC NAME MAP ──────────────────────────────────────────────────────────────
# Read POC names from the uploaded template if available,
# otherwise fall back to the POC number itself.

def _load_poc_names(template_bytes: bytes) -> dict:
    """
    Extract POC_number -> POC_name mapping from the template.
    The name is the text in col A of the header row (e.g. "FOUREZ ETS.").
    """
    names = {}
    try:
        wb = openpyxl.load_workbook(BytesIO(template_bytes))
        ws = wb.active
        for r in range(1, (ws.max_row or 50) + 1):
            col_a = ws.cell(r, 1).value
            if col_a and str(col_a).startswith("29"):
                poc_num = str(col_a).strip()
                # Name is the text in same cell's row but further left — 
                # actually it's on the row ABOVE (the group header row before data)
                # From analysis: header row has name in col A, POC number also in col A of NEXT row
                # Re-check: row 6 has name "FOUREZ ETS." and row 7 has POC 29060006
                # So name_row = data_start_row - 1? No...
                # Actually: row 6 = header (col A = "FOUREZ ETS.", col B = "Account", ...)
                #           row 7 = first data row (col A = 29060006, col B = account, ...)
                # The header row (pink) that precedes this POC has the name in col A
                # Let's look for pink rows above
                # Simpler: just map poc_num -> poc_num, let caller supply names
                names[poc_num] = poc_num
        # Now find actual names: header rows are pink+bold with 'Account' in col B
        for r in range(1, (ws.max_row or 50) + 1):
            col_b = ws.cell(r, 2).value
            col_a = ws.cell(r, 1).value
            fill  = ws.cell(r, 1).fill
            # Header row: col B = 'Account', col A = name text (not a 29-number)
            if col_b == "Account" and col_a and not str(col_a).startswith("29"):
                name = str(col_a).strip()
                # The POC number is in the NEXT row's col A
                poc_val = ws.cell(r + 1, 1).value
                if poc_val and str(poc_val).startswith("29"):
                    names[str(poc_val).strip()] = name
    except Exception:
        pass
    return names


# ── MAIN BUILDER ──────────────────────────────────────────────────────────────

def build_poc_sheet(acc_df: pd.DataFrame, account_id: str,
                    template_bytes: bytes = None, today=None) -> bytes:
    """
    Build a POC-grouped Excel sheet for a NEGOBOISSONS-style account.

    Groups rows by Reference Key 3 (values starting with "29").
    Produces: customer header → per-POC section (header+data+subtotal+gap) → grand total.
    """
    if today is None:
        today = datetime.date.today()
    today_str = pd.Timestamp(today).strftime("%d/%m/%Y")

    # Load POC names from template
    poc_names = {}
    if template_bytes:
        poc_names = _load_poc_names(template_bytes)

    amount_col = next(
        (c for c in acc_df.columns
         if "amount" in c.lower() or "bedrag" in c.lower()), None
    )
    ref_col = next(
        (c for c in acc_df.columns
         if "reference key 3" in c.lower() or "ref" in c.lower()
         and "key" in c.lower()), None
    )

    if amount_col:
        acc_df = acc_df.copy()
        acc_df[amount_col] = pd.to_numeric(acc_df[amount_col], errors="coerce").fillna(0)

    # ── Identify date columns ──────────────────────────────────────────────────
    date_cols = {
        ci for ci, col in enumerate(acc_df.columns, 1)
        if any(kw in col.lower() for kw in ["date", "datum"])
        or pd.api.types.is_datetime64_any_dtype(acc_df[col])
    }
    amount_ci = (
        list(acc_df.columns).index(amount_col) + 1 if amount_col else None
    )

    # ── Group by POC (Reference Key 3 starting with "29") ─────────────────────
    if ref_col and ref_col in acc_df.columns:
        acc_df["_ref3"] = acc_df[ref_col].astype(str).str.strip()
    else:
        # Fallback: try to find a column with 29xxxxx values
        acc_df["_ref3"] = ""
        for col in acc_df.columns:
            sample = acc_df[col].astype(str).str.strip()
            if sample.str.startswith("29").sum() > len(acc_df) * 0.3:
                acc_df["_ref3"] = sample
                break

    poc_groups = {}
    no_poc_rows = []

    for _, row in acc_df.iterrows():
        ref = str(row["_ref3"]).strip()
        if ref.startswith("29") and len(ref) >= 7:
            poc_groups.setdefault(ref, []).append(row)
        else:
            no_poc_rows.append(row)

    # Sort POC groups by POC number
    sorted_pocs = sorted(poc_groups.keys())

    # ── Build workbook ─────────────────────────────────────────────────────────
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = str(account_id)[:31]

    # Column widths matching template
    col_widths = [14, 10, 14, 14, 12, 13, 13, 8, 16, 2, 12, 2, 16]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Customer header block (rows 1-4) ──────────────────────────────────────
    _write(ws, 1, 3, account_id, bold=True, size=10)
    _write(ws, 2, 3, "NEGOBOISSONS", bold=False, size=10)
    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 14
    ws.row_dimensions[4].height = 14

    # Grand total placeholder — will be filled at end (row 8, cols K and M)
    _write(ws, 8, 11, "Grand total", bold=True, size=10)

    # ── Data columns header list ───────────────────────────────────────────────
    display_cols = [c for c in acc_df.columns if not c.startswith("_")]
    ncols = len(display_cols)

    # ── Write each POC group ───────────────────────────────────────────────────
    r = 6  # start at row 6 matching template
    subtotal_rows = []
    grand_total = 0.0

    for poc in sorted_pocs:
        rows = poc_groups[poc]
        poc_name = poc_names.get(poc, "")

        # ── Pink header row ────────────────────────────────────────────────────
        # Col A: POC name (e.g. "FOUREZ ETS."), then column headers
        _write(ws, r, 1, poc_name or poc, bold=True,
               fill_rgb=PINK_FILL, size=9)
        for ci, col_name in enumerate(display_cols, 2):
            _write(ws, r, ci, col_name, bold=True, fill_rgb=PINK_FILL,
                   size=9, ha="center")
        ws.row_dimensions[r].height = 15
        r += 1

        # ── Data rows ─────────────────────────────────────────────────────────
        poc_total = 0.0
        for row_idx, data_row in enumerate(rows):
            is_first = (row_idx == 0)
            row_fill = PINK_FILL if is_first else (GREY_FILL if row_idx % 2 == 0 else WHITE_FILL)
            bold_row = is_first

            # Col A: POC number on first row, blank thereafter
            _write(ws, r, 1,
                   int(poc) if is_first else None,
                   bold=bold_row, fill_rgb=row_fill, size=9)

            for ci, col_name in enumerate(display_cols, 2):
                val  = data_row.get(col_name, "")
                ci_actual = ci  # col B onwards
                col_idx   = ci_actual - 1  # 0-based index into display_cols

                is_amt  = (col_name == amount_col)
                is_date = (col_idx + 1) in date_cols  # date_cols is 1-based from acc_df

                if is_amt:
                    cell_val = float(val) if pd.notna(val) else 0.0
                    poc_total += cell_val
                elif is_date:
                    try:
                        cell_val = pd.Timestamp(val).to_pydatetime() if pd.notna(val) else ""
                    except Exception:
                        cell_val = str(val) if pd.notna(val) else ""
                elif pd.isna(val):
                    cell_val = ""
                elif isinstance(val, float) and val == int(val):
                    cell_val = int(val)
                else:
                    cell_val = val

                fg = (GREEN_FG if is_amt and isinstance(cell_val, (int, float)) and cell_val >= 0
                      else RED_FG if is_amt and isinstance(cell_val, (int, float)) and cell_val < 0
                      else BLACK_FG)

                cell = ws.cell(row=r, column=ci_actual, value=cell_val)
                cell.font      = _font(bold=bold_row, color=fg, size=9)
                cell.fill      = _fill(row_fill)
                cell.alignment = _align("right" if is_amt else "left")
                if is_amt:
                    cell.number_format = "#,##0.00"
                elif is_date and isinstance(cell_val, datetime.datetime):
                    cell.number_format = "DD/MM/YYYY"

            ws.row_dimensions[r].height = 13
            r += 1

        # ── Subtotal row ───────────────────────────────────────────────────────
        # Only col I (amount) has a value, all other cols blank
        for ci in range(1, ncols + 2):
            ws.cell(r, ci).fill = _fill(WHITE_FILL)
        if amount_ci:
            cell = ws.cell(row=r, column=amount_ci + 1, value=poc_total)
            cell.font          = _font(bold=True, size=9,
                                       color=GREEN_FG if poc_total >= 0 else RED_FG)
            cell.fill          = _fill(WHITE_FILL)
            cell.alignment     = _align("right")
            cell.number_format = "#,##0.00"
        ws.row_dimensions[r].height = 13
        subtotal_rows.append(r)
        grand_total += poc_total
        r += 1

        # ── 2 blank separator rows ─────────────────────────────────────────────
        ws.row_dimensions[r].height = 6
        r += 1
        ws.row_dimensions[r].height = 6
        r += 1

    # ── Catch-all for rows with no POC number ─────────────────────────────────
    if no_poc_rows:
        _write(ws, r, 1, "OTHER", bold=True, fill_rgb=PINK_FILL, size=9)
        for ci, col_name in enumerate(display_cols, 2):
            _write(ws, r, ci, col_name, bold=True, fill_rgb=PINK_FILL,
                   size=9, ha="center")
        ws.row_dimensions[r].height = 15
        r += 1

        other_total = 0.0
        for row_idx, data_row in enumerate(no_poc_rows):
            row_fill = GREY_FILL if row_idx % 2 == 0 else WHITE_FILL
            ws.cell(r, 1).fill = _fill(row_fill)
            for ci, col_name in enumerate(display_cols, 2):
                val    = data_row.get(col_name, "")
                is_amt = (col_name == amount_col)
                if is_amt:
                    cell_val = float(val) if pd.notna(val) else 0.0
                    other_total += cell_val
                elif pd.isna(val):
                    cell_val = ""
                else:
                    cell_val = val
                cell = ws.cell(r, ci, value=cell_val)
                cell.fill = _fill(row_fill)
                if is_amt:
                    cell.number_format = "#,##0.00"
                    cell.alignment = _align("right")
            ws.row_dimensions[r].height = 13
            r += 1

        # Subtotal
        if amount_ci:
            cell = ws.cell(r, amount_ci + 1, value=other_total)
            cell.font          = _font(bold=True, size=9)
            cell.number_format = "#,##0.00"
            cell.alignment     = _align("right")
        subtotal_rows.append(r)
        grand_total += other_total
        r += 1

    # ── Grand total (col K label, col M value at row 8) ───────────────────────
    cell_gt = ws.cell(row=8, column=13, value=grand_total)
    cell_gt.font          = _font(bold=True, size=11,
                                  color=GREEN_FG if grand_total >= 0 else RED_FG)
    cell_gt.number_format = "#,##0.00"
    cell_gt.alignment     = _align("right")

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
