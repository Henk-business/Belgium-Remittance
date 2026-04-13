"""
Merged account builder.

When a customer has multiple SAP accounts that should be delivered as one
combined Excel file, this builder:
  1. Reads the structure from a reference template (column order, widths,
     header style, row heights).
  2. Produces a workbook with one sheet per account + a Summary sheet,
     matching the template exactly.
  3. Updates all titles, dates, line counts, and amounts with fresh data.

Account groups are stored in GitHub as JSON under
  templates/group_<primary_account>.json
  e.g. {"accounts": ["30172457", "30521289"], "label": "North and South Beverages"}
"""

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
from io import BytesIO
import datetime
import re
import warnings
warnings.filterwarnings("ignore")

# ── COLOURS (match standard splitter output) ──────────────────────────────────
DK_BLUE  = "1F3864"
MD_BLUE  = "2E75B6"
WHITE    = "FFFFFF"
GREY     = "F2F2F2"
POS_FG   = "C00000"   # red  = positive (invoices)
NEG_FG   = "375623"   # green = negative (credits)
BLACK_FG = "000000"


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
    if fmt:    cell.number_format = fmt
    if border: cell.border = _thin()
    return cell

def _mw(ws, row, c1, c2, val=None, bold=False, fill=WHITE,
        fg=BLACK_FG, size=10, ha="left"):
    ws.merge_cells(f"{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}")
    cell = ws.cell(row=row, column=c1, value=val)
    cell.font      = _font(bold=bold, color=fg, size=size)
    cell.fill      = _fill(fill)
    cell.alignment = _align(ha)
    cell.border    = _thin()
    for c in range(c1+1, c2+1):
        ws.cell(row=row, column=c).fill   = _fill(fill)
        ws.cell(row=row, column=c).border = _thin()
    return cell


def _read_template_structure(template_bytes: bytes) -> dict:
    """
    Extract structure info from a multi-sheet template:
    - Column headers (from first account sheet)
    - Column widths (from first account sheet)
    - Row heights
    - Header row number
    - Summary sheet column widths
    """
    wb = openpyxl.load_workbook(BytesIO(template_bytes))
    info = {
        "account_sheets": [],
        "has_summary":    False,
        "summary_cols":   {},
        "data_cols":      [],
        "col_widths":     {},
        "header_row":     4,
        "ncols":          9,
    }

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if sheet_name.lower() == "summary":
            info["has_summary"] = True
            for i in range(1, ws.max_column + 1):
                ltr = get_column_letter(i)
                info["summary_cols"][i] = ws.column_dimensions[ltr].width or 15
        else:
            info["account_sheets"].append(sheet_name)
            if not info["data_cols"]:
                # Find header row
                for r in range(1, min(10, ws.max_row + 1)):
                    row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                    non_empty = [v for v in row_vals if v and str(v).strip()]
                    if len(non_empty) >= 3:
                        info["header_row"] = r
                        info["data_cols"]  = [v for v in row_vals if v]
                        info["ncols"]      = len(row_vals)
                        break
                # Column widths
                for i in range(1, ws.max_column + 1):
                    ltr = get_column_letter(i)
                    info["col_widths"][i] = ws.column_dimensions[ltr].width or 12

    return info


def _write_account_sheet(wb, ws, acc_id: str, acc_df: pd.DataFrame,
                          tmpl_info: dict, today_str: str,
                          amount_col: str, lang: str = "en"):
    """Write one account sheet matching the template layout."""
    from splitter_engine import translate_doc_types
    acc_df = translate_doc_types(acc_df, lang)

    header_row = tmpl_info["header_row"]
    data_cols  = tmpl_info["data_cols"]
    ncols      = tmpl_info["ncols"]
    n_lines    = len(acc_df)

    # Set column widths
    for i, w in tmpl_info["col_widths"].items():
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Rows above header ─────────────────────────────────────────────────────
    # Row 1: Account title (merged, dark blue)
    _mw(ws, 1, 1, ncols, f"Account: {acc_id}",
        bold=True, fill=DK_BLUE, fg=WHITE, size=13, ha="left")
    ws.row_dimensions[1].height = 32

    # Row 2: Subtitle (merged, medium blue)
    subtitle = f"{today_str}  ·  {n_lines} lines  ·  Invoices not yet due removed"
    _mw(ws, 2, 1, ncols, subtitle,
        bold=False, fill=MD_BLUE, fg=WHITE, size=9, ha="left")
    ws.row_dimensions[2].height = 16

    # Row 3: blank gap
    ws.row_dimensions[3].height = 6

    # ── Column headers ────────────────────────────────────────────────────────
    for ci, col_name in enumerate(data_cols, 1):
        cell = ws.cell(row=header_row, column=ci, value=col_name)
        cell.font      = _font(bold=True, color=WHITE, size=9)
        cell.fill      = _fill(MD_BLUE)
        cell.alignment = _align("center")
        cell.border    = _thin()
    ws.row_dimensions[header_row].height = 18

    # ── Build column map: data_col name → sap_df column ───────────────────────
    def _norm(s): return re.sub(r"[^a-z0-9]", " ", str(s).lower()).split()
    col_map = {}
    for ci, tmpl_hdr in enumerate(data_cols, 1):
        t_words = set(_norm(tmpl_hdr))
        best, best_s = None, 0
        for sap_col in acc_df.columns:
            s_words = set(_norm(sap_col))
            overlap = len(t_words & s_words) / max(len(t_words | s_words), 1)
            if overlap > best_s and overlap >= 0.3:
                best_s, best = overlap, sap_col
        if best:
            col_map[ci] = best

    amt_ci = next((ci for ci, col in col_map.items()
                   if amount_col and col == amount_col), None)
    date_cols_ci = {ci for ci, col in col_map.items()
                    if any(kw in col.lower() for kw in ["date","datum"])
                    or (col in acc_df.columns and
                        pd.api.types.is_datetime64_any_dtype(acc_df[col]))}

    # ── Data rows ─────────────────────────────────────────────────────────────
    for di, (_, row_data) in enumerate(acc_df.iterrows()):
        r        = header_row + 1 + di
        row_fill = WHITE if di % 2 == 0 else GREY
        for ci in range(1, ncols + 1):
            sap_col = col_map.get(ci)
            is_amt  = (ci == amt_ci)
            is_date = (ci in date_cols_ci)

            if sap_col and sap_col in row_data.index:
                raw = row_data[sap_col]
                if is_amt:
                    val = float(raw) if pd.notna(raw) else 0.0
                    fg  = POS_FG if val >= 0 else NEG_FG
                elif is_date:
                    try:
                        val = pd.Timestamp(raw).to_pydatetime() if pd.notna(raw) else ""
                    except Exception:
                        val = str(raw) if pd.notna(raw) else ""
                    fg = BLACK_FG
                elif pd.isna(raw):
                    val, fg = "", BLACK_FG
                elif isinstance(raw, float) and raw == int(raw):
                    val, fg = int(raw), BLACK_FG
                else:
                    val, fg = raw, BLACK_FG
            else:
                val, fg = "", BLACK_FG

            cell = ws.cell(row=r, column=ci, value=val)
            cell.font      = _font(color=fg, size=9)
            cell.fill      = _fill(row_fill)
            cell.alignment = _align("right" if is_amt else "left")
            cell.border    = _thin()
            if is_amt:
                cell.number_format = "#,##0.00"
            elif is_date and isinstance(val, datetime.datetime):
                cell.number_format = "DD/MM/YYYY"

        ws.row_dimensions[r].height = 13

    return acc_df[amount_col].sum() if amount_col and amount_col in acc_df.columns else 0


def build_merged_workbook(account_dfs: dict, template_bytes: bytes,
                           amount_col: str, today=None,
                           group_label: str = "",
                           lang: str = "en") -> bytes:
    """
    Build a combined workbook with one sheet per account + Summary sheet.

    account_dfs: {account_id: DataFrame}
    template_bytes: the reference template to match
    """
    if today is None:
        today = datetime.date.today()
    today_str = pd.Timestamp(today).strftime("%d/%m/%Y")

    tmpl_info = _read_template_structure(template_bytes)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    account_totals = {}

    # ── One sheet per account ─────────────────────────────────────────────────
    for acc_id, acc_df in account_dfs.items():
        ws = wb.create_sheet(title=str(acc_id)[:31])
        total = _write_account_sheet(
            wb, ws, str(acc_id), acc_df, tmpl_info,
            today_str, amount_col, lang
        )
        account_totals[str(acc_id)] = {"total": total, "lines": len(acc_df)}

    # ── Summary sheet (first) ─────────────────────────────────────────────────
    ws_sum = wb.create_sheet(title="Summary", index=0)

    # Column widths from template
    sum_widths = tmpl_info.get("summary_cols", {1:18, 2:10, 3:18, 4:12, 5:14})
    for i, w in sum_widths.items():
        ws_sum.column_dimensions[get_column_letter(i)].width = w

    ncols_sum = 5
    n_accounts = len(account_dfs)
    total_lines = sum(v["lines"] for v in account_totals.values())
    grand_total = sum(v["total"] for v in account_totals.values())

    # Row 1: Title
    title = (f"SAP ACCOUNT OVERVIEW  —  {n_accounts} accounts  ·  {today_str}"
             if not group_label else
             f"{group_label}  —  {today_str}")
    _mw(ws_sum, 1, 1, ncols_sum, title,
        bold=True, fill=DK_BLUE, fg=WHITE, size=13, ha="left")
    ws_sum.row_dimensions[1].height = 34

    # Row 2: subtitle
    _mw(ws_sum, 2, 1, ncols_sum,
        "Invoices not yet due have been removed.",
        bold=False, fill=MD_BLUE, fg=WHITE, size=9, ha="left")
    ws_sum.row_dimensions[2].height = 16

    # Row 3: blank
    ws_sum.row_dimensions[3].height = 8

    # Row 4: headers
    sum_headers = ["Account", "Lines", "Total Amount (€)", "Open Items", "Sheet Name"]
    for ci, h in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=4, column=ci, value=h)
        cell.font      = _font(bold=True, color=WHITE, size=9)
        cell.fill      = _fill(MD_BLUE)
        cell.alignment = _align("center")
        cell.border    = _thin()
    ws_sum.row_dimensions[4].height = 15

    # Data rows
    for ri, (acc_id, info) in enumerate(account_totals.items()):
        r        = 5 + ri
        row_fill = GREY if ri % 2 == 0 else WHITE
        total    = info["total"]
        fg_amt   = POS_FG if total >= 0 else NEG_FG

        for ci, val in enumerate(
            [acc_id, info["lines"], total, info["lines"], acc_id], 1
        ):
            cell = ws_sum.cell(row=r, column=ci, value=val)
            is_amt = (ci == 3)
            cell.font      = _font(bold=True,
                                   color=fg_amt if is_amt else BLACK_FG,
                                   size=10)
            cell.fill      = _fill(row_fill)
            cell.alignment = _align("right" if is_amt or ci == 2 else "left")
            cell.border    = _thin()
            if is_amt:
                cell.number_format = "#,##0.00"
        ws_sum.row_dimensions[r].height = 16

    # Total row
    r_total = 5 + len(account_totals)
    gt_fg = POS_FG if grand_total >= 0 else NEG_FG
    for ci, val in enumerate(["TOTAL", total_lines, grand_total, "", ""], 1):
        cell = ws_sum.cell(row=r_total, column=ci, value=val)
        is_amt = (ci == 3)
        cell.font      = _font(bold=True,
                               color=gt_fg if is_amt else WHITE,
                               size=10)
        cell.fill      = _fill(DK_BLUE)
        cell.alignment = _align("right" if is_amt or ci == 2 else "left")
        cell.border    = _thin()
        if is_amt:
            cell.number_format = "#,##0.00"
    ws_sum.row_dimensions[r_total].height = 18

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
