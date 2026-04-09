"""Account splitter engine."""
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime
import warnings
warnings.filterwarnings("ignore")

from common import BG, FG, c, mr, col_w, hdr_row, fd, auto_widths, clean_id
from template_manager import apply_template

CUSTOMER_CONFIG_KEY = "customer_configs"

# Columns that are always stripped from the output — internal SAP fields
# that are not relevant to send to customers.
STRIP_COLS = {
    "Reason code",
    "Clerk Abbreviation",
    "Cleared/open items symbol",
    "Case ID",
    "Status",
    "Dunning Block",
    "Disputed item",
    "Payment Block",
    "Payment Method",
    "Net due date symbol",
    "G/L Account",
    "Text",
    "Clearing date",
    "Clearing Document",
    "Dunning Level",
    "Last Dunned",
    "Reversed with",
    "Document Header Text",
    "User Name",
    "Special G/L ind.",
    "Billing Document",
    "Reference Key 1",
    # normalised versions (after SAP column map)
    "text",
    "clearing_date",
    "clearing_doc",
    "header_text",
    "sap_class",
    "is_open",
    "ref",
    "doc_number_str",
}


def get_configs(state):
    return state.get(CUSTOMER_CONFIG_KEY, {})


def save_config(state, account_id, config):
    if CUSTOMER_CONFIG_KEY not in state:
        state[CUSTOMER_CONFIG_KEY] = {}
    state[CUSTOMER_CONFIG_KEY][str(account_id)] = config


def split_accounts(df, account_col, amount_col, due_date_col,
                   remove_not_due=True, reference_date=None,
                   customer_configs=None):
    if reference_date is None:
        reference_date = datetime.date.today()
    ref_ts = pd.Timestamp(reference_date)

    # ── Strip SAP-inserted subtotal/summary rows FIRST
    # Real transactions always have a Document Number. Rows without one are
    # SAP-generated subtotals (account totals, grand totals) that must be removed
    # before any splitting or totalling happens.
    doc_num_col = next(
        (c for c in df.columns
         if c.lower() in ("document number", "belegnummer", "doc_number", "doc number")),
        None
    )
    if doc_num_col:
        has_doc = (
            df[doc_num_col].notna() &
            ~df[doc_num_col].astype(str).str.strip().isin(["", "nan", "0", "0.0"])
        )
        df = df[has_doc].copy()

    accounts = sorted([
        clean_id(a) for a in df[account_col].dropna().unique()
        if clean_id(a) is not None
    ])

    result = {}
    for acc in accounts:
        mask = df[account_col].apply(clean_id) == acc
        acc_df = df[mask].copy()

        # ── Remove invoices not yet due so total reflects only displayed rows
        if remove_not_due and due_date_col and due_date_col in acc_df.columns:
            due = pd.to_datetime(acc_df[due_date_col], errors="coerce")
            acc_df = acc_df[due.isna() | (due <= ref_ts)].copy()

        # ── Strip unwanted SAP columns
        cols_to_drop = [col for col in acc_df.columns if col in STRIP_COLS]
        if cols_to_drop:
            acc_df = acc_df.drop(columns=cols_to_drop)

        if customer_configs:
            cfg = customer_configs.get(str(acc), {})
            cols = cfg.get("columns")
            if cols:
                valid = [col for col in cols if col in acc_df.columns]
                extra = [col for col in acc_df.columns if col not in valid]
                acc_df = acc_df[valid + extra]

        result[acc] = acc_df
    return result


def _safe_tab(name, idx):
    for ch in r"\/*?:[]":
        name = name.replace(ch, "_")
    return (name[:28] if len(name) > 28 else name) or f"Acct_{idx+1}"


def build_split_workbook(account_data, amount_col, today=None, title_prefix="", templates=None):
    if today is None:
        today = datetime.date.today()
    today_str = pd.Timestamp(today).strftime("%d/%m/%Y")

    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"

    accounts = list(account_data.keys())
    col_w(ws_sum, [18, 10, 18, 12, 14])

    r = 1
    mr(ws_sum, r, 1, 5,
       title_prefix + "SAP ACCOUNT OVERVIEW  \u2014  " + str(len(accounts)) + " accounts  \u00b7  " + today_str,
       bold=True, bg="dk_blue", fg="white", sz=13, ha="center")
    ws_sum.row_dimensions[r].height = 34
    r += 1
    mr(ws_sum, r, 1, 5, "Invoices not yet due have been removed.",
       bg="md_blue", fg="white", sz=9, ha="center", italic=True)
    ws_sum.row_dimensions[r].height = 16
    r += 2

    hdr_row(ws_sum, r, ["Account", "Lines", "Total Amount (\u20ac)", "Open Items", "Sheet Name"])
    r += 1

    grand_total = 0.0
    grand_lines = 0

    from openpyxl.styles import Border, Side
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for idx, (acc, acc_df) in enumerate(account_data.items()):
        tab_name = _safe_tab(str(acc), idx)
        n = len(acc_df)
        total = acc_df[amount_col].sum() if amount_col and amount_col in acc_df.columns else 0.0
        clearing_col = next(
            (col for col in acc_df.columns
             if "clearing" in col.lower() and "doc" in col.lower()), None
        )
        n_open = acc_df[clearing_col].isna().sum() if clearing_col else n
        bg = "grey" if idx % 2 == 0 else "white"

        for ci, val in enumerate([acc, n, total, n_open, tab_name], 1):
            cell = ws_sum.cell(row=r, column=ci, value=val)
            fg_col = "166534" if ci == 3 and isinstance(val, float) and val >= 0 else \
                     "B91C1C" if ci == 3 and isinstance(val, float) and val < 0 else \
                     "1D4ED8" if ci == 5 else "000000"
            cell.font  = Font(name="Arial", bold=(ci == 1), size=10, color=fg_col)
            cell.fill  = PatternFill("solid", fgColor=BG.get(bg, "FFFFFF"))
            cell.alignment = Alignment(
                horizontal="right" if ci in (2, 3, 4) else ("center" if ci == 5 else "left"),
                vertical="center"
            )
            if ci == 3:
                cell.number_format = "#,##0.00"
            cell.border = border
        ws_sum.row_dimensions[r].height = 16
        r += 1
        grand_total += total
        grand_lines += n

    for ci, val in enumerate(["TOTAL", grand_lines, grand_total, "", ""], 1):
        cell = ws_sum.cell(row=r, column=ci, value=val)
        cell.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=BG["dk_blue"])
        cell.alignment = Alignment(
            horizontal="right" if ci in (2, 3) else "left", vertical="center"
        )
        if ci == 3:
            cell.number_format = "#,##0.00"
    ws_sum.row_dimensions[r].height = 18

    for idx, (acc, acc_df) in enumerate(account_data.items()):
        tab_name = _safe_tab(str(acc), idx)

        # Template accounts get their own separate workbook (see build_template_sheet)
        # so we skip them here but still show them in the summary
        if templates and str(acc) in templates:
            # Just add a note row in summary (account already counted in grand totals above)
            pass

        ws = wb.create_sheet(title=tab_name)
        ncols = len(acc_df.columns)

        r2 = 1
        mr(ws, r2, 1, ncols, "Account: " + str(acc),
           bold=True, bg="dk_blue", fg="white", sz=13, ha="center")
        ws.row_dimensions[r2].height = 32
        r2 += 1
        mr(ws, r2, 1, ncols,
           today_str + "  \u00b7  " + f"{len(acc_df):,}" + " lines  \u00b7  Invoices not yet due removed",
           bg="md_blue", fg="white", sz=9, ha="center")
        ws.row_dimensions[r2].height = 16
        r2 += 2

        for ci, col_name in enumerate(acc_df.columns, 1):
            cell = ws.cell(row=r2, column=ci, value=str(col_name))
            cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
            cell.fill = PatternFill("solid", fgColor=BG["md_blue"])
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[r2].height = 18
        r2 += 1

        amount_ci = None
        if amount_col and amount_col in acc_df.columns:
            amount_ci = list(acc_df.columns).index(amount_col) + 1

        # Identify date columns
        date_cols = set()
        for ci_chk, col_name in enumerate(acc_df.columns, 1):
            if any(kw in col_name.lower() for kw in ["date", "datum"]):
                date_cols.add(ci_chk)
            elif pd.api.types.is_datetime64_any_dtype(acc_df[col_name]):
                date_cols.add(ci_chk)

        for ri, (_, row) in enumerate(acc_df.iterrows(), r2):
            bg = "grey" if ri % 2 == 0 else "white"
            for ci, (col_name, val) in enumerate(row.items(), 1):
                is_amt  = (ci == amount_ci)
                is_date = (ci in date_cols)
                if is_amt:
                    cell_val = float(val) if pd.notna(val) else 0.0
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
                cell = ws.cell(row=ri, column=ci, value=cell_val)
                fg_col = (
                    "166534" if is_amt and isinstance(cell_val, (int, float)) and cell_val >= 0
                    else "B91C1C" if is_amt and isinstance(cell_val, (int, float)) and cell_val < 0
                    else "000000"
                )
                cell.font = Font(name="Arial", size=9, color=fg_col)
                cell.fill = PatternFill("solid", fgColor=BG.get(bg, "FFFFFF"))
                cell.alignment = Alignment(
                    horizontal="right" if is_amt else "left", vertical="center"
                )
                if is_amt:
                    cell.number_format = "#,##0.00"
                elif is_date and isinstance(cell_val, datetime.datetime):
                    cell.number_format = "DD/MM/YYYY"
            ws.row_dimensions[ri].height = 13

        total_r = r2 + len(acc_df)
        c(ws, total_r, 1, "TOTAL", bold=True, bg="dk_blue", fg="white", sz=10)
        for ci in range(2, ncols + 1):
            c(ws, total_r, ci, None, bg="dk_blue")
        if amount_ci:
            total_val = acc_df[amount_col].sum()
            cell = ws.cell(row=total_r, column=amount_ci, value=total_val)
            cell.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=BG["dk_blue"])
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[total_r].height = 18

        auto_widths(ws, acc_df, start_col=1)
        ws.freeze_panes = "A5"

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def build_template_sheet(account_id: str, acc_df: pd.DataFrame,
                         template_bytes: bytes, amount_col: str,
                         today=None) -> bytes:
    """
    Build a single-account workbook using the customer's custom template.
    Returns the filled workbook as bytes.
    Falls back to the standard layout if template application fails.
    """
    from template_manager import apply_template
    today_str = pd.Timestamp(today or datetime.date.today()).strftime("%d/%m/%Y")

    try:
        return apply_template(template_bytes, acc_df)
    except Exception:
        # Fallback: standard layout for this account
        data = build_split_workbook(
            {account_id: acc_df}, amount_col,
            today=today,
            title_prefix=f"Account {account_id} — ",
        )
        return data.getvalue()
