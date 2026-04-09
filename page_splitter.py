import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import traceback

from splitter_engine import split_accounts, build_split_workbook, build_template_sheet
from template_manager import template_preview
from template_manager import (
    save_template as _sess_save,
    delete_template as _sess_del,
    get_template as _sess_get,
    list_templates as _sess_list,
    export_templates_json as _sess_export,
    import_templates_json as _sess_import,
)
from github_storage import (
    github_configured, list_github_templates,
    get_template_cached, save_github_template, delete_github_template,
    invalidate_cache,
)
from common import clean_id, get_email, LANG_LABELS, mailto_link
from customer_rules import (
    load_rule_github as _load_rule_direct,
    get_rule_cached, save_rule_github, delete_rule_github,
    invalidate_rule_cache, merge_rule, DEFAULT_RULE,
    _gh_ok as rules_gh_ok,
)
from chunked_builder import build_chunked_sheet

ACCOUNT_COLS  = ["Account", "Customer", "Debtor", "Klant",
                 "Debiteurnummer", "Debiteur", "Konto", "Kundennummer"]
AMOUNT_COLS   = ["Amount in local currency", "Bedrag in lokale valuta",
                 "Betrag in Hauswährung", "Amount in document currency"]
DUE_DATE_COLS = ["Net due date", "Netto-vervaldatum", "Nettofälligkeitsdatum"]


def _find(df, candidates):
    for n in candidates:
        if n in df.columns:
            return n
    for col in df.columns:
        for cand in candidates:
            if cand.lower() in col.lower():
                return col
    return None


def show():
    st.markdown("## 📂 Account Splitter")
    st.caption(
        "Split a multi-account SAP export into one sheet per customer. "
        "Removes invoices not yet due and applies customer templates automatically."
    )

    # ── UPLOAD ────────────────────────────────────────────────────────────────
    st.markdown("### 1 · Upload SAP export")
    st.markdown("**SAP Multi-Account Export** — FBL5N or any open items report (.xlsx)")
    uploaded = st.file_uploader(
        "SAP export", type=["xlsx", "xls"],
        label_visibility="collapsed", key="spl_file",
    )

    if not uploaded:
        st.info("Export from SAP (FBL5N) with your customer account range, save as .xlsx, and upload here.")
        _template_manager()
        return

    # ── PARSE FILE ────────────────────────────────────────────────────────────
    try:
        df_raw = pd.read_excel(uploaded, sheet_name=0, header=0, dtype=str)
        df_raw.columns = [str(col).strip() for col in df_raw.columns]
    except Exception as e:
        st.error(f"Could not read file: {e}")
        return

    account_col  = _find(df_raw, ACCOUNT_COLS)
    amount_col   = _find(df_raw, AMOUNT_COLS)
    due_date_col = _find(df_raw, DUE_DATE_COLS)

    if not account_col:
        st.error("Could not detect an account column. Use the override below.")

    # ── SETTINGS ──────────────────────────────────────────────────────────────
    st.markdown("### 2 · Confirm settings")

    with st.expander("Column detection — click to override if needed"):
        col_opts = df_raw.columns.tolist()
        account_col = st.selectbox(
            "Account column", col_opts,
            index=col_opts.index(account_col) if account_col in col_opts else 0,
            key="spl_acc_col",
        )
        amount_col = st.selectbox(
            "Amount column", ["(none)"] + col_opts,
            index=col_opts.index(amount_col) + 1 if amount_col in col_opts else 0,
            key="spl_amt_col",
        )
        due_date_col = st.selectbox(
            "Due date column", ["(none)"] + col_opts,
            index=col_opts.index(due_date_col) + 1 if due_date_col in col_opts else 0,
            key="spl_due_col",
        )
        if amount_col == "(none)":
            amount_col = None
        if due_date_col == "(none)":
            due_date_col = None

    accounts = sorted([
        clean_id(a) for a in df_raw[account_col].dropna().unique()
        if clean_id(a) is not None
    ])

    if accounts:
        st.markdown(f"**{len(accounts)} accounts detected:** " + "  ".join(f"`{a}`" for a in accounts))
    else:
        st.warning("No accounts found in the selected column.")

    c1, c2 = st.columns(2)
    with c1:
        remove_not_due = st.checkbox("Remove invoices not yet due", value=True, key="spl_remove")
    with c2:
        ref_date = st.date_input("Reference date", value=datetime.date.today(), key="spl_refdate")

    # ── GENERATE ──────────────────────────────────────────────────────────────
    st.markdown("### 3 · Generate")
    gen_col, _ = st.columns([1, 2])
    with gen_col:
        generate = st.button(
            "▶  Split into separate sheets",
            use_container_width=True,
            type="primary",
            key="spl_run",
        )

    if generate:
        if not accounts:
            st.error("No accounts detected — check your column selection.")
        else:
            with st.spinner(f"Splitting {len(accounts)} accounts…"):
                try:
                    if due_date_col:
                        df_raw[due_date_col] = pd.to_datetime(df_raw[due_date_col], errors="coerce")
                    if amount_col:
                        df_raw[amount_col] = pd.to_numeric(df_raw[amount_col], errors="coerce").fillna(0)

                    account_data = split_accounts(
                        df_raw, account_col, amount_col, due_date_col,
                        remove_not_due=remove_not_due,
                        reference_date=ref_date,
                    )
                    # Build the combined workbook — standard layout for all accounts.
                    # Templated accounts get their own individual sheets in the email section.
                    wb_bytes = build_split_workbook(account_data, amount_col, today=ref_date)
                    st.session_state["spl_result"]       = wb_bytes
                    st.session_state["spl_account_data"] = account_data
                    st.session_state["spl_ref_date"]     = ref_date
                    st.session_state["spl_amount_col"]   = amount_col
                    # Pre-fetch which accounts have templates so we can show a notice
                    has_templates = []
                    for acc_check in account_data.keys():
                        if github_configured():
                            tb = get_template_cached(str(acc_check))
                        else:
                            tb = _sess_get(st.session_state, str(acc_check))
                        if tb:
                            has_templates.append(str(acc_check))
                    st.session_state["spl_has_templates"] = has_templates
                except Exception as e:
                    st.error(f"Error: {e}")
                    with st.expander("Detail"):
                        st.code(traceback.format_exc())

    if "spl_result" not in st.session_state:
        _template_manager()
        return

    # ── RESULTS ───────────────────────────────────────────────────────────────
    account_data = st.session_state["spl_account_data"]
    ref_date     = st.session_state["spl_ref_date"]
    amount_col   = st.session_state["spl_amount_col"]
    today_str    = pd.Timestamp(ref_date).strftime("%d/%m/%Y")
    safe_date    = str(ref_date).replace("-", "")

    st.markdown("---")
    st.success(f"Done — {len(account_data)} account sheets generated")

    summary_rows = []
    for acc, acc_df in account_data.items():
        total = acc_df[amount_col].sum() if amount_col and amount_col in acc_df.columns else None
        summary_rows.append({
            "Account": acc, "Lines": len(acc_df),
            "Total (€)": f"{total:,.2f}" if total is not None else "—",
        })
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    # ── SECTION 1: DOWNLOADS ────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📥 Downloads")
    st.caption("Custom rules and templates are applied automatically per account.")

    st.download_button(
        "⬇  Download all accounts (standard layout)",
        data=st.session_state["spl_result"].getvalue(),
        file_name=f"Accounts_{safe_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, key="spl_dl_all",
    )

    st.markdown("**Individual downloads (with custom rules/templates applied):**")

    for acc, acc_df_sel in account_data.items():
        total_sel = acc_df_sel[amount_col].sum() if amount_col and amount_col in acc_df_sel.columns else 0

        rule = None
        if rules_gh_ok():
            try:
                rule = _load_rule_direct(str(acc))
            except Exception:
                pass

        tmpl_bytes = get_template_cached(str(acc)) if github_configured() else _sess_get(st.session_state, str(acc))

        if rule and (rule.get("chunk_size", 0) > 0 or rule.get("columns")):
            try:
                acc_wb_bytes = build_chunked_sheet(acc_df_sel, str(acc), rule, today=ref_date)
                layout_label = f"✓ chunked €{rule.get('chunk_size',0):,.0f} batches"
            except Exception as e:
                acc_wb_bytes = build_split_workbook({acc: acc_df_sel}, amount_col, today=ref_date).getvalue()
                layout_label = f"standard (rule error: {e})"
        elif tmpl_bytes:
            try:
                acc_wb_bytes = build_template_sheet(str(acc), acc_df_sel, tmpl_bytes, amount_col, today=ref_date)
                layout_label = "✓ custom template"
            except Exception:
                acc_wb_bytes = build_split_workbook({acc: acc_df_sel}, amount_col, today=ref_date).getvalue()
                layout_label = "standard"
        else:
            acc_wb_bytes = build_split_workbook({acc: acc_df_sel}, amount_col, today=ref_date, title_prefix=f"Account {acc} — ").getvalue()
            layout_label = "standard layout"

        dl_c, info_c = st.columns([2, 3])
        with dl_c:
            st.download_button(
                f"⬇  Account {acc}",
                data=acc_wb_bytes,
                file_name=f"Account_{acc}_{safe_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key=f"spl_dl_{acc}",
            )
        with info_c:
            st.caption(f"{len(acc_df_sel)} lines  ·  €{total_sel:,.2f}  ·  {layout_label}")

    # ── SECTION 2: EMAILS ───────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📧 Payment reminder emails")

    e1, e2, e3 = st.columns(3)
    with e1:
        lang    = st.selectbox("Language", list(LANG_LABELS.keys()), format_func=lambda x: LANG_LABELS[x], key="spl_lang")
    with e2:
        sender  = st.text_input("Your name",  key="spl_sender",  placeholder="Your Name")
    with e3:
        company = st.text_input("Company",    key="spl_company", placeholder="Your Company")

    for acc_idx, (acc, acc_df_sel) in enumerate(account_data.items()):
        total_sel = acc_df_sel[amount_col].sum() if amount_col and amount_col in acc_df_sel.columns else 0
        subject, body = get_email(
            "account", lang,
            customer_name=f"Account {acc}", account_id=str(acc), date=today_str,
            total_amount=f"€{total_sel:,.2f}",
            sender_name=sender or "[Your Name]", company_name=company or "[Your Company]",
        )
        with st.expander(f"Account {acc}  ·  €{total_sel:,.2f}", expanded=(acc_idx == 0)):
            st.text_input("Subject", value=subject, key=f"spl_subj_{acc}")
            st.text_area("Body",    value=body,    height=200, key=f"spl_body_{acc}")
            to_email = st.text_input("Customer email", key=f"spl_to_{acc}", placeholder="customer@example.com")
            if to_email:
                mailto = mailto_link(to_email, subject, body)
                st.markdown(
                    f'<a href="{mailto}" style="display:block;text-align:center;'
                    f'background:linear-gradient(135deg,#0f2942,#1d4ed8);color:white;'
                    f'font-weight:600;padding:12px;border-radius:8px;'
                    f'text-decoration:none;font-size:13px;margin-top:4px;">'
                    f'📧 Open in Email Client</a>',
                    unsafe_allow_html=True,
                )
            else:
                st.caption("Enter email above for mailto link.")

    # ── TEMPLATE MANAGER ──────────────────────────────────────────────────────
    st.markdown("---")
    _template_manager()


def _template_manager():
    """Persistent customer template manager — GitHub-backed if configured."""
    st.markdown("### 🎨 Customer templates")

    if not github_configured():
        _template_manager_session()
        return

    st.caption(
        "Templates are stored permanently in your GitHub repo and applied "
        "automatically when a matching account number is detected."
    )

    with st.spinner("Loading templates from GitHub…"):
        saved = list_github_templates()

    if saved:
        st.markdown(f"**{len(saved)} template(s) stored in GitHub:**")
        for tmpl_info in saved:
            acc_id     = tmpl_info["account_id"]
            size_kb    = tmpl_info["size"] / 1024
            tmpl_bytes = get_template_cached(acc_id)

            try:
                info   = template_preview(tmpl_bytes) if tmpl_bytes else {}
                detail = (
                    f"{info.get('layout_type','?')}  ·  "
                    f"{info.get('max_col','?')} cols  ·  "
                    f"header row {info.get('header_row','?')}"
                ) if info else f"{size_kb:.1f} KB"
            except Exception:
                detail = f"{size_kb:.1f} KB"

            ca, cb, cc, cd = st.columns([3, 1, 1, 1])
            with ca:
                st.markdown(f"**Account {acc_id}** — {detail}")
            with cb:
                if tmpl_bytes:
                    st.download_button(
                        "⬇ View", data=tmpl_bytes,
                        file_name=f"template_{acc_id}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"gh_dl_{acc_id}", use_container_width=True,
                    )
            with cc:
                if st.button("🔄 Replace", key=f"gh_replace_{acc_id}", use_container_width=True):
                    st.session_state[f"replacing_{acc_id}"] = True
            with cd:
                if st.button("🗑 Delete", key=f"gh_del_{acc_id}", use_container_width=True):
                    with st.spinner(f"Deleting {acc_id}…"):
                        ok, msg = delete_github_template(acc_id)
                    if ok:
                        invalidate_cache()
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

            if st.session_state.get(f"replacing_{acc_id}"):
                rep_file = st.file_uploader(
                    f"New template for {acc_id}",
                    type=["xlsx", "xls"],
                    key=f"gh_rep_file_{acc_id}",
                )
                if rep_file:
                    with st.spinner("Uploading…"):
                        ok, msg = save_github_template(acc_id, rep_file.read())
                    if ok:
                        invalidate_cache()
                        del st.session_state[f"replacing_{acc_id}"]
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
    else:
        st.info("No templates saved yet. Upload one below.")

    with st.expander("➕  Upload a new customer template"):
        st.caption(
            "Supports plain tables (custom column headers) and full custom layouts "
            "(merged cells, logos, branded headers). Upload the customer's actual Excel file."
        )
        acc_input = st.text_input(
            "Customer account number", key="gh_tmpl_acc",
            placeholder="e.g. 30113601",
        )
        new_file = st.file_uploader(
            "Customer template (.xlsx)", type=["xlsx", "xls"], key="gh_tmpl_file",
        )
        if new_file:
            raw = new_file.read()
            try:
                info = template_preview(raw)
                c1, c2, c3 = st.columns(3)
                c1.metric("Layout",     info["layout_type"])
                c2.metric("Columns",    info["max_col"])
                c3.metric("Header row", info["header_row"])
                if info["headers"]:
                    st.markdown(
                        "**Detected headers:** " +
                        "  ".join(f"`{h}`" for h in info["headers"][:8]) +
                        ("…" if len(info["headers"]) > 8 else "")
                    )
            except Exception as e:
                st.warning(f"Could not preview: {e}")

            if not acc_input:
                st.warning("Enter the account number above first.")
            else:
                if st.button(
                    f"💾  Save template for account {acc_input} to GitHub",
                    key="gh_tmpl_save", type="primary",
                ):
                    with st.spinner("Saving to GitHub…"):
                        ok, msg = save_github_template(acc_input, raw)
                    if ok:
                        invalidate_cache()
                        st.success(msg + " Applied automatically next time you split.")
                        st.rerun()
                    else:
                        st.error(msg)


    # ── CUSTOM RULES EDITOR ───────────────────────────────────────────────────
    with st.expander("\u2699\ufe0f  Custom output rules per customer"):
        st.caption(
            "Define how a customer sheet is structured: column order, "
            "grouping into payment batches, and grand total position. "
            "Rules are stored permanently in GitHub."
        )
        rule_acc = st.text_input(
            "Account number to configure",
            key="rule_acc_input", placeholder="e.g. 30111788",
        )
        if rule_acc:
            from github_storage import _repo as _gr2
            existing = get_rule_cached(rule_acc, _gr2()) if rules_gh_ok() else None
            base = existing or DEFAULT_RULE.copy()
            if existing:
                st.success("This account already has saved rules.")
            chunk_size = st.number_input(
                "Chunk size (\u20ac) - group rows into batches. 0 = disabled.",
                min_value=0.0, value=float(base.get("chunk_size", 0)),
                step=1000.0, format="%.0f", key="rule_chunk",
            )
            show_account = st.checkbox(
                "Include Account column",
                value=base.get("show_account", True), key="rule_show_acc",
            )
            total_position = st.radio(
                "Grand total position", ["bottom", "right"],
                index=0 if base.get("total_position", "bottom") == "bottom" else 1,
                key="rule_total_pos",
                help="right = yellow box to the right of the data"
            )
            cols_text = st.text_area(
                "Column order (one per line - leave blank for default)",
                value="\n".join(base.get("columns", [])),
                height=140, key="rule_cols",
                help="Exact SAP column names one per line in the order you want"
            )
            sort_col = st.text_input(
                "Sort rows by this column before chunking",
                value=(base.get("sort_by") or ["Net due date"])[0],
                key="rule_sort",
            )
            rule_obj = {
                "chunk_size":     chunk_size,
                "show_account":   show_account,
                "total_position": total_position,
                "columns":        [c.strip() for c in cols_text.strip().splitlines() if c.strip()],
                "sort_by":        [sort_col] if sort_col.strip() else ["Net due date"],
            }
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                if st.button(f"\U0001f4be  Save rules for {rule_acc}", key="rule_save", type="primary"):
                    ok, msg = save_rule_github(rule_acc, rule_obj)
                    if ok:
                        invalidate_rule_cache()
                        st.success(msg)
                    else:
                        st.error(msg)
            with sc2:
                if st.button(f"\U0001f50d  Verify saved rule", key="rule_verify"):
                    loaded = _load_rule_direct(rule_acc)
                    if loaded:
                        st.success(
                            f"Rule found in GitHub for account {rule_acc}:\n"
                            f"chunk_size={loaded.get('chunk_size',0)}  "
                            f"columns={loaded.get('columns',[])}  "
                            f"total_position={loaded.get('total_position','bottom')}"
                        )
                    else:
                        st.error(
                            f"No rule found in GitHub for account {rule_acc}. "
                            "Save it first, or check the account number matches exactly."
                        )
            with sc3:
                if existing and st.button(f"\U0001f5d1  Delete", key="rule_del"):
                    ok, msg = delete_rule_github(rule_acc)
                    if ok:
                        invalidate_rule_cache()
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)


def _template_manager_session():
    """Session-only fallback when GitHub is not configured."""
    st.warning(
        "GitHub storage is not configured — templates last this session only. "
        "Follow the setup steps below to make them permanent."
    )

    with st.expander("⚙️  One-time setup — permanent templates"):
        st.markdown("""
**Step 1 — Create a GitHub token**
1. Go to github.com → Settings → Developer settings → Personal access tokens → Fine-grained tokens
2. Click **Generate new token** · Name: `AR Suite` · Expiration: **No expiration**
3. Repository access → Only select repositories → pick your repo
4. Permissions → **Contents → Read and write**
5. Generate and copy the token

**Step 2 — Add to Streamlit secrets**

Streamlit Cloud → your app → ⋮ → Settings → Secrets:
```toml
[github]
token = "github_pat_your_token_here"
repo  = "your-username/your-repo-name"
```
        """)

    saved_ids = _sess_list(st.session_state)
    if saved_ids:
        st.markdown(f"**{len(saved_ids)} template(s) saved this session:**")
        for acc_id in saved_ids:
            ca, cb = st.columns([4, 1])
            with ca:
                st.markdown(f"**Account {acc_id}**")
            with cb:
                if st.button("Delete", key=f"sess_del_{acc_id}"):
                    _sess_del(st.session_state, acc_id)
                    st.rerun()
        st.download_button(
            "💾  Download backup (restore next session)",
            data=_sess_export(st.session_state),
            file_name="ar_suite_templates_backup.json",
            mime="application/json",
            key="sess_export",
        )

    with st.expander("➕  Upload a template (session only)"):
        acc = st.text_input("Account number", key="sess_acc", placeholder="e.g. 30113601")
        f   = st.file_uploader("Template (.xlsx)", type=["xlsx", "xls"], key="sess_file")
        if f and acc:
            raw = f.read()
            try:
                info = template_preview(raw)
                st.caption(f"{info['layout_type']} · {info['max_col']} columns")
            except Exception:
                pass
            if st.button("Save", key="sess_save"):
                _sess_save(st.session_state, acc, raw)
                st.success(f"Saved for account {acc} (this session).")
                st.rerun()

    with st.expander("📂  Restore from backup"):
        bf = st.file_uploader("Backup .json", type=["json"], key="sess_restore")
        if bf:
            if st.button("Restore", key="sess_restore_btn"):
                n = _sess_import(st.session_state, bf.read())
                st.success(f"Restored {n} template(s).")
                st.rerun()
