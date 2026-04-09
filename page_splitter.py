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
            "Account":   acc,
            "Lines":     len(acc_df),
            "Total (€)": f"{total:,.2f}" if total is not None else "—",
        })
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    st.download_button(
        "⬇  Download Split Excel (all accounts — standard layout)",
        data=st.session_state["spl_result"].getvalue(),
        file_name=f"Accounts_{safe_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="spl_dl_all",
    )
    has_templates = st.session_state.get("spl_has_templates", [])
    if has_templates:
        st.info(
            f"📋  **{len(has_templates)} account(s) have custom templates: "
            + ", ".join(has_templates) +
            "**  \n"
            "The download above uses the standard layout. "
            "Scroll down to the email section to download each templated account individually — "
            "those downloads use the custom template."
        )

    # ── EMAIL DRAFTS — one section per account ────────────────────────────────
    st.markdown("---")
    st.markdown("### 📧 Payment reminder emails")

    e1, e2, e3 = st.columns(3)
    with e1:
        lang    = st.selectbox("Language", list(LANG_LABELS.keys()),
                               format_func=lambda x: LANG_LABELS[x], key="spl_lang")
    with e2:
        sender  = st.text_input("Your name",  key="spl_sender",  placeholder="Your Name")
    with e3:
        company = st.text_input("Company",    key="spl_company", placeholder="Your Company")

    st.caption("One draft per account. Expand to preview, enter email, and open in your email client.")

    for acc_idx, (acc, acc_df_sel) in enumerate(account_data.items()):
        total_sel = acc_df_sel[amount_col].sum() if amount_col and amount_col in acc_df_sel.columns else 0

        subject, body = get_email(
            "account", lang,
            customer_name=f"Account {acc}",
            account_id=str(acc),
            date=today_str,
            total_amount=f"\u20ac{total_sel:,.2f}",
            sender_name=sender  or "[Your Name]",
            company_name=company or "[Your Company]",
        )

        with st.expander(
            f"Account {acc}  ·  {len(acc_df_sel)} lines  ·  \u20ac{total_sel:,.2f}",
            expanded=(acc_idx == 0),
        ):
            # Get template for this account (GitHub first, session fallback)
            tmpl_bytes = None
            if github_configured():
                tmpl_bytes = get_template_cached(str(acc))
            else:
                tmpl_bytes = _sess_get(st.session_state, str(acc))

            if tmpl_bytes:
                acc_wb_bytes = build_template_sheet(
                    str(acc), acc_df_sel, tmpl_bytes, amount_col, today=ref_date
                )
                st.caption("✓ Custom template applied for this account")
            else:
                acc_wb_bytes = build_split_workbook(
                    {acc: acc_df_sel}, amount_col, today=ref_date,
                    title_prefix=f"Account {acc} — ",
                ).getvalue()

            st.text_input("Subject", value=subject, key=f"spl_subj_{acc}")
            st.text_area("Body",    value=body,    height=200, key=f"spl_body_{acc}")

            to_email = st.text_input(
                "Customer email (optional)", key=f"spl_to_{acc}",
                placeholder="customer@example.com",
            )

            dl_col, link_col = st.columns(2)
            with dl_col:
                st.download_button(
                    f"⬇  Download account {acc} sheet",
                    data=acc_wb_bytes,
                    file_name=f"Account_{acc}_{safe_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"spl_dl_{acc}",
                )
            with link_col:
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
