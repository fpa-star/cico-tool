"""
app.py  —  CICO Classification Tool  (Streamlit GUI)
Run:  streamlit run app.py
"""

import json, logging, os, sys, time
from datetime import datetime
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

st.set_page_config(
    page_title="CICO Classification Tool",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Hide sidebar completely
st.markdown("""
<style>
[data-testid="collapsedControl"] { display: none; }
section[data-testid="stSidebar"] { display: none; }
/* Tighten top padding */
.block-container { padding-top: 1rem !important; }
/* Tab styling */
.stTabs [data-baseweb="tab-list"] {
    gap: 0px;
    background: #1e3a5f;
    border-radius: 8px 8px 0 0;
    padding: 4px 8px 0 8px;
}
.stTabs [data-baseweb="tab"] {
    color: #cce0ff;
    font-weight: 600;
    font-size: 15px;
    padding: 10px 28px;
    border-radius: 6px 6px 0 0;
}
.stTabs [aria-selected="true"] {
    background: white !important;
    color: #1e3a5f !important;
}
/* Toolbar button row */
.toolbar-btn { display: inline-flex; gap: 8px; flex-wrap: wrap; align-items: center; }
</style>
""", unsafe_allow_html=True)

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from classifier      import CICOClassifier, _to_float
from contra_matcher  import run_contra_matching
from excel_handler   import load_bank_dump, validate_columns, write_classified_output, list_dump_sheets
from rule_parser     import parse_rules_from_excel, save_rules
from validator       import run_gotcha_checks
import parse_all_tabs as _pat

logging.basicConfig(level=logging.INFO)

# ── Paths ─────────────────────────────────────────────────────────────────────
CONFIG_PATH = os.path.join(_HERE, "config.json")
RULES_JSON  = os.path.join(_HERE, "data", "classification_rules.json")
HEADS_JSON  = os.path.join(_HERE, "data", "valid_heads.json")

# ── Helpers ───────────────────────────────────────────────────────────────────
def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH) as f: return json.load(f)
    return {"api_key": "", "default_folder": ""}

def save_config(cfg):
    with open(CONFIG_PATH, "w") as f: json.dump(cfg, f, indent=2)

def load_valid_heads():
    if os.path.exists(HEADS_JSON):
        d = json.load(open(HEADS_JSON, encoding="utf-8"))
        return d.get("heads", []), d.get("subheads", {})
    return [], {}

def _safe_float(v):
    try: return float(str(v).replace(",","").strip())
    except: return 0.0

def _build_head_summary(df):
    out = []
    for head, grp in df.groupby("Head"):
        if not head or head in ("","DSA-SKIP"): continue
        out.append({
            "Head":            head,
            "Rows":            len(grp),
            "Credits (INR L)": round(grp["Credit"].apply(_safe_float).sum()/1e5, 2),
            "Debits (INR L)":  round(grp["Debit"].apply(_safe_float).sum()/1e5, 2),
            "Net (INR L)":     round(grp["Net"].apply(_safe_float).sum()/1e5, 2),
        })
    return sorted(out, key=lambda x: -abs(x.get("Net (INR L)",0)))

# ── Session state ─────────────────────────────────────────────────────────────
def _init():
    defaults = {
        "config": load_config(), "df": None, "source_path": None,
        "classified_df": None, "edited_df": None, "stats": None,
        "summary": None, "output_path": None, "step": 0,
        "undo_stack": [], "redo_stack": [], "grid_version": 0,
    }
    for k,v in defaults.items():
        if k not in st.session_state: st.session_state[k] = v
_init()

def _push_undo():
    """Save current state to undo stack before a mutating action."""
    if st.session_state.edited_df is not None:
        st.session_state.undo_stack.append(st.session_state.edited_df.copy())
        if len(st.session_state.undo_stack) > 30:
            st.session_state.undo_stack.pop(0)
        st.session_state.redo_stack.clear()
        st.session_state.grid_version += 1

rules_count = len(json.load(open(RULES_JSON, encoding="utf-8"))) if os.path.exists(RULES_JSON) else 0

# ── Header bar ────────────────────────────────────────────────────────────────
hcol1, hcol2 = st.columns([6, 2])
hcol1.markdown("## 💰 CICO Classification Tool")
hcol2.markdown(f"<div style='text-align:right;padding-top:14px;color:#555;font-size:13px'>Rules loaded: <b>{rules_count:,}</b></div>", unsafe_allow_html=True)

# ── Top navigation tabs ───────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "⚙️  Setup",
    "📂  Load File",
    "🔍  Classify & Edit",
    "📤  Export",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — SETUP
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("### ⚙️ Setup")
    cfg = st.session_state.config

    with st.form("setup_form"):
        api_key     = st.text_input("Anthropic API Key (optional)", value=cfg.get("api_key",""), type="password")
        folder_path = st.text_input("Default Output Folder Path", value=cfg.get("default_folder",""))
        if st.form_submit_button("💾 Save Settings"):
            cfg["api_key"] = api_key; cfg["default_folder"] = folder_path
            save_config(cfg); st.session_state.config = cfg
            st.success("Settings saved.")

    st.markdown("---")
    st.subheader("🔄 Refresh All Rules from Excel")
    st.markdown("Upload `CICO_ claude.xlsx` to parse **all 4 tabs** — Classification Rules, Account Master, Valid Heads, Gotchas.")

    rules_file = st.file_uploader("Upload CICO master Excel (.xlsx)", type=["xlsx"], key="rules_upload")
    if st.button("🔄 Refresh All Rules", type="primary") and rules_file:
        with st.spinner("Parsing all tabs…"):
            try:
                import tempfile
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(rules_file.read()); tmp_path = tmp.name
                orig = _pat.INPUT; _pat.INPUT = tmp_path
                _pat.main(); _pat.INPUT = orig
                os.unlink(tmp_path)
                rc = len(json.load(open(RULES_JSON, encoding="utf-8")))
                st.success(f"✅ Parsed successfully — **{rc:,} classification rules**, Account master, Valid heads, Gotchas updated.")
            except Exception as e:
                st.error(f"Error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — LOAD FILE
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("### 📂 Load Bank Statement File")

    bank_file = st.file_uploader("Bank Statement Dump (.xlsx)", type=["xlsx"], key="bank_upload")

    selected_sheet = None
    if bank_file:
        import tempfile
        tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp2.write(bank_file.read()); tmp2.close()
        st.session_state["_tmp_bank_path"] = tmp2.name
        dump_sheets = list_dump_sheets(tmp2.name)
        if not dump_sheets:
            import openpyxl as _ox
            _wb = _ox.load_workbook(tmp2.name, read_only=True)
            dump_sheets = _wb.sheetnames; _wb.close()
        selected_sheet = st.selectbox("Select the bank dump sheet", dump_sheets)

    if bank_file and st.button("📋 Load & Validate", type="primary"):
        with st.spinner("Loading…"):
            try:
                tmp_path = st.session_state.get("_tmp_bank_path")
                df = load_bank_dump(tmp_path, sheet_name=selected_sheet)
                st.session_state.df = df
                st.session_state.source_path = tmp_path
                st.session_state.classified_df = None
                st.session_state.edited_df = None

                col_check = validate_columns(df)
                c1, c2, c3 = st.columns(3)
                c1.metric("Rows", f"{len(df):,}")
                c2.metric("Accounts", f"{df['Account Name'].nunique()}")
                if "Value Date" in df.columns:
                    c3.metric("Dates", f"{df['Value Date'].nunique()}")
                if not col_check["ok"]:
                    st.warning(f"Missing columns: {col_check['missing']}")
                else:
                    st.success(f"✅ Loaded {len(df):,} rows — all expected columns present.")
                with st.expander("Account list"):
                    st.dataframe(df["Account Name"].value_counts().reset_index(), use_container_width=True)
            except Exception as e:
                st.error(f"Failed: {e}")

    if st.session_state.df is not None:
        st.info(f"✅ File loaded — **{len(st.session_state.df):,} rows** ready. Switch to **Classify & Edit** tab.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — CLASSIFY & EDIT
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    if st.session_state.df is None:
        st.warning("No file loaded. Go to the **Load File** tab first.")
        st.stop()

    valid_heads, subhead_map = load_valid_heads()
    all_subheads = sorted(set(s for sl in subhead_map.values() for s in sl if s))

    # ── Run Classification ────────────────────────────────────────────────────
    if st.session_state.classified_df is None:
        st.markdown("### ▶ Run Classification")
        cfg = st.session_state.config
        use_ai = st.checkbox("Enable AI fallback (Layer 3)", value=bool(cfg.get("api_key","")))
        if st.button("▶ Run Classification", type="primary"):
            api_key    = cfg.get("api_key","") if use_ai else ""
            classifier = CICOClassifier(api_key=api_key or None)
            df = st.session_state.df.copy()

            pb   = st.progress(0)
            stat = st.empty()
            cc   = st.columns(4)

            def upd(done, total, s):
                pct = int(done/total*100) if total else 0
                pb.progress(pct); stat.text(f"{done:,}/{total:,} rows…")
                cc[0].metric("RULE",     s.get("layer1",0))
                cc[1].metric("AI",       s.get("ai",0))
                cc[2].metric("DSA-SKIP", s.get("dsa_skip",0))
                cc[3].metric("TBD",      s.get("tbd",0))

            df, stats = classifier.classify_dataframe(df, progress_callback=upd)
            pb.progress(100); stat.text("Classification complete ✓")

            with st.spinner("Gotcha checks…"):   gotcha = run_gotcha_checks(df)
            with st.spinner("Contra matching…"): contra = run_contra_matching(df)

            if "Net" not in df.columns:
                df["Net"] = df.apply(lambda r: _safe_float(r.get("Credit",0)) - _safe_float(r.get("Debit",0)), axis=1)
            if "Row Color" not in df.columns:
                df["Row Color"] = ""

            st.session_state.classified_df = df
            st.session_state.edited_df     = df.copy()
            st.session_state.stats         = stats
            st.session_state.summary       = {
                **stats,
                "head_summary":   _build_head_summary(df),
                "gotcha_fixes":   gotcha.get("fixes",[]),
                "gotcha_flags":   gotcha.get("flags",[]),
                "contra":         contra,
                "tbd_narrations": [],
            }
            st.rerun()

    # ── Interactive Editor ────────────────────────────────────────────────────
    if st.session_state.edited_df is not None:

        # ══ Process pending actions FIRST (before any rendering) ══
        # This is the correct Streamlit pattern: mutations happen at the top,
        # then the UI renders with the updated state.

        action_msg = ""

        if st.session_state.get("_action") == "add_row":
            st.session_state.pop("_action")
            _push_undo()
            new_row = {c: "" for c in st.session_state.edited_df.columns}
            new_row["Sub-head 2"] = "MANUAL"
            st.session_state.edited_df = pd.concat(
                [pd.DataFrame([new_row]), st.session_state.edited_df],
                ignore_index=True
            )
            action_msg = f"✅ Blank row added at top. Total rows: {len(st.session_state.edited_df):,}"

        elif st.session_state.get("_action") == "delete":
            st.session_state.pop("_action")
            sel = st.session_state.pop("_selected_rows", [])
            if sel:
                _push_undo()
                sel_df = pd.DataFrame(sel)
                drop_keys = set(zip(
                    sel_df.get("Narration",    pd.Series(dtype=str)).astype(str).tolist(),
                    sel_df.get("Account Name", pd.Series(dtype=str)).astype(str).tolist(),
                ))
                mask = st.session_state.edited_df.apply(
                    lambda r: (str(r.get("Narration","")), str(r.get("Account Name",""))) in drop_keys, axis=1
                )
                st.session_state.edited_df = st.session_state.edited_df[~mask].reset_index(drop=True)
                action_msg = f"✅ Deleted {mask.sum()} row(s)."
            else:
                action_msg = "⚠️ No rows were selected. Tick checkboxes in the table first."

        elif st.session_state.get("_action") == "color":
            color_val = st.session_state.pop("_action_color", "")
            sel = st.session_state.pop("_selected_rows", [])
            st.session_state.pop("_action")
            if sel:
                _push_undo()
                sel_df = pd.DataFrame(sel)
                keys = set(zip(
                    sel_df.get("Narration",    pd.Series(dtype=str)).astype(str).tolist(),
                    sel_df.get("Account Name", pd.Series(dtype=str)).astype(str).tolist(),
                ))
                new_color = "" if color_val == "Clear" else color_val
                def _color_fn(r):
                    if (str(r.get("Narration","")), str(r.get("Account Name",""))) in keys:
                        return new_color
                    return r.get("Row Color","")
                st.session_state.edited_df["Row Color"] = st.session_state.edited_df.apply(_color_fn, axis=1)
                action_msg = f"✅ Color '{new_color or 'cleared'}' applied to {len(sel_df)} row(s)."
            else:
                action_msg = "⚠️ No rows were selected. Tick checkboxes in the table first."

        elif st.session_state.get("_action") == "undo":
            st.session_state.pop("_action")
            if st.session_state.undo_stack:
                st.session_state.redo_stack.append(st.session_state.edited_df.copy())
                st.session_state.edited_df = st.session_state.undo_stack.pop()
                st.session_state.grid_version += 1
                action_msg = "↩ Undo applied."

        elif st.session_state.get("_action") == "redo":
            st.session_state.pop("_action")
            if st.session_state.redo_stack:
                st.session_state.undo_stack.append(st.session_state.edited_df.copy())
                st.session_state.edited_df = st.session_state.redo_stack.pop()
                st.session_state.grid_version += 1
                action_msg = "↪ Redo applied."

        elif st.session_state.get("_action") == "save":
            upd_data = st.session_state.pop("_save_data", None)
            saved_idx = st.session_state.pop("_save_idx", None)
            st.session_state.pop("_action")
            if upd_data is not None:
                _push_undo()
                upd_df = pd.DataFrame(upd_data)
                for col in ["Head","Sub-head 1","Row Color"]:
                    if col in upd_df.columns and saved_idx is not None:
                        for arr_pos, orig_idx in enumerate(saved_idx):
                            if arr_pos < len(upd_df):
                                st.session_state.edited_df.at[orig_idx, col] = upd_df.iloc[arr_pos][col]
                st.session_state.summary["head_summary"] = _build_head_summary(st.session_state.edited_df)
                action_msg = f"✅ Changes saved — {len(upd_df)} rows updated."

        df = st.session_state.edited_df.copy()

        # ── Summary metrics ───────────────────────────────────────────────
        stats = st.session_state.stats or {}
        m = st.columns(5)
        m[0].metric("Total Rows",  stats.get("total", len(df)))
        m[1].metric("RULE",        stats.get("layer1", 0))
        m[2].metric("AI",          stats.get("ai", 0))
        m[3].metric("DSA-SKIP",    stats.get("dsa_skip", 0))
        m[4].metric("TBD",         stats.get("tbd", 0))

        if action_msg:
            if action_msg.startswith("⚠️"):
                st.warning(action_msg)
            else:
                st.success(action_msg)

        with st.expander("📊 Head-wise Summary (click to expand)", expanded=False):
            hs = _build_head_summary(df)
            if hs: st.dataframe(pd.DataFrame(hs), use_container_width=True, hide_index=True)

        st.markdown("---")

        # ── Toolbar above table ───────────────────────────────────────────
        st.markdown("#### 📋 Transaction Table")

        tb1, tb2, tb3, tb4, tb5, tb6, tb7, tb8 = st.columns([2, 1.2, 1.2, 2, 2, 2, 1.5, 1.5])

        save_clicked       = tb1.button("💾 Save Changes",   type="primary", key="btn_save")
        undo_clicked       = tb2.button("↩ Undo",  disabled=len(st.session_state.undo_stack)==0, key="btn_undo")
        redo_clicked       = tb3.button("↪ Redo",  disabled=len(st.session_state.redo_stack)==0, key="btn_redo")
        reclassify_clicked = tb4.button("🔄 Re-classify",                    key="btn_reclassify")
        add_clicked        = tb5.button("➕ Add Row",                         key="btn_add")
        delete_clicked     = tb6.button("🗑️ Delete Selected",                key="btn_delete")
        color_choice       = tb7.selectbox("color", ["—","Red","Yellow","Green","Blue","Orange","Clear"],
                                key="color_picker", label_visibility="collapsed")
        apply_color        = tb8.button("🎨 Apply Color",                    key="btn_color")

        st.caption("✏️ Click **Head**, **Sub-head 1**, or **Row Color** cell to edit inline  |  Tick checkbox to select rows for Delete / Color")

        # ── Filters ───────────────────────────────────────────────────────
        with st.expander("🔽 Panel Filters (use column header ▼ triangles in table for per-column filter)", expanded=False):
            filter_cols = st.columns(4)
            acct_opts  = sorted(df["Account Name"].dropna().unique().tolist())
            head_opts  = sorted(df["Head"].dropna().unique().tolist())
            sub_opts   = sorted(df["Sub-head 1"].dropna().unique().tolist())
            src_opts   = sorted(df["Sub-head 2"].dropna().unique().tolist())

            sel_accts = filter_cols[0].multiselect("Account Name", acct_opts, placeholder="All")
            sel_heads = filter_cols[1].multiselect("Head",         head_opts, placeholder="All")
            sel_subs  = filter_cols[2].multiselect("Sub-head",     sub_opts,  placeholder="All")
            sel_srcs  = filter_cols[3].multiselect("Source",       src_opts,  placeholder="All")

            filter_cols2 = st.columns(4)
            if "Bank" in df.columns:
                bank_opts = sorted(df["Bank"].dropna().unique().tolist())
                sel_banks = filter_cols2[0].multiselect("Bank", bank_opts, placeholder="All")
            else:
                sel_banks = []
            nar_kw    = filter_cols2[1].text_input("Narration contains", placeholder="Search…")
            sel_dir   = filter_cols2[2].selectbox("Direction", ["All","Debit only","Credit only"])
            sel_color_f = filter_cols2[3].multiselect("Row Color",
                           ["Red","Yellow","Green","Blue","Orange"], placeholder="All")

        # Apply panel filters
        fdf = df.copy()
        if sel_accts: fdf = fdf[fdf["Account Name"].isin(sel_accts)]
        if sel_heads: fdf = fdf[fdf["Head"].isin(sel_heads)]
        if sel_subs:  fdf = fdf[fdf["Sub-head 1"].isin(sel_subs)]
        if sel_srcs:  fdf = fdf[fdf["Sub-head 2"].isin(sel_srcs)]
        if sel_banks and "Bank" in fdf.columns: fdf = fdf[fdf["Bank"].isin(sel_banks)]
        if nar_kw:    fdf = fdf[fdf["Narration"].astype(str).str.upper().str.contains(nar_kw.upper(), na=False)]
        if sel_dir == "Debit only":  fdf = fdf[fdf["Debit"].apply(_safe_float)  > 0]
        if sel_dir == "Credit only": fdf = fdf[fdf["Credit"].apply(_safe_float) > 0]
        if sel_color_f: fdf = fdf[fdf["Row Color"].isin(sel_color_f)]

        st.caption(f"Showing **{len(fdf):,}** of **{len(df):,}** rows")

        # ── Build display dataframe ───────────────────────────────────────
        display_cols = []
        for c in ["Bank","Account Name","Account No","Value Date","Transaction Type",
                  "Narration","CCY","Debit","Credit","Net",
                  "Head","Sub-head 1","Sub-head 2","Row Color"]:
            if c in fdf.columns:
                display_cols.append(c)

        show_df = fdf[display_cols].copy()

        # Fix duplicate column names
        seen = {}
        new_cols = []
        for c in show_df.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        show_df.columns = new_cols
        # Drop duplicate columns (e.g. Account No_1)
        show_df = show_df[[c for c in show_df.columns if not (c.endswith("_1") or c.endswith("_2")) or c in ("Sub-head 1", "Sub-head 2")]]
        display_cols = list(show_df.columns)

        for col in ["Head", "Sub-head 1", "Row Color"]:
            if col in show_df.columns:
                show_df[col] = show_df[col].astype(str).replace("nan", "")

        # ── AgGrid — Excel-style ──────────────────────────────────────────
        gb = GridOptionsBuilder.from_dataframe(show_df)
        gb.configure_default_column(
            editable=False, resizable=True, sortable=True,
            filter=True, floatingFilter=True,
            minWidth=80,
        )

        gb.configure_column("Head", editable=True,
            cellEditor="agSelectCellEditor",
            cellEditorParams={"values": [""] + valid_heads},
            width=200)

        gb.configure_column("Sub-head 1", editable=True,
            cellEditor="agSelectCellEditor",
            cellEditorParams={"values": [""] + all_subheads},
            width=180)

        gb.configure_column("Row Color", editable=True,
            cellEditor="agSelectCellEditor",
            cellEditorParams={"values": ["","Red","Yellow","Green","Blue","Orange"]},
            width=110,
            cellStyle=JsCode("""
            function(params) {
                const m = {'Red':'#FFB3B3','Yellow':'#FFF3B3','Green':'#B3FFB3','Blue':'#B3D9FF','Orange':'#FFD9B3'};
                return m[params.value] ? {backgroundColor: m[params.value], fontWeight:'bold'} : {};
            }"""))

        row_style_fn = JsCode("""
        function(params) {
            const m = {'Red':'#FFE5E5','Yellow':'#FFFBE5','Green':'#E5FFE5','Blue':'#E5F2FF','Orange':'#FFF2E5'};
            return m[params.data && params.data['Row Color']] ? {background: m[params.data['Row Color']]} : {};
        }""")

        gb.configure_column("Narration",    width=260)
        gb.configure_column("Account Name", width=200)
        gb.configure_column("Value Date",   width=120)
        gb.configure_column("Debit",        width=115, type=["numericColumn"])
        gb.configure_column("Credit",       width=115, type=["numericColumn"])
        gb.configure_column("Net",          width=115, type=["numericColumn"])

        gb.configure_selection("multiple", use_checkbox=True, header_checkbox=True)
        gb.configure_grid_options(
            rowStyle=row_style_fn,
            suppressRowClickSelection=True,
            enableRangeSelection=True,
        )
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=100)
        gb.configure_side_bar(filters_panel=True, columns_panel=True)

        grid_response = AgGrid(
            show_df,
            gridOptions=gb.build(),
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            allow_unsafe_jscode=True,
            theme="balham",
            height=520,
            fit_columns_on_grid_load=False,
            key=f"main_grid_{st.session_state.grid_version}",
        )

        # Always persist selected rows so they survive button-click reruns
        sel_rows = grid_response.get("selected_rows")
        if sel_rows is not None and len(sel_rows) > 0:
            st.session_state["_selected_rows"] = sel_rows

        # ── Toolbar button handlers — set flags, then rerun ───────────────
        if add_clicked:
            st.session_state["_action"] = "add_row"
            st.session_state.grid_version += 1
            st.rerun()

        if delete_clicked:
            st.session_state["_action"] = "delete"
            st.rerun()

        if apply_color and color_choice != "—":
            st.session_state["_action"] = "color"
            st.session_state["_action_color"] = color_choice
            st.rerun()

        if undo_clicked:
            st.session_state["_action"] = "undo"
            st.rerun()

        if redo_clicked:
            st.session_state["_action"] = "redo"
            st.rerun()

        if save_clicked:
            updated = grid_response["data"]
            if updated is not None and len(updated) > 0:
                st.session_state["_action"]    = "save"
                st.session_state["_save_data"] = pd.DataFrame(updated).to_dict("records")
                st.session_state["_save_idx"]  = list(fdf.index)
                st.rerun()

        if reclassify_clicked:
            st.session_state.classified_df = None
            st.session_state.edited_df     = None
            st.rerun()

        # ── Gotcha flags ──────────────────────────────────────────────────
        flags = (st.session_state.summary or {}).get("gotcha_flags", [])
        if flags:
            with st.expander(f"⚠️ {len(flags)} rows flagged for review (Gotcha checks)"):
                st.dataframe(pd.DataFrame(flags)[["gotcha","account","narration","old_head","action"]], use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — EXPORT
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown("### 📤 Export Classified File")

    if st.session_state.edited_df is None:
        st.warning("No classified data yet. Go to **Classify & Edit** tab first.")
        st.stop()

    cfg     = st.session_state.config
    out_dir = st.text_input("Output folder", value=cfg.get("default_folder",""))

    if st.button("💾 Export to Excel", type="primary"):
        with st.spinner("Writing Excel output…"):
            try:
                out_path = write_classified_output(
                    source_path   = st.session_state.source_path,
                    df            = st.session_state.edited_df,
                    summary_data  = st.session_state.summary or {},
                    output_folder = out_dir,
                )
                st.session_state.output_path = out_path
                st.success(f"✅ Saved: `{out_path}`")
            except Exception as e:
                st.error(f"Export failed: {e}")

    if st.session_state.output_path and os.path.exists(st.session_state.output_path):
        with open(st.session_state.output_path,"rb") as f:
            st.download_button(
                "⬇ Download Classified File",
                data      = f.read(),
                file_name = os.path.basename(st.session_state.output_path),
                mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.markdown("---")
    summary = st.session_state.summary or {}
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total",    summary.get("total","—"))
    c2.metric("RULE",     summary.get("layer1",0))
    c3.metric("AI",       summary.get("ai",0))
    c4.metric("DSA-SKIP", summary.get("dsa_skip",0))
    c5.metric("TBD",      summary.get("tbd",0))

    hs = summary.get("head_summary",[])
    if hs:
        st.subheader("Head-wise Summary (INR L)")
        st.dataframe(pd.DataFrame(hs), use_container_width=True, hide_index=True)

    st.markdown("---")
    if st.button("🔁 Start New Classification"):
        for k in ["df","classified_df","edited_df","stats","summary","output_path","source_path"]:
            st.session_state[k] = None
        st.rerun()
