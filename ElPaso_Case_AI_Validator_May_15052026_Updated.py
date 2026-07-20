import io
import os
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

import core

st.set_page_config(page_title="El Paso Case Validator", page_icon="⚖️", layout="wide")

# =========================
# SESSION STATE
# =========================
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "html_df" not in st.session_state:
    st.session_state.html_df = None
if "excel_df" not in st.session_state:
    st.session_state.excel_df = None
if "client_df" not in st.session_state:
    st.session_state.client_df = None
if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None
if "output_name" not in st.session_state:
    st.session_state.output_name = None
if "log_lines" not in st.session_state:
    st.session_state.log_lines = []

st.title("⚖️ El Paso Case Validator")
st.caption("HTML → Excel comparison tool")

# =========================
# SIDEBAR: INPUTS
# =========================
with st.sidebar:
    st.header("Inputs")

    html_files = st.file_uploader(
        "HTML files (list_*.html and detail_*.html)",
        type=["html", "htm"],
        accept_multiple_files=True,
        help="Upload both the list_ main files and their matching detail_ files.",
    )
    excel_files = st.file_uploader(
        "Excel output files",
        type=["xlsx"],
        accept_multiple_files=True,
    )
    client_file = st.file_uploader(
        "Client Input file (optional)",
        type=["xlsx", "xls"],
        accept_multiple_files=False,
    )

    st.divider()
    st.subheader("Validation Fields")
    select_all = st.checkbox("Select all fields", value=True)
    active_fields = []
    with st.expander("Choose fields", expanded=False):
        for f in core.FIELDS:
            checked = st.checkbox(f, value=select_all, key=f"field_{f}")
            if checked:
                active_fields.append(f)

    run_clicked = st.button("▶  Run Validation", type="primary", use_container_width=True)

# =========================
# LOGGING HELPER
# =========================
def log_fn(msg, tag=""):
    st.session_state.log_lines.append((msg, tag))

# =========================
# RUN PIPELINE
# =========================
if run_clicked:
    st.session_state.log_lines = []
    if not html_files or not excel_files:
        st.error("Please upload at least one HTML file set and one Excel file.")
    else:
        with st.spinner("Running validation…"):
            with tempfile.TemporaryDirectory() as tmpdir:
                html_dir = os.path.join(tmpdir, "html")
                excel_dir = os.path.join(tmpdir, "excel")
                os.makedirs(html_dir, exist_ok=True)
                os.makedirs(excel_dir, exist_ok=True)

                for uf in html_files:
                    with open(os.path.join(html_dir, uf.name), "wb") as f:
                        f.write(uf.getbuffer())
                for uf in excel_files:
                    with open(os.path.join(excel_dir, uf.name), "wb") as f:
                        f.write(uf.getbuffer())

                client_path = ""
                if client_file is not None:
                    client_path = os.path.join(tmpdir, client_file.name)
                    with open(client_path, "wb") as f:
                        f.write(client_file.getbuffer())

                log_fn("Loading HTML files…", "info")
                html_df = core.load_html(html_dir, log_fn)
                log_fn("Loading Excel files…", "info")
                excel_df = core.load_excel(excel_dir, log_fn)
                client_cases, client_df, norm_to_orig = core.load_client(client_path, log_fn)

                if html_df.empty or excel_df.empty:
                    log_fn("No data found in the uploaded files.", "err")
                    st.session_state.log_lines = st.session_state.log_lines
                    st.error("No data found — check that the HTML/Excel files uploaded are in the expected format.")
                else:
                    result_df = core.validate(html_df, excel_df, active_fields, client_cases, norm_to_orig)

                    # --- Safety net: ensure result_df has a clean, unique
                    # index and unique column names as soon as it's created,
                    # so downstream display/styling never trips on duplicates
                    # coming from concat/merge operations inside core.py.
                    result_df = result_df.reset_index(drop=True)
                    result_df = result_df.loc[:, ~result_df.columns.duplicated()]

                    out_name = f"Result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    out_path = os.path.join(tmpdir, out_name)
                    core.save_output(
                        out_path, html_df, excel_df, result_df,
                        client_df if not client_df.empty else None,
                        active_fields,
                    )
                    with open(out_path, "rb") as f:
                        output_bytes = f.read()

                    st.session_state.result_df = result_df
                    st.session_state.html_df = html_df
                    st.session_state.excel_df = excel_df
                    st.session_state.client_df = client_df
                    st.session_state.output_bytes = output_bytes
                    st.session_state.output_name = out_name
                    log_fn(f"Saved: {out_name}", "ok")

# =========================
# METRICS
# =========================
result_df = st.session_state.result_df
col1, col2, col3, col4, col5 = st.columns(5)
if result_df is not None and not result_df.empty:
    matches = int((result_df["ValidationStatus"] == "MATCH").sum())
    mismatches = int((result_df["ValidationStatus"] == "MISMATCH").sum())
    notfound = int((result_df["ValidationStatus"] == "NOT FOUND").sum())
    notfound_out = int((result_df["ValidationStatus"] == "No cases matched your search criteria").sum())
    col1.metric("HTML Rows", len(st.session_state.html_df))
    col2.metric("Excel Rows", len(st.session_state.excel_df))
    col3.metric("Matched", matches)
    col4.metric("Mismatched", mismatches + notfound)
    col5.metric("Not Found in Output", notfound_out)
else:
    col1.metric("HTML Rows", "—")
    col2.metric("Excel Rows", "—")
    col3.metric("Matched", "—")
    col4.metric("Mismatched", "—")
    col5.metric("Not Found in Output", "—")

if st.session_state.output_bytes:
    st.download_button(
        "⬇ Download Result Workbook",
        data=st.session_state.output_bytes,
        file_name=st.session_state.output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

# =========================
# TABS: LOG / RESULTS
# =========================
tab_log, tab_results = st.tabs(["Run Log", "Results"])

with tab_log:
    if st.session_state.log_lines:
        log_text = "\n".join(f"[{tag.upper() or 'INFO'}] {msg}" for msg, tag in st.session_state.log_lines)
        st.code(log_text, language=None)
    else:
        st.info("Run a validation to see the log here.")

with tab_results:
    if result_df is None or result_df.empty:
        st.info("No results yet — upload files and click Run Validation.")
    else:
        colf1, colf2 = st.columns([2, 3])
        with colf1:
            status_options = ["All"] + sorted(result_df["ValidationStatus"].dropna().unique().tolist())
            status_filter = st.selectbox("Filter by status", status_options)
        with colf2:
            search_q = st.text_input("Search (Case Number / Source HTML)")

        df = result_df.copy()
        if status_filter != "All":
            df = df[df["ValidationStatus"] == status_filter]
        if search_q:
            q = search_q.lower()
            mask = df.apply(
                lambda r: q in str(r.get("CaseNumber", "")).lower() or q in str(r.get("SourceHTML", "")).lower(),
                axis=1,
            )
            df = df[mask]

        # Styler.apply/.map require a unique index and unique columns —
        # filtering can leave duplicate index labels behind, and upstream
        # merges can occasionally leave duplicate column names.
        df = df.reset_index(drop=True)
        df = df.loc[:, ~df.columns.duplicated()]

        # Temporary debug output — remove once confirmed fixed.
        with st.expander("Debug info (safe to remove later)", expanded=False):
            st.write("Columns:", df.columns.tolist())
            st.write("Index is unique:", df.index.is_unique)
            st.write("Columns is unique:", df.columns.is_unique)

        mismatch_cols = [c for c in df.columns if c.endswith("_Result")]

        def highlight_status(row):
            status = row.get("ValidationStatus", "")
            color = ""
            if status == "MATCH":
                color = "background-color: #EAF3DE"
            elif status == "MISMATCH":
                color = "background-color: #FCEBEB"
            elif status == "NOT FOUND":
                color = "background-color: #FAEEDA"
            elif status == "No cases matched your search criteria":
                color = "background-color: #FEF3C7"
            return [color] * len(row)

        display_cols = ["SourceHTML", "CaseNumber", "ClientInputCaseNumber",
                         "ClientCaseNumber_Result", "SourceExcel", "ValidationStatus",
                         "ClientMatch"] + mismatch_cols
        # de-duplicate while preserving order: ClientCaseNumber_Result also
        # matches the "_Result" suffix filter used to build mismatch_cols
        seen = set()
        display_cols = [c for c in display_cols if c in df.columns and not (c in seen or seen.add(c))]

        show_df = df[display_cols].copy()
        # Final guard: even after the earlier dedup, slicing by display_cols
        # could theoretically reintroduce a repeated label if display_cols
        # itself had a repeat that slipped past — this makes absolutely sure.
        show_df = show_df.loc[:, ~show_df.columns.duplicated()]
        show_df = show_df.reset_index(drop=True)

        try:
            st.dataframe(
                show_df.style.apply(highlight_status, axis=1),
                use_container_width=True,
                height=500,
            )
        except KeyError as e:
            st.error(f"Styling failed ({e}); showing unstyled table instead.")
            st.dataframe(show_df, use_container_width=True, height=500)

        st.caption(f"{len(df)} of {len(result_df)} rows shown")
