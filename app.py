"""
DHC Working Automation — Streamlit app.

Upload the 4 source files, click Process, preview summaries, download the final workbook.
"""
import traceback
import streamlit as st
import etl

st.set_page_config(page_title="DHC Working Automation", layout="wide", initial_sidebar_state="expanded")

# --- Custom CSS for styling ---
st.markdown(
    """
    <style>
    .main { max-width: 1400px; margin: 0 auto; }
    .stTitle { color: #1F3864; font-size: 2.5rem; margin-bottom: 0.5rem; }
    h2 { color: #2E5395; border-bottom: 2px solid #D9E1F2; padding-bottom: 0.5rem; }
    .stAlert { border-radius: 8px; }
    [data-baseweb="button"] button { border-radius: 6px; font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 DHC Working Automation")
st.markdown("**Automated receipt processing & summary generation** — Upload your 4 monthly files, process, and download.")

# --- Sidebar Help ---
with st.sidebar:
    st.header("📋 Quick Guide")
    with st.expander("File Requirements", expanded=False):
        st.markdown(
            """
**1. DCR file (.xlsb)**
- Daily Collection Report extract
- Required sheets: `Sheet1` (receipts), `Sheet2` (agreement master)

**2. To be Disabled (.xlsb)**
- Cash-mode compliance lists
- Required sheets: `CIF Level Disable`, `Agreement Level Disble`

**3. Employee Master (.xlsx)**
- Must contain a `Mobile Number` column
- Used to flag receipts entered by agents vs. customers

**4. Previous DHC Working (.xlsx)**
- Last month's output workbook
- The `Look Up` sheet is carried forward (columns J:N)
- **New agreements will show blank CIF** — fill in and use as next month's input
            """
        )
    with st.expander("Output Sheets", expanded=False):
        st.markdown(
            """
**Receipt made summary** — Updated/Pending vs Bounced/Cancelled by mode & status

**RTGS Summary** — Zone × Receipt Type × TAT (Turn-Around Time) matrix + Online payment breakdown

**Cash Mode Validat Summary** — Compliance violations: customers on disable lists paying via cash/Airtel

**Delay in RCPTING Summary** — Full month: Zone × Receipt Type × TAT aging buckets

**RCPT CXN** — Cancelled-receipt log (remarks column empty for you to fill)
            """
        )

st.markdown("---")

# --- File uploads ---
col1, col2 = st.columns(2)
with col1:
    dcr_file = st.file_uploader("1️⃣ DCR file (.xlsb)", type=["xlsb"], help="Daily Collection Report")
    employee_file = st.file_uploader("3️⃣ Employee Master (.xlsx)", type=["xlsx"], help="Mobile number list")
with col2:
    disable_file = st.file_uploader("2️⃣ To be Disabled (.xlsb)", type=["xlsb"], help="Cash-mode compliance lists")
    prior_working_file = st.file_uploader("4️⃣ Previous DHC Working (.xlsx)", type=["xlsx"], help="Carried-forward master")

all_files_ready = all([dcr_file, disable_file, employee_file, prior_working_file])
col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1])
with col_btn1:
    run = st.button("🚀 Process Files", type="primary", disabled=not all_files_ready, use_container_width=True)

if run:
    with st.spinner("Processing... this may take 1–2 minutes"):
        try:
            with st.status("Loading & transforming data...", expanded=True) as status:
                st.write("📥 Loading DCR extract...")
                receipts, dcr_master = etl.load_dcr(dcr_file)

                st.write("📥 Loading disable lists...")
                cif_disable, agr_disable = etl.load_disable_lists(disable_file)

                st.write("📥 Loading employee mobiles...")
                mobiles = etl.load_employee_mobiles(employee_file)

                st.write("📥 Loading carried-forward Look Up master...")
                prior_master = etl.load_prior_lookup_master(prior_working_file)

                st.write("🔧 Refreshing Look Up master...")
                lookup_master = etl.build_lookup_master(prior_master, dcr_master)

                st.write("🔧 Building DCR tab...")
                dcr_tab = etl.build_dcr_tab(receipts, lookup_master, agr_disable, cif_disable, mobiles)

                st.write("🔧 Building RTGS tab...")
                rtgs_tab = etl.build_rtgs_tab(dcr_tab)

                st.write("📊 Generating summary sheets...")
                receipt_made_summary = etl.build_receipt_made_summary(dcr_tab)
                rtgs_summary = etl.build_rtgs_summary(rtgs_tab, dcr_tab)
                cash_mode_validation_summary = etl.build_cash_mode_validation_summary(dcr_tab)
                delay_summary = etl.build_delay_in_rcpting_summary(dcr_tab)
                rcpt_cxn = etl.build_rcpt_cxn(dcr_tab)

                st.write("💾 Writing Excel workbook...")
                workbook_buffer = etl.write_output_workbook(
                    rtgs_summary, delay_summary, receipt_made_summary,
                    cash_mode_validation_summary, rcpt_cxn
                )
                status.update(label="✅ Done!", state="complete")

            # --- Store in session ---
            st.session_state["workbook_buffer"] = workbook_buffer
            st.session_state["receipt_made_summary"] = receipt_made_summary
            st.session_state["rtgs_summary"] = rtgs_summary
            st.session_state["cash_mode_validation"] = cash_mode_validation_summary
            st.session_state["delay_summary"] = delay_summary
            st.session_state["rcpt_cxn"] = rcpt_cxn
            st.session_state["lookup_master"] = lookup_master
            st.session_state["dcr_tab"] = dcr_tab
            st.session_state["rtgs_tab"] = rtgs_tab

            # --- Stats banner ---
            st.markdown("---")
            st.success("✅ Processing complete!")
            metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
            with metric_col1:
                st.metric("Total Receipts", f"{len(dcr_tab):,}", delta="DCR rows processed")
            with metric_col2:
                st.metric("RTGS Receipts", f"{len(rtgs_tab):,}", delta="RTGS-mode subset")
            with metric_col3:
                needs_cif = int(lookup_master["NEEDS_CIF_MAPPING"].sum())
                st.metric("New Agreements", f"{needs_cif:,}", delta="missing CIF mapping")
            with metric_col4:
                cxn_count = len(rcpt_cxn)
                st.metric("Cancelled Receipts", f"{cxn_count:,}", delta="for review")

            # --- Warning for new agreements ---
            if needs_cif > 0:
                st.warning(
                    f"⚠️ **{needs_cif:,} agreement(s) need CIF mapping.** These are brand-new to this month's DCR. "
                    f"Fill in their CIF/Zone/Sub Region in the 'Look Up' section below and use the next output "
                    f"as your 'Previous DHC Working' file next month.",
                    icon="⚠️"
                )

        except Exception as e:
            st.error(f"❌ Error: {e}")
            with st.expander("Traceback"):
                st.code(traceback.format_exc())

# --- Preview tabs (only show after successful run) ---
if "workbook_buffer" in st.session_state:
    st.markdown("---")
    st.subheader("📑 Sheet Previews")

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
        ["Receipt Made Summary", "RTGS Summary", "Cash Mode Validation", "Delay in RCPTING", "RCPT CXN", "Look Up Master", "Debug: DCR"]
    )

    with tab1:
        summary = st.session_state["receipt_made_summary"]
        st.markdown("**Updated / Pending**")
        df_left = etl.receipt_made_table_to_dataframe(summary["left"], ["Cleared", "Deposit", "Pending"])
        st.dataframe(df_left, use_container_width=True)
        st.markdown("**Updated / Bounced or Cancelled**")
        df_right = etl.receipt_made_table_to_dataframe(summary["right"], ["Cleared", "Deposit", "Bounced", "Cxn"])
        st.dataframe(df_right, use_container_width=True)

    with tab2:
        summary = st.session_state["rtgs_summary"]
        st.markdown("**Zone × Receipt Type × TAT Matrix**")
        df_matrix = etl.zone_tat_matrix_to_dataframe(summary["matrix"])
        st.dataframe(df_matrix, use_container_width=True)
        st.markdown("**Online Payment Sources (Full Month)**")
        sources = summary["online_source_block"]["rows"]
        if sources:
            st.dataframe(
                {name: cnt for name, cnt in sources},
                use_container_width=True,
                height=200
            )
        else:
            st.info("No online payment receipts this period.")

    with tab3:
        df = st.session_state["cash_mode_validation"]
        if df.empty:
            st.info("✅ No compliance violations (no flagged customers paid via cash/Airtel).")
        else:
            st.dataframe(df.head(100), use_container_width=True)
            st.caption(f"Showing first 100 of {len(df)} rows")

    with tab4:
        summary = st.session_state["delay_summary"]
        st.markdown("**Zone × Receipt Type × TAT Aging**")
        df_matrix = etl.zone_tat_matrix_to_dataframe(summary["matrix"])
        st.dataframe(df_matrix, use_container_width=True)

    with tab5:
        df = st.session_state["rcpt_cxn"]
        if df.empty:
            st.info("✅ No cancelled receipts this period.")
        else:
            st.dataframe(df, use_container_width=True)

    with tab6:
        df = st.session_state["lookup_master"]
        st.dataframe(df.head(100), use_container_width=True)
        st.caption(f"Showing first 100 of {len(df)} agreements")

    with tab7:
        df = st.session_state["dcr_tab"]
        cols_to_show = ["AGREEMENTNO", "RECEIPTSOURCE", "MODEOFPAYMENT", "Status", "Mode", "AMOUNTPAID", "RECEIPT ENTER DATE"]
        cols_available = [c for c in cols_to_show if c in df.columns]
        st.dataframe(df[cols_available].head(50), use_container_width=True)

    # --- Download button ---
    st.markdown("---")
    st.subheader("💾 Download Output")
    st.download_button(
        "📥 Download DHC Working Output.xlsx",
        data=st.session_state["workbook_buffer"],
        file_name="DHC_Working_Automated_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
    st.markdown(
        "✅ The workbook includes 5 sheets with automated summaries. "
        "**RCPT CXN** remarks column is left blank for your review and notes.",
        help="The 'RCPT CXN' sheet is the only manual input — fill in remarks for each cancelled receipt."
    )
