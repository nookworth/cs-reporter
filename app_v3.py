# CS Reporter Web UI
# Simple Streamlit interface for generating reports
# **Date:** 2026-02-20
# **Status:** ✅ Completed
# **Component:** app

import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# Import V2 modules
from src.config_v2 import load_config
from src.excel_reader_v2 import ExcelReader
from src.ppt_writer_v3 import PowerPointWriterV3

# Page config
st.set_page_config(page_title="CS Reporter", page_icon="📊", layout="wide")

# Title
st.title("📊 CS Reporter - Web Interface")
st.markdown("Generate PowerPoint reports from Excel ticket data")

# Sidebar - Configuration
with st.sidebar:
    st.header("⚙️ Configuration")

    # Version selection
    version = st.radio("Version", ["V2 (Recommended)", "V1 (Legacy)"], index=0)

    # Config file selection
    if version == "V2 (Recommended)":
        config_options = ["demo_mapping_v2.yaml", "mapping_v2.yaml"]
    else:
        config_options = ["demo_mapping.yaml", "mapping.yaml"]

    config_file = st.selectbox("Config File", config_options)

    st.markdown("---")
    st.markdown("### 📖 Quick Guide")
    st.markdown("""
    1. Upload current month Excel
    2. Upload previous month Excel
    3. Preview extracted data
    4. Generate report
    5. Download PowerPoint
    """)

# Main content
col1, col2 = st.columns(2)

with col1:
    st.header("📁 Current Month")
    current_file = st.file_uploader(
        "Upload current month Excel file", type=["xlsx", "xls"], key="current"
    )

    if current_file:
        st.success(f"✓ Loaded: {current_file.name}")

        # Show preview
        with st.expander("Preview Data"):
            try:
                df = pd.read_excel(current_file, nrows=5)
                st.dataframe(df)
                st.info(f"Sheets available: {pd.ExcelFile(current_file).sheet_names}")
            except Exception as e:
                st.error(f"Error reading file: {e}")

with col2:
    st.header("📁 Previous Month")
    previous_file = st.file_uploader(
        "Upload previous month Excel file", type=["xlsx", "xls"], key="previous"
    )

    if previous_file:
        st.success(f"✓ Loaded: {previous_file.name}")

        # Show preview
        with st.expander("Preview Data"):
            try:
                df = pd.read_excel(previous_file, nrows=5)
                st.dataframe(df)
                st.info(f"Sheets available: {pd.ExcelFile(previous_file).sheet_names}")
            except Exception as e:
                st.error(f"Error reading file: {e}")

# Generate button
st.markdown("---")

if current_file and previous_file:
    if st.button("🚀 Generate Report", type="primary", use_container_width=True):
        with st.spinner("Generating report..."):
            try:
                # Save uploaded files temporarily
                with tempfile.TemporaryDirectory() as tmpdir:
                    current_path = Path(tmpdir) / current_file.name
                    previous_path = Path(tmpdir) / previous_file.name

                    with open(current_path, "wb") as f:
                        f.write(current_file.getbuffer())
                    with open(previous_path, "wb") as f:
                        f.write(previous_file.getbuffer())

                    # Load config
                    config_path = f"config/{config_file}"
                    config = load_config(config_path)

                    st.info("✓ Configuration loaded")

                    # Extract data
                    reader = ExcelReader(str(current_path), config, str(previous_path))
                    data = reader.extract_data()

                    st.info(f"✓ Extracted {len(data)} fields")

                    # Show extracted data preview
                    with st.expander("📊 Extracted Data Preview", expanded=True):
                        # Show scalar fields
                        scalar_data = {
                            k: v for k, v in data.items() if not isinstance(v, list)
                        }
                        if scalar_data:
                            st.subheader("Metrics")
                            df_metrics = pd.DataFrame([scalar_data])
                            st.dataframe(df_metrics)

                            # Add charts
                            st.subheader("📈 Visual Analytics")

                            col_chart1, col_chart2 = st.columns(2)

                            with col_chart1:
                                # Ticket volume comparison
                                st.markdown("**Ticket Volume Comparison**")
                                total_current = data.get("re_req", 0) + data.get(
                                    "su_req", 0
                                )
                                total_previous = data.get("re_prev_req", 0) + data.get(
                                    "su_prev_req", 0
                                )

                                chart_data = pd.DataFrame(
                                    {
                                        "Month": ["Previous", "Current"],
                                        "Tickets": [total_previous, total_current],
                                    }
                                )
                                st.bar_chart(chart_data.set_index("Month"))

                                # Show percentage change
                                if total_previous > 0:
                                    change = (
                                        (total_current - total_previous)
                                        / total_previous
                                    ) * 100
                                    st.metric("Change", f"{change:+.1f}%")

                            with col_chart2:
                                # Satisfaction ratings
                                st.markdown("**Satisfaction Ratings**")
                                sat_data = pd.DataFrame(
                                    {
                                        "Type": ["Good", "Good with Comment"],
                                        "Count": [
                                            data.get("re_sat", 0)
                                            + data.get("su_sat", 0),
                                            data.get("re_sat_c", 0)
                                            + data.get("su_sat_c", 0),
                                        ],
                                    }
                                )
                                st.bar_chart(sat_data.set_index("Type"))

                        # Show tables with charts
                        for table_name, table_data in data.items():
                            if isinstance(table_data, list) and table_data:
                                st.subheader(f"Table: {table_name}")
                                df_table = pd.DataFrame(table_data)

                                col_table, col_chart = st.columns([1, 1])

                                with col_table:
                                    st.dataframe(df_table)

                                with col_chart:
                                    # Create chart based on table type
                                    if "count" in df_table.columns.str.lower().tolist():
                                        count_col = [
                                            c
                                            for c in df_table.columns
                                            if "count" in c.lower()
                                        ][0]
                                        name_col = [
                                            c
                                            for c in df_table.columns
                                            if c != count_col
                                        ][0]

                                        chart_df = df_table.head(5).set_index(name_col)
                                        st.bar_chart(chart_df[count_col])

                    # Generate PowerPoint
                    writer = PowerPointWriterV3(config)
                    output_path = writer.generate_report(data)

                    st.success(f"✓ Report generated: {output_path}")

                    # Provide download button
                    with open(output_path, "rb") as f:
                        ppt_bytes = f.read()

                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    download_name = f"cs_report_{timestamp}.pptx"

                    st.download_button(
                        label="📥 Download PowerPoint Report",
                        data=ppt_bytes,
                        file_name=download_name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True,
                    )

            except Exception as e:
                st.error(f"❌ Error generating report: {e}")
                st.exception(e)
else:
    st.info("👆 Please upload both Excel files to generate a report")

# Footer
st.markdown("---")
st.caption("CS Reporter v2.0 | Built with Streamlit")
