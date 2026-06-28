import streamlit as st
from processors.dm_stats_processor import generate_demografik

st.set_page_config(
    page_title="DM Stats",
    page_icon="📊",
    layout="wide"
)

st.title("📊 DM Stats / DEMOGRAFIK Generator")

uploaded_files = st.file_uploader(
    "Upload Excel file(s)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Generate DEMOGRAFIK"):
        try:
            excel_bytes, out_name, logs = generate_demografik(uploaded_files)

            st.success(f"Generated: {out_name}")

            with st.expander("Processing log"):
                for log in logs:
                    st.write(log)

            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(str(e))
else:
    st.info("Upload one or more Excel files to start.")
