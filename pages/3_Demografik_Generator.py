import streamlit as st
import processors.dm_stats_processor as dm_stats

st.set_page_config(
    page_title="DM Stats",
    page_icon="📊",
    layout="wide"
)

st.title("📊 DEMOGRAFIK Generator")

uploaded_files = st.file_uploader(
    "Upload Excel file(s)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="dm_stats_excel_uploader"
)

st.subheader("Age Group Configuration")

age_group_text = st.text_input(
    "Age groups",
    value="18-21, 22-30, 31-40, 41-50, 51-60, 61+",
    help="Enter age groups separated by commas. Example: 18-24, 25-30, 31-40, 41-50, 51-60, 61+"
)

age_groups = dm_stats.parse_age_groups(age_group_text)

if age_groups:
    st.caption("Active age groups: " + " | ".join(age_groups))
else:
    st.error("Invalid age group configuration. Example: 18-21, 22-30, 31-40, 41-50, 51-60, 61+")

if uploaded_files and age_groups:
    if st.button("Generate DEMOGRAFIK", key="dm_stats_generate_button"):
        try:
            excel_bytes, out_name, logs = dm_stats.generate_demografik(
                uploaded_files,
                age_groups=age_groups
            )

            st.success(f"Generated: {out_name}")

            with st.expander("Processing log"):
                for log in logs:
                    st.write(log)

            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dm_stats_download_button"
            )

        except Exception as e:
            st.error(str(e))
elif not uploaded_files:
    st.info("Upload one or more Excel files to start.")
