import streamlit as st
from processors.dm_stats_processor import generate_demografik


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


# ============================================================
# AGE GROUP INPUT
# ============================================================

st.subheader("Age Group Settings")

st.caption(
    "Enter the age ranges to be used in the DEMOGRAFIK calculation."
)

age_col1, age_col2, age_col3 = st.columns(3)

with age_col1:
    age_group_1 = st.text_input(
        "Age Group 1",
        value="18-21",
        key="dm_age_group_1"
    )

with age_col2:
    age_group_2 = st.text_input(
        "Age Group 2",
        value="22-30",
        key="dm_age_group_2"
    )

with age_col3:
    age_group_3 = st.text_input(
        "Age Group 3",
        value="31-40",
        key="dm_age_group_3"
    )


age_col4, age_col5, age_col6 = st.columns(3)

with age_col4:
    age_group_4 = st.text_input(
        "Age Group 4",
        value="41-50",
        key="dm_age_group_4"
    )

with age_col5:
    age_group_5 = st.text_input(
        "Age Group 5",
        value="51-60",
        key="dm_age_group_5"
    )

with age_col6:
    age_group_6 = st.text_input(
        "Age Group 6",
        value="61+",
        key="dm_age_group_6"
    )


age_groups = [
    age_group_1.strip(),
    age_group_2.strip(),
    age_group_3.strip(),
    age_group_4.strip(),
    age_group_5.strip(),
    age_group_6.strip()
]


# ============================================================
# GENERATE
# ============================================================

if uploaded_files:

    if st.button(
        "Generate DEMOGRAFIK",
        key="dm_stats_generate_button"
    ):

        try:

            # ------------------------------------------------
            # Validate age groups
            # ------------------------------------------------

            if any(not age for age in age_groups):
                st.error("All 6 age groups must be filled in.")
                st.stop()

            if len(set(age_groups)) != len(age_groups):
                st.error("Age groups must be unique.")
                st.stop()

            # ------------------------------------------------
            # Generate
            # ------------------------------------------------

            with st.spinner("Generating DEMOGRAFIK..."):

                excel_bytes, out_name, logs = generate_demografik(
                    uploaded_files,
                    age_groups=age_groups
                )

            # ------------------------------------------------
            # Result
            # ------------------------------------------------

            st.success(f"Generated: {out_name}")

            with st.expander("Processing log"):

                for log in logs:
                    st.write(log)

            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=out_name,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key="dm_stats_download_button"
            )

        except Exception as e:

            st.error(f"Error: {e}")

else:

    st.info("Upload one or more Excel files to start.")
