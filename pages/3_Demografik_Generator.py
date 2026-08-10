import streamlit as st

from processors.dm_stats_processor import generate_demografik


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DM Stats",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# PAGE TITLE
# ============================================================

st.title("📊 DEMOGRAFIK Generator")


# ============================================================
# FILE UPLOADER
# ============================================================

uploaded_files = st.file_uploader(
    "Upload Excel file(s)",
    type=[
        "xlsx",
        "xls"
    ],
    accept_multiple_files=True,
    key="dm_stats_excel_uploader"
)


# ============================================================
# FILE PREVIEW
# ============================================================

if uploaded_files:

    st.write(
        f"**{len(uploaded_files)} file(s) selected**"
    )

    for uploaded_file in uploaded_files:

        st.write(
            f"- `{uploaded_file.name}`"
        )


# ============================================================
# GENERATE
# ============================================================

if uploaded_files:

    if st.button(
        "Generate DEMOGRAFIK",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "Processing demographic statistics..."
            ):

                excel_bytes, out_name, logs = (
                    generate_demografik(
                        uploaded_files
                    )
                )

            # ------------------------------------------------
            # Success
            # ------------------------------------------------

            st.success(
                f"Generated successfully: {out_name}"
            )

            # ------------------------------------------------
            # Workbook structure
            # ------------------------------------------------

            st.info(
                "Output worksheets: "
                "**PARLIMEN** and **DM**. "
                "The DUN worksheet has been removed."
            )

            # ------------------------------------------------
            # Processing log
            # ------------------------------------------------

            with st.expander(
                "Processing log"
            ):

                for log in logs:

                    st.write(
                        log
                    )

            # ------------------------------------------------
            # Download
            # ------------------------------------------------

            st.download_button(
                label="⬇️ Download Excel",
                data=excel_bytes,
                file_name=out_name,
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        except Exception as e:

            st.error(
                f"Generation failed: {e}"
            )

else:

    st.info(
        "Upload one or more Excel files to start."
    )
