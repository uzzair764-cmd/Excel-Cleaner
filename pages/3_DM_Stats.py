import os
import io
import zipfile
import streamlit as st

from processors.dm_stats_processor import (
    DRIVE_ROOT,
    list_dirs,
    list_excel_files,
    process_stats_file,
    make_excel_bytes,
)

st.set_page_config(
    page_title="DM Stats",
    page_icon="📊",
    layout="wide"
)

st.title("📊 DM Stats")

st.caption(f"Drive root: {DRIVE_ROOT}")

mode = st.radio(
    "Choose input method",
    ["Upload Excel files", "Browse Google Drive folder"],
    horizontal=True
)

selected_drive_files = []
uploaded_files = []

if mode == "Browse Google Drive folder":
    if not os.path.exists(DRIVE_ROOT):
        st.error("Google Drive path not found. This only works when Streamlit is running inside Colab or a machine with Drive mounted.")
        st.stop()

    current_path = DRIVE_ROOT

    folders_1 = list_dirs(current_path)
    selected_1 = st.selectbox("Select folder", [""] + folders_1)

    if selected_1:
        current_path = os.path.join(current_path, selected_1)

    folders_2 = list_dirs(current_path)
    selected_2 = st.selectbox("Select negeri / folder", [""] + folders_2)

    if selected_2:
        current_path = os.path.join(current_path, selected_2)

    folders_3 = list_dirs(current_path)
    selected_3 = st.selectbox("Select subfolder", [""] + folders_3)

    if selected_3:
        current_path = os.path.join(current_path, selected_3)

    folders_4 = list_dirs(current_path)
    selected_4 = st.selectbox("Select deeper subfolder", [""] + folders_4)

    if selected_4:
        current_path = os.path.join(current_path, selected_4)

    st.info(f"Current folder: {current_path}")

    excel_files = list_excel_files(current_path)

    if not excel_files:
        st.warning("No Excel files found in this folder.")
        st.stop()

    select_all = st.checkbox("Select all Excel files")

    selected_files = st.multiselect(
        "Select Excel file(s)",
        excel_files,
        default=excel_files if select_all else []
    )

    selected_drive_files = [
        (fname, os.path.join(current_path, fname))
        for fname in selected_files
    ]

else:
    uploaded_files = st.file_uploader(
        "Upload Excel file(s)",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

if st.button("Generate DM Stats"):
    zip_buffer = io.BytesIO()
    success_count = 0

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:

        if mode == "Browse Google Drive folder":
            files_to_process = selected_drive_files

            if not files_to_process:
                st.warning("Please select at least one Excel file.")
                st.stop()

            for fname, fpath in files_to_process:
                try:
                    with open(fpath, "rb") as f:
                        df_parl, df_dun, df_dm = process_stats_file(f)

                    excel_bytes = make_excel_bytes(df_parl, df_dun, df_dm)

                    out_name = fname.rsplit(".", 1)[0] + "_stats.xlsx"
                    zipf.writestr(out_name, excel_bytes)

                    success_count += 1
                    st.success(f"Done: {fname}")

                except Exception as e:
                    st.error(f"Skipped {fname}: {e}")

        else:
            if not uploaded_files:
                st.warning("Please upload at least one Excel file.")
                st.stop()

            for uploaded_file in uploaded_files:
                try:
                    df_parl, df_dun, df_dm = process_stats_file(uploaded_file)
                    excel_bytes = make_excel_bytes(df_parl, df_dun, df_dm)

                    out_name = uploaded_file.name.rsplit(".", 1)[0] + "_stats.xlsx"
                    zipf.writestr(out_name, excel_bytes)

                    success_count += 1
                    st.success(f"Done: {uploaded_file.name}")

                except Exception as e:
                    st.error(f"Skipped {uploaded_file.name}: {e}")

    if success_count:
        zip_buffer.seek(0)

        st.download_button(
            label="Download DM Stats ZIP",
            data=zip_buffer.getvalue(),
            file_name="DM_STATS_OUTPUT.zip",
            mime="application/zip"
        )
