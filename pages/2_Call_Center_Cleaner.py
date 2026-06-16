import streamlit as st
from processors.call_center_processor import read_file, standardize_columns, clean_text, run_cleaner

st.set_page_config(
    page_title="Call Center Cleaner",
    page_icon="📞",
    layout="wide"
)

st.markdown("""
<style>
.block-container {
    padding-top: 2rem;
}
.big-title {
    font-size: 2.3rem;
    font-weight: 800;
    margin-bottom: 0.2rem;
}
.subtext {
    color: #666;
    margin-bottom: 2rem;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="big-title">Call Center Number Cleaner</div>', unsafe_allow_html=True)
st.markdown('<div class="subtext">Clean Excel/CSV call center files, split CSVs, and export ZIP.</div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Upload Excel or CSV files",
    type=["xlsx", "csv"],
    accept_multiple_files=True
)

with st.sidebar:
    st.header("Options")

    start_id = st.text_input(
        "Starting ID",
        value="CJ1000",
        help="Example: CJ1000 will start output from CJ1001"
    )

    chunk_size = st.number_input(
        "CSV chunk size",
        min_value=1000,
        max_value=500000,
        value=50000,
        step=1000
    )

    remove_invalid = st.toggle("Remove invalid phone numbers", value=True)
    remove_duplicates = st.toggle("Remove duplicate numbers", value=True)
    prefix_6 = st.toggle("Prefix 6 in CSV", value=True)

if uploaded_files:
    st.subheader("Preview")

    selected_file = st.selectbox(
        "Choose preview file",
        uploaded_files,
        format_func=lambda f: f.name
    )

    try:
        preview_df = read_file(selected_file)
        preview_df = standardize_columns(preview_df)
        preview_df = clean_text(preview_df)

        c1, c2 = st.columns(2)
        c1.metric("Rows", len(preview_df))
        c2.metric("Columns", len(preview_df.columns))

        with st.expander("Detected columns"):
            st.write(list(preview_df.columns))

        st.dataframe(preview_df.head(30), use_container_width=True)

    except Exception as e:
        st.error(f"Preview error: {e}")

    st.divider()

    if st.button("Process Files", type="primary", use_container_width=True):
        try:
            with st.spinner("Processing..."):
                zip_bytes, summary_df = run_cleaner(
                    uploaded_files=uploaded_files,
                    start_id=start_id,
                    chunk_size=int(chunk_size),
                    remove_invalid=remove_invalid,
                    remove_duplicates=remove_duplicates,
                    prefix_6=prefix_6
                )

            st.success("Done.")

            st.subheader("Summary")
            st.dataframe(summary_df, use_container_width=True)

            zip_name = (
                uploaded_files[0].name.rsplit(".", 1)[0] + ".zip"
                if len(uploaded_files) == 1
                else "OUTPUT.zip"
            )

            st.download_button(
                "Download ZIP",
                data=zip_bytes,
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True
            )

        except Exception as e:
            st.error(str(e))

else:
    st.info("Upload one or more files to start.")
