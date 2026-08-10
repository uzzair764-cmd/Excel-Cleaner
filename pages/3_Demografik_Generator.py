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

# ============================================================
# PARTY COLUMN ORDER
# ============================================================
# Required output order:
# UMNO, UMNO (%), PKR, PKR (%), PAS, PAS (%), PPBM, PPBM (%)
#
# dm_stats_processor.py uses these module-level constants when
# building and writing the demographic summary, so update them
# before generate_demografik() is called.
# ============================================================

PARTY_ORDER = ['UMNO', 'PKR', 'PAS', 'PPBM']


def apply_party_order():
    dm_stats.PARTY_COLS = PARTY_ORDER

    party_headers = []
    for party in PARTY_ORDER:
        party_headers.extend([party, f'{party} (%)'])

    party_start = dm_stats.HEADERS.index('PAS')
    party_end = dm_stats.HEADERS.index('UMNO (%)') + 1

    dm_stats.HEADERS = (
        dm_stats.HEADERS[:party_start]
        + party_headers
        + dm_stats.HEADERS[party_end:]
    )


if uploaded_files:
    if st.button("Generate DEMOGRAFIK", key="dm_stats_generate_button"):
        try:
            apply_party_order()

            excel_bytes, out_name, logs = dm_stats.generate_demografik(uploaded_files)

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
else:
    st.info("Upload one or more Excel files to start.")
