"""
NoKp Lookup Page
-----------------
Loads a large voter dataset (~889,490 rows) once, caches it in memory,
and lets the user look up a record instantly by NoKp (Malaysian IC number).

WHY THIS APPROACH:
- st.cache_data loads the file only once per app session (not on every rerun/search)
- Setting NoKp as the DataFrame index makes lookups O(1) instead of scanning
  all 889K rows with a boolean filter on every search
- Parquet is used for the on-disk source because it loads MUCH faster than
  .xlsx for large row counts (xlsx parsing is the main bottleneck at this scale)

SETUP:
1. Convert your source Excel file to Parquet once (see convert_to_parquet() below,
   run it locally/in Colab before deploying — don't do this conversion inside
   the Streamlit app itself).
2. Place the resulting .parquet file next to this script (or update DATA_PATH).
3. Run: streamlit run nokp_lookup_page.py
"""

import streamlit as st
import pandas as pd
from pathlib import Path

DATA_PATH = Path("voter_data.parquet")  # <-- update to your actual file


# ---------------------------------------------------------------------------
# One-time conversion helper (run this separately, NOT inside the Streamlit app)
# ---------------------------------------------------------------------------
def convert_to_parquet(xlsx_path: str, parquet_path: str):
    """
    Run this once (e.g. in Colab) to convert your source Excel file to Parquet.
    Parquet loads roughly 10-20x faster than xlsx at this row count.
    """
    df = pd.read_excel(xlsx_path, dtype=str)  # dtype=str preserves leading zeros in NoKp/phone numbers
    df.columns = [c.strip() for c in df.columns]  # guard against stray whitespace in headers
    df.to_parquet(parquet_path, index=False)
    print(f"Wrote {len(df):,} rows to {parquet_path}")


# ---------------------------------------------------------------------------
# Cached data load — this only runs once per session, not on every interaction
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Loading voter data...")
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_parquet(path)
    # Ensure NoKp is a clean string column (no leading/trailing spaces, no float artifacts)
    df["NoKp"] = df["NoKp"].astype(str).str.strip()
    # Indexing on NoKp turns lookup into a fast hash-based operation
    df = df.set_index("NoKp", drop=False)
    return df


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------
st.set_page_config(page_title="NoKp Lookup", layout="wide")
st.title("NoKp Lookup")

if not DATA_PATH.exists():
    st.error(
        f"Data file not found at `{DATA_PATH}`. "
        "Run convert_to_parquet() first to generate it from your source .xlsx file."
    )
    st.stop()

df = load_data(DATA_PATH)
st.caption(f"Loaded {len(df):,} records.")

nokp_input = st.text_input("Enter NoKp (IC Number)", max_chars=12, placeholder="e.g. 900101071234")

if nokp_input:
    query = nokp_input.strip()
    if query in df.index:
        result = df.loc[[query]]  # double brackets guard against duplicate NoKp values
        st.success(f"Found {len(result)} matching record(s).")
        st.dataframe(result, use_container_width=True)
    else:
        st.warning("No record found for that NoKp.")
