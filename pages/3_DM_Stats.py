import streamlit as st
from processors.dm_stats_processor import render_dm_stats_tool

st.set_page_config(
    page_title="DM Stats",
    page_icon="📊",
    layout="wide"
)

render_dm_stats_tool()
