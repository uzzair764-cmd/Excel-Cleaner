import streamlit as st

st.set_page_config(
    page_title="Data Tools",
    page_icon="whatsapp.png",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {
        display: none;
    }

    [data-testid="collapsedControl"] {
        display: none;
    }

    .block-container {
        padding-top: 4rem;
        max-width: 1300px;
    }

    .main-title {
        text-align: center;
        font-size: 3.4rem;
        font-weight: 850;
        margin-bottom: 0.4rem;
    }

    .subtitle {
        text-align: center;
        color: #777;
        font-size: 1.15rem;
        margin-bottom: 3.5rem;
    }

    div[data-testid="stPageLink"] > a {
        height: 420px;
        border-radius: 34px;
        border: 2px solid #d9d9d9;
        background: #ffffff;
        box-shadow: 0 14px 40px rgba(0,0,0,0.06);
        display: flex;
        align-items: center;
        justify-content: center;
        text-align: center;
        font-size: 2.7rem;
        font-weight: 900;
        color: #111;
        text-decoration: none;
        transition: all 0.18s ease-in-out;
    }

    div[data-testid="stPageLink"] > a:hover {
        transform: translateY(-5px);
        box-shadow: 0 22px 56px rgba(0,0,0,0.12);
        border-color: #bfbfbf;
        background: #fcfcfc;
        color: #111;
        text-decoration: none;
    }

    div[data-testid="stPageLink"] p {
        font-size: 2.7rem;
        font-weight: 900;
        letter-spacing: -0.03em;
    }

    div[data-testid="stPageLink"] svg {
        width: 3rem;
        height: 3rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="main-title">Data Tools</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Choose a tool to start processing.</div>',
    unsafe_allow_html=True
)

left, right = st.columns(2, gap="large")

with left:
    st.page_link(
        "pages/1_WA_Exporter.py",
        label="WA EXPORTER",
        icon="📤",
        use_container_width=True
    )

with right:
    st.page_link(
        "pages/2_Call_Center_Cleaner.py",
        label="CALL CENTER CLEANER",
        icon="📞",
        use_container_width=True
    )
