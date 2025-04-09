import streamlit as st

st.set_page_config(
    page_title="Hello Navler",
    page_icon="👋",
)

st.write("# Welcome to Navla SEO Tools! 👋")

st.sidebar.success("Select a tool above.")

# Sidebar per la configurazione globale
with st.sidebar:
    st.header("🔐 Configurazione")
    if 'openai_api_key' not in st.session_state:
        st.session_state['openai_api_key'] = ""

    st.session_state['openai_api_key'] = st.text_input(
        "Inserisci la tua OpenAI API Key",
        type="password",
        value=st.session_state['openai_api_key']
    )