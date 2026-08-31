import streamlit as st

# Page configuration
st.set_page_config(
    page_title="VisualDeck-AI",
    page_icon="📊",
    layout="wide"
)

# App Title & Subtitle
st.title("📊 VisualDeck-AI")
st.subheader("Transform Visuals and Documents into Editable PowerPoint Decks")

# Sidebar for System Status & Settings
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # Check if Gemini API key exists in Streamlit Secrets
    api_key = st.secrets.get("GEMINI_API_KEY")
    
    if api_key:
        st.success("API Key Detected", icon="✅")
    else:
        st.error("API Key Missing", icon="⚠️")
        st.info("Add `GEMINI_API_KEY` inside Streamlit Cloud -> Settings -> Secrets.")

st.write("Welcome to VisualDeck-AI. Upload your visual materials or documents below to get started.")

# File uploader section
uploaded_files = st.file_uploader(
    label="Upload Images (PNG, JPG, JPEG) or Documents (PDF)",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"Successfully selected {len(uploaded_files)} file(s).")
    for file in uploaded_files:
        st.write(f"- **{file.name}** ({file.type})")
