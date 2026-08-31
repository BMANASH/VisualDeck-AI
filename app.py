import streamlit as st

# Set page title and layout
st.set_page_config(
    page_title="VisualDeck-AI",
    page_icon="📊",
    layout="wide"
)

# App Title & Subtitle
st.title("📊 VisualDeck-AI")
st.subheader("Transform Visuals and Documents into Editable PowerPoint Decks")

st.write("Welcome to VisualDeck-AI. Upload your visual materials or documents below to get started.")

# File uploader section supporting both images and documents
uploaded_files = st.file_uploader(
    label="Upload Images (PNG, JPG, JPEG) or Documents (PDF)",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True
)

# Simple confirmation message when files are uploaded
if uploaded_files:
    st.success(f"Successfully selected {len(uploaded_files)} file(s).")
    for file in uploaded_files:
        st.write(f"- **{file.name}** ({file.type})")
