import streamlit as st
from PIL import Image

# 1. Page Configuration
st.set_page_config(
    page_title="VisualDeck AI | Presentation Engine",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Custom Futuristic Styling (CSS)
st.markdown("""
<style>
    .kpi-card {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.05), rgba(255, 255, 255, 0.01));
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 12px;
        padding: 16px 20px;
        backdrop-filter: blur(10px);
        box-shadow: 0 4px 20px 0 rgba(0, 0, 0, 0.3);
    }
    .kpi-title {
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #8E9CAE;
        margin-bottom: 4px;
        font-weight: 600;
    }
    .kpi-value {
        font-size: 1.4rem;
        font-weight: 700;
        color: #00D2FF;
        letter-spacing: -0.5px;
    }
    .kpi-desc {
        font-size: 0.78rem;
        color: #6C7A89;
        margin-top: 4px;
    }
    .badge-pill {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 50px;
        font-size: 0.75rem;
        font-weight: 600;
        letter-spacing: 0.5px;
        background: rgba(0, 210, 255, 0.1);
        border: 1px solid rgba(0, 210, 255, 0.3);
        color: #00D2FF;
        margin-bottom: 10px;
    }
    .hero-title {
        font-size: 2.2rem;
        font-weight: 800;
        background: linear-gradient(90deg, #FFFFFF, #92B4EC);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0px;
    }
    .hero-subtitle {
        color: #94A3B8;
        font-size: 1rem;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# 3. Sidebar: Deck Customization Controls
with st.sidebar:
    st.markdown('<div class="badge-pill">DECK CONFIGURATION</div>', unsafe_allow_html=True)
    st.markdown("### 🎨 Presentation Setup")
    
    deck_theme = st.selectbox(
        "Visual Deck Theme",
        options=[
            "Modern Futuristic (Dark Neon)",
            "Executive Leadership (Deep Navy)",
            "Clean Minimalist (Monochrome Light)",
            "Emerald Growth (Tech Green)"
        ],
        index=0
    )
    
    slide_length = st.select_slider(
        "Target Slide Count",
        options=["3-5 Slides (Brief)", "6-8 Slides (Standard)", "9-12 Slides (Detailed)"],
        value="6-8 Slides (Standard)"
    )

    presentation_tone = st.selectbox(
        "Audience Tone",
        options=["Investor / Board Ready", "Technical Deep Dive", "Product Showcase", "Strategic Summary"],
        index=0
    )
    
    st.divider()
    
    st.markdown("### ⚙️ Extraction Parameters")
    chart_recreation = st.toggle("Convert Charts to Native Data Tables", value=True)
    extract_logos = st.toggle("Extract & Reposition Brand Logos", value=True)

# 4. Main Hero Section
st.markdown('<div class="badge-pill">MULTIMODAL AI DECK SYNTHESIZER</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-title">VisualDeck AI</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-subtitle">Transform visual reports, charts, and documents into structured, editable PowerPoint presentations.</div>', unsafe_allow_html=True)

# 5. File Upload Section
st.markdown("### 📤 Upload Source Materials")
st.caption("Supported formats: PNG, JPG, JPEG diagrams, and multi-page PDF documents.")

uploaded_files = st.file_uploader(
    label="Upload Images or Documents",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

# 6. Dynamic KPI Cards (Calculated from uploaded files)
file_count = len(uploaded_files) if uploaded_files else 0
total_size_mb = sum([f.size for f in uploaded_files]) / (1024 * 1024) if uploaded_files else 0.0

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">AI Vision Model</div>
        <div class="kpi-value">Gemini 3 Flash</div>
        <div class="kpi-desc">Fast Multimodal Pipeline</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Source Assets</div>
        <div class="kpi-value">{file_count} File{'s' if file_count != 1 else ''}</div>
        <div class="kpi-desc">{total_size_mb:.2f} MB Loaded</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Target Output</div>
        <div class="kpi-value">{slide_length.split(' ')[0]}</div>
        <div class="kpi-desc">{presentation_tone.split(' /')[0]}</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    theme_name = deck_theme.split(' (')[0]
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Design Theme</div>
        <div class="kpi-value">{theme_name}</div>
        <div class="kpi-desc">Native Shape Styling</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# 7. Uploaded Asset Preview Gallery
if uploaded_files:
    st.markdown(f"#### 📑 Source Assets Ready for Analysis ({len(uploaded_files)})")
    
    preview_cols = st.columns(min(len(uploaded_files), 4))
    
    for idx, file in enumerate(uploaded_files):
        col = preview_cols[idx % 4]
        with col:
            st.markdown(f"**Asset {idx + 1}:** `{file.name}`")
            if file.type.startswith("image/"):
                img = Image.open(file)
                st.image(img, use_container_width=True)
            elif file.type == "application/pdf":
                st.info(f"📄 PDF Document\n\nSize: {round(file.size / 1024, 1)} KB")
    
    st.markdown("---")
    
    # Action Trigger
    btn_col1, btn_col2 = st.columns([2, 5])
    with btn_col1:
        generate_btn = st.button("⚡ Synthesize Presentation Deck", type="primary", use_container_width=True)
    
    if generate_btn:
        st.info("AI Analysis and Slide Generation Engine will run in the next step.")
