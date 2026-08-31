import streamlit as st
from PIL import Image
import io

# 1. Page Configuration
st.set_page_config(
    page_title="VisualDeck AI | Intelligent Presentation Engine",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Custom Futuristic Styling (CSS)
st.markdown("""
<style>
    /* Metric / KPI Card Styling */
    .kpi-card {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.05), rgba(255, 255, 255, 0.01));
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 12px;
        backdrop-filter: blur(10px);
        box-shadow: 0 4px 20px 0 rgba(0, 0, 0, 0.3);
    }
    .kpi-title {
        font-size: 0.82rem;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #8E9CAE;
        margin-bottom: 6px;
        font-weight: 600;
    }
    .kpi-value {
        font-size: 1.45rem;
        font-weight: 700;
        color: #00D2FF;
        letter-spacing: -0.5px;
    }
    .kpi-desc {
        font-size: 0.78rem;
        color: #6C7A89;
        margin-top: 4px;
    }

    /* Badge Pills */
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

    /* Header Accent */
    .hero-title {
        font-size: 2.3rem;
        font-weight: 800;
        background: linear-gradient(90deg, #FFFFFF, #92B4EC);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0px;
    }
    .hero-subtitle {
        color: #94A3B8;
        font-size: 1.05rem;
        margin-bottom: 24px;
    }
</style>
""", unsafe_allow_html=True)

# 3. Secure Backend Check (Silently handles API key)
api_key = st.secrets.get("GEMINI_API_KEY")

# 4. Sidebar: Deck Customization Panel
with st.sidebar:
    st.markdown('<div class="badge-pill">SYSTEM CONTROLS</div>', unsafe_allow_html=True)
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
        options=["Concise (3-5 slides)", "Standard (6-8 slides)", "Detailed (9-12 slides)"],
        value="Standard (6-8 slides)"
    )

    presentation_tone = st.selectbox(
        "Audience Tone",
        options=["Investor / Board Ready", "Technical Deep Dive", "Product Showcase", "Strategic Summary"],
        index=0
    )
    
    st.divider()
    
    st.markdown("### ⚙️ Engine Parameters")
    chart_recreation = st.toggle("Convert Visual Charts to Native Data Tables", value=True)
    extract_logos = st.toggle("Isolate & Reposition Brand Logos", value=True)

# 5. Main Hero Section
st.markdown('<div class="badge-pill">MULTIMODAL AI DECK SYNTHESIZER</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-title">VisualDeck AI</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-subtitle">Transform unstructured visual documents, diagrams, and reports into structured, native PowerPoint presentations.</div>', unsafe_allow_html=True)

# 6. Futuristic KPI Overview Cards
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">Vision Model</div>
        <div class="kpi-value">Gemini 2.5</div>
        <div class="kpi-desc">Multimodal Parsing Engine</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">Output Format</div>
        <div class="kpi-value">Native .PPTX</div>
        <div class="kpi-desc">Fully Editable Shapes & Text</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">Active Theme</div>
        <div class="kpi-value">""" + deck_theme.split(' (')[0] + """</div>
        <div class="kpi-desc">Adaptive Color Hierarchy</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    file_count_display = "0 Files"
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Input Buffer</div>
        <div class="kpi-value">{file_count_display}</div>
        <div class="kpi-desc">Ready for Synthesis</div>
    </div>
    """, unsafe_allow_html=True)

# Missing Key Warning (Only shows if secrets are empty)
if not api_key:
    st.warning("⚠️ Google Gemini API Key not detected. Please verify your Streamlit Secrets configuration.")

st.markdown("---")

# 7. Document & Visual Upload Section
st.markdown("### 📤 Upload Source Materials")
st.caption("Supported formats: High-resolution PNG, JPG, JPEG diagrams, charts, and multi-page PDF documents.")

uploaded_files = st.file_uploader(
    label="Drag and drop or browse files",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True,
    help="Upload one or multiple visual documents to convert into a PowerPoint deck."
)

# 8. Uploaded Content Preview Gallery
if uploaded_files:
    st.markdown(f"#### 📑 Uploaded Source Assets ({len(uploaded_files)})")
    
    preview_cols = st.columns(min(len(uploaded_files), 4))
    
    for idx, file in enumerate(uploaded_files):
        col = preview_cols[idx % 4]
        with col:
            st.markdown(f"**Asset {idx + 1}:** `{file.name}`")
            if file.type.startswith("image/"):
                img = Image.open(file)
                st.image(img, use_container_width=True)
            elif file.type == "application/pdf":
                st.info(f"📄 PDF Document\n\n({round(file.size / 1024, 1)} KB)")
    
    st.markdown("---")
    
    # Action Trigger
    btn_col1, btn_col2 = st.columns([2, 5])
    with btn_col1:
        generate_btn = st.button("Generate PowerPoint Presentation", type="primary", use_container_width=True)
    
    if generate_btn:
        st.info("Presentation generation engine will execute in the next step.")
