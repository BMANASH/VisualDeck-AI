import streamlit as st
from PIL import Image

# 1. Page Configuration
st.set_page_config(
    page_title="VisualDeck AI | Presentation Generator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Modern Futuristic UI Styling
st.markdown("""
<style>
    /* Metric & KPI Cards */
    .kpi-card {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.05), rgba(255, 255, 255, 0.01));
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 12px;
        padding: 16px 20px;
        backdrop-filter: blur(10px);
        box-shadow: 0 4px 20px 0 rgba(0, 0, 0, 0.3);
    }
    .kpi-title {
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #8E9CAE;
        margin-bottom: 4px;
        font-weight: 600;
    }
    .kpi-value {
        font-size: 1.35rem;
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
        margin-bottom: 8px;
    }

    /* Hero Header */
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

    /* Section Container Card */
    .settings-box {
        background: rgba(255, 255, 255, 0.02);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        padding: 20px;
        margin-top: 15px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# 3. Sidebar: Helpful User Guide & Overview
with st.sidebar:
    st.markdown('<div class="badge-pill">ABOUT VISUALDECK AI</div>', unsafe_allow_html=True)
    st.markdown("### 💡 How It Works")
    st.write(
        "VisualDeck AI turns your raw visual documents, charts, tables, and photos "
        "into clean, professionally designed, fully editable PowerPoint presentations."
    )
    
    st.markdown("---")
    st.markdown("### 📋 Quick 3-Step Process")
    st.markdown("""
    1. **Upload Files:** Add your screenshots, diagrams, photos, or PDF reports.
    2. **Set Preferences:** Pick your desired slide count, tone, and color style.
    3. **Generate & Download:** The AI reads the information, extracts logos/colors, and creates an editable `.pptx` file.
    """)
    
    st.markdown("---")
    st.markdown("### 🛡️ Smart Features")
    st.markdown("""
    - **Adaptive Colors:** Matches your brand colors automatically from images and logos.
    - **Editable Content:** Builds real PowerPoint shapes, tables, and text (never flat screenshots).
    - **Plain Language:** Transforms dense data into easy-to-read business summaries.
    """)

# 4. Main Page Header
st.markdown('<div class="badge-pill">INTELLIGENT SLIDE GENERATOR</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-title">VisualDeck AI</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-subtitle">Upload any image or document, and let AI transform it into a structured, editable PowerPoint presentation.</div>', unsafe_allow_html=True)

# 5. Calculate File Stats for KPI Cards
uploaded_files = st.file_uploader(
    label="Upload Source Materials",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True,
    key="file_uploader_main"
)

file_count = len(uploaded_files) if uploaded_files else 0
total_size_mb = sum([f.size for f in uploaded_files]) / (1024 * 1024) if uploaded_files else 0.0

# 6. Top Status & KPI Overview Cards
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">AI Engine</div>
        <div class="kpi-value">Gemini 3 Flash</div>
        <div class="kpi-desc">Vision & Layout Synthesizer</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Uploaded Files</div>
        <div class="kpi-value">{file_count} File{'s' if file_count != 1 else ''}</div>
        <div class="kpi-desc">{total_size_mb:.2f} MB Ready for Processing</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">Slide Format</div>
        <div class="kpi-value">Native .PPTX</div>
        <div class="kpi-desc">Fully Editable in PowerPoint</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">Design Engine</div>
        <div class="kpi-value">Adaptive AI</div>
        <div class="kpi-desc">Auto-Extracts Brand Colors</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# 7. Uploaded Assets Preview Gallery
if uploaded_files:
    st.markdown(f"### 📑 Uploaded Files Ready for Conversion ({len(uploaded_files)})")
    
    # Display preview in responsive grid
    preview_cols = st.columns(min(len(uploaded_files), 4))
    for idx, file in enumerate(uploaded_files):
        col = preview_cols[idx % 4]
        with col:
            st.caption(f"**File {idx + 1}:** {file.name}")
            if file.type.startswith("image/"):
                img = Image.open(file)
                st.image(img, use_container_width=True)
            elif file.type == "application/pdf":
                st.info(f"📄 PDF Document\n\nSize: {round(file.size / 1024, 1)} KB")
    
    st.markdown("---")

    # 8. Presentation Customization (Located Below Uploads)
    st.markdown("### ⚙️ Presentation Preferences")
    st.caption("Tell the AI how to structure your presentation:")
    
    custom_col1, custom_col2 = st.columns(2)
    
    with custom_col1:
        slide_count = st.slider(
            label="🎯 Number of Slides to Generate",
            min_value=3,
            max_value=15,
            value=6,
            help="Choose how many slides you want in your final presentation."
        )
        
        presentation_tone = st.selectbox(
            label="🗣️ Presentation Style & Tone",
            options=[
                "Executive Summary (Clear, concise, high-level points)",
                "Detailed Overview (In-depth analysis with full data)",
                "Visual & Data Focused (Emphasizes numbers, charts, and metrics)",
                "Story & Pitch Deck (Problem, Solution, Value Proposition)"
            ],
            index=0
        )

    with custom_col2:
        theme_preference = st.selectbox(
            label="🎨 Slide Color & Theme Style",
            options=[
                "🤖 Auto-Detect (AI reads colors from your logos and images)",
                "Modern Dark Mode (Dark background with high-contrast accents)",
                "Executive Navy (Deep corporate navy with clean white cards)",
                "Clean Minimalist (Crisp white background with dark slate text)",
                "Tech Emerald (Deep modern green with mint highlights)"
            ],
            index=0,
            help="Choose Auto-Detect for the AI to dynamically extract colors from your images."
        )
        
        recreate_tables = st.checkbox(
            label="Recreate visual tables and charts as editable PowerPoint elements",
            value=True
        )

    st.markdown("---")

    # 9. Main Action Button
    generate_col1, generate_col2 = st.columns([1, 2])
    with generate_col1:
        generate_clicked = st.button("⚡ Generate AI Presentation", type="primary", use_container_width=True)
    
    if generate_clicked:
        st.success(f"Creating a {slide_count}-slide presentation using the {theme_preference.split(' (')[0]} style...")
        st.info("The AI extraction and slide-building engine will connect here in the next step.")

else:
    # Friendly helper notice when no files are uploaded yet
    st.info("👆 Upload one or more images or PDF documents above to begin customizing your presentation.")
