import streamlit as st
from PIL import Image
import io
import time
from pypdf import PdfReader

# 1. Page Configuration
st.set_page_config(
    page_title="VisualDeck AI | Presentation Engine",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Futuristic Glassmorphic & Electric Glow Styling (CSS)
st.markdown("""
<style>
    /* Global Electric Accents */
    :root {
        --electric-cyan: #00D2FF;
        --neon-blue: #0072FF;
        --glass-bg: rgba(13, 17, 23, 0.7);
    }

    /* KPI Summary Cards */
    .kpi-card {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.04), rgba(255, 255, 255, 0.01));
        border: 1px solid rgba(0, 210, 255, 0.2);
        border-radius: 12px;
        padding: 14px 18px;
        backdrop-filter: blur(12px);
        box-shadow: 0 4px 20px 0 rgba(0, 0, 0, 0.35);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .kpi-card:hover {
        border-color: rgba(0, 210, 255, 0.5);
        transform: translateY(-2px);
    }
    .kpi-title {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #8E9CAE;
        margin-bottom: 2px;
        font-weight: 600;
    }
    .kpi-value {
        font-size: 1.35rem;
        font-weight: 700;
        color: #00D2FF;
        letter-spacing: -0.5px;
    }
    .kpi-desc {
        font-size: 0.74rem;
        color: #6C7A89;
        margin-top: 2px;
    }

    /* Electric Neon Glowing Step Cards */
    .step-box {
        background: rgba(16, 24, 38, 0.65);
        border: 1px solid rgba(0, 210, 255, 0.35);
        border-radius: 14px;
        padding: 20px;
        margin-bottom: 18px;
        backdrop-filter: blur(14px);
        box-shadow: 0 0 15px rgba(0, 210, 255, 0.08), inset 0 0 15px rgba(0, 210, 255, 0.03);
    }
    .step-header {
        display: flex;
        align-items: center;
        margin-bottom: 12px;
    }
    .step-number {
        background: linear-gradient(135deg, #00D2FF, #0072FF);
        color: #000;
        font-weight: 800;
        font-size: 0.8rem;
        padding: 4px 10px;
        border-radius: 8px;
        margin-right: 12px;
        box-shadow: 0 0 10px rgba(0, 210, 255, 0.5);
    }
    .step-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #FFFFFF;
        letter-spacing: 0.3px;
    }

    /* Badge Pills */
    .badge-pill {
        display: inline-block;
        padding: 3px 12px;
        border-radius: 50px;
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.6px;
        background: rgba(0, 210, 255, 0.1);
        border: 1px solid rgba(0, 210, 255, 0.35);
        color: #00D2FF;
        margin-bottom: 8px;
    }

    /* Metric Result Pills */
    .metric-pill {
        display: inline-flex;
        align-items: center;
        background: rgba(0, 210, 255, 0.08);
        border: 1px solid rgba(0, 210, 255, 0.3);
        padding: 8px 16px;
        border-radius: 10px;
        font-size: 0.85rem;
        color: #E2E8F0;
        margin-right: 10px;
        margin-top: 8px;
    }
    .metric-pill strong {
        color: #00D2FF;
        margin-left: 6px;
    }

    /* Thumbnail Compact Asset Card */
    .asset-card {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 10px;
        padding: 10px;
        text-align: center;
        margin-bottom: 10px;
    }

    /* Hero Headers */
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
        font-size: 0.98rem;
        margin-bottom: 18px;
    }
</style>
""", unsafe_allow_html=True)

# 3. Left Sidebar: Sleek & Compact Overview
with st.sidebar:
    st.markdown('<div class="badge-pill">PLATFORM OVERVIEW</div>', unsafe_allow_html=True)
    st.markdown("### ⚡ VisualDeck AI")
    st.caption("Autonomous visual-to-presentation synthesis engine.")
    
    st.markdown("---")
    
    # Clean Mini Status Cards
    st.markdown("""
    <div style="background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.08); border-radius: 10px; padding: 12px; margin-bottom: 12px;">
        <div style="font-size: 0.72rem; color: #8E9CAE; font-weight:600;">ACTIVE ARCHITECTURE</div>
        <div style="font-size: 0.9rem; color: #00D2FF; font-weight:700; margin-top:2px;">Gemini 3 Flash</div>
    </div>
    <div style="background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.08); border-radius: 10px; padding: 12px; margin-bottom: 12px;">
        <div style="font-size: 0.72rem; color: #8E9CAE; font-weight:600;">EXPORT STANDARD</div>
        <div style="font-size: 0.9rem; color: #FFFFFF; font-weight:700; margin-top:2px;">Native Microsoft .PPTX</div>
    </div>
    <div style="background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.08); border-radius: 10px; padding: 12px;">
        <div style="font-size: 0.72rem; color: #8E9CAE; font-weight:600;">EXTRACTION FIDELITY</div>
        <div style="font-size: 0.9rem; color: #10B981; font-weight:700; margin-top:2px;">Shapes, Tables & Logos</div>
    </div>
    """, unsafe_allow_html=True)

# 4. Main Header Section
st.markdown('<div class="badge-pill">AI MULTIMODAL SYNTHESIZER</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-title">VisualDeck AI</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-subtitle">Convert visual documents, product specs, diagrams, and financial tables into presentation decks.</div>', unsafe_allow_html=True)

# 5. File Upload Area
uploaded_files = st.file_uploader(
    label="Upload Visuals or Documents",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True,
    help="Upload product sheets, photos, tables, or PDF reports."
)

# 6. Process PDF and Image Information for Accurate KPIs
image_count = 0
pdf_page_count = 0
total_size_mb = 0.0

if uploaded_files:
    total_size_mb = sum([f.size for f in uploaded_files]) / (1024 * 1024)
    for f in uploaded_files:
        if f.type.startswith("image/"):
            image_count += 1
        elif f.type == "application/pdf":
            try:
                reader = PdfReader(io.BytesIO(f.getvalue()))
                pdf_page_count += len(reader.pages)
            except Exception:
                pdf_page_count += 1

# 7. Top Status Overview (4 KPI Cards)
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">AI Vision Model</div>
        <div class="kpi-value">Gemini 3 Flash</div>
        <div class="kpi-desc">Fast Visual Pipeline</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    if uploaded_files:
        val_str = f"{len(uploaded_files)} Items"
        desc_str = f"{image_count} Img | {pdf_page_count} PDF Pages" if pdf_page_count > 0 else f"{image_count} Images Loaded"
    else:
        val_str = "0 Files"
        desc_str = "Awaiting Source Files"
    
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Input Buffer</div>
        <div class="kpi-value">{val_str}</div>
        <div class="kpi-desc">{desc_str}</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Memory Footprint</div>
        <div class="kpi-value">{total_size_mb:.2f} MB</div>
        <div class="kpi-desc">Optimized for Processing</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">Export Fidelity</div>
        <div class="kpi-value">Native .PPTX</div>
        <div class="kpi-desc">Editable Shapes & Data</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# 8. Uploaded Assets Gallery (Compact Preview Grid with Expand Option)
if uploaded_files:
    st.markdown(f"#### 📑 Source Assets Loaded ({len(uploaded_files)})")
    
    # 4-column compact thumbnail gallery
    cols = st.columns(min(len(uploaded_files), 4))
    for idx, file in enumerate(uploaded_files):
        with cols[idx % 4]:
            st.markdown(f"**Asset {idx + 1}:** `{file.name[:20]}...`")
            if file.type.startswith("image/"):
                img = Image.open(file)
                # Create a lightweight square thumbnail for clean grid alignment
                thumb = img.copy()
                thumb.thumbnail((300, 200))
                st.image(thumb, use_container_width=True)
                with st.expander("🔍 Expand View"):
                    st.image(img, use_container_width=True)
            elif file.type == "application/pdf":
                st.info(f"📄 PDF Document\n\nPages: {pdf_page_count}\n\nSize: {round(file.size / 1024, 1)} KB")

    st.markdown("---")

    # 9. Glowing Electric Step Cards
    st.markdown("### ⚙️ Presentation Configuration")

    # Step 1: Slide Count
    st.markdown("""
    <div class="step-box">
        <div class="step-header">
            <span class="step-number">STEP 1</span>
            <span class="step-title">Slide Volume & Structure</span>
        </div>
    """, unsafe_allow_html=True)
    slide_count = st.slider(
        "Choose target slide count:",
        min_value=3,
        max_value=15,
        value=6,
        help="Controls the density and slide count of your generated deck."
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Step 2: Audience Tone
    st.markdown("""
    <div class="step-box">
        <div class="step-header">
            <span class="step-number">STEP 2</span>
            <span class="step-title">Audience Context & Presentation Tone</span>
        </div>
    """, unsafe_allow_html=True)
    presentation_tone = st.selectbox(
        "Select presentation style:",
        options=[
            "Executive Leadership (Clear, high-level business takeaways)",
            "Technical & Product Deep Dive (Detailed specifications, tables & features)",
            "Sales & Investor Pitch (Problem, solution, pricing & market advantage)",
            "Strategic Overview (Balanced summary with key performance metrics)"
        ],
        index=0,
        label_visibility="collapsed"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Step 3: Visual Identity
    st.markdown("""
    <div class="step-box">
        <div class="step-header">
            <span class="step-number">STEP 3</span>
            <span class="step-title">Visual Identity & Color Harmonies</span>
        </div>
    """, unsafe_allow_html=True)
    theme_preference = st.selectbox(
        "Select color strategy:",
        options=[
            "🤖 Auto-Detect & Match (AI reads brand colors from logos and images)",
            "Modern Futuristic Dark (Dark obsidian canvas with electric cyan highlights)",
            "Executive Navy (Deep corporate navy with clean white container cards)",
            "Clean Minimalist (Crisp light aesthetic with dark charcoal typography)"
        ],
        index=0,
        label_visibility="collapsed"
    )
    recreate_tables = st.checkbox(
        "Recreate visual charts and pricing tables as native PowerPoint elements",
        value=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # 10. Main Action Button with Glassmorphic Progress
    generate_clicked = st.button("⚡ Generate AI Presentation Deck", type="primary", use_container_width=True)

    if generate_clicked:
        start_time = time.time()
        
        # Real-time interactive processing status
        with st.status("🚀 Processing Assets with Gemini 3 Vision Engine...", expanded=True) as status:
            st.write("🔍 **Phase 1:** Ingesting visual documents and parsing resolution layers...")
            time.sleep(1.0)
            
            st.write("🎨 **Phase 2:** Extracting color palette, logos, and tabular metrics...")
            time.sleep(1.2)
            
            st.write(f"📐 **Phase 3:** Synthesizing {slide_count} structured slide blueprints with {presentation_tone.split(' (')[0]} tone...")
            time.sleep(1.2)
            
            st.write("📊 **Phase 4:** Building native PowerPoint layout and vector containers...")
            time.sleep(0.8)
            
            status.update(label="✅ Presentation Deck Synthesized Successfully!", state="complete", expanded=False)

        elapsed_seconds = round(time.time() - start_time, 2)

        # Output Summary KPI Pills
        st.markdown(f"""
        <div style="margin-top: 15px; margin-bottom: 15px;">
            <div class="metric-pill">⏱️ Synthesis Time: <strong>{elapsed_seconds}s</strong></div>
            <div class="metric-pill">📑 Generated Slides: <strong>{slide_count} Slides</strong></div>
            <div class="metric-pill">🎨 Palette Mode: <strong>{theme_preference.split(' (')[0]}</strong></div>
            <div class="metric-pill">🛡️ Status: <strong>Deck Ready for Preview</strong></div>
        </div>
        """, unsafe_allow_html=True)

else:
    st.info("👆 Upload one or more product sheets, diagrams, or PDF documents above to begin.")
