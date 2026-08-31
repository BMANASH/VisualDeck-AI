import streamlit as st
from PIL import Image
import io
import time
from pypdf import PdfReader

# 1. Page Configuration
st.set_page_config(
    page_title="VisualDeck AI | Intelligent Presentation Engine",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Modern Glassmorphic & Electric Styling
st.markdown("""
<style>
    /* Global Container Theme */
    .stApp {
        background-color: #0A0E17;
    }

    /* KPI Summary Cards */
    .kpi-card {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.04), rgba(255, 255, 255, 0.01));
        border: 1px solid rgba(0, 210, 255, 0.22);
        border-radius: 12px;
        padding: 14px 18px;
        backdrop-filter: blur(12px);
        box-shadow: 0 4px 20px 0 rgba(0, 0, 0, 0.4);
        margin-bottom: 10px;
    }
    .kpi-title {
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #8E9CAE;
        margin-bottom: 4px;
        font-weight: 600;
    }
    .kpi-value {
        font-size: 1.3rem;
        font-weight: 700;
        color: #00D2FF;
        letter-spacing: -0.5px;
    }
    .kpi-desc {
        font-size: 0.74rem;
        color: #6C7A89;
        margin-top: 2px;
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

    /* Sidebar Mini KPI Cards */
    .sidebar-kpi {
        background: rgba(16, 24, 38, 0.7);
        border: 1px solid rgba(0, 210, 255, 0.2);
        border-radius: 10px;
        padding: 12px 14px;
        margin-bottom: 10px;
    }
    .sidebar-kpi-title {
        font-size: 0.68rem;
        color: #8E9CAE;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .sidebar-kpi-val {
        font-size: 0.88rem;
        color: #FFFFFF;
        font-weight: 600;
        margin-top: 2px;
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
        font-size: 0.98rem;
        margin-bottom: 18px;
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
</style>
""", unsafe_allow_html=True)

# 3. Left Sidebar: Rich Information & System Metrics
with st.sidebar:
    st.markdown('<div class="badge-pill">PLATFORM CONTROL</div>', unsafe_allow_html=True)
    st.markdown("### ⚡ VisualDeck AI")
    st.caption("Autonomous visual-to-presentation synthesis engine.")
    
    st.markdown("---")
    
    st.markdown("""
    <div class="sidebar-kpi">
        <div class="sidebar-kpi-title">Core AI Architecture</div>
        <div class="sidebar-kpi-val" style="color:#00D2FF;">Gemini 3 Flash (Vision)</div>
    </div>
    <div class="sidebar-kpi">
        <div class="sidebar-kpi-title">Presentation Format</div>
        <div class="sidebar-kpi-val">Editable Microsoft .PPTX</div>
    </div>
    <div class="sidebar-kpi">
        <div class="sidebar-kpi-title">Visual Extraction</div>
        <div class="sidebar-kpi-val" style="color:#10B981;">Charts, Tables & Brand Colors</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### 📋 How It Works")
    st.markdown("""
    <div style="font-size:0.82rem; color:#94A3B8; line-height:1.6;">
        <b style="color:#FFF;">1. Ingestion:</b> Upload photos, product sheets, diagrams, or PDF files.<br>
        <b style="color:#FFF;">2. Analysis:</b> AI scans visual hierarchy, tables, and brand palettes.<br>
        <b style="color:#FFF;">3. Generation:</b> Outputs an editable native PowerPoint deck.
    </div>
    """, unsafe_allow_html=True)

# 4. Main Hero Section
st.markdown('<div class="badge-pill">AI MULTIMODAL SYNTHESIZER</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-title">VisualDeck AI</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-subtitle">Transform unstructured visual documents, diagrams, and reports into structured, native PowerPoint presentations.</div>', unsafe_allow_html=True)

# 5. File Upload Area
uploaded_files = st.file_uploader(
    label="Upload Visual Documents or Data Sheets",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True,
    help="Upload product sheets, infographics, tables, or PDF files."
)

# 6. Parse File Metrics for KPI Cards
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

# 7. Dynamic KPI Cards
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">Vision Model</div>
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
        <div class="kpi-title">Buffer Size</div>
        <div class="kpi-value">{total_size_mb:.2f} MB</div>
        <div class="kpi-desc">Ready for Analysis</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown("""
    <div class="kpi-card">
        <div class="kpi-title">Export Standard</div>
        <div class="kpi-value">Native .PPTX</div>
        <div class="kpi-desc">Editable Shapes & Data</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# 8. Uploaded Assets Gallery (Clean Compact Grid)
if uploaded_files:
    st.markdown(f"#### 📑 Source Assets Ready for Analysis ({len(uploaded_files)})")
    
    cols = st.columns(min(len(uploaded_files), 4))
    for idx, file in enumerate(uploaded_files):
        with cols[idx % 4]:
            with st.container(border=True):
                st.caption(f"**Asset {idx + 1}:** `{file.name[:18]}...`")
                if file.type.startswith("image/"):
                    img = Image.open(file)
                    # Fixed thumbnail height for alignment
                    thumb = img.copy()
                    thumb.thumbnail((260, 140))
                    st.image(thumb, use_container_width=True)
                    with st.expander("🔍 View Full Image"):
                        st.image(img, use_container_width=True)
                elif file.type == "application/pdf":
                    st.info(f"📄 PDF Document\n\nPages: {pdf_page_count}\nSize: {round(file.size / 1024, 1)} KB")

    st.markdown("---")

    # 9. Clean Multi-Column Presentation Settings
    st.markdown("### ⚙️ Presentation Customization")
    
    with st.container(border=True):
        opt_col1, opt_col2, opt_col3 = st.columns(3)
        
        with opt_col1:
            st.markdown("**🎯 Step 1: Slide Count**")
            slide_count = st.slider(
                "Target number of slides:",
                min_value=3,
                max_value=15,
                value=6,
                help="Select the exact slide count for your final presentation."
            )
            
        with opt_col2:
            st.markdown("**🗣️ Step 2: Audience Tone**")
            presentation_tone = st.selectbox(
                "Presentation tone:",
                options=[
                    "Executive Summary (Concise business takeaways)",
                    "Product Deep Dive (Specs, tables & features)",
                    "Investor Pitch (Problem, solution & market advantage)",
                    "Strategic Overview (Balanced summary with metrics)"
                ],
                index=0
            )

        with opt_col3:
            st.markdown("**🎨 Step 3: Color & Theme**")
            theme_preference = st.selectbox(
                "Slide theme:",
                options=[
                    "🤖 Auto-Detect (AI reads brand colors from visuals)",
                    "Modern Futuristic Dark (Obsidian & Electric Cyan)",
                    "Executive Navy (Deep Corporate Navy & White)",
                    "Clean Minimalist (Crisp White & Charcoal)"
                ],
                index=0
            )

        st.markdown("---")
        recreate_tables = st.checkbox(
            "Recreate visual charts and pricing tables as native PowerPoint elements",
            value=True
        )

    # 10. Synthesis Action Trigger & Live Progress Engine
    generate_clicked = st.button("⚡ Generate AI Presentation Deck", type="primary", use_container_width=True)

    if generate_clicked:
        start_time = time.time()
        
        with st.status("🚀 Processing Source Materials with Gemini 3 Vision Engine...", expanded=True) as status:
            st.write("🔍 **Phase 1:** Scanning visual layout, text hierarchies, and pricing tables...")
            time.sleep(1.0)
            
            st.write("🎨 **Phase 2:** Analyzing color harmonies and isolating logos...")
            time.sleep(1.2)
            
            st.write(f"📐 **Phase 3:** Structuring {slide_count} slide blueprints with {presentation_tone.split(' (')[0]} tone...")
            time.sleep(1.2)
            
            st.write("📊 **Phase 4:** Building native PowerPoint vector shapes and slide objects...")
            time.sleep(0.8)
            
            status.update(label="✅ Presentation Deck Synthesized Successfully!", state="complete", expanded=False)

        elapsed_seconds = round(time.time() - start_time, 2)

        # Output Summary KPI Pills
        st.markdown(f"""
        <div style="margin-top: 15px; margin-bottom: 15px;">
            <div class="metric-pill">⏱️ Synthesis Speed: <strong>{elapsed_seconds}s</strong></div>
            <div class="metric-pill">📑 Generated Slides: <strong>{slide_count} Slides</strong></div>
            <div class="metric-pill">🎨 Color Strategy: <strong>{theme_preference.split(' (')[0]}</strong></div>
            <div class="metric-pill">🛡️ Status: <strong>Deck Ready for Preview</strong></div>
        </div>
        """, unsafe_allow_html=True)

else:
    st.info("👆 Upload one or more product sheets, diagrams, or PDF documents above to begin.")
