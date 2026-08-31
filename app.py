import streamlit as st
from PIL import Image
import io
import json
import time
from pypdf import PdfReader
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

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
    .stApp {
        background-color: #0A0E17;
    }
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
    .slide-preview-card {
        background: rgba(16, 24, 38, 0.75);
        border: 1px solid rgba(0, 210, 255, 0.25);
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 16px;
    }
</style>
""", unsafe_allow_html=True)

# 3. Helper Functions: Colors & PowerPoint Engine
def hex_to_rgb(hex_code, default_hex="00D2FF"):
    try:
        clean_hex = hex_code.lstrip('#')
        if len(clean_hex) != 6:
            clean_hex = default_hex.lstrip('#')
        return RGBColor(*(int(clean_hex[i:i+2], 16) for i in (0, 2, 4)))
    except Exception:
        return RGBColor(0, 210, 255)

def build_powerpoint(deck_data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    palette = deck_data.get("theme", {})
    bg_color = hex_to_rgb(palette.get("background_hex", "0A0E17"))
    primary_color = hex_to_rgb(palette.get("primary_accent_hex", "00D2FF"))
    card_color = hex_to_rgb(palette.get("card_bg_hex", "161B22"))
    text_color = hex_to_rgb(palette.get("text_color_hex", "FFFFFF"))
    secondary_text = hex_to_rgb(palette.get("secondary_text_hex", "94A3B8"))

    for slide_info in deck_data.get("slides", []):
        slide = prs.slides.add_slide(blank_layout)
        
        # 1. Slide Canvas Background
        bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = bg_color
        bg_shape.line.color.rgb = bg_color

        # 2. Slide Header Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.0))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = slide_info.get("title", "Executive Overview")
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = primary_color

        if slide_info.get("subtitle"):
            p_sub = tf.add_paragraph()
            p_sub.text = slide_info.get("subtitle")
            p_sub.font.size = Pt(14)
            p_sub.font.color.rgb = secondary_text

        layout = slide_info.get("layout_type", "cards")

        # 3. Layout Type A: KPI Metric Cards
        if layout == "kpi_grid" and slide_info.get("kpis"):
            kpis = slide_info["kpis"]
            card_width = Inches(3.6)
            card_height = Inches(4.2)
            for idx, kpi in enumerate(kpis[:3]):
                left = Inches(0.8 + idx * 4.0)
                top = Inches(1.8)
                
                # Card Background Shape
                c_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_width, card_height)
                c_shape.fill.solid()
                c_shape.fill.fore_color.rgb = card_color
                c_shape.line.color.rgb = primary_color
                
                c_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.4), card_width - Inches(0.4), card_height - Inches(0.8))
                c_tf = c_box.text_frame
                c_tf.word_wrap = True
                
                p_val = c_tf.paragraphs[0]
                p_val.text = kpi.get("value", "")
                p_val.font.size = Pt(32)
                p_val.font.bold = True
                p_val.font.color.rgb = primary_color
                
                p_lbl = c_tf.add_paragraph()
                p_lbl.text = kpi.get("label", "")
                p_lbl.font.size = Pt(14)
                p_lbl.font.bold = True
                p_lbl.font.color.rgb = text_color
                
                if kpi.get("desc"):
                    p_desc = c_tf.add_paragraph()
                    p_desc.text = kpi.get("desc")
                    p_desc.font.size = Pt(11)
                    p_desc.font.color.rgb = secondary_text

        # 4. Layout Type B: Native Data Table
        elif layout == "table" and slide_info.get("table"):
            table_data = slide_info["table"]
            headers = table_data.get("headers", [])
            rows = table_data.get("rows", [])
            
            if headers and rows:
                num_rows = len(rows) + 1
                num_cols = len(headers)
                t_shape = slide.shapes.add_table(num_rows, num_cols, Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.5))
                table = t_shape.table
                
                for c_idx, head in enumerate(headers):
                    cell = table.cell(0, c_idx)
                    cell.text = str(head)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = card_color
                    for p in cell.text_frame.paragraphs:
                        p.font.size = Pt(13)
                        p.font.bold = True
                        p.font.color.rgb = primary_color
                
                for r_idx, row in enumerate(rows):
                    for c_idx, val in enumerate(row):
                        cell = table.cell(r_idx + 1, c_idx)
                        cell.text = str(val)
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = bg_color
                        for p in cell.text_frame.paragraphs:
                            p.font.size = Pt(11)
                            p.font.color.rgb = text_color

        # 5. Layout Type C: Structured Bullet Cards
        else:
            bullets = slide_info.get("bullets", [])
            card_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.8))
            card_shape.fill.solid()
            card_shape.fill.fore_color.rgb = card_color
            card_shape.line.color.rgb = primary_color
            
            b_box = slide.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(11.1), Inches(4.2))
            b_tf = b_box.text_frame
            b_tf.word_wrap = True
            
            for b_idx, bullet in enumerate(bullets):
                p = b_tf.paragraphs[0] if b_idx == 0 else b_tf.add_paragraph()
                p.text = f"• {bullet}"
                p.font.size = Pt(14)
                p.font.color.rgb = text_color
                p.space_after = Pt(12)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream

# 4. Left Sidebar Overview
with st.sidebar:
    st.markdown('<div class="badge-pill">PLATFORM CONTROL</div>', unsafe_allow_html=True)
    st.markdown("### ⚡ VisualDeck AI")
    st.caption("Autonomous visual-to-presentation synthesis engine.")
    st.markdown("---")
    
    st.markdown("""
    <div class="sidebar-kpi">
        <div class="sidebar-kpi-title">Core AI Engine</div>
        <div class="sidebar-kpi-val" style="color:#00D2FF;">Gemini 3 Flash (Vision)</div>
    </div>
    <div class="sidebar-kpi">
        <div class="sidebar-kpi-title">Export Format</div>
        <div class="sidebar-kpi-val">Editable Microsoft .PPTX</div>
    </div>
    <div class="sidebar-kpi">
        <div class="sidebar-kpi-title">Extraction Fidelity</div>
        <div class="sidebar-kpi-val" style="color:#10B981;">Shapes, Tables & Colors</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### 📋 How It Works")
    st.markdown("""
    <div style="font-size:0.82rem; color:#94A3B8; line-height:1.6;">
        <b style="color:#FFF;">1. Ingest:</b> Upload product sheets, diagrams, tables, or PDF files.<br>
        <b style="color:#FFF;">2. Analyze:</b> AI scans visual hierarchy, tables, and brand palettes.<br>
        <b style="color:#FFF;">3. Build:</b> Generates a native editable PowerPoint presentation.
    </div>
    """, unsafe_allow_html=True)

# 5. Main Hero Section
st.markdown('<div class="badge-pill">AI MULTIMODAL SYNTHESIZER</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-title">VisualDeck AI</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-subtitle">Transform unstructured visual documents, diagrams, and reports into structured, native PowerPoint presentations.</div>', unsafe_allow_html=True)

# 6. File Upload Section
uploaded_files = st.file_uploader(
    label="Upload Visual Documents or Data Sheets",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True
)

# Parse file data for KPI stats
image_count = 0
pdf_page_count = 0
total_size_mb = 0.0
loaded_images = []

if uploaded_files:
    total_size_mb = sum([f.size for f in uploaded_files]) / (1024 * 1024)
    for f in uploaded_files:
        if f.type.startswith("image/"):
            image_count += 1
            loaded_images.append(Image.open(f))
        elif f.type == "application/pdf":
            try:
                reader = PdfReader(io.BytesIO(f.getvalue()))
                pdf_page_count += len(reader.pages)
            except Exception:
                pdf_page_count += 1

# 7. Top KPI Summary Cards
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown("""<div class="kpi-card"><div class="kpi-title">Vision Model</div><div class="kpi-value">Gemini 3 Flash</div><div class="kpi-desc">Fast Visual Pipeline</div></div>""", unsafe_allow_html=True)
with col2:
    val_str = f"{len(uploaded_files)} Items" if uploaded_files else "0 Files"
    desc_str = f"{image_count} Img | {pdf_page_count} PDF Pages" if pdf_page_count > 0 else f"{image_count} Images Loaded"
    st.markdown(f"""<div class="kpi-card"><div class="kpi-title">Input Buffer</div><div class="kpi-value">{val_str}</div><div class="kpi-desc">{desc_str}</div></div>""", unsafe_allow_html=True)
with col3:
    st.markdown(f"""<div class="kpi-card"><div class="kpi-title">Buffer Size</div><div class="kpi-value">{total_size_mb:.2f} MB</div><div class="kpi-desc">Ready for Synthesis</div></div>""", unsafe_allow_html=True)
with col4:
    st.markdown("""<div class="kpi-card"><div class="kpi-title">Export Standard</div><div class="kpi-value">Native .PPTX</div><div class="kpi-desc">Editable Shapes & Data</div></div>""", unsafe_allow_html=True)

st.markdown("---")

# 8. Uploaded Assets Gallery (Compact Preview Grid)
if uploaded_files:
    st.markdown(f"#### 📑 Source Assets Ready for Analysis ({len(uploaded_files)})")
    cols = st.columns(min(len(uploaded_files), 4))
    for idx, file in enumerate(uploaded_files):
        with cols[idx % 4]:
            with st.container(border=True):
                st.caption(f"**Asset {idx + 1}:** `{file.name[:18]}...`")
                if file.type.startswith("image/"):
                    img = Image.open(file)
                    thumb = img.copy()
                    thumb.thumbnail((260, 120))
                    st.image(thumb, use_container_width=True)
                    with st.expander("🔍 View Full Image"):
                        st.image(img, use_container_width=True)
                elif file.type == "application/pdf":
                    st.info(f"📄 PDF Document\n\nPages: {pdf_page_count}\nSize: {round(file.size / 1024, 1)} KB")

    st.markdown("---")

    # 9. Unified Presentation Customization Console
    st.markdown("### ⚙️ Presentation Customization")
    with st.container(border=True):
        opt_col1, opt_col2, opt_col3 = st.columns(3)
        with opt_col1:
            st.markdown("**🎯 Step 1: Slide Count**")
            slide_count = st.slider("Target number of slides:", min_value=3, max_value=12, value=4)
        with opt_col2:
            st.markdown("**🗣️ Step 2: Audience Tone**")
            presentation_tone = st.selectbox(
                "Presentation tone:",
                ["Executive Summary", "Product Deep Dive", "Investor Pitch", "Strategic Overview"],
                index=0
            )
        with opt_col3:
            st.markdown("**🎨 Step 3: Color & Theme**")
            theme_preference = st.selectbox(
                "Slide theme:",
                ["Auto-Detect (Extract from Visuals)", "Modern Dark", "Executive Navy", "Clean Minimalist"],
                index=0
            )

    # 10. AI Generation Engine Execution
    generate_clicked = st.button("⚡ Generate AI Presentation Deck", type="primary", use_container_width=True)

    if generate_clicked:
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("⚠️ Gemini API Key not found in Streamlit Secrets. Please configure `GEMINI_API_KEY`.")
        elif not loaded_images:
            st.warning("Please upload at least one image to synthesize slides.")
        else:
            start_time = time.time()
            with st.status("🚀 Processing with Gemini 3 Vision Engine...", expanded=True) as status:
                st.write("🔍 **Phase 1:** Scanning visual layout, tables, and specifications...")
                
                # Setup Gemini Client
                client = genai.Client(api_key=api_key)
                
                # Structured Presentation Prompt
                prompt = f"""
                You are an expert executive presentation designer. Analyze these uploaded source images carefully.
                Extract key data, specifications, tables, pricing, and visual themes.
                Generate a {slide_count}-slide presentation outline matching this audience tone: '{presentation_tone}'.
                
                Output ONLY a valid JSON object matching this exact schema:
                {{
                  "theme": {{
                    "background_hex": "#0A0E17",
                    "primary_accent_hex": "#00D2FF",
                    "card_bg_hex": "#161B22",
                    "text_color_hex": "#FFFFFF",
                    "secondary_text_hex": "#94A3B8"
                  }},
                  "deck_title": "Presentation Title",
                  "slides": [
                    {{
                      "slide_number": 1,
                      "layout_type": "kpi_grid",
                      "title": "Executive Summary & Core Metrics",
                      "subtitle": "High-level overview",
                      "kpis": [
                        {{"label": "Key Metric", "value": "Value", "desc": "Brief note"}}
                      ]
                    }},
                    {{
                      "slide_number": 2,
                      "layout_type": "table",
                      "title": "Specifications & Pricing Breakdown",
                      "subtitle": "Detailed comparison",
                      "table": {{
                        "headers": ["Item / Model", "Capacity / Spec", "Investment / Price"],
                        "rows": [
                          ["Model A", "Specs", "Price"],
                          ["Model B", "Specs", "Price"]
                        ]
                      }}
                    }},
                    {{
                      "slide_number": 3,
                      "layout_type": "bullets",
                      "title": "Strategic Advantages & Value Proposition",
                      "subtitle": "Key highlights",
                      "bullets": [
                        "Highlight point 1 extracted from materials",
                        "Highlight point 2 extracted from materials",
                        "Highlight point 3 extracted from materials"
                      ]
                    }}
                  ]
                }}
                Do not include markdown backticks or any conversational text. Return only the raw JSON.
                """

                # Call Gemini Vision with model fallback
                response_json = None
                models_to_try = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash']
                for m in models_to_try:
                    try:
                        content_payload = [prompt] + loaded_images[:5]
                        res = client.models.generate_content(
                            model=m,
                            contents=content_payload
                        )
                        raw_text = res.text.strip()
                        if raw_text.startswith("```json"):
                            raw_text = raw_text[7:-3].strip()
                        elif raw_text.startswith("```"):
                            raw_text = raw_text[3:-3].strip()
                        response_json = json.loads(raw_text)
                        break
                    except Exception:
                        continue

                if not response_json:
                    status.update(label="❌ Failed to parse presentation data. Please retry.", state="error")
                    st.error("Could not process presentation schema. Please check API quota or try with fewer images.")
                else:
                    st.write("📊 **Phase 2:** Synthesizing native PowerPoint vector shapes and slide objects...")
                    pptx_stream = build_powerpoint(response_json)
                    st.session_state["generated_pptx"] = pptx_stream
                    st.session_state["deck_data"] = response_json
                    status.update(label="✅ Presentation Deck Synthesized Successfully!", state="complete", expanded=False)

    # 11. Display Slide Preview & Download Options
    if "generated_pptx" in st.session_state and "deck_data" in st.session_state:
        deck = st.session_state["deck_data"]
        slides = deck.get("slides", [])
        
        st.markdown("---")
        st.markdown(f"### 📑 Presentation Preview: **{deck.get('deck_title', 'Synthesized Deck')}**")
        
        # Download Button
        st.download_button(
            label="📥 Download Native .PPTX Presentation",
            data=st.session_state["generated_pptx"],
            file_name="VisualDeck_AI_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary"
        )
        
        # Interactive Slide Previews
        for s in slides:
            with st.container(border=True):
                st.markdown(f"#### Slide {s.get('slide_number', '')}: {s.get('title', '')}")
                if s.get("subtitle"):
                    st.caption(f"_{s.get('subtitle')}_")
                
                if s.get("layout_type") == "kpi_grid" and s.get("kpis"):
                    kpi_cols = st.columns(len(s["kpis"]))
                    for idx, k in enumerate(s["kpis"]):
                        with kpi_cols[idx]:
                            st.metric(label=k.get("label", "Metric"), value=k.get("value", "N/A"))
                
                elif s.get("layout_type") == "table" and s.get("table"):
                    headers = s["table"].get("headers", [])
                    rows = s["table"].get("rows", [])
                    if headers and rows:
                        import pandas as pd
                        df = pd.DataFrame(rows, columns=headers)
                        st.dataframe(df, use_container_width=True)
                
                elif s.get("bullets"):
                    for b in s["bullets"]:
                        st.markdown(f"- {b}")

else:
    st.info("👆 Upload one or more product sheets, diagrams, or PDF documents above to begin.")
