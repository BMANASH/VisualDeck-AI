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
from pptx.enum.shapes import MSO_SHAPE
import pandas as pd

# 1. Page Configuration
st.set_page_config(
    page_title="VisualDeck AI | Intelligent Presentation Engine",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Modern Glassmorphic & Electric Styling (CSS)
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
    .slide-deck-frame {
        background: #0D131F;
        border: 1px solid rgba(0, 210, 255, 0.3);
        border-radius: 14px;
        padding: 24px;
        margin-bottom: 22px;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.45);
    }
</style>
""", unsafe_allow_html=True)

# 3. Helper Functions: Colors & PowerPoint Engine
def hex_to_rgb(hex_code, default_hex="00D2FF"):
    try:
        clean_hex = str(hex_code).lstrip('#')
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
    card_color = hex_to_rgb(palette.get("card_bg_hex", "141C2E"))
    text_color = hex_to_rgb(palette.get("text_color_hex", "FFFFFF"))
    secondary_text = hex_to_rgb(palette.get("secondary_text_hex", "94A3B8"))

    for slide_info in deck_data.get("slides", []):
        slide = prs.slides.add_slide(blank_layout)
        
        # 1. Slide Canvas Background
        bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = bg_color
        bg_shape.line.color.rgb = bg_color

        # 2. Modern Accent Line Under Canvas Top
        accent_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.04))
        accent_line.fill.solid()
        accent_line.fill.fore_color.rgb = primary_color
        accent_line.line.color.rgb = primary_color

        # 3. Slide Header Title & Subtitle
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.733), Inches(1.1))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = slide_info.get("title", "Executive Overview")
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = primary_color

        if slide_info.get("subtitle"):
            p_sub = tf.add_paragraph()
            p_sub.text = slide_info.get("subtitle")
            p_sub.font.size = Pt(12)
            p_sub.font.color.rgb = secondary_text

        layout = slide_info.get("layout_type", "cards")

        # 4. Dynamic Layout: KPI Metric Cards
        if layout == "kpi_grid" and slide_info.get("kpis"):
            kpis = slide_info["kpis"]
            card_count = min(len(kpis), 3)
            total_avail_width = 11.733
            card_width = (total_avail_width - (card_count - 1) * 0.4) / card_count
            card_height = 4.4
            
            for idx, kpi in enumerate(kpis[:card_count]):
                left = 0.8 + idx * (card_width + 0.4)
                top = 1.9
                
                c_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(card_width), Inches(card_height))
                c_shape.fill.solid()
                c_shape.fill.fore_color.rgb = card_color
                c_shape.line.color.rgb = primary_color
                
                c_box = slide.shapes.add_textbox(Inches(left + 0.25), Inches(top + 0.4), Inches(card_width - 0.5), Inches(card_height - 0.8))
                c_tf = c_box.text_frame
                c_tf.word_wrap = True
                
                p_val = c_tf.paragraphs[0]
                p_val.text = str(kpi.get("value", ""))
                p_val.font.size = Pt(30)
                p_val.font.bold = True
                p_val.font.color.rgb = primary_color
                
                p_lbl = c_tf.add_paragraph()
                p_lbl.text = str(kpi.get("label", ""))
                p_lbl.font.size = Pt(13)
                p_lbl.font.bold = True
                p_lbl.font.color.rgb = text_color
                
                if kpi.get("desc"):
                    p_desc = c_tf.add_paragraph()
                    p_desc.text = str(kpi.get("desc"))
                    p_desc.font.size = Pt(10)
                    p_desc.font.color.rgb = secondary_text

        # 5. Dynamic Layout: Native PowerPoint Table
        elif layout == "table" and slide_info.get("table"):
            table_data = slide_info["table"]
            headers = table_data.get("headers", [])
            rows = table_data.get("rows", [])
            
            if headers and rows:
                num_rows = len(rows) + 1
                num_cols = len(headers)
                t_shape = slide.shapes.add_table(num_rows, num_cols, Inches(0.8), Inches(1.9), Inches(11.733), Inches(4.5))
                table = t_shape.table
                
                for c_idx, head in enumerate(headers):
                    cell = table.cell(0, c_idx)
                    cell.text = str(head)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = card_color
                    for p in cell.text_frame.paragraphs:
                        p.font.size = Pt(12)
                        p.font.bold = True
                        p.font.color.rgb = primary_color
                
                for r_idx, row in enumerate(rows):
                    for c_idx, val in enumerate(row):
                        cell = table.cell(r_idx + 1, c_idx)
                        cell.text = str(val)
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = bg_color
                        for p in cell.text_frame.paragraphs:
                            p.font.size = Pt(10)
                            p.font.color.rgb = text_color

        # 6. Dynamic Layout: Visual Cards & Highlight Bullets
        else:
            bullets = slide_info.get("bullets", [])
            card_count = min(len(bullets), 4)
            if card_count > 0:
                card_height = (4.5 - (card_count - 1) * 0.25) / card_count
                for idx, bullet in enumerate(bullets[:card_count]):
                    top = 1.9 + idx * (card_height + 0.25)
                    c_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(top), Inches(11.733), Inches(card_height))
                    c_shape.fill.solid()
                    c_shape.fill.fore_color.rgb = card_color
                    c_shape.line.color.rgb = primary_color
                    
                    b_box = slide.shapes.add_textbox(Inches(1.1), Inches(top + 0.1), Inches(11.1), Inches(card_height - 0.2))
                    b_tf = b_box.text_frame
                    b_tf.word_wrap = True
                    p = b_tf.paragraphs[0]
                    p.text = f"• {bullet}"
                    p.font.size = Pt(12)
                    p.font.color.rgb = text_color

        # 7. Slide Footer
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.85), Inches(11.733), Inches(0.4))
        f_tf = footer_box.text_frame
        f_p = f_tf.paragraphs[0]
        f_p.text = f"VisualDeck AI  •  Slide {slide_info.get('slide_number', '')}"
        f_p.font.size = Pt(9)
        f_p.font.color.rgb = secondary_text

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream

# 4. Dialog for Clean Full-Image Inspection Popups
@st.dialog("🖼️ Source Image Preview", width="large")
def preview_modal(img_obj, filename):
    st.image(img_obj, caption=filename, use_container_width=True)

# 5. Left Sidebar Overview
with st.sidebar:
    st.markdown('<div class="badge-pill">PLATFORM CONTROL</div>', unsafe_allow_html=True)
    st.markdown("### ⚡ VisualDeck AI")
    st.caption("Autonomous visual-to-presentation synthesis engine.")
    st.markdown("---")
    
    st.markdown("""
    <div class="sidebar-kpi">
        <div class="sidebar-kpi-title">Core AI Architecture</div>
        <div class="sidebar-kpi-val" style="color:#00D2FF;">Gemini 3 Series</div>
    </div>
    <div class="sidebar-kpi">
        <div class="sidebar-kpi-title">Presentation Format</div>
        <div class="sidebar-kpi-val">16:9 Widescreen .PPTX</div>
    </div>
    <div class="sidebar-kpi">
        <div class="sidebar-kpi-title">Visual Extraction</div>
        <div class="sidebar-kpi-val" style="color:#10B981;">Shapes, Tables & Colors</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### 📋 How It Works")
    st.markdown("""
    <div style="font-size:0.82rem; color:#94A3B8; line-height:1.6;">
        <b style="color:#FFF;">1. Ingest:</b> Upload brochures, specs, diagrams, or PDFs.<br>
        <b style="color:#FFF;">2. Analyze:</b> AI extracts tables, key metrics, and brand palettes.<br>
        <b style="color:#FFF;">3. Edit & Download:</b> Review and tweak slides before downloading .PPTX.
    </div>
    """, unsafe_allow_html=True)

# 6. Main Hero Section
st.markdown('<div class="badge-pill">AI MULTIMODAL SYNTHESIZER</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-title">VisualDeck AI</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-subtitle">Transform unstructured visual documents, diagrams, and reports into structured, native PowerPoint presentations.</div>', unsafe_allow_html=True)

# 7. File Upload Section
uploaded_files = st.file_uploader(
    label="Upload Visual Documents or Data Sheets",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True
)

image_count = 0
pdf_page_count = 0
total_size_mb = 0.0
raw_image_parts = []
original_files = []

if uploaded_files:
    total_size_mb = sum([f.size for f in uploaded_files]) / (1024 * 1024)
    for f in uploaded_files:
        if f.type.startswith("image/"):
            image_count += 1
            img = Image.open(f)
            original_files.append({"type": "image", "name": f.name, "size": f.size, "obj": img})
            
            buf = io.BytesIO()
            img.convert('RGB').save(buf, format='JPEG', quality=85)
            part = types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg")
            raw_image_parts.append(part)

        elif f.type == "application/pdf":
            try:
                reader = PdfReader(io.BytesIO(f.getvalue()))
                pages = len(reader.pages)
            except Exception:
                pages = 1
            pdf_page_count += pages
            original_files.append({"type": "pdf", "name": f.name, "size": f.size, "pages": pages})

# 8. Top KPI Summary Cards
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown("""<div class="kpi-card"><div class="kpi-title">AI Engine</div><div class="kpi-value">Gemini 3 Series</div><div class="kpi-desc">Multimodal Vision Pipeline</div></div>""", unsafe_allow_html=True)
with col2:
    val_str = f"{len(uploaded_files)} Items" if uploaded_files else "0 Files"
    desc_str = f"{image_count} Img | {pdf_page_count} PDF Pages" if pdf_page_count > 0 else f"{image_count} Images Loaded"
    st.markdown(f"""<div class="kpi-card"><div class="kpi-title">Input Buffer</div><div class="kpi-value">{val_str}</div><div class="kpi-desc">{desc_str}</div></div>""", unsafe_allow_html=True)
with col3:
    st.markdown(f"""<div class="kpi-card"><div class="kpi-title">Buffer Size</div><div class="kpi-value">{total_size_mb:.2f} MB</div><div class="kpi-desc">Ready for Synthesis</div></div>""", unsafe_allow_html=True)
with col4:
    st.markdown("""<div class="kpi-card"><div class="kpi-title">Export Standard</div><div class="kpi-value">16:9 .PPTX</div><div class="kpi-desc">Editable Shapes & Data</div></div>""", unsafe_allow_html=True)

st.markdown("---")

# 9. Source Assets Compact List (Matching Your Reference Design)
if uploaded_files:
    st.markdown(f"### 📄 Uploaded Files ({len(uploaded_files)})")
    
    for idx, item in enumerate(original_files):
        with st.container(border=True):
            fcol1, fcol2 = st.columns([5, 1])
            with fcol1:
                size_kb = round(item["size"] / 1024, 1)
                icon = "🖼️" if item["type"] == "image" else "📄"
                st.markdown(f"**{icon} {item['name']}**")
                st.caption(f"{size_kb} KB" + (f" • {item['pages']} Pages" if item["type"] == "pdf" else ""))
            with fcol2:
                if item["type"] == "image":
                    if st.button("👁️ View", key=f"view_btn_{idx}", use_container_width=True):
                        preview_modal(item["obj"], item["name"])
                else:
                    st.button("📄 PDF", key=f"view_btn_{idx}", disabled=True, use_container_width=True)

    st.markdown("---")

    # 10. Unified Presentation Customization Console
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

    # 11. AI Generation Engine Execution
    generate_clicked = st.button("⚡ Generate AI Presentation Deck", type="primary")

    if generate_clicked:
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("⚠️ Gemini API Key not found in Streamlit Secrets. Please configure `GEMINI_API_KEY`.")
        elif not raw_image_parts:
            st.warning("Please upload at least one image to synthesize slides.")
        else:
            start_time = time.time()
            with st.status("🚀 Synthesizing Deck with Gemini 3 Engine...", expanded=True) as status:
                st.write("🔍 **Phase 1:** Reading visual layout, tables, certifications, and product details...")
                
                client = genai.Client(api_key=api_key)
                
                prompt = f"""
                You are an executive presentation designer. Analyze the uploaded source images carefully.
                Guidelines:
                1. DYNAMIC FLOW: Generate a clean {slide_count}-slide presentation outline matching this tone: '{presentation_tone}'.
                2. DATA EXTRACTION: Extract exact model names, capacities, PESO certifications, and pricing into native tables and KPI metrics.
                3. BRAND PALETTE: Extract the dominant brand colors from logos and graphics for the deck theme.
                
                Output ONLY a JSON object conforming strictly to this structure:
                {{
                  "theme": {{
                    "background_hex": "#0A0E17",
                    "primary_accent_hex": "#00D2FF",
                    "card_bg_hex": "#141C2E",
                    "text_color_hex": "#FFFFFF",
                    "secondary_text_hex": "#94A3B8"
                  }},
                  "deck_title": "Executive Presentation Title",
                  "slides": [
                    {{
                      "slide_number": 1,
                      "layout_type": "kpi_grid",
                      "title": "Slide Title",
                      "subtitle": "Subtitle or core insight",
                      "kpis": [
                        {{"label": "Metric Name", "value": "Extracted Value", "desc": "Context note"}}
                      ]
                    }},
                    {{
                      "slide_number": 2,
                      "layout_type": "table",
                      "title": "Model Comparison & Specifications",
                      "subtitle": "Detailed breakdown",
                      "table": {{
                        "headers": ["Model / Item", "Capacity / Spec", "Investment / Price"],
                        "rows": [
                          ["Model A", "2000L", "₹7,49,300/-"]
                        ]
                      }}
                    }},
                    {{
                      "slide_number": 3,
                      "layout_type": "cards",
                      "title": "Key Advantages & Features",
                      "subtitle": "Value summary",
                      "bullets": [
                        "Clear bullet point 1",
                        "Clear bullet point 2"
                      ]
                    }}
                  ]
                }}
                """

                gemini_3_models = [
                    'gemini-3.1-flash-lite',
                    'gemini-3.5-flash-lite',
                    'gemini-3.7-flash',
                    'gemini-3.5-flash'
                ]
                
                response_json = None
                last_error = None
                
                for model_name in gemini_3_models:
                    try:
                        content_payload = [prompt] + raw_image_parts[:5]
                        res = client.models.generate_content(
                            model=model_name,
                            contents=content_payload,
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json"
                            )
                        )
                        raw_text = res.text.strip()
                        response_json = json.loads(raw_text)
                        break
                    except Exception as e:
                        last_error = str(e)
                        continue

                if not response_json:
                    status.update(label="❌ Generation Encountered an Error", state="error")
                    st.error(f"API Error Log: {last_error}")
                else:
                    st.write("📊 **Phase 2:** Synthesizing 16:9 native PowerPoint vector shapes and slide tables...")
                    pptx_stream = build_powerpoint(response_json)
                    st.session_state["generated_pptx"] = pptx_stream
                    st.session_state["deck_data"] = response_json
                    status.update(label="✅ Presentation Deck Synthesized Successfully!", state="complete", expanded=False)

    # 12. Interactive Slide Preview & Live Deck Editor
    if "deck_data" in st.session_state and st.session_state["deck_data"]:
        deck = st.session_state["deck_data"]
        slides = deck.get("slides", [])
        
        st.markdown("---")
        st.markdown(f"### 📑 Synthesized Presentation: **{deck.get('deck_title', 'Synthesized Deck')}**")
        
        # Download Button
        st.download_button(
            label="📥 Download Native 16:9 .PPTX Presentation",
            data=st.session_state["generated_pptx"],
            file_name="VisualDeck_AI_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary"
        )
        
        st.markdown("#### ✏️ Interactive Slide Preview & Editor")
        st.caption("Review your generated slides below. You can edit any text, metrics, or tables before downloading.")
        
        for idx, s in enumerate(slides):
            with st.container(border=True):
                st.markdown(f"### 🖥️ Slide {s.get('slide_number', idx + 1)}: {s.get('title', '')}")
                
                # Visual Render Box
                with st.expander("👁️ View Rendered Slide Card", expanded=True):
                    st.markdown(f"**{s.get('title', '')}**")
                    if s.get("subtitle"):
                        st.caption(f"_{s.get('subtitle')}_")
                    
                    if s.get("layout_type") == "kpi_grid" and s.get("kpis"):
                        kpi_cols = st.columns(len(s["kpis"]))
                        for k_idx, k in enumerate(s["kpis"]):
                            with kpi_cols[k_idx]:
                                st.metric(label=k.get("label", "Metric"), value=k.get("value", "N/A"))
                    
                    elif s.get("layout_type") == "table" and s.get("table"):
                        headers = s["table"].get("headers", [])
                        rows = s["table"].get("rows", [])
                        if headers and rows:
                            df = pd.DataFrame(rows, columns=headers)
                            st.dataframe(df, use_container_width=True)
                    
                    elif s.get("bullets"):
                        for b in s["bullets"]:
                            st.markdown(f"- {b}")
                
                # Inline Editing Controls
                with st.expander(f"✏️ Edit Slide {idx + 1} Content"):
                    s["title"] = st.text_input("Slide Title:", value=s.get("title", ""), key=f"title_{idx}")
                    s["subtitle"] = st.text_input("Subtitle / Insight:", value=s.get("subtitle", ""), key=f"sub_{idx}")
                    
                    if s.get("layout_type") == "table" and s.get("table"):
                        st.markdown("**Edit Table Data:**")
                        headers = s["table"].get("headers", [])
                        rows = s["table"].get("rows", [])
                        if headers and rows:
                            df_edit = pd.DataFrame(rows, columns=headers)
                            edited_df = st.data_editor(df_edit, key=f"tbl_edit_{idx}")
                            s["table"]["headers"] = list(edited_df.columns)
                            s["table"]["rows"] = edited_df.values.tolist()
                    
                    elif s.get("layout_type") == "cards" and s.get("bullets"):
                        st.markdown("**Edit Feature Points (One per line):**")
                        bullet_text = "\n".join(s.get("bullets", []))
                        edited_bullets = st.text_area("Bullets:", value=bullet_text, key=f"bullets_{idx}", height=120)
                        s["bullets"] = [b.strip() for b in edited_bullets.split("\n") if b.strip()]

        # Rebuild PPTX Button
        if st.button("💾 Apply Edits & Rebuild Presentation", type="secondary"):
            st.session_state["generated_pptx"] = build_powerpoint(st.session_state["deck_data"])
            st.success("✅ Presentation updated with your changes! Click Download above to get your updated .pptx file.")
            st.rerun()

else:
    st.info("👆 Upload one or more product sheets, diagrams, or PDF documents above to begin.")
