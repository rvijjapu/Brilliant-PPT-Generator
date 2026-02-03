"""
Brilliant PPT Generator - Enhanced Streamlit App
Transform bullet points into stunning PowerPoint presentations with AI
"""

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import io
import json
import re

# For AI integration - install with: pip install openai
# Or use emergentintegrations if available
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

# Page configuration
st.set_page_config(
    page_title="Brilliant PPT Generator",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Professional Light Themes - No Dark Colors
THEMES = {
    'Executive Blue': {
        'primary': (37, 99, 235),
        'secondary': (239, 246, 255),
        'accent': (59, 130, 246),
        'title_text': (255, 255, 255),
        'body_text': (30, 41, 59),
        'slide_bg': (255, 255, 255),
        'header_bg': (37, 99, 235),
        'gradient': 'linear-gradient(135deg, #2563EB 0%, #3B82F6 100%)'
    },
    'Modern Teal': {
        'primary': (13, 148, 136),
        'secondary': (240, 253, 250),
        'accent': (20, 184, 166),
        'title_text': (255, 255, 255),
        'body_text': (30, 41, 59),
        'slide_bg': (255, 255, 255),
        'header_bg': (13, 148, 136),
        'gradient': 'linear-gradient(135deg, #0D9488 0%, #14B8A6 100%)'
    },
    'Vibrant Orange': {
        'primary': (234, 88, 12),
        'secondary': (255, 247, 237),
        'accent': (249, 115, 22),
        'title_text': (255, 255, 255),
        'body_text': (30, 41, 59),
        'slide_bg': (255, 255, 255),
        'header_bg': (234, 88, 12),
        'gradient': 'linear-gradient(135deg, #EA580C 0%, #F97316 100%)'
    },
    'Royal Purple': {
        'primary': (124, 58, 237),
        'secondary': (245, 243, 255),
        'accent': (139, 92, 246),
        'title_text': (255, 255, 255),
        'body_text': (30, 41, 59),
        'slide_bg': (255, 255, 255),
        'header_bg': (124, 58, 237),
        'gradient': 'linear-gradient(135deg, #7C3AED 0%, #8B5CF6 100%)'
    },
    'Fresh Green': {
        'primary': (22, 163, 74),
        'secondary': (240, 253, 244),
        'accent': (34, 197, 94),
        'title_text': (255, 255, 255),
        'body_text': (30, 41, 59),
        'slide_bg': (255, 255, 255),
        'header_bg': (22, 163, 74),
        'gradient': 'linear-gradient(135deg, #16A34A 0%, #22C55E 100%)'
    },
    'Coral Rose': {
        'primary': (225, 29, 72),
        'secondary': (255, 241, 242),
        'accent': (244, 63, 94),
        'title_text': (255, 255, 255),
        'body_text': (30, 41, 59),
        'slide_bg': (255, 255, 255),
        'header_bg': (225, 29, 72),
        'gradient': 'linear-gradient(135deg, #E11D48 0%, #F43F5E 100%)'
    },
    'Sky Breeze': {
        'primary': (2, 132, 199),
        'secondary': (240, 249, 255),
        'accent': (14, 165, 233),
        'title_text': (255, 255, 255),
        'body_text': (30, 41, 59),
        'slide_bg': (255, 255, 255),
        'header_bg': (2, 132, 199),
        'gradient': 'linear-gradient(135deg, #0284C7 0%, #0EA5E9 100%)'
    },
    'Amber Gold': {
        'primary': (217, 119, 6),
        'secondary': (255, 251, 235),
        'accent': (245, 158, 11),
        'title_text': (255, 255, 255),
        'body_text': (30, 41, 59),
        'slide_bg': (255, 255, 255),
        'header_bg': (217, 119, 6),
        'gradient': 'linear-gradient(135deg, #D97706 0%, #F59E0B 100%)'
    }
}


def enhance_content_with_ai(title: str, content: str, api_key: str) -> dict:
    """Use OpenAI to enhance content into professional slide structure"""
    if not OPENAI_AVAILABLE or not api_key:
        return None
    
    try:
        client = OpenAI(api_key=api_key)
        
        system_prompt = """You are a world-class presentation expert who creates CEO-impressive, professional PowerPoint content. 
Your job is to transform raw bullet points or ideas into a compelling, structured presentation narrative.

Rules:
1. Create clear, impactful slide titles that tell a story
2. Transform bullet points into professional, concise statements
3. Auto-correct any spelling/grammar errors
4. Add context and flow to connect ideas
5. Keep each bullet to 1-2 lines max for readability
6. Ensure 3-5 bullets per slide for visual balance
7. Create a logical narrative flow across slides
8. Use action-oriented language
9. Make every word count - be concise yet impactful

Respond ONLY with valid JSON in this exact format:
{
  "title": "Enhanced presentation title",
  "slides": [
    {
      "title": "Slide 1 Title",
      "bullets": ["Bullet 1", "Bullet 2", "Bullet 3"]
    }
  ]
}"""

        user_prompt = f"""Transform this into a professional presentation:

Presentation Title: {title}

Raw Content:
{content}

Create 3-7 well-structured slides that tell a compelling story. Make it CEO-impressive!"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7,
            max_tokens=2000
        )
        
        response_text = response.choices[0].message.content.strip()
        
        # Clean markdown code blocks if present
        if response_text.startswith('```json'):
            response_text = response_text[7:]
        if response_text.startswith('```'):
            response_text = response_text[3:]
        if response_text.endswith('```'):
            response_text = response_text[:-3]
        response_text = response_text.strip()
        
        return json.loads(response_text)
    
    except Exception as e:
        st.error(f"AI Enhancement Error: {str(e)}")
        return None


def parse_content_simple(text: str) -> list:
    """Parse text into slides with titles and bullet points (without AI)"""
    lines = text.split('\n')
    slides = []
    current_slide = None

    for line in lines:
        line = line.strip()
        
        if not line:
            if current_slide and (current_slide['title'] or current_slide['bullets']):
                slides.append(current_slide)
                current_slide = None
            continue

        if line.startswith('-') or line.startswith('*') or line.startswith('•'):
            if not current_slide:
                current_slide = {'title': 'Key Points', 'bullets': []}
            current_slide['bullets'].append(line[1:].strip())
        else:
            if current_slide and current_slide.get('bullets'):
                slides.append(current_slide)
            current_slide = {'title': line, 'bullets': []}

    if current_slide and (current_slide.get('title') or current_slide.get('bullets')):
        slides.append(current_slide)

    return slides if slides else [{'title': 'Introduction', 'bullets': ['Add your content here']}]


def create_presentation(title: str, slides_data: list, theme_name: str) -> io.BytesIO:
    """Create a professional PowerPoint presentation"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # 16:9 widescreen
    prs.slide_height = Inches(7.5)
    
    theme = THEMES[theme_name]

    # ========== TITLE SLIDE ==========
    title_slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(title_slide_layout)
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*theme['header_bg'])
    
    # Title text
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.333), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_frame.paragraphs[0].font.size = Pt(54)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['title_text'])
    title_frame.paragraphs[0].font.name = 'Calibri Light'
    
    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(12.333), Inches(0.6))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = 'Professional Presentation'
    subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    subtitle_frame.paragraphs[0].font.size = Pt(24)
    subtitle_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['secondary'])
    subtitle_frame.paragraphs[0].font.name = 'Calibri'

    # ========== CONTENT SLIDES ==========
    for idx, slide_content in enumerate(slides_data):
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        # White background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*theme['slide_bg'])
        
        # Header bar
        header = slide.shapes.add_shape(
            1, Inches(0), Inches(0), Inches(13.333), Inches(1.2)
        )
        header.fill.solid()
        header.fill.fore_color.rgb = RGBColor(*theme['header_bg'])
        header.line.fill.background()
        
        # Slide title
        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.25), Inches(12), Inches(0.7))
        title_frame = title_box.text_frame
        title_frame.text = slide_content.get('title', f'Slide {idx + 1}')
        title_frame.paragraphs[0].font.size = Pt(36)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['title_text'])
        title_frame.paragraphs[0].font.name = 'Calibri Light'
        
        # Accent line
        accent = slide.shapes.add_shape(
            1, Inches(0.7), Inches(1.5), Inches(0.12), Inches(4.5)
        )
        accent.fill.solid()
        accent.fill.fore_color.rgb = RGBColor(*theme['accent'])
        accent.line.fill.background()
        
        # Bullet points
        bullets = slide_content.get('bullets', [])
        if bullets:
            content_box = slide.shapes.add_textbox(Inches(1.3), Inches(1.7), Inches(11.5), Inches(5))
            text_frame = content_box.text_frame
            text_frame.word_wrap = True
            text_frame.vertical_anchor = MSO_ANCHOR.TOP
            
            for i, bullet in enumerate(bullets):
                if i > 0:
                    p = text_frame.add_paragraph()
                else:
                    p = text_frame.paragraphs[0]
                
                p.text = f"• {bullet}"
                p.level = 0
                p.font.size = Pt(22)
                p.font.color.rgb = RGBColor(*theme['body_text'])
                p.font.name = 'Calibri'
                p.space_after = Pt(18)
        
        # Page number
        footer_box = slide.shapes.add_textbox(Inches(12.3), Inches(7.0), Inches(0.8), Inches(0.3))
        footer_frame = footer_box.text_frame
        footer_frame.text = f'{idx + 1}'
        footer_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
        footer_frame.paragraphs[0].font.size = Pt(14)
        footer_frame.paragraphs[0].font.color.rgb = RGBColor(150, 150, 150)

    # ========== THANK YOU SLIDE ==========
    end_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(end_slide_layout)
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*theme['header_bg'])
    
    thanks_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.333), Inches(1.5))
    thanks_frame = thanks_box.text_frame
    thanks_frame.text = 'Thank You'
    thanks_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    thanks_frame.paragraphs[0].font.size = Pt(60)
    thanks_frame.paragraphs[0].font.bold = True
    thanks_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['title_text'])
    thanks_frame.paragraphs[0].font.name = 'Calibri Light'
    
    questions_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(12.333), Inches(0.6))
    questions_frame = questions_box.text_frame
    questions_frame.text = 'Questions & Discussion'
    questions_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    questions_frame.paragraphs[0].font.size = Pt(28)
    questions_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['secondary'])
    questions_frame.paragraphs[0].font.name = 'Calibri'
    
    # Save to BytesIO
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    
    return pptx_io


# ========== CUSTOM CSS ==========
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    .stApp {
        background: linear-gradient(180deg, #F8FAFC 0%, #F1F5F9 100%);
        font-family: 'Inter', sans-serif;
    }
    
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }
    
    /* Header */
    .main-header {
        text-align: center;
        padding: 2rem 0;
        margin-bottom: 2rem;
    }
    
    .main-header h1 {
        font-size: 3rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 0.5rem;
    }
    
    .main-header p {
        font-size: 1.2rem;
        color: #64748B;
    }
    
    /* Cards */
    .content-card {
        background: white;
        padding: 1.5rem;
        border-radius: 16px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06);
        border: 1px solid #E2E8F0;
        margin-bottom: 1.5rem;
    }
    
    .card-title {
        font-size: 1.1rem;
        font-weight: 600;
        color: #0F172A;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Theme grid */
    .theme-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 0.75rem;
    }
    
    .theme-item {
        text-align: center;
        padding: 0.5rem;
        border-radius: 8px;
        cursor: pointer;
        transition: transform 0.2s;
    }
    
    .theme-item:hover {
        transform: translateY(-2px);
    }
    
    .theme-preview {
        width: 100%;
        height: 40px;
        border-radius: 6px;
        margin-bottom: 0.25rem;
    }
    
    .theme-name {
        font-size: 0.75rem;
        color: #64748B;
    }
    
    /* Buttons */
    .stButton > button {
        width: 100%;
        padding: 0.875rem 1.5rem;
        font-size: 1rem;
        font-weight: 600;
        border-radius: 12px;
        border: none;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
    }
    
    .stDownloadButton > button {
        width: 100%;
        padding: 0.875rem 1.5rem;
        font-size: 1rem;
        font-weight: 600;
        border-radius: 12px;
        background: linear-gradient(135deg, #10B981 0%, #059669 100%);
        color: white;
        border: none;
    }
    
    /* Preview */
    .slide-preview {
        background: white;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.12);
    }
    
    .slide-header {
        padding: 1rem 1.5rem;
        color: white;
    }
    
    .slide-body {
        padding: 1.5rem;
        min-height: 200px;
    }
    
    /* Info box */
    .info-box {
        background: linear-gradient(135deg, #EFF6FF 0%, #F5F3FF 100%);
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    
    .info-box p {
        margin: 0.25rem 0;
        font-size: 0.9rem;
        color: #475569;
    }
    
    /* Hide Streamlit elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)


# ========== MAIN APP ==========
st.markdown("""
    <div class="main-header">
        <h1>✨ Brilliant PPT Generator</h1>
        <p>Transform your bullet points into CEO-impressive presentations with AI</p>
    </div>
""", unsafe_allow_html=True)

# Layout
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">📝 Your Content</div>', unsafe_allow_html=True)
    
    pres_title = st.text_input(
        "Presentation Title",
        placeholder="e.g., Q4 Business Strategy 2026",
        help="Give your presentation a catchy title"
    )
    
    content = st.text_area(
        "Bullet Points & Ideas",
        height=250,
        placeholder="""Enter your content here...

Example:

Market Overview
- Global market size: $50B
- Growth rate: 15% annually
- Key competitors

Our Strategy
- Innovative approach
- Cost-effective solutions
- Scalable architecture

Revenue Projections
- Q4 target: $2M
- Year-end forecast: $8M""",
        help="Lines starting with -, *, or • become bullet points. Empty lines create new slides."
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    # AI Enhancement
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">🤖 AI Enhancement</div>', unsafe_allow_html=True)
    
    enhance_with_ai = st.toggle("Enhance with AI", value=True, help="Use GPT-4o-mini to create compelling narratives")
    
    if enhance_with_ai:
        api_key = st.text_input(
            "OpenAI API Key",
            type="password",
            placeholder="sk-...",
            help="Enter your OpenAI API key for AI enhancement"
        )
        
        st.markdown("""
            <div class="info-box">
                <p>✅ Auto-correct spelling & grammar</p>
                <p>✅ Create compelling narratives</p>
                <p>✅ Structure content professionally</p>
                <p>✅ Generate 3-7 optimized slides</p>
            </div>
        """, unsafe_allow_html=True)
    else:
        api_key = None
    
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    # Theme Selection
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">🎨 Choose Theme</div>', unsafe_allow_html=True)
    
    selected_theme = st.selectbox(
        "Select Color Theme",
        options=list(THEMES.keys()),
        help="Choose a professional theme for your presentation"
    )
    
    # Theme preview
    theme = THEMES[selected_theme]
    st.markdown(f"""
        <div style="
            background: {theme['gradient']};
            height: 80px;
            border-radius: 12px;
            margin: 1rem 0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        "></div>
    """, unsafe_allow_html=True)
    
    # Theme grid preview
    cols = st.columns(4)
    for idx, (name, t) in enumerate(THEMES.items()):
        with cols[idx % 4]:
            is_selected = "✓ " if name == selected_theme else ""
            r, g, b = t['primary']
            st.markdown(f"""
                <div style="text-align: center; margin-bottom: 0.5rem;">
                    <div style="
                        background: rgb({r},{g},{b});
                        height: 30px;
                        border-radius: 6px;
                        margin-bottom: 0.25rem;
                        border: {3 if name == selected_theme else 0}px solid #0F172A;
                    "></div>
                    <span style="font-size: 0.7rem; color: #64748B;">{is_selected}{name}</span>
                </div>
            """, unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Tips
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">💡 Pro Tips</div>', unsafe_allow_html=True)
    st.info("""
    **Formatting Guide:**
    - Use `-`, `*`, or `•` for bullet points
    - Empty lines create new slides
    - Keep bullets concise (1-2 lines)
    - Use 3-5 bullets per slide
    
    **AI Enhancement:**
    - Fixes spelling & grammar automatically
    - Creates professional narratives
    - Structures content for impact
    """)
    st.markdown('</div>', unsafe_allow_html=True)

# Generate Button
st.markdown("---")

col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    generate_clicked = st.button(
        "🚀 Generate Presentation",
        type="primary",
        use_container_width=True
    )

# Generation Logic
if generate_clicked:
    if not pres_title:
        st.error("⚠️ Please enter a presentation title")
    elif not content:
        st.error("⚠️ Please enter some content")
    else:
        try:
            with st.spinner("✨ Creating your brilliant presentation..."):
                # Get slides data
                if enhance_with_ai and api_key:
                    ai_result = enhance_content_with_ai(pres_title, content, api_key)
                    if ai_result:
                        slides_data = ai_result.get('slides', [])
                        pres_title = ai_result.get('title', pres_title)
                        st.success("🤖 AI enhanced your content!")
                    else:
                        slides_data = parse_content_simple(content)
                        st.warning("⚠️ AI enhancement failed. Using basic parsing.")
                else:
                    slides_data = parse_content_simple(content)
                
                if not slides_data:
                    st.error("❌ No valid content found. Please check your formatting.")
                else:
                    # Generate PPTX
                    pptx_file = create_presentation(pres_title, slides_data, selected_theme)
                    filename = re.sub(r'[^\w\s-]', '', pres_title.lower()).replace(' ', '_')[:50] + '.pptx'
                    
                    st.success(f"✅ Generated {len(slides_data)} slides successfully!")
                    
                    # Preview
                    st.markdown("### 📊 Slide Preview")
                    for idx, slide in enumerate(slides_data):
                        r, g, b = theme['header_bg']
                        with st.expander(f"Slide {idx + 1}: {slide.get('title', 'Untitled')}", expanded=(idx == 0)):
                            st.markdown(f"""
                                <div class="slide-preview">
                                    <div class="slide-header" style="background: rgb({r},{g},{b});">
                                        <strong>{slide.get('title', 'Untitled')}</strong>
                                    </div>
                                    <div class="slide-body">
                                        {''.join([f'<p>• {b}</p>' for b in slide.get('bullets', [])])}
                                    </div>
                                </div>
                            """, unsafe_allow_html=True)
                    
                    # Download button
                    st.markdown("---")
                    col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
                    with col_dl2:
                        st.download_button(
                            label="📥 Download PPTX",
                            data=pptx_file,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                    
                    st.balloons()
                    
        except Exception as e:
            st.error(f"❌ Error generating presentation: {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: #64748B; padding: 1rem;">
        Made with ❤️ for creating CEO-impressive presentations
    </div>
""", unsafe_allow_html=True)
