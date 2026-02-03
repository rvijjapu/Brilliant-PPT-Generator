"""
Brilliant PPT Generator - Streamlit App
Transform bullet points into stunning PowerPoint presentations
"""

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import io

# Page configuration
st.set_page_config(
    page_title="Brilliant PPT Generator",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Theme definitions
THEMES = {
    'Midnight Executive': {
        'primary': (30, 39, 97),
        'secondary': (202, 220, 252),
        'accent': (255, 255, 255),
        'text': (255, 255, 255),
        'bodyText': (30, 39, 97)
    },
    'Forest & Moss': {
        'primary': (44, 95, 45),
        'secondary': (151, 188, 98),
        'accent': (245, 245, 245),
        'text': (255, 255, 255),
        'bodyText': (44, 95, 45)
    },
    'Coral Energy': {
        'primary': (249, 97, 103),
        'secondary': (249, 231, 149),
        'accent': (47, 60, 126),
        'text': (255, 255, 255),
        'bodyText': (47, 60, 126)
    },
    'Warm Terracotta': {
        'primary': (184, 80, 66),
        'secondary': (231, 232, 209),
        'accent': (167, 190, 174),
        'text': (255, 255, 255),
        'bodyText': (61, 61, 61)
    },
    'Ocean Gradient': {
        'primary': (6, 90, 130),
        'secondary': (28, 114, 147),
        'accent': (33, 41, 92),
        'text': (255, 255, 255),
        'bodyText': (6, 90, 130)
    },
    'Teal Trust': {
        'primary': (2, 128, 144),
        'secondary': (2, 195, 154),
        'accent': (0, 168, 150),
        'text': (255, 255, 255),
        'bodyText': (2, 128, 144)
    }
}

def parse_content(text):
    """Parse text into slides with titles and bullet points"""
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
                current_slide = {'title': '', 'bullets': []}
            current_slide['bullets'].append(line[1:].strip())
        else:
            if current_slide and current_slide['bullets']:
                slides.append(current_slide)
            current_slide = {'title': line, 'bullets': []}

    if current_slide and (current_slide['title'] or current_slide['bullets']):
        slides.append(current_slide)

    return slides

def create_presentation(title, content, theme_name):
    """Create a professional PowerPoint presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    
    theme = THEMES[theme_name]
    slides_data = parse_content(content)
    
    if not slides_data:
        raise ValueError("No valid content found. Please check your formatting.")

    # Title Slide
    title_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(title_slide_layout)
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*theme['primary'])
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.0), Inches(9), Inches(1.2))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_frame.paragraphs[0].font.size = Pt(48)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['text'])
    title_frame.paragraphs[0].font.name = 'Arial Black'
    
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(9), Inches(0.5))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = 'Generated with ✨'
    subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    subtitle_frame.paragraphs[0].font.size = Pt(20)
    subtitle_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['secondary'])

    # Content Slides
    for idx, slide_data in enumerate(slides_data):
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)
        
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        header = slide.shapes.add_shape(
            1,
            Inches(0), Inches(0), Inches(10), Inches(0.8)
        )
        header.fill.solid()
        header.fill.fore_color.rgb = RGBColor(*theme['primary'])
        header.line.fill.background()
        
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(9), Inches(0.5))
        title_frame = title_box.text_frame
        title_frame.text = slide_data['title'] or f'Slide {idx + 1}'
        title_frame.paragraphs[0].font.size = Pt(32)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['text'])
        title_frame.paragraphs[0].font.name = 'Arial Black'
        title_frame.margin_top = Inches(0)
        title_frame.margin_bottom = Inches(0)
        
        accent = slide.shapes.add_shape(
            1,
            Inches(0.5), Inches(1.2), Inches(0.08), Inches(3.5)
        )
        accent.fill.solid()
        accent.fill.fore_color.rgb = RGBColor(*theme['accent'])
        accent.line.fill.background()
        
        if slide_data['bullets']:
            content_box = slide.shapes.add_textbox(Inches(1.0), Inches(1.4), Inches(8.5), Inches(3.2))
            text_frame = content_box.text_frame
            text_frame.word_wrap = True
            text_frame.vertical_anchor = MSO_ANCHOR.TOP
            
            for i, bullet in enumerate(slide_data['bullets']):
                if i > 0:
                    p = text_frame.add_paragraph()
                else:
                    p = text_frame.paragraphs[0]
                
                p.text = bullet
                p.level = 0
                p.font.size = Pt(18)
                p.font.color.rgb = RGBColor(*theme['bodyText'])
                p.font.name = 'Calibri'
        
        footer_box = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.8), Inches(0.3))
        footer_frame = footer_box.text_frame
        footer_frame.text = f'{idx + 1} / {len(slides_data)}'
        footer_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
        footer_frame.paragraphs[0].font.size = Pt(12)
        footer_frame.paragraphs[0].font.color.rgb = RGBColor(153, 153, 153)

    # Thank You Slide
    end_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(end_slide_layout)
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*theme['primary'])
    
    thanks_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.0), Inches(9), Inches(1.2))
    thanks_frame = thanks_box.text_frame
    thanks_frame.text = 'Thank You!'
    thanks_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    thanks_frame.paragraphs[0].font.size = Pt(54)
    thanks_frame.paragraphs[0].font.bold = True
    thanks_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['text'])
    thanks_frame.paragraphs[0].font.name = 'Arial Black'
    
    questions_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(9), Inches(0.5))
    questions_frame = questions_box.text_frame
    questions_frame.text = 'Questions?'
    questions_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    questions_frame.paragraphs[0].font.size = Pt(24)
    questions_frame.paragraphs[0].font.color.rgb = RGBColor(*theme['secondary'])
    
    # Save to BytesIO
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    
    return pptx_io

# Custom CSS
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    div[data-testid="stMarkdownContainer"] > h1 {
        color: white;
        text-align: center;
        font-size: 3rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    .subtitle {
        text-align: center;
        color: #f0f0f0;
        font-size: 1.3rem;
        margin-bottom: 2rem;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
    }
    div[data-testid="stHorizontalBlock"] {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.3);
    }
    .stButton>button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-size: 1.2rem;
        font-weight: 600;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 2rem;
        width: 100%;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        box-shadow: 0 8px 20px rgba(102, 126, 234, 0.5);
        transform: translateY(-2px);
    }
    .stDownloadButton>button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        font-size: 1.1rem;
        font-weight: 600;
        border-radius: 10px;
        padding: 0.75rem 2rem;
        width: 100%;
    }
    .stDownloadButton>button:hover {
        box-shadow: 0 8px 20px rgba(16, 185, 129, 0.5);
        transform: translateY(-2px);
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.title("✨ Brilliant PPT Generator")
st.markdown('<p class="subtitle">Transform your bullet points into stunning presentations</p>', unsafe_allow_html=True)

# Main content
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("### 📝 Your Content")
    
    pres_title = st.text_input(
        "Presentation Title",
        placeholder="Enter your presentation title...",
        help="Give your presentation a catchy title"
    )
    
    content = st.text_area(
        "Content",
        height=400,
        placeholder="""Enter your content here...

Example format:

Introduction
- Welcome to our presentation
- Today's agenda
- Key objectives

Market Analysis
- Current market size: $50B
- Growth rate: 15% annually
- Key competitors

Our Solution
- Innovative approach
- Cost-effective
- Scalable architecture""",
        help="Lines starting with -, *, or • become bullet points. Empty lines create new slides."
    )

with col2:
    st.markdown("### 🎨 Choose Theme")
    
    theme_colors = {
        'Midnight Executive': '135deg, #1E2761 0%, #CADCFC 100%',
        'Forest & Moss': '135deg, #2C5F2D 0%, #97BC62 100%',
        'Coral Energy': '135deg, #F96167 0%, #F9E795 100%',
        'Warm Terracotta': '135deg, #B85042 0%, #E7E8D1 100%',
        'Ocean Gradient': '135deg, #065A82 0%, #1C7293 100%',
        'Teal Trust': '135deg, #028090 0%, #02C39A 100%'
    }
    
    selected_theme = st.selectbox(
        "Select a color palette",
        options=list(THEMES.keys()),
        help="Choose a theme that matches your content"
    )
    
    # Show theme preview
    st.markdown(f"""
        <div style="
            background: linear-gradient({theme_colors[selected_theme]});
            height: 100px;
            border-radius: 10px;
            margin: 1rem 0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        "></div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 💡 Pro Tips")
    st.info("""
    ✅ Use `-`, `*`, or `•` for bullet points
    
    ✅ Empty lines create new slides
    
    ✅ Keep bullets concise (1-2 lines)
    
    ✅ Use 3-5 bullets per slide for best readability
    """)

# Generate button
st.markdown("---")
if st.button("🚀 Generate Beautiful Presentation", use_container_width=True):
    if not pres_title:
        st.error("⚠️ Please enter a presentation title")
    elif not content:
        st.error("⚠️ Please enter some content")
    else:
        try:
            with st.spinner("✨ Generating your beautiful presentation..."):
                pptx_file = create_presentation(pres_title, content, selected_theme)
                
                filename = pres_title.lower().replace(' ', '_').replace('/', '_')[:50] + '.pptx'
                
                st.success("✅ Presentation generated successfully!")
                
                st.download_button(
                    label="📥 Download Presentation",
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
    <div style="text-align: center; color: white; padding: 1rem;">
        Made with ❤️ for creating beautiful presentations quickly!
    </div>
""", unsafe_allow_html=True)
