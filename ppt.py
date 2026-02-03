"""
Brilliant PPT Generator - Enhanced Version with Proper AI Story Generation
"""

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import io
import json
import re

# OpenAI for AI content generation
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

st.set_page_config(
    page_title="Brilliant PPT Generator",
    page_icon="✨",
    layout="wide"
)

# Professional Light Themes
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
    """Use OpenAI to enhance content into professional slide structure with storytelling"""
    
    if not api_key:
        st.error("Please provide an OpenAI API key")
        return None
    
    try:
        client = OpenAI(api_key=api_key)
        
        system_prompt = """You are an expert presentation designer and storyteller who creates CEO-level, impressive PowerPoint presentations.

YOUR TASK:
Transform the user's raw notes/bullet points into a compelling, professional presentation with a clear narrative flow.

CRITICAL RULES:
1. CREATE A STORY: Don't just list facts. Build a narrative that flows logically from slide to slide.
2. ENHANCE THE CONTENT: Take simple bullet points and make them powerful, action-oriented statements.
3. AUTO-CORRECT: Fix any spelling, grammar, or unclear language.
4. ADD VALUE: If the content is vague or incomplete, intelligently expand it with relevant professional content.
5. STRUCTURE: Create 4-7 slides with clear progression (Introduction → Problem/Context → Solution/Main Points → Benefits/Results → Conclusion/Call to Action)
6. EACH SLIDE MUST HAVE:
   - A compelling, concise title (5-8 words max)
   - 3-5 bullet points that are clear and impactful
   - Each bullet should be 1-2 lines, professional language

SLIDE STRUCTURE TEMPLATE:
- Slide 1: Opening Hook / Executive Summary
- Slide 2-3: Context / Problem / Current Situation  
- Slide 3-5: Solution / Key Points / Main Content
- Slide 6: Benefits / Impact / Results
- Slide 7: Conclusion / Next Steps / Call to Action

OUTPUT FORMAT - Return ONLY valid JSON, no markdown, no explanation:
{
    "title": "Enhanced Presentation Title",
    "slides": [
        {
            "title": "Compelling Slide Title",
            "bullets": [
                "First impactful point with clear value",
                "Second point that builds on the narrative",
                "Third point that drives action"
            ]
        }
    ]
}"""

        user_prompt = f"""Create a CEO-impressive presentation from this content:

TITLE: {title}

RAW CONTENT/NOTES:
{content}

Transform this into a professional, story-driven presentation. Make it compelling and impressive. Fix any errors and enhance weak points. Return ONLY the JSON output."""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7,
            max_tokens=3000
        )
        
        response_text = response.choices[0].message.content.strip()
        
        # Clean any markdown formatting
        if '```json' in response_text:
            response_text = response_text.split('```json')[1]
        if '```' in response_text:
            response_text = response_text.split('```')[0]
        response_text = response_text.strip()
        
        # Parse JSON
        result = json.loads(response_text)
        
        # Validate structure
        if 'slides' not in result or not result['slides']:
            raise ValueError("No slides generated")
        
        for slide in result['slides']:
            if 'title' not in slide:
                slide['title'] = 'Key Points'
            if 'bullets' not in slide or not slide['bullets']:
                slide['bullets'] = ['Content to be added']
        
        return result
    
    except json.JSONDecodeError as e:
        st.error(f"Failed to parse AI response. Using fallback parser.")
        st.text(f"Raw response: {response_text[:500]}...")
        return None
    except Exception as e:
        st.error(f"AI Error: {str(e)}")
        return None


def generate_smart_content_without_api(title: str, content: str) -> dict:
    """Generate structured slides without API - intelligent parsing with enhancement"""
    
    lines = [l.strip() for l in content.split('\n') if l.strip()]
    
    if not lines:
        # Generate default professional structure
        return {
            "title": title,
            "slides": [
                {
                    "title": "Executive Overview",
                    "bullets": [
                        f"Introduction to {title}",
                        "Key objectives and goals",
                        "Strategic importance and value proposition"
                    ]
                },
                {
                    "title": "Key Highlights",
                    "bullets": [
                        "Primary focus areas identified",
                        "Critical success factors defined",
                        "Measurable outcomes expected"
                    ]
                },
                {
                    "title": "Next Steps & Action Items",
                    "bullets": [
                        "Immediate priorities to address",
                        "Timeline for implementation",
                        "Resources and support required"
                    ]
                }
            ]
        }
    
    slides = []
    current_slide = None
    
    for line in lines:
        # Check if it's a bullet point
        is_bullet = line.startswith(('-', '*', '•', '>', '→')) or (len(line) > 2 and line[0].isdigit() and line[1] in '.)')
        
        if is_bullet:
            # Clean the bullet
            bullet_text = re.sub(r'^[-*•>\→\d.)\s]+', '', line).strip()
            if bullet_text:
                if current_slide is None:
                    current_slide = {'title': 'Key Points', 'bullets': []}
                current_slide['bullets'].append(bullet_text)
        else:
            # It's a title/heading
            if current_slide and current_slide.get('bullets'):
                slides.append(current_slide)
            current_slide = {'title': line, 'bullets': []}
    
    # Add last slide
    if current_slide:
        if current_slide.get('bullets'):
            slides.append(current_slide)
        elif current_slide.get('title'):
            # Title without bullets - add some default content
            current_slide['bullets'] = [
                f"Details about {current_slide['title']}",
                "Key points to discuss",
                "Important considerations"
            ]
            slides.append(current_slide)
    
    # If still no slides, create from raw content
    if not slides:
        # Split content into chunks and create slides
        all_points = [l for l in lines if l]
        chunks = [all_points[i:i+4] for i in range(0, len(all_points), 4)]
        
        for idx, chunk in enumerate(chunks):
            slide_title = chunk[0] if not chunk[0].startswith(('-', '*', '•')) else f"Key Points {idx + 1}"
            bullets = chunk[1:] if not chunk[0].startswith(('-', '*', '•')) else chunk
            bullets = [re.sub(r'^[-*•>\→\d.)\s]+', '', b).strip() for b in bullets]
            bullets = [b for b in bullets if b]
            
            if not bullets:
                bullets = [slide_title]
                slide_title = f"Section {idx + 1}"
            
            slides.append({
                'title': slide_title,
                'bullets': bullets[:5]  # Max 5 bullets
            })
    
    # Ensure each slide has content
    for slide in slides:
        if not slide.get('bullets') or len(slide['bullets']) == 0:
            slide['bullets'] = [
                f"Overview of {slide.get('title', 'this topic')}",
                "Key details and information",
                "Important points to remember"
            ]
    
    return {
        "title": title,
        "slides": slides
    }


def create_presentation(title: str, slides_data: list, theme_name: str) -> io.BytesIO:
    """Create a professional PowerPoint presentation"""
    
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    theme = THEMES.get(theme_name, THEMES['Executive Blue'])

    # ========== TITLE SLIDE ==========
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # Background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*theme['header_bg'])
    
    # Main title
    title_box = slide.shapes.add_textbox(Inches(0.75), Inches(2.3), Inches(11.833), Inches(1.8))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(52)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.font.name = 'Calibri Light'
    
    # Subtitle
    sub_box = slide.shapes.add_textbox(Inches(0.75), Inches(4.3), Inches(11.833), Inches(0.8))
    tf = sub_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Professional Presentation"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(24)
    p.font.color.rgb = RGBColor(*theme['secondary'])
    p.font.name = 'Calibri'

    # ========== CONTENT SLIDES ==========
    for idx, slide_content in enumerate(slides_data):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # White background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        # Header bar
        header = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(1.3))
        header.fill.solid()
        header.fill.fore_color.rgb = RGBColor(*theme['header_bg'])
        header.line.fill.background()
        
        # Slide title
        title_box = slide.shapes.add_textbox(Inches(0.75), Inches(0.3), Inches(11.833), Inches(0.8))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = slide_content.get('title', f'Slide {idx + 1}')
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.name = 'Calibri Light'
        
        # Accent bar
        accent = slide.shapes.add_shape(1, Inches(0.75), Inches(1.6), Inches(0.15), Inches(4.8))
        accent.fill.solid()
        accent.fill.fore_color.rgb = RGBColor(*theme['accent'])
        accent.line.fill.background()
        
        # Bullet points
        bullets = slide_content.get('bullets', [])
        if bullets:
            content_box = slide.shapes.add_textbox(Inches(1.3), Inches(1.8), Inches(11.3), Inches(5.2))
            tf = content_box.text_frame
            tf.word_wrap = True
            
            for i, bullet in enumerate(bullets):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                
                p.text = f"•  {bullet}"
                p.font.size = Pt(22)
                p.font.color.rgb = RGBColor(*theme['body_text'])
                p.font.name = 'Calibri'
                p.space_before = Pt(6)
                p.space_after = Pt(14)
                p.level = 0
        
        # Page number
        num_box = slide.shapes.add_textbox(Inches(12.3), Inches(7.0), Inches(0.8), Inches(0.4))
        tf = num_box.text_frame
        p = tf.paragraphs[0]
        p.text = str(idx + 1)
        p.alignment = PP_ALIGN.RIGHT
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(150, 150, 150)

    # ========== THANK YOU SLIDE ==========
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*theme['header_bg'])
    
    # Thank you text
    thanks_box = slide.shapes.add_textbox(Inches(0.75), Inches(2.5), Inches(11.833), Inches(1.5))
    tf = thanks_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Thank You"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(60)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.font.name = 'Calibri Light'
    
    # Questions text
    q_box = slide.shapes.add_textbox(Inches(0.75), Inches(4.2), Inches(11.833), Inches(0.8))
    tf = q_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Questions & Discussion"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(28)
    p.font.color.rgb = RGBColor(*theme['secondary'])
    p.font.name = 'Calibri'
    
    # Save
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


# ========== CUSTOM STYLES ==========
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

.stApp {
    background: linear-gradient(180deg, #F8FAFC 0%, #EEF2FF 100%);
}

.main-title {
    text-align: center;
    padding: 1.5rem 0;
}

.main-title h1 {
    font-size: 2.8rem;
    font-weight: 700;
    color: #1E293B;
    margin: 0;
}

.main-title p {
    font-size: 1.1rem;
    color: #64748B;
    margin-top: 0.5rem;
}

.card {
    background: white;
    border-radius: 16px;
    padding: 1.5rem;
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
    margin-bottom: 1rem;
}

.card-header {
    font-size: 1rem;
    font-weight: 600;
    color: #1E293B;
    margin-bottom: 1rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

.slide-preview-container {
    background: white;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 8px 30px rgba(0,0,0,0.12);
    margin: 0.75rem 0;
}

.slide-preview-header {
    padding: 0.75rem 1rem;
    color: white;
    font-weight: 600;
}

.slide-preview-body {
    padding: 1rem;
    min-height: 120px;
    background: white;
}

.slide-preview-body p {
    margin: 0.4rem 0;
    color: #334155;
    font-size: 0.9rem;
}

.theme-pill {
    display: inline-block;
    padding: 0.25rem 0.75rem;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 500;
    margin: 0.25rem;
}

.success-box {
    background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%);
    border: 1px solid #6EE7B7;
    border-radius: 12px;
    padding: 1rem;
    margin: 1rem 0;
}

.info-list {
    background: #F8FAFC;
    border-radius: 8px;
    padding: 1rem;
    margin: 0.5rem 0;
}

.info-list p {
    margin: 0.3rem 0;
    font-size: 0.9rem;
    color: #475569;
}
</style>
""", unsafe_allow_html=True)


# ========== MAIN UI ==========
st.markdown("""
<div class="main-title">
    <h1>✨ Brilliant PPT Generator</h1>
    <p>Transform your ideas into CEO-impressive presentations with AI storytelling</p>
</div>
""", unsafe_allow_html=True)

# Two column layout
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    # Content Input Section
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-header">📝 Your Content</div>', unsafe_allow_html=True)
    
    pres_title = st.text_input(
        "Presentation Title *",
        placeholder="e.g., Q4 Business Strategy 2026",
        help="Enter a clear, professional title"
    )
    
    content = st.text_area(
        "Your Notes / Bullet Points *",
        height=280,
        placeholder="""Enter your raw content, notes, or bullet points...

The AI will transform this into a compelling story!

Example:
Market Analysis
- Market size $50 billion
- Growing 15% per year
- Main competitors: Company A, B, C

Our Solution
- AI-powered platform
- 50% cost reduction
- Easy to implement

Results Expected
- Revenue increase 30%
- Customer satisfaction up
- Market share growth""",
        help="Enter bullet points, notes, or any text. AI will enhance and structure it."
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    # AI Settings
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-header">🤖 AI Enhancement</div>', unsafe_allow_html=True)
    
    use_ai = st.toggle("Enable AI Story Generation", value=True, 
                       help="Uses GPT-4o-mini to create compelling narratives")
    
    api_key = ""
    if use_ai:
        api_key = st.text_input(
            "OpenAI API Key",
            type="password",
            placeholder="sk-...",
            help="Required for AI enhancement. Get key from platform.openai.com"
        )
        
        st.markdown("""
        <div class="info-list">
            <p>✅ Creates compelling story narrative</p>
            <p>✅ Auto-corrects spelling & grammar</p>
            <p>✅ Enhances weak bullet points</p>
            <p>✅ Structures content professionally</p>
            <p>✅ Generates 4-7 optimized slides</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("AI disabled. Will use smart parsing to structure your content.")
    
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    # Theme Selection
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-header">🎨 Choose Theme</div>', unsafe_allow_html=True)
    
    selected_theme = st.selectbox(
        "Color Theme",
        options=list(THEMES.keys()),
        index=0
    )
    
    theme = THEMES[selected_theme]
    r, g, b = theme['header_bg']
    
    # Theme preview
    st.markdown(f"""
        <div style="
            background: {theme['gradient']};
            height: 70px;
            border-radius: 12px;
            margin: 0.75rem 0;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: 600;
            font-size: 1.1rem;
        ">{selected_theme}</div>
    """, unsafe_allow_html=True)
    
    # All themes preview
    st.markdown("**All Themes:**")
    theme_cols = st.columns(4)
    for idx, (name, t) in enumerate(THEMES.items()):
        with theme_cols[idx % 4]:
            r2, g2, b2 = t['header_bg']
            selected_border = "3px solid #1E293B" if name == selected_theme else "none"
            st.markdown(f"""
                <div style="
                    background: rgb({r2},{g2},{b2});
                    height: 25px;
                    border-radius: 4px;
                    margin-bottom: 0.25rem;
                    border: {selected_border};
                "></div>
                <p style="font-size: 0.65rem; color: #64748B; margin: 0; text-align: center;">
                    {"✓ " if name == selected_theme else ""}{name}
                </p>
            """, unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Tips
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-header">💡 Tips for Best Results</div>', unsafe_allow_html=True)
    st.markdown("""
    **Content Tips:**
    - Include topic headings (without bullets)
    - Use `-`, `*`, or `•` for bullet points
    - Separate sections with blank lines
    - Include numbers and data when available
    
    **AI will:**
    - Fix any typos or grammar issues
    - Create a logical flow/story
    - Add professional language
    - Structure into 4-7 impactful slides
    """)
    st.markdown('</div>', unsafe_allow_html=True)

# Generate Button
st.markdown("---")
col_b1, col_b2, col_b3 = st.columns([1, 2, 1])
with col_b2:
    generate = st.button("🚀 Generate Presentation", type="primary", use_container_width=True)

# Store results in session state
if 'generated_slides' not in st.session_state:
    st.session_state.generated_slides = None
if 'generated_title' not in st.session_state:
    st.session_state.generated_title = None
if 'pptx_file' not in st.session_state:
    st.session_state.pptx_file = None

# Generation Logic
if generate:
    if not pres_title.strip():
        st.error("⚠️ Please enter a presentation title")
    elif not content.strip():
        st.error("⚠️ Please enter some content or bullet points")
    else:
        with st.spinner("✨ Creating your brilliant presentation..."):
            result = None
            
            # Try AI enhancement
            if use_ai and api_key:
                st.info("🤖 AI is analyzing your content and creating a compelling story...")
                result = enhance_content_with_ai(pres_title, content, api_key)
                
                if result:
                    st.success(f"✅ AI created {len(result['slides'])} professional slides!")
            
            # Fallback to smart parsing
            if not result:
                if use_ai and api_key:
                    st.warning("⚠️ AI enhancement failed. Using smart parsing instead.")
                else:
                    st.info("📝 Processing your content with smart parsing...")
                result = generate_smart_content_without_api(pres_title, content)
            
            if result and result.get('slides'):
                st.session_state.generated_slides = result['slides']
                st.session_state.generated_title = result.get('title', pres_title)
                
                # Generate PPTX
                st.session_state.pptx_file = create_presentation(
                    st.session_state.generated_title,
                    st.session_state.generated_slides,
                    selected_theme
                )
                
                st.balloons()
            else:
                st.error("❌ Failed to generate slides. Please try again.")

# Display Results
if st.session_state.generated_slides:
    st.markdown("---")
    st.markdown(f"### 📊 Preview: {st.session_state.generated_title}")
    st.markdown(f"**{len(st.session_state.generated_slides)} slides generated** | Theme: {selected_theme}")
    
    # Show slides preview
    for idx, slide in enumerate(st.session_state.generated_slides):
        r, g, b = theme['header_bg']
        with st.expander(f"📄 Slide {idx + 1}: {slide.get('title', 'Untitled')}", expanded=(idx < 2)):
            st.markdown(f"""
            <div class="slide-preview-container">
                <div class="slide-preview-header" style="background: rgb({r},{g},{b});">
                    {slide.get('title', 'Untitled')}
                </div>
                <div class="slide-preview-body">
                    {''.join([f'<p>• {bullet}</p>' for bullet in slide.get('bullets', [])])}
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    # Download section
    st.markdown("---")
    col_d1, col_d2, col_d3 = st.columns([1, 2, 1])
    with col_d2:
        if st.session_state.pptx_file:
            filename = re.sub(r'[^\w\s-]', '', st.session_state.generated_title.lower())
            filename = filename.replace(' ', '_')[:50] + '.pptx'
            
            st.download_button(
                label="📥 Download PowerPoint (PPTX)",
                data=st.session_state.pptx_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
            
            st.markdown("""
            <div class="success-box">
                <p style="margin: 0; text-align: center;">
                    ✅ <strong>Your presentation is ready!</strong><br>
                    Click above to download your CEO-impressive PowerPoint.
                </p>
            </div>
            """, unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("""
<p style="text-align: center; color: #64748B; font-size: 0.9rem;">
    Made with ❤️ for creating CEO-impressive presentations in seconds
</p>
""", unsafe_allow_html=True)
