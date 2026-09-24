import streamlit as st
import requests
import json
import io
import urllib.parse

# کتێبخانەیێن وۆرد و پاوەرپۆینت
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

st.set_page_config(page_title="سیستەمێ ئەکادیمی یێ ڕاپۆرت و سمیناران", page_icon="🎓", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    p, h1, h2, h3, label, .stMarkdown { text-align: right !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.title("🎓 سیستەمێ ئەکادیمی یێ دروستکرنا ڕاپۆرت (تێر و تەسەل) و سمیناران")
st.write("ڕاپۆرتێن قووڵ و درێژ (هەتا پتر ژ ١٠ لاپەڕان) دگەل پاوەرپۆینتەکا مۆدێرن کو وێنەیێن گونجای بۆ هەر سلایدەکی لەخۆدگریت.")

# کۆنترۆڵکرنا کلیلێ
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    api_key = st.text_input("کلیلا Gemini API لێرە بنڤیسە:", type="password")

# فۆڕما پێزانینان
with st.container():
    col1, col2 = st.columns(2)
    with col1:
        student_name = st.text_input("👤 ناڤێ قوتابی:")
        department = st.text_input("🏛️ پشک یان کۆلێژ:")
        language = st.selectbox("🌐 زمانێ نڤیسینێ:", ["کوردی (بادینی)", "کوردی (سۆرانی)", "العربية", "English"])
    with col2:
        teacher_name = st.text_input("👨‍🏫 ناڤێ مامۆستایێ بابەتی:")
        topic = st.text_input("📝 بابەتێ سەرەکی یێ ڕاپۆرتێ:")
        pages_count = st.slider("📄 ژمارا لاپەڕێن پێدڤی بۆ ڕاپۆرتێ:", min_value=3, max_value=20, value=12)

# گۆڕینا ژمارەیان بۆ شێوازێ عەرەبی
def convert_numbers(text, is_rtl):
    if not is_rtl or not text:
        return text
    western_to_eastern = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")
    return str(text).translate(western_to_eastern)

# ڕێکخستنا ئاراستەیا دەقی د وۆرد دا بۆ RTL
def set_docx_rtl(paragraph, is_rtl=True):
    if not is_rtl:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

# ڕێکخستنا فۆنت و ستایلێ وۆرد
def format_run(run, font_name="Calibri", size_pt=12, bold=False, color_rgb=(0, 0, 0), is_rtl=True):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)
    if is_rtl:
        rPr = run._r.get_or_add_rPr()
        rtl = OxmlElement('w:rtl')
        rtl.set(qn('w:val'), '1')
        rPr.append(rtl)

# ئینانا وێنەیێ گونجای بۆ پاوەرپۆینتێ
def fetch_slide_image(keyword):
    clean_keyword = urllib.parse.quote(str(keyword).strip() or "education")
    url = f"https://images.unsplash.com/photo-1497633762265-9d179a990aa6?w=600&auto=format&fit=crop&q=80"
    try:
        source_url = f"https://source.unsplash.com/featured/600x450/?{clean_keyword}"
        res = requests.get(source_url, timeout=5)
        if res.status_code == 200 and len(res.content) > 1000:
            return io.BytesIO(res.content)
    except Exception:
        pass
    
    # وێنەیەکێ یەدەگ ئەگەر پەیوەندی لاواز بوو
    try:
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            return io.BytesIO(res.content)
    except Exception:
        return None
    return None

def generate_academic_content(topic, lang, pages, student, dept, teacher, key):
    candidate_models = ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-2.5-flash"]
    headers = {"Content-Type": "application/json"}
    
    prompt = f"""
    You are an esteemed university professor and academic thesis supervisor.
    Write an exhaustive, university-level, in-depth academic research paper on the topic: "{topic}".
    Language to write in: {lang}.
    Requested Length: At least {pages} full academic pages. DO NOT summarize or write short paragraphs. Every single section must contain deep academic arguments, historical context, technical explanations, real-world case studies, and critical evaluations.
    
    Student: "{student}", Department: "{dept}", Supervisor: "{teacher}".

    For the presentation, provide 6 to 9 detailed slides. For each slide, also provide a single relevant English keyword (e.g. "artificial intelligence", "laboratory", "business finance") that matches the slide content so that an image can be fetched.

    Return ONLY a single valid JSON object (no markdown fences, no formatting backticks) with this structure:
    {{
        "title": "Comprehensive Academic Title",
        "abstract": "Deep, comprehensive abstract (250-350 words) detailing background, methodology, discussion, and findings.",
        "sections": [
            {{
                "heading": "Section 1: Theoretical Framework & Literature Review",
                "content": "Deep, multi-paragraph, extensive analysis..."
            }},
            {{
                "heading": "Section 2: Detailed Methodology & Analytical Core",
                "content": "Comprehensive, exhaustive academic discussion..."
            }},
            {{
                "heading": "Section 3: Practical Applications, Data & Case Studies",
                "content": "Deep technical analysis and practical implications..."
            }},
            {{
                "heading": "Section 4: Critical Challenges, Limitations & Ethical Dimensions",
                "content": "Extensive academic breakdown..."
            }},
            {{
                "heading": "Section 5: Future Horizons & Strategic Recommendations",
                "content": "Future analysis and structured insights..."
            }}
        ],
        "conclusion": "A comprehensive, formal academic conclusion (multiple paragraphs) tying together all findings and proposing future research directions.",
        "references": [
            "Official academic reference 1 in full APA format",
            "Official academic reference 2 in full APA format",
            "Official academic reference 3 in full APA format",
            "Official academic reference 4 in full APA format",
            "Official academic reference 5 in full APA format",
            "Official academic reference 6 in full APA format"
        ],
        "slides": [
            {{
                "slide_title": "Slide Title",
                "image_keyword": "technology",
                "bullet_points": [
                    "Direct, comprehensive point 1",
                    "Direct, comprehensive point 2",
                    "Direct, comprehensive point 3"
                ]
            }}
        ]
    }}
    """
    
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"response_mime_type": "application/json"}
    }
    
    last_error = ""
    for model_name in candidate_models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={key}"
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=120)
            if response.status_code == 200:
                result = response.json()
                text_content = result["candidates"][0]["content"]["parts"][0]["text"]
                return json.loads(text_content)
            else:
                last_error = response.text
                continue
        except Exception as e:
            last_error = str(e)
            continue
            
    raise Exception(f"API Error: {last_error}")

def build_docx(data, student, dept, teacher, is_rtl):
    doc = Document()
    
    # ڕێکخستنا حاشیەیا لاپەڕەیان (Standard 1-inch Margins)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # 1. لاپەڕا بەرگی فەرمی (Cover Page)
    p_uni = doc.add_paragraph()
    set_docx_rtl(p_uni, is_rtl)
    run_dept = p_uni.add_run(convert_numbers(dept or "زانکۆ / پەیمانگەهـ", is_rtl))
    format_run(run_dept, size_pt=14, bold=True, color_rgb=(70, 80, 95), is_rtl=is_rtl)
    
    p_title = doc.add_paragraph()
    set_docx_rtl(p_title, is_rtl)
    p_title.paragraph_format.space_before = Pt(48)
    p_title.paragraph_format.space_after = Pt(28)
    run_title = p_title.add_run(convert_numbers(data.get("title", "ڕاپۆرتا زانستی"), is_rtl))
    format_run(run_title, size_pt=24, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
    
    p_meta = doc.add_paragraph()
    set_docx_rtl(p_meta, is_rtl)
    p_meta.paragraph_format.space_before = Pt(36)
    lbl_s = "ئامادەکرن ژ لایێ: " if is_rtl else "Prepared by: "
    lbl_t = "سەرپەرشتیا: " if is_rtl else "Supervised by: "
    
    r_s = p_meta.add_run(f"{lbl_s}{student or '-'}\n{lbl_t}{teacher or '-'}\n\n")
    format_run(r_s, size_pt=13, bold=False, color_rgb=(80, 90, 100), is_rtl=is_rtl)
    
    doc.add_page_break()
    
    # 2. پوختە (Abstract)
    p_abs_h = doc.add_paragraph()
    set_docx_rtl(p_abs_h, is_rtl)
    r_abs_h = p_abs_h.add_run("پوختە (Abstract)" if is_rtl else "Abstract")
    format_run(r_abs_h, size_pt=16, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
    
    p_abs = doc.add_paragraph()
    set_docx_rtl(p_abs, is_rtl)
    p_abs.paragraph_format.line_spacing = 1.3
    r_abs = p_abs.add_run(convert_numbers(data.get("abstract", ""), is_rtl))
    format_run(r_abs, size_pt=11.5, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
    
    doc.add_page_break()
    
    # 3. پشکێن بەرفرەهـ و تێر و تەسەل (Body Sections)
    for idx, sec in enumerate(data.get("sections", [])):
        p_sec_h = doc.add_paragraph()
        set_docx_rtl(p_sec_h, is_rtl)
        p_sec_h.paragraph_format.space_before = Pt(20)
        p_sec_h.paragraph_format.space_after = Pt(8)
        r_sec_h = p_sec_h.add_run(convert_numbers(f"{idx+1}. {sec.get('heading', '')}", is_rtl))
        format_run(r_sec_h, size_pt=15, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
        
        # پاراگرافێن ناڤەرۆکێ
        raw_content = sec.get('content', '')
        paragraphs = raw_content.split("\n\n") if "\n\n" in raw_content else [raw_content]
        for para_text in paragraphs:
            if not para_text.strip():
                continue
            p_sec = doc.add_paragraph()
            set_docx_rtl(p_sec, is_rtl)
            p_sec.paragraph_format.line_spacing = 1.35
            p_sec.paragraph_format.space_after = Pt(10)
            r_sec = p_sec.add_run(convert_numbers(para_text.strip(), is_rtl))
            format_run(r_sec, size_pt=12, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
            
    # 4. دەرئەنجام (Conclusion)
    p_con_h = doc.add_paragraph()
    set_docx_rtl(p_con_h, is_rtl)
    p_con_h.paragraph_format.space_before = Pt(24)
    p_con_h.paragraph_format.space_after = Pt(8)
    r_con_h = p_con_h.add_run("دەرئەنجام و ڕاسپاردە (Conclusion & Recommendations)" if is_rtl else "Conclusion & Recommendations")
    format_run(r_con_h, size_pt=15, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
    
    p_con = doc.add_paragraph()
    set_docx_rtl(p_con, is_rtl)
    p_con.paragraph_format.line_spacing = 1.35
    r_con = p_con.add_run(convert_numbers(data.get("conclusion", ""), is_rtl))
    format_run(r_con, size_pt=12, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
    
    # 5. ژێدەر و سەرچاوەکان (References - APA)
    p_ref_h = doc.add_paragraph()
    set_docx_rtl(p_ref_h, is_rtl)
    p_ref_h.paragraph_format.space_before = Pt(26)
    p_ref_h.paragraph_format.space_after = Pt(10)
    r_ref_h = p_ref_h.add_run("سەرچاوەکان ب شێوازێ ئەکادیمی (References - APA)" if is_rtl else "References (APA)")
    format_run(r_ref_h, size_pt=15, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
    
    for ref in data.get("references", []):
        p_ref = doc.add_paragraph()
        set_docx_rtl(p_ref, is_rtl)
        p_ref.paragraph_format.space_after = Pt(6)
        r_ref = p_ref.add_run(f"• {convert_numbers(ref, is_rtl)}")
        format_run(r_ref, size_pt=11, bold=False, color_rgb=(70, 80, 95), is_rtl=is_rtl)
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def build_pptx(data, student, dept, teacher, is_rtl):
    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    DARK_BG = PptxRGBColor(15, 23, 42)
    ACCENT_GOLD = PptxRGBColor(245, 158, 11)
    WHITE = PptxRGBColor(255, 255, 255)
    LIGHT_GRAY = PptxRGBColor(203, 213, 225)
    CARD_BG = PptxRGBColor(30, 41, 59)
    
    # سلایدا دەستپێکێ (Title Slide)
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = DARK_BG
    bg1.line.fill.background()
    
    box1 = s1.shapes.add_textbox(PptxInches(1), PptxInches(1.8), PptxInches(11.333), PptxInches(4.5))
    tf1 = box1.text_frame
    tf1.word_wrap = True
    
    p_t = tf1.paragraphs[0]
    p_t.text = convert_numbers(data.get("title", "پریزێنتەیشن"), is_rtl)
    p_t.font.size = PptxPt(38)
    p_t.font.bold = True
    p_t.font.color.rgb = ACCENT_GOLD
    p_t.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
    if is_rtl: p_t._pPr.set('rtl', '1')
    
    p_sub = tf1.add_paragraph()
    p_sub.text = convert_numbers(f"{dept or 'زانکۆ / پەیمانگەهـ'}", is_rtl)
    p_sub.font.size = PptxPt(22)
    p_sub.font.color.rgb = WHITE
    p_sub.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
    if is_rtl: p_sub._pPr.set('rtl', '1')
    
    p_inf = tf1.add_paragraph()
    lbl_s = "پێشکێشکار: " if is_rtl else "Presenter: "
    lbl_t = " | سەرپەرشت: " if is_rtl else " | Supervisor: "
    p_inf.text = convert_numbers(f"\n{lbl_s}{student or '-'} {lbl_t}{teacher or '-'}", is_rtl)
    p_inf.font.size = PptxPt(16)
    p_inf.font.color.rgb = LIGHT_GRAY
    p_inf.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
    if is_rtl: p_inf._pPr.set('rtl', '1')
    
    # سلایدێن ناڤەرۆکێ دگەل وێنەیان
    for s_item in data.get("slides", []):
        sl = prs.slides.add_slide(blank_layout)
        
        # باکگراوند
        bg = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = DARK_BG
        bg.line.fill.background()
        
        # سەردێڕێ سەرەکی یێ سلایدێ
        t_box = sl.shapes.add_textbox(PptxInches(1), PptxInches(0.6), PptxInches(11.333), PptxInches(1))
        t_frame = t_box.text_frame
        t_para = t_frame.paragraphs[0]
        t_para.text = convert_numbers(s_item.get("slide_title", ""), is_rtl)
        t_para.font.size = PptxPt(28)
        t_para.font.bold = True
        t_para.font.color.rgb = ACCENT_GOLD
        t_para.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
        if is_rtl: t_para._pPr.set('rtl', '1')
        
        # دابەشکرنا سلایدێ بۆ دوو بەشان: دەق و وێنە
        if is_rtl:
            # بۆ کوردی و عەرەبی: دەق ل لایێ ڕاستێ، وێنە ل لایێ چەپێ
            text_left, text_width = PptxInches(6.8), PptxInches(5.5)
            img_left, img_top, img_width = PptxInches(1.0), PptxInches(1.8), PptxInches(5.2)
        else:
            # بۆ ئینگلیزی: دەق ل چەپێ، وێنە ل ڕاستێ
            text_left, text_width = PptxInches(1.0), PptxInches(5.5)
            img_left, img_top, img_width = PptxInches(7.0), PptxInches(1.8), PptxInches(5.2)
            
        # کارتا دەقی
        card = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, text_left, PptxInches(1.8), text_width, PptxInches(5.0))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = ACCENT_GOLD
        card.line.width = PptxPt(1)
        
        c_box = sl.shapes.add_textbox(text_left + PptxInches(0.2), PptxInches(2.0), text_width - PptxInches(0.4), PptxInches(4.6))
        c_frame = c_box.text_frame
        c_frame.word_wrap = True
        
        for idx, pt in enumerate(s_item.get("bullet_points", [])):
            para = c_frame.paragraphs[0] if idx == 0 else c_frame.add_paragraph()
            para.text = f"•  {convert_numbers(pt, is_rtl)}"
            para.font.size = PptxPt(17)
            para.font.color.rgb = WHITE
            para.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
            para.space_after = PptxPt(16)
            if is_rtl: para._pPr.set('rtl', '1')
            
        # دابەزاندن و دانانا وێنەیێ گونجای بۆ ڤێ سلایدێ
        img_keyword = s_item.get("image_keyword", "education")
        img_data = fetch_slide_image(img_keyword)
        if img_data:
            try:
                sl.shapes.add_picture(img_data, img_left, img_top, width=img_width)
            except Exception:
                pass
            
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

is_rtl_lang = language != "English"

# دەستپێکرنا دروستکردنێ
if st.button("🚀 دروستکرنا ڕاپۆرت و سمینارێ", type="primary"):
    if not api_key:
        st.error("تکایە کلیلا API بنڤیسە.")
    elif not topic:
        st.warning("تکایە بابەتێ ڕاپۆرتێ بنڤیسە.")
    else:
        with st.spinner("داتایێن زانستی ب تێر و تەسەلی دهێنە ئامادەکرن و وێنەیێن سمینارێ دادبەزن... چەند چرکەیان بگرە"):
            try:
                content = generate_academic_content(topic, language, pages_count, student_name, department, teacher_name, api_key)
                
                st.session_state["docx_file"] = build_docx(content, student_name, department, teacher_name, is_rtl_lang).getvalue()
                st.session_state["pptx_file"] = build_pptx(content, student_name, department, teacher_name, is_rtl_lang).getvalue()
                st.session_state["topic_name"] = topic
                st.session_state["generated"] = True
                st.success("✅ ڕاپۆرتا بەرفرەهـ و سمینارا وێنەدار ب سەرکەفتیانە هاتنە دروستکرن!")
            except Exception as e:
                st.error(f"کێشەیەک ڕویدا: {e}")

# بەشێ دانلۆدکرنێ
if st.session_state.get("generated", False):
    st.markdown("### 📥 فایلێن خو داونلۆد بکە:")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            label="📄 داگرتنا فایلا Word (ڕاپۆرتا تێر و تەسەل)",
            data=st.session_state["docx_file"],
            file_name=f"{st.session_state['topic_name']}_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    with col_d2:
        st.download_button(
            label="📊 داگرتنا فایلا PowerPoint (دگەل وێنەیێن تایبەت)",
            data=st.session_state["pptx_file"],
            file_name=f"{st.session_state['topic_name']}_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
