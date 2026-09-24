import streamlit as st
import requests
import json
import io
import time

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

st.set_page_config(page_title="دروستکەرێ ڕاپۆرت و سمیناران", page_icon="🎓", layout="centered")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    p, h1, h2, h3, label { text-align: right !important; }
    </style>
""", unsafe_allow_html=True)

st.title("🎓 پلاتفۆڕمێ دروستکرنا ڕاپۆرت و سمیناران")
st.write("زانیارییان تژی بکە دا کو ڕاپۆرتەکا ئەکادیمی یا تێر و تەسەل دگەل پاوەرپۆینتەکا دیزاینکری یا پێشکەفتی ب دەست بێخی.")

# وەرگرتنا کلیلا API
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    api_key = st.text_input("کلیلا Gemini API لێرە دابنێ:", type="password")

# فۆڕما سەرەکی
with st.container():
    topic = st.text_input("بابەتێ سەرەکی یێ ڕاپۆرتێ (Topic) *", placeholder="بۆ نموونە: Artificial Intelligence in Modern Medicine")
    
    col1, col2 = st.columns(2)
    with col1:
        student_name = st.text_input("ناڤێ قوتابی:", placeholder="بۆ نموونە: ئەحمەد سەعید")
    with col2:
        department = st.text_input("پشک یان کۆلێژ:", placeholder="بۆ نموونە: زانستێن کۆمپیۆتەری")
        
    col3, col4 = st.columns(2)
    with col3:
        supervisor = st.text_input("ناڤێ مامۆستایێ بابەتی / سەرپەرشتیار:", placeholder="بۆ نموونە: د. ئاراس خالد")
    with col4:
        language = st.selectbox("زمانێ نڤیسینێ:", ["English", "کوردی (بادینی)", "کوردی (سۆرانی)", "العربية"])
        
    pages_count = st.slider("ژمارا لاپەڕێن پێدڤی بۆ ڕاپۆرتێ:", min_value=2, max_value=10, value=4)

def generate_academic_content(topic, name, dept, sup, lang, pages, key):
    # لیستا مۆدێلێن نوو و فەرمی
    candidate_models = ["gemini-3.8-flash", "gemini-3.5-flash", "gemini-3.5-flash-lite"]
    headers = {"Content-Type": "application/json"}
    
    prompt = f"""
    You are an expert university professor and professional academic writer. Write an extensive, rigorous, deeply technical academic research report and a structured presentation on:
    Topic: "{topic}"
    Target Language: {lang}
    Target Depth & Length: Very thorough, formatted across approximately {pages} full pages.
    Student: {name} | Department: {dept} | Supervisor: {sup}

    REQUIREMENTS:
    1. Introduction must be comprehensive with clear objectives, historical/theoretical context, and research scope.
    2. Provide at least 4 to 6 detailed, multi-paragraph core sections explaining mechanisms, methodologies, case examples, practical implications, and future outlook.
    3. Conclusion must summarize critical findings and future research directions.
    4. Provide 4-7 formal references in strict APA format.
    5. Provide at least 5 to 7 presentation slides. For each slide, write a title and 3-4 rich, meaningful bullet points.

    Output STRICTLY valid raw JSON without markdown markers:
    {{
        "title": "Academic Research Title",
        "student_name": "{name}",
        "department": "{dept}",
        "supervisor": "{sup}",
        "introduction": "Comprehensive introduction text...",
        "sections": [
            {{"heading": "Section Heading", "content": "Thorough multi-paragraph content..."}}
        ],
        "conclusion": "Comprehensive conclusion text...",
        "references": ["Reference 1", "Reference 2", "Reference 3"],
        "slides": [
            {{"slide_title": "Slide Title", "bullet_points": ["Point 1", "Point 2", "Point 3"]}}
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
            res = requests.post(url, headers=headers, json=payload, timeout=90)
            if res.status_code == 200:
                clean_json = res.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                return json.loads(clean_json)
            elif res.status_code == 503:
                time.sleep(1.5)
                continue
            else:
                last_error = res.text
        except Exception as e:
            last_error = str(e)
            continue
            
    raise Exception(f"API Error: {last_error}")

def create_rich_docx(data):
    doc = Document()
    
    # 1. بەرگێ فەرمی (Cover Page)
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = p_title.add_run("\n\n" + data.get("title", "Academic Report") + "\n")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(26)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(15, 23, 42) # Slate Dark
    
    p_divider = doc.add_paragraph()
    p_divider.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_divider.add_run("________________________________________\n\n")
    r_div.font.color.rgb = RGBColor(37, 99, 235)
    
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if data.get("student_name"):
        p_meta.add_run(f"ئامادەکرن: {data['student_name']}\n").font.size = Pt(14)
    if data.get("department"):
        p_meta.add_run(f"پشک / کۆلێژ: {data['department']}\n").font.size = Pt(13)
    if data.get("supervisor"):
        p_meta.add_run(f"سەرپەرشتیار: {data['supervisor']}\n").font.size = Pt(13)
    p_meta.add_run(f"\nڕاپۆرتا زانستی یا ئەکادیمی\n").font.size = Pt(12)
    
    doc.add_page_break()
    
    # 2. پێشەکی
    h1 = doc.add_heading("1. پێشەکی (Introduction)", level=1)
    h1.runs[0].font.color.rgb = RGBColor(30, 64, 175)
    p_intro = doc.add_paragraph(data.get("introduction", ""))
    p_intro.paragraph_format.line_spacing = 1.25
    p_intro.paragraph_format.space_after = Pt(12)
    
    # 3. بەش و تەوەرێن سەرەکی
    sections = data.get("sections", [])
    for idx, sec in enumerate(sections, start=2):
        h = doc.add_heading(f"{idx}. {sec.get('heading', '')}", level=1)
        h.runs[0].font.color.rgb = RGBColor(30, 64, 175)
        p = doc.add_paragraph(sec.get("content", ""))
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.space_after = Pt(12)
        
    # 4. دەرئەنجام
    h_c = doc.add_heading(f"{len(sections) + 2}. دەرئەنجام (Conclusion)", level=1)
    h_c.runs[0].font.color.rgb = RGBColor(30, 64, 175)
    p_c = doc.add_paragraph(data.get("conclusion", ""))
    p_c.paragraph_format.line_spacing = 1.25
    p_c.paragraph_format.space_after = Pt(12)
    
    # 5. سەرچاوە
    h_r = doc.add_heading(f"{len(sections) + 3}. سەرچاوەکان (References)", level=1)
    h_r.runs[0].font.color.rgb = RGBColor(30, 64, 175)
    for ref in data.get("references", []):
        doc.add_paragraph(ref, style='List Bullet')
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def create_styled_pptx(data):
    prs = Presentation()
    prs.slide_width = PptInches(13.33)  # شاشەیا پان 16:9
    prs.slide_height = PptInches(7.5)
    
    # ڕەنگێن مۆدێرن (Dark Navy Theme)
    BG_COLOR = PptRGBColor(15, 23, 42)       # ڕەش/شینێ تاری
    ACCENT_COLOR = PptRGBColor(56, 189, 248) # شینێ گەش (Sky Blue)
    CARD_COLOR = PptRGBColor(30, 41, 59)     # ڕەنگێ باکگراوندێ کاردێ
    TEXT_WHITE = PptRGBColor(255, 255, 255)
    TEXT_MUTED = PptRGBColor(203, 213, 225)
    
    blank_layout = prs.slide_layouts[6]
    
    # 1. سلایدا سەرەکی (Cover Slide)
    s0 = prs.slides.add_slide(blank_layout)
    bg0 = s0.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, PptInches(13.33), PptInches(7.5))
    bg0.fill.solid()
    bg0.fill.fore_color.rgb = BG_COLOR
    bg0.line.fill.background()
    
    # کاردێ ناڤەند
    card0 = s0.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptInches(1.2), PptInches(1.2), PptInches(10.93), PptInches(5.1))
    card0.fill.solid()
    card0.fill.fore_color.rgb = CARD_COLOR
    card0.line.color.rgb = ACCENT_COLOR
    card0.line.width = PptPt(2)
    
    tf0 = card0.text_frame
    tf0.word_wrap = True
    
    p0 = tf0.paragraphs[0]
    p0.text = data.get("title", "پریزێنتەیشن")
    p0.font.bold = True
    p0.font.size = PptPt(36)
    p0.font.color.rgb = TEXT_WHITE
    p0.alignment = PP_ALIGN.CENTER
    
    p_meta = tf0.add_paragraph()
    p_meta.text = f"\nئامادەکرن: {data.get('student_name', '')}  |  پشک: {data.get('department', '')}\nسەرپەرشتیار: {data.get('supervisor', '')}"
    p_meta.font.size = PptPt(20)
    p_meta.font.color.rgb = ACCENT_COLOR
    p_meta.alignment = PP_ALIGN.CENTER
    
    # 2. سلایدێن ناڤەرۆکێ
    for s_info in data.get("slides", []):
        slide = prs.slides.add_slide(blank_layout)
        
        # باکگراوندێ تاری
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, PptInches(13.33), PptInches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_COLOR
        bg.line.fill.background()
        
        # سەرنڤیسێ سلایدێ
        title_box = slide.shapes.add_textbox(PptInches(1.0), PptInches(0.6), PptInches(11.33), PptInches(1.0))
        tf_t = title_box.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.text = s_info.get("slide_title", "")
        p_t.font.bold = True
        p_t.font.size = PptPt(30)
        p_t.font.color.rgb = ACCENT_COLOR
        
        # کاردێ ناڤەرۆکێ
        content_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptInches(1.0), PptInches(1.8), PptInches(11.33), PptInches(4.8))
        content_card.fill.solid()
        content_card.fill.fore_color.rgb = CARD_COLOR
        content_card.line.fill.background()
        
        tf_c = content_card.text_frame
        tf_c.word_wrap = True
        
        points = s_info.get("bullet_points", [])
        for i, pt in enumerate(points):
            p = tf_c.paragraphs[0] if i == 0 else tf_c.add_paragraph()
            p.text = f"•  {pt}\n"
            p.font.size = PptPt(20)
            p.font.color.rgb = TEXT_MUTED
            
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

# دوگمەیا کارپێکرنێ
if st.button("🚀 دروستکرنا ڕاپۆرت و سمینارێ", type="primary", use_container_width=True):
    if not api_key:
        st.error("تکایە دەستپێکێ کلیلا API بنڤیسە یان د بەشێ سێرڤەری دا دابنێ.")
    elif not topic:
        st.warning("تکایە ناڤێ بابەتێ ڕاپۆرتێ بنڤیسە.")
    else:
        with st.spinner("داتایێن زانستی دهێنە کۆمکرن، دیزاینا سلایدان دهێتە دانان و فایل ئامادە دبن... چەند چرکەیان بگرە"):
            try:
                data = generate_academic_content(topic, student_name, department, supervisor, language, pages_count, api_key)
                
                docx_file = create_rich_docx(data)
                pptx_file = create_styled_pptx(data)
                
                st.success("✅ ب سەرکەفتیانە ڕاپۆرت و سمینار هاتنە دروستکرن!")
                
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.download_button(
                        label="📄 داگرتنا فایلا Word (.docx)",
                        data=docx_file,
                        file_name=f"{topic[:25]}_report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with col_d2:
                    st.download_button(
                        label="📊 داگرتنا فایلا PowerPoint (.pptx)",
                        data=pptx_file,
                        file_name=f"{topic[:25]}_presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"کێشەیەک ڕویدا: {e}")
