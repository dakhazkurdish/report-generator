import streamlit as st
import requests
import json
import io
import urllib.parse

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

st.set_page_config(page_title="سیستەمێ زیرەک یێ دروستکرنا راپورت و سمیناران", page_icon="🎓", layout="wide", initial_sidebar_state="collapsed")

# ڤەشارتنا هەمی ئایکۆنێن گیت‌هاب، پێنوس، هێدەر و فۆتەرێن ستریملیت دا کو کەس کۆدی نەبینیت
st.markdown("""
    <style>
    #MainMenu {visibility: hidden !important;}
    header {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    .viewerBadge_container__1QSob {display: none !important;}
    .styles_viewerBadge__1yB5G {display: none !important;}
    [data-testid="stToolbar"] {visibility: hidden !important; display: none !important;}
    [data-testid="stDecoration"] {visibility: hidden !important; display: none !important;}
    [data-testid="stStatusWidget"] {visibility: hidden !important; display: none !important;}
    
    .stApp { direction: rtl; text-align: right; }
    p, h1, h2, h3, label, div { text-align: right !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; background-color: #1E3A8A; color: white; height: 3.2em; font-size: 16px; }
    textarea { direction: rtl !important; text-align: right !important; font-family: 'Segoe UI', Tahoma, sans-serif !important; }
    </style>
""", unsafe_allow_html=True)

st.title("سیستەمێ زیرەک یێ دروستکرنا راپورت و سمیناران")

api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    api_key = st.text_input("کلیلا Gemini API لێرە بنڤیسە:", type="password")

with st.container():
    col1, col2 = st.columns(2)
    with col1:
        student_name = st.text_input("👤 ناڤێ قوتابی:")
        department = st.text_input("🏛️ پەیمانگەهـ یان کۆلێژ / پشک:")
        language = st.selectbox("🌐 زمانێ نڤیسینێ:", ["کوردی (بادینی)", "کوردی (سۆرانی)", "العربية", "English"])
    with col2:
        teacher_name = st.text_input("👨‍🏫 ناڤێ مامۆستایێ بابەتی:")
        topic = st.text_input("📝 بابەتێ سەرەکی یێ ڕاپۆرتێ:")
        pages_count = st.slider("📄 ژمارا لاپەڕێن پێدڤی بۆ ڕاپۆرتێ:", min_value=3, max_value=25, value=12)

# بژاردەیا چوارچێوەیێ ڕاپۆرتێ
enable_border = st.checkbox("🖼️ چوارچێوە (Border) بۆ لاپەڕێن ڕاپۆرتا Word بهێتە دانان؟", value=True)

# سێرچ بۆکسێ تێبینی و ڕێنماییێن تایبەت
custom_notes = st.text_area(
    "💡 تێبینی یان داخوازیێن تایبەت (بتنێ فەرمانە، ناچیتە ناڤ ڕاپۆرتێ):",
    placeholder="بۆ نموونە: گرنگیێ ب مێژوویا بابەتی بدە، نموونەیێن کرداری ل سەر عێراقێ بینە، ئاستێ زمانێ ئەکادیمی گەلەک بلند بیت...",
    height=80
)

def convert_numbers(text, is_rtl):
    if not is_rtl or not text:
        return text
    western_to_eastern = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")
    return str(text).translate(western_to_eastern)

def set_docx_rtl(paragraph, is_rtl=True):
    if not is_rtl:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

def format_run(run, font_name="Calibri", size_pt=14, bold=False, color_rgb=(0, 0, 0), is_rtl=True):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)
    if is_rtl:
        rPr = run._r.get_or_add_rPr()
        rtl = OxmlElement('w:rtl')
        rtl.set(qn('w:val'), '1')
        rPr.append(rtl)

# زێدەکرنا ژمارا لاپەڕەی د بنی دا
def add_page_number_to_section(section, is_rtl):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)
    format_run(run, size_pt=10, bold=False, color_rgb=(120, 130, 140), is_rtl=is_rtl)

# زێدەکرنا چوارچێوەیێ لاپەڕەی (Page Border)
def add_page_borders(section):
    sectPr = section._sectPr
    pgBorders = parse_xml(r'''
        <w:pgBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:offsetFrom="page">
            <w:top w:val="single" w:sz="12" w:space="24" w:color="1E3A8A"/>
            <w:left w:val="single" w:sz="12" w:space="24" w:color="1E3A8A"/>
            <w:bottom w:val="single" w:sz="12" w:space="24" w:color="1E3A8A"/>
            <w:right w:val="single" w:sz="12" w:space="24" w:color="1E3A8A"/>
        </w:pgBorders>
    ''')
    sectPr.append(pgBorders)

# ئینانا لۆگۆیێ ئەکادیمی یان زانکۆیی
def fetch_academic_logo(dept_name):
    headers = {"User-Agent": "AcademicSlideGen/5.0"}
    if dept_name:
        try:
            clean_dept = urllib.parse.quote(dept_name.strip())
            search_url = f"https://en.wikipedia.org/w/api.php?action=query&format=json&generator=search&gsrsearch={clean_dept}%20logo&gsrlimit=1&prop=pageimages&pithumbsize=400"
            r = requests.get(search_url, headers=headers, timeout=4)
            if r.status_code == 200:
                pages = r.json().get("query", {}).get("pages", {})
                for _, p_info in pages.items():
                    thumb = p_info.get("thumbnail", {}).get("source")
                    if thumb:
                        img_r = requests.get(thumb, headers=headers, timeout=4)
                        if img_r.status_code == 200 and len(img_r.content) > 1500:
                            return io.BytesIO(img_r.content)
        except Exception:
            pass
    return None

def fetch_slide_image(keyword, topic_context=""):
    search_terms = [keyword, topic_context]
    headers = {"User-Agent": "AcademicSlideGen/5.0"}
    for term in search_terms:
        if not term: continue
        clean_kw = urllib.parse.quote(str(term).strip())
        try:
            search_url = f"https://en.wikipedia.org/w/api.php?action=query&format=json&generator=search&gsrsearch={clean_kw}&gsrlimit=3&prop=pageimages&pithumbsize=900"
            r = requests.get(search_url, headers=headers, timeout=5)
            if r.status_code == 200:
                pages = r.json().get("query", {}).get("pages", {})
                for _, p_info in pages.items():
                    thumb = p_info.get("thumbnail", {}).get("source")
                    if thumb:
                        img_r = requests.get(thumb, headers=headers, timeout=5)
                        if img_r.status_code == 200 and len(img_r.content) > 3000:
                            return io.BytesIO(img_r.content)
        except Exception:
            continue
    return None

def call_gemini(prompt, key, as_json=True):
    candidate_models = ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-2.5-flash"]
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "response_mime_type": "application/json" if as_json else "text/plain",
            "maxOutputTokens": 8192
        }
    }
    last_error = ""
    for model_name in candidate_models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={key}"
        try:
            res = requests.post(url, headers=headers, json=payload, timeout=90)
            if res.status_code == 200:
                result = res.json()
                text_content = result["candidates"][0]["content"]["parts"][0]["text"]
                return json.loads(text_content) if as_json else text_content
            else:
                last_error = res.text
        except Exception as e:
            last_error = str(e)
    raise Exception(f"API Error: {last_error}")

def generate_multi_step_report(topic, lang, pages, student, dept, teacher, notes, key, progress_bar, status_text):
    num_sections = max(4, pages - 2)
    
    notes_prompt_part = ""
    if notes.strip():
        notes_prompt_part = f"""
        CRITICAL OPERATIONAL INSTRUCTIONS FROM USER:
        "{notes.strip()}"
        Apply these instructions into the content and focus of the research SILENTLY. 
        DO NOT quote, mention, or print these instructions or phrases anywhere in the generated output text.
        """
        
    status_text.write("قۆناغا ١: پلان و نەخشەڕێیا ڕاپۆرتێ و سمینارێ دهێتە دارشتن...")
    progress_bar.progress(10)
    
    plan_prompt = f"""
    You are an esteemed university professor and thesis advisor.
    Create a comprehensive academic research outline and presentation slides for the topic: "{topic}".
    Target Language: {lang}.
    Student: "{student}", Department: "{dept}", Supervisor: "{teacher}".
    Required Sections count: {num_sections}.
    {notes_prompt_part}

    CRITICAL RULES:
    1. Presentation slides MUST be 100% written in {lang}. Do not write bullet points in English unless target language is English.
    2. For each slide, provide an exact, highly specific English query for `image_search_query` that specifically describes the topic of that slide.

    Return strictly a JSON object:
    {{
        "title": "Full Academic Title in {lang}",
        "abstract": "Extensive abstract in {lang} (200-300 words)",
        "english_main_topic": "English translation of the main topic for image searching",
        "section_titles": [
            "Section Title 1",
            "Section Title 2"
        ],
        "slides": [
            {{
                "slide_title": "Slide Title in {lang}",
                "image_search_query": "specific english search terms for exact image",
                "bullet_points": ["Point 1 in {lang}", "Point 2 in {lang}", "Point 3 in {lang}"]
            }}
        ]
    }}
    """
    plan = call_gemini(plan_prompt, key, as_json=True)
    
    sections = []
    sec_titles = plan.get("section_titles", [])
    total_secs = len(sec_titles)
    main_en_topic = plan.get("english_main_topic", topic)
    
    for i, s_title in enumerate(sec_titles):
        status_text.write(f"قۆناغا ٢: نڤیسینا بەرفرەهـ یا تەوەرێ ({i+1} ژ {total_secs}): {s_title}...")
        pct = 15 + int((i + 1) / total_secs * 65)
        progress_bar.progress(pct)
        
        sec_prompt = f"""
        You are writing Section {i+1} of a comprehensive academic thesis on the topic: "{topic}".
        Language: {lang}.
        Section Title: "{s_title}".
        {notes_prompt_part}
        
        Write an EXTREMELY IN-DEPTH, MULTI-PARAGRAPH scholarly text for this section alone.
        DO NOT summarize. Include historical depth, scientific definitions, analytical breakdowns, practical examples, and real-world implications.
        Length: Write at least 4 to 6 large paragraphs for this section.
        Return ONLY plain text for this section's content.
        """
        sec_content = call_gemini(sec_prompt, key, as_json=False)
        sections.append({
            "heading": s_title,
            "content": sec_content.strip()
        })
        
    status_text.write("قۆناغا ٣: دەرئەنجام و لیستا سەرچاوەیان (APA) دهێنە دارشتن...")
    progress_bar.progress(85)
    
    ending_prompt = f"""
    Write a formal conclusion and APA academic references for the research "{topic}".
    Language: {lang}.
    {notes_prompt_part}
    Return strictly a JSON object:
    {{
        "conclusion": "Detailed multi-paragraph conclusion in {lang}",
        "references": [
            "Full APA Reference 1",
            "Full APA Reference 2",
            "Full APA Reference 3",
            "Full APA Reference 4",
            "Full APA Reference 5"
        ]
    }}
    """
    ending = call_gemini(ending_prompt, key, as_json=True)
    
    progress_bar.progress(100)
    status_text.write("فایلێن Word و PowerPoint ب شێوەیێ ستاندارد ئامادە دبن...")
    
    return {
        "title": plan.get("title", topic),
        "abstract": plan.get("abstract", ""),
        "sections": sections,
        "conclusion": ending.get("conclusion", ""),
        "references": ending.get("references", []),
        "slides": plan.get("slides", []),
        "main_en_topic": main_en_topic
    }

def build_docx(data, student, dept, teacher, is_rtl, with_border=True):
    doc = Document()
    
    section_cover = doc.sections[0]
    section_cover.top_margin = Inches(1)
    section_cover.bottom_margin = Inches(1)
    section_cover.left_margin = Inches(1)
    section_cover.right_margin = Inches(1)
    
    if with_border:
        add_page_borders(section_cover)

    # 1. لاپەڕا سەرەکی (Cover Page)
    # ئینانا لۆگۆیێ ئەکادیمی ئەگەر هەبیت
    logo_data = fetch_academic_logo(dept)
    if logo_data:
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            doc.add_picture(logo_data, width=Inches(1.6))
        except Exception:
            pass
            
    p_uni = doc.add_paragraph()
    set_docx_rtl(p_uni, is_rtl)
    p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_dept = p_uni.add_run(convert_numbers(dept or "پەیمانگەهـ / زانکۆ", is_rtl))
    format_run(run_dept, size_pt=18, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
    
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_div.add_run("______________________________")
    format_run(r_div, size_pt=12, bold=True, color_rgb=(203, 213, 225), is_rtl=is_rtl)
    
    p_title = doc.add_paragraph()
    set_docx_rtl(p_title, is_rtl)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(36)
    p_title.paragraph_format.space_after = Pt(28)
    run_title = p_title.add_run(convert_numbers(data.get("title", "ڕاپۆرتا زانستی"), is_rtl))
    format_run(run_title, size_pt=24, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
    
    p_box = doc.add_paragraph()
    set_docx_rtl(p_box, is_rtl)
    p_box.paragraph_format.space_before = Pt(40)
    lbl_s = "ئامادەکرن ژ لایێ قوتابی: " if is_rtl else "Prepared by: "
    lbl_t = "سەرپەرشتیا مامۆستا: " if is_rtl else "Supervised by: "
    
    r_meta = p_box.add_run(f"📋 {lbl_s}{student or '-'}\n\n👨‍🏫 {lbl_t}{teacher or '-'}\n\n📅 ساڵا ئەکادیمی: {convert_numbers('2025 - 2026', is_rtl)}")
    format_run(r_meta, size_pt=14, bold=True, color_rgb=(51, 65, 85), is_rtl=is_rtl)
    
    # بەشێ دووێ: ناڤەرۆک (دگەل ژمارا لاپەڕەیان)
    section_body = doc.add_section()
    section_body.top_margin = Inches(1)
    section_body.bottom_margin = Inches(1)
    section_body.left_margin = Inches(1)
    section_body.right_margin = Inches(1)
    
    if with_border:
        add_page_borders(section_body)
    add_page_number_to_section(section_body, is_rtl)
    
    # Abstract
    p_abs_h = doc.add_paragraph()
    set_docx_rtl(p_abs_h, is_rtl)
    r_abs_h = p_abs_h.add_run("پوختە (Abstract)" if is_rtl else "Abstract")
    format_run(r_abs_h, size_pt=16, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
    
    p_abs = doc.add_paragraph()
    set_docx_rtl(p_abs, is_rtl)
    p_abs.paragraph_format.line_spacing = 1.3
    r_abs = p_abs.add_run(convert_numbers(data.get("abstract", ""), is_rtl))
    format_run(r_abs, size_pt=14, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
    doc.add_page_break()
    
    # Body Sections
    for idx, sec in enumerate(data.get("sections", [])):
        p_sec_h = doc.add_paragraph()
        set_docx_rtl(p_sec_h, is_rtl)
        p_sec_h.paragraph_format.space_before = Pt(20)
        p_sec_h.paragraph_format.space_after = Pt(8)
        r_sec_h = p_sec_h.add_run(convert_numbers(f"{idx+1}. {sec.get('heading', '')}", is_rtl))
        format_run(r_sec_h, size_pt=16, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
        
        paras = sec.get('content', '').split("\n\n")
        for p_t in paras:
            if not p_t.strip(): continue
            p_sec = doc.add_paragraph()
            set_docx_rtl(p_sec, is_rtl)
            p_sec.paragraph_format.line_spacing = 1.3
            p_sec.paragraph_format.space_after = Pt(12)
            r_sec = p_sec.add_run(convert_numbers(p_t.strip(), is_rtl))
            format_run(r_sec, size_pt=14, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
            
    # Conclusion
    p_con_h = doc.add_paragraph()
    set_docx_rtl(p_con_h, is_rtl)
    p_con_h.paragraph_format.space_before = Pt(22)
    r_con_h = p_con_h.add_run("دەرئەنجام (Conclusion)" if is_rtl else "Conclusion")
    format_run(r_con_h, size_pt=16, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
    
    p_con = doc.add_paragraph()
    set_docx_rtl(p_con, is_rtl)
    p_con.paragraph_format.line_spacing = 1.3
    r_con = p_con.add_run(convert_numbers(data.get("conclusion", ""), is_rtl))
    format_run(r_con, size_pt=14, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
    
    # References
    p_ref_h = doc.add_paragraph()
    set_docx_rtl(p_ref_h, is_rtl)
    p_ref_h.paragraph_format.space_before = Pt(24)
    r_ref_h = p_ref_h.add_run("سەرچاوەکان (References)" if is_rtl else "References")
    format_run(r_ref_h, size_pt=16, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
    
    for ref in data.get("references", []):
        p_ref = doc.add_paragraph()
        set_docx_rtl(p_ref, is_rtl)
        p_ref.paragraph_format.space_after = Pt(6)
        r_ref = p_ref.add_run(f"• {convert_numbers(ref, is_rtl)}")
        format_run(r_ref, size_pt=13, bold=False, color_rgb=(70, 80, 95), is_rtl=is_rtl)
        
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
    
    # سلایدا ئێکێ
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
    p_t.font.size = PptxPt(36)
    p_t.font.bold = True
    p_t.font.color.rgb = ACCENT_GOLD
    p_t.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
    if is_rtl: p_t._pPr.set('rtl', '1')
    
    p_sub = tf1.add_paragraph()
    p_sub.text = convert_numbers(f"{dept or 'پەیمانگەهـ / زانکۆ'}", is_rtl)
    p_sub.font.size = PptxPt(20)
    p_sub.font.color.rgb = WHITE
    p_sub.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
    if is_rtl: p_sub._pPr.set('rtl', '1')
    
    p_inf = tf1.add_paragraph()
    lbl_s = "قوتابی: " if is_rtl else "Student: "
    lbl_t = " | مامۆستا: " if is_rtl else " | Lecturer: "
    p_inf.text = convert_numbers(f"\n{lbl_s}{student or '-'} {lbl_t}{teacher or '-'}", is_rtl)
    p_inf.font.size = PptxPt(16)
    p_inf.font.color.rgb = LIGHT_GRAY
    p_inf.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
    if is_rtl: p_inf._pPr.set('rtl', '1')
    
    main_en_topic = data.get("main_en_topic", "")
    
    # سلایدێن ناڤەرۆکێ
    for idx_s, s_item in enumerate(data.get("slides", [])):
        sl = prs.slides.add_slide(blank_layout)
        bg = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = DARK_BG
        bg.line.fill.background()
        
        t_box = sl.shapes.add_textbox(PptxInches(1), PptxInches(0.6), PptxInches(11.333), PptxInches(1))
        t_frame = t_box.text_frame
        t_para = t_frame.paragraphs[0]
        t_para.text = convert_numbers(s_item.get("slide_title", ""), is_rtl)
        t_para.font.size = PptxPt(26)
        t_para.font.bold = True
        t_para.font.color.rgb = ACCENT_GOLD
        t_para.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
        if is_rtl: t_para._pPr.set('rtl', '1')
        
        img_query = s_item.get("image_search_query", "")
        img_data = fetch_slide_image(img_query, main_en_topic)
        
        if img_data:
            if is_rtl:
                text_left, text_width = PptxInches(6.8), PptxInches(5.5)
                img_left, img_top, img_width = PptxInches(1.0), PptxInches(1.8), PptxInches(5.2)
            else:
                text_left, text_width = PptxInches(1.0), PptxInches(5.5)
                img_left, img_top, img_width = PptxInches(7.0), PptxInches(1.8), PptxInches(5.2)
        else:
            text_left, text_width = PptxInches(1.5), PptxInches(10.333)
            
        card = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, text_left, PptxInches(1.8), text_width, PptxInches(5.0))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = ACCENT_GOLD
        card.line.width = PptxPt(1)
        
        c_box = sl.shapes.add_textbox(text_left + PptxInches(0.3), PptxInches(2.0), text_width - PptxInches(0.6), PptxInches(4.6))
        c_frame = c_box.text_frame
        c_frame.word_wrap = True
        
        for idx, pt in enumerate(s_item.get("bullet_points", [])):
            para = c_frame.paragraphs[0] if idx == 0 else c_frame.add_paragraph()
            para.text = f"•  {convert_numbers(pt, is_rtl)}"
            para.font.size = PptxPt(18)
            para.font.color.rgb = WHITE
            para.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
            para.space_after = PptxPt(16)
            if is_rtl: para._pPr.set('rtl', '1')
            
        if img_data:
            try:
                sl.shapes.add_picture(img_data, img_left, img_top, width=img_width)
            except Exception:
                pass
            
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

def build_plain_text(data, student, dept, teacher, is_rtl):
    txt = f"=========================================\n"
    txt += f"{data.get('title', 'ڕاپۆرت')}\n"
    txt += f"پەیمانگەهـ / زانکۆ: {dept}\n"
    txt += f"قوتابی: {student} | سەرپەرشتیا: {teacher}\n"
    txt += f"=========================================\n\n"
    txt += f"پوختە (Abstract):\n{data.get('abstract', '')}\n\n"
    for idx, sec in enumerate(data.get("sections", [])):
        txt += f"-----------------------------------------\n"
        txt += f"{idx+1}. {sec.get('heading', '')}\n"
        txt += f"-----------------------------------------\n"
        txt += f"{sec.get('content', '')}\n\n"
    txt += f"-----------------------------------------\n"
    txt += f"دەرئەنجام (Conclusion):\n"
    txt += f"{data.get('conclusion', '')}\n\n"
    txt += f"سەرچاوەکان (References):\n"
    for ref in data.get("references", []):
        txt += f"• {ref}\n"
    return convert_numbers(txt, is_rtl)

is_rtl_lang = language != "English"

if st.button("🚀 دروستکرنا ڕاپۆرت و سمینارێ", type="primary"):
    if not api_key:
        st.error("تکایە کلیلا API بنڤیسە.")
    elif not topic:
        st.warning("تکایە بابەتێ ڕاپۆرتێ بنڤیسە.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        try:
            content = generate_multi_step_report(topic, language, pages_count, student_name, department, teacher_name, custom_notes, api_key, progress_bar, status_text)
            
            st.session_state["docx_file"] = build_docx(content, student_name, department, teacher_name, is_rtl_lang, enable_border).getvalue()
            st.session_state["pptx_file"] = build_pptx(content, student_name, department, teacher_name, is_rtl_lang).getvalue()
            st.session_state["plain_text"] = build_plain_text(content, student_name, department, teacher_name, is_rtl_lang)
            st.session_state["topic_name"] = topic
            st.session_state["generated"] = True
            
            progress_bar.empty()
            status_text.empty()
            st.success("✅ ڕاپۆرت و سمینار ب شێوازێ ستاندارد ئامادە بوون!")
        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error(f"کێشەیەک ڕویدا: {e}")

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
            label="📊 داگرتنا فایلا PowerPoint (سلایدێن کوردی دگەل وێنەیێن تایبەت)",
            data=st.session_state["pptx_file"],
            file_name=f"{st.session_state['topic_name']}_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    
    st.markdown("---")
    st.markdown("### 📋 دەقێ ڕاپۆرتێ بۆ کۆپیکردنا ڕاستەوخۆ:")
    st.caption("دشێی ڤی دەقی دیاربکەی (Ctrl+A پاشان Ctrl+C) و پەیست بکەیە ناڤ وۆردێ خو بێی داگرتن:")
    st.text_area("", value=st.session_state["plain_text"], height=400)
