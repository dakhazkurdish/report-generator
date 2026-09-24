import streamlit as st
import requests
import json
import io
import urllib.parse

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

st.set_page_config(page_title="سیستەمێ زیرەک یێ دروستکرنا راپورت و سمیناران", page_icon="🎓", layout="wide")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    p, h1, h2, h3, label, div { text-align: right !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; background-color: #1E3A8A; color: white; height: 3.2em; font-size: 16px; }
    </style>
""", unsafe_allow_html=True)

# ناڤنیشانێ سەرەکی بتنێ
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

# وێنەگرتن ب ڕێکا ڕاستەوخۆ یا مسۆگەر (LoremPicsum + Wikipedia Search)
def fetch_slide_image(keyword, seed=1):
    clean_kw = urllib.parse.quote(str(keyword).strip() or "science")
    try:
        search_url = f"https://en.wikipedia.org/w/api.php?action=query&format=json&generator=search&gsrsearch={clean_kw}&gsrlimit=1&prop=pageimages&pithumbsize=800"
        headers = {"User-Agent": "AcademicSlideGen/2.0"}
        r = requests.get(search_url, headers=headers, timeout=4)
        if r.status_code == 200:
            pages = r.json().get("query", {}).get("pages", {})
            for _, p_info in pages.items():
                thumb = p_info.get("thumbnail", {}).get("source")
                if thumb:
                    img_r = requests.get(thumb, headers=headers, timeout=5)
                    if img_r.status_code == 200 and len(img_r.content) > 2000:
                        return io.BytesIO(img_r.content)
    except Exception:
        pass
    
    # ئەگەر ژ ویکیپیدیا نەهات، وێنەیەکێ باڵاکێش و سەردەمیانە بینە دا سلاید بەتاڵ نەمینیت
    try:
        fallback_url = f"https://picsum.photos/seed/{abs(hash(clean_kw)) % 1000 + seed}/800/600"
        fb_r = requests.get(fallback_url, timeout=4)
        if fb_r.status_code == 200:
            return io.BytesIO(fb_r.content)
    except Exception:
        return None
    return None

def generate_academic_content(topic, lang, pages, student, dept, teacher, key):
    candidate_models = ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-2.5-flash"]
    headers = {"Content-Type": "application/json"}
    
    # گەلەک ب توندی داخوازا دەقێ درێژ و پڕ هاتییە کرن
    prompt = f"""
    You are a distinguished university professor writing a formal, extensive, in-depth academic research paper.
    TOPIC: "{topic}"
    LANGUAGE: {lang}
    STUDENT: "{student}", DEPARTMENT: "{dept}", SUPERVISOR: "{teacher}".
    TARGET PAGES: At least {pages} pages of rigorous scholarly content.
    
    MANDATORY RULES:
    1. Write LONG, COMPREHENSIVE multi-paragraph texts for each section. Do not give summaries or short bullet points in the report. Each section must thoroughly analyze the principles, history, real-world case studies, technical aspects, and debates around the topic.
    2. Provide 6 to 8 structured presentation slides. Each slide MUST have a concrete English keyword for `image_keyword` (e.g. "computer network", "renewable energy", "biotechnology", "medical hospital").

    Return ONLY a raw JSON object (NO markdown fences ```json ```):
    {{
        "title": "Full Academic Research Title",
        "abstract": "Extensive academic abstract (at least 200-300 words).",
        "sections": [
            {{
                "heading": "Section 1: Detailed Title",
                "content": "Paragraph 1...\\n\\nParagraph 2...\\n\\nParagraph 3...\\n\\nParagraph 4..."
            }},
            {{
                "heading": "Section 2: Detailed Title",
                "content": "Paragraph 1...\\n\\nParagraph 2...\\n\\nParagraph 3...\\n\\nParagraph 4..."
            }},
            {{
                "heading": "Section 3: Detailed Title",
                "content": "Paragraph 1...\\n\\nParagraph 2...\\n\\nParagraph 3...\\n\\nParagraph 4..."
            }},
            {{
                "heading": "Section 4: Detailed Title",
                "content": "Paragraph 1...\\n\\nParagraph 2...\\n\\nParagraph 3..."
            }},
            {{
                "heading": "Section 5: Detailed Title",
                "content": "Paragraph 1...\\n\\nParagraph 2...\\n\\nParagraph 3..."
            }}
        ],
        "conclusion": "Detailed, formal academic conclusion (multiple in-depth paragraphs).",
        "references": [
            "Full APA Reference 1",
            "Full APA Reference 2",
            "Full APA Reference 3",
            "Full APA Reference 4",
            "Full APA Reference 5"
        ],
        "slides": [
            {{
                "slide_title": "Slide Title",
                "image_keyword": "concrete_english_concept",
                "bullet_points": [
                    "Direct takeaway 1",
                    "Direct takeaway 2",
                    "Direct takeaway 3"
                ]
            }}
        ]
    }}
    """
    
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "response_mime_type": "application/json",
            "maxOutputTokens": 8192  # دەرگەهـ بهێلە ڤەکری دا کورت نەکەت
        }
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
    
    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    # لاپەڕا سەرەکی (Cover Page)
    p_uni = doc.add_paragraph()
    set_docx_rtl(p_uni, is_rtl)
    r_uni = p_uni.add_run(convert_numbers(dept or "پەیمانگەهـ / زانکۆ", is_rtl))
    format_run(r_uni, size_pt=16, bold=True, color_rgb=(30, 41, 59), is_rtl=is_rtl)
    
    p_title = doc.add_paragraph()
    set_docx_rtl(p_title, is_rtl)
    p_title.paragraph_format.space_before = Pt(54)
    p_title.paragraph_format.space_after = Pt(36)
    r_title = p_title.add_run(convert_numbers(data.get("title", "ڕاپۆرتا زانستی"), is_rtl))
    format_run(r_title, size_pt=22, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
    
    p_box = doc.add_paragraph()
    set_docx_rtl(p_box, is_rtl)
    p_box.paragraph_format.space_before = Pt(40)
    lbl_s = "ئامادەکرن ژ لایێ قوتابی: " if is_rtl else "Prepared by: "
    lbl_t = "سەرپەرشتیا مامۆستا: " if is_rtl else "Supervised by: "
    
    r_meta = p_box.add_run(f"📋 {lbl_s}{student or '-'}\n\n👨‍🏫 {lbl_t}{teacher or '-'}\n\n📅 ساڵا ئەکادیمی: {convert_numbers('2025 - 2026', is_rtl)}")
    format_run(r_meta, size_pt=14, bold=True, color_rgb=(51, 65, 85), is_rtl=is_rtl)
    
    doc.add_page_break()
    
    # پوختە (Abstract)
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
    
    # پشکێن سەرەکی (Body)
    for idx, sec in enumerate(data.get("sections", [])):
        p_sec_h = doc.add_paragraph()
        set_docx_rtl(p_sec_h, is_rtl)
        p_sec_h.paragraph_format.space_before = Pt(20)
        p_sec_h.paragraph_format.space_after = Pt(8)
        r_sec_h = p_sec_h.add_run(convert_numbers(f"{idx+1}. {sec.get('heading', '')}", is_rtl))
        format_run(r_sec_h, size_pt=16, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
        
        raw_text = sec.get('content', '')
        paras = raw_text.split("\n\n") if "\n\n" in raw_text else [raw_text]
        for p_t in paras:
            if not p_t.strip(): continue
            p_sec = doc.add_paragraph()
            set_docx_rtl(p_sec, is_rtl)
            p_sec.paragraph_format.line_spacing = 1.3
            p_sec.paragraph_format.space_after = Pt(12)
            r_sec = p_sec.add_run(convert_numbers(p_t.strip(), is_rtl))
            format_run(r_sec, size_pt=14, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
            
    # دەرئەنجام
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
    
    # سەرچاوەکان
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
    
    # سلایدا دەستپێکێ
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
        
        if is_rtl:
            text_left, text_width = PptxInches(6.8), PptxInches(5.5)
            img_left, img_top, img_width = PptxInches(1.0), PptxInches(1.8), PptxInches(5.2)
        else:
            text_left, text_width = PptxInches(1.0), PptxInches(5.5)
            img_left, img_top, img_width = PptxInches(7.0), PptxInches(1.8), PptxInches(5.2)
            
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
            
        # وێنە ب مسۆگەری
        img_keyword = s_item.get("image_keyword", "")
        img_data = fetch_slide_image(img_keyword, seed=idx_s)
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

if st.button("🚀 دروستکرنا ڕاپۆرت و سمینارێ", type="primary"):
    if not api_key:
        st.error("تکایە کلیلا API بنڤیسە.")
    elif not topic:
        st.warning("تکایە بابەتێ ڕاپۆرتێ بنڤیسە.")
    else:
        with st.spinner("داتایێن زانستی ب تێر و تەسەلی دهێنە ئامادەکرن... تکایە تا خولەکەکێ بگرە"):
            try:
                content = generate_academic_content(topic, language, pages_count, student_name, department, teacher_name, api_key)
                st.session_state["docx_file"] = build_docx(content, student_name, department, teacher_name, is_rtl_lang).getvalue()
                st.session_state["pptx_file"] = build_pptx(content, student_name, department, teacher_name, is_rtl_lang).getvalue()
                st.session_state["topic_name"] = topic
                st.session_state["generated"] = True
                st.success("✅ ڕاپۆرت و سمینار ب سەرکەفتیانە و ب شێوازێ ستاندارد ئامادە بوون!")
            except Exception as e:
                st.error(f"کێشەیەک ڕویدا: {e}")

if st.session_state.get("generated", False):
    st.markdown("### 📥 فایلێن خو داونلۆد بکە:")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            label="📄 داگرتنا فایلا Word (ڕاپۆرتا ستاندارد)",
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
