import streamlit as st
import requests
import json
import io
import re

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

st.title("🎓 سیستەمێ زیرەک و پیشەیی یێ دروستکرنا ڕاپۆرت و سمیناران")
st.write("ناڤەرۆکا ئەکادیمی یا تێر و تەسەل، دیزاینا ستاندارد، و دەرهێنانا فایلان ب شێوازێ فەرمی (RTL).")

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
        pages_count = st.slider("📄 ژمارا لاپەڕێن پێدڤی بۆ ڕاپۆرتێ:", min_value=2, max_value=8, value=3)

# گۆڕینا ژمارەیان بۆ عەرەبی ئەگەر پێدڤی بوو
def convert_numbers(text, is_rtl):
    if not is_rtl or not text:
        return text
    western_to_eastern = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")
    return str(text).translate(western_to_eastern)

# ڕێکخستنا RTL د Word دا
def set_docx_rtl(paragraph, is_rtl=True):
    if not is_rtl:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

# ڕێکخستنا فۆنتێ Word
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

def generate_academic_content(topic, lang, pages, student, dept, teacher, key):
    candidate_models = ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-2.5-flash"]
    headers = {"Content-Type": "application/json"}
    
    prompt = f"""
    You are an expert university professor and senior researcher. 
    Write an extensive, highly comprehensive, formal academic research report and a structured presentation for the following details:
    - Topic: "{topic}"
    - Target Language: {lang}
    - Student Name: "{student}"
    - Department/College: "{dept}"
    - Supervising Professor: "{teacher}"
    - Desired Depth: {pages} detailed academic pages (each section must contain thorough explanations, data, context, and depth).

    Return ONLY a single valid JSON object (no markdown quotes, no ```json ``` fences) with this structure:
    {{
        "title": "Clear, professional academic research title",
        "abstract": "A complete abstract summarizing research context, objectives, methodology, and key takeaways (150-250 words)",
        "sections": [
            {{
                "heading": "Section Title (e.g. Introduction & Background)",
                "content": "Deep, multi-paragraph, professional academic analysis and breakdown of this section."
            }},
            {{
                "heading": "Section Title (e.g. Analytical Dimensions / Practical Implementations)",
                "content": "Deep academic discussion with detailed examples and comprehensive points."
            }},
            {{
                "heading": "Section Title (e.g. Future Perspectives & Key Challenges)",
                "content": "Thorough analytical evaluation."
            }}
        ],
        "conclusion": "A detailed, structured conclusion summarizing conclusions and practical recommendations.",
        "references": [
            "Official academic reference 1 in APA format",
            "Official academic reference 2 in APA format",
            "Official academic reference 3 in APA format"
        ],
        "slides": [
            {{
                "slide_title": "Slide Title",
                "bullet_points": [
                    "Direct, concise presentation takeaway 1",
                    "Direct, concise presentation takeaway 2",
                    "Direct, concise presentation takeaway 3"
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
            response = requests.post(url, headers=headers, json=payload, timeout=90)
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
    
    # لاپەڕا سەرەکی (Cover Page)
    p_uni = doc.add_paragraph()
    set_docx_rtl(p_uni, is_rtl)
    run_dept = p_uni.add_run(convert_numbers(dept or "زانکۆ / پەیمانگەهـ", is_rtl))
    format_run(run_dept, size_pt=14, bold=True, color_rgb=(70, 80, 95), is_rtl=is_rtl)
    
    p_title = doc.add_paragraph()
    set_docx_rtl(p_title, is_rtl)
    p_title.paragraph_format.space_before = Pt(36)
    p_title.paragraph_format.space_after = Pt(24)
    run_title = p_title.add_run(convert_numbers(data.get("title", "ڕاپۆرتا زانستی"), is_rtl))
    format_run(run_title, size_pt=22, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
    
    p_meta = doc.add_paragraph()
    set_docx_rtl(p_meta, is_rtl)
    lbl_s = "ئامادەکرن ژ لایێ: " if is_rtl else "Prepared by: "
    lbl_t = "سەرپەرشتیا: " if is_rtl else "Supervised by: "
    
    r_s = p_meta.add_run(f"{lbl_s}{student or '-'}\n{lbl_t}{teacher or '-'}\n\n")
    format_run(r_s, size_pt=12, bold=False, color_rgb=(90, 100, 110), is_rtl=is_rtl)
    
    doc.add_page_break()
    
    # Abstract
    p_abs_h = doc.add_paragraph()
    set_docx_rtl(p_abs_h, is_rtl)
    r_abs_h = p_abs_h.add_run("پوختە (Abstract)" if is_rtl else "Abstract")
    format_run(r_abs_h, size_pt=16, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
    
    p_abs = doc.add_paragraph()
    set_docx_rtl(p_abs, is_rtl)
    p_abs.paragraph_format.line_spacing = 1.25
    r_abs = p_abs.add_run(convert_numbers(data.get("abstract", ""), is_rtl))
    format_run(r_abs, size_pt=11, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
    
    # پشکێن سەرەکی (Body Sections)
    for idx, sec in enumerate(data.get("sections", [])):
        p_sec_h = doc.add_paragraph()
        set_docx_rtl(p_sec_h, is_rtl)
        p_sec_h.paragraph_format.space_before = Pt(16)
        p_sec_h.paragraph_format.space_after = Pt(6)
        r_sec_h = p_sec_h.add_run(convert_numbers(f"{idx+1}. {sec.get('heading', '')}", is_rtl))
        format_run(r_sec_h, size_pt=14, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
        
        p_sec = doc.add_paragraph()
        set_docx_rtl(p_sec, is_rtl)
        p_sec.paragraph_format.line_spacing = 1.25
        p_sec.paragraph_format.space_after = Pt(10)
        r_sec = p_sec.add_run(convert_numbers(sec.get("content", ""), is_rtl))
        format_run(r_sec, size_pt=11.5, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
        
    # Conclusion
    p_con_h = doc.add_paragraph()
    set_docx_rtl(p_con_h, is_rtl)
    p_con_h.paragraph_format.space_before = Pt(16)
    r_con_h = p_con_h.add_run("دەرئەنجام (Conclusion)" if is_rtl else "Conclusion")
    format_run(r_con_h, size_pt=14, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
    
    p_con = doc.add_paragraph()
    set_docx_rtl(p_con, is_rtl)
    p_con.paragraph_format.line_spacing = 1.25
    r_con = p_con.add_run(convert_numbers(data.get("conclusion", ""), is_rtl))
    format_run(r_con, size_pt=11.5, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
    
    # References
    p_ref_h = doc.add_paragraph()
    set_docx_rtl(p_ref_h, is_rtl)
    p_ref_h.paragraph_format.space_before = Pt(18)
    r_ref_h = p_ref_h.add_run("سەرچاوەکان (References)" if is_rtl else "References")
    format_run(r_ref_h, size_pt=14, bold=True, color_rgb=(24, 43, 73), is_rtl=is_rtl)
    
    for ref in data.get("references", []):
        p_ref = doc.add_paragraph()
        set_docx_rtl(p_ref, is_rtl)
        p_ref.paragraph_format.space_after = Pt(4)
        r_ref = p_ref.add_run(f"• {convert_numbers(ref, is_rtl)}")
        format_run(r_ref, size_pt=10.5, bold=False, color_rgb=(70, 80, 95), is_rtl=is_rtl)
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def build_pptx(data, student, dept, teacher, is_rtl):
    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    DARK_BG = PptxRGBColor(15, 23, 42)      # شینێ تاری یێ مۆدێرن
    ACCENT_GOLD = PptxRGBColor(245, 158, 11) # ڕەنگێ زێڕی
    WHITE = PptxRGBColor(255, 255, 255)
    LIGHT_GRAY = PptxRGBColor(203, 213, 225)
    CARD_BG = PptxRGBColor(30, 41, 59)
    
    # سلایدا ئێکێ (Title Slide)
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = DARK_BG
    bg1.line.fill.background()
    
    box1 = s1.shapes.add_textbox(PptxInches(1), PptxInches(1.5), PptxInches(11.333), PptxInches(4.5))
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
    p_sub.text = convert_numbers(f"{dept or 'زانکۆ / پەیمانگەهـ'}", is_rtl)
    p_sub.font.size = PptxPt(20)
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
    
    # سلایدێن ناڤەرۆکێ
    for s_item in data.get("slides", []):
        sl = prs.slides.add_slide(blank_layout)
        
        # باکگراوند
        bg = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = DARK_BG
        bg.line.fill.background()
        
        # کارتا ناڤەڕاست
        card = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(1), PptxInches(1), PptxInches(11.333), PptxInches(5.5))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = ACCENT_GOLD
        card.line.width = PptxPt(1.5)
        
        # سەردێڕێ سلایدێ
        t_box = sl.shapes.add_textbox(PptxInches(1.3), PptxInches(1.2), PptxInches(10.7), PptxInches(1))
        t_frame = t_box.text_frame
        t_para = t_frame.paragraphs[0]
        t_para.text = convert_numbers(s_item.get("slide_title", ""), is_rtl)
        t_para.font.size = PptxPt(26)
        t_para.font.bold = True
        t_para.font.color.rgb = ACCENT_GOLD
        t_para.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
        if is_rtl: t_para._pPr.set('rtl', '1')
        
        # خاڵێن سمینارێ
        c_box = sl.shapes.add_textbox(PptxInches(1.3), PptxInches(2.3), PptxInches(10.7), PptxInches(3.8))
        c_frame = c_box.text_frame
        c_frame.word_wrap = True
        
        for idx, pt in enumerate(s_item.get("bullet_points", [])):
            para = c_frame.paragraphs[0] if idx == 0 else c_frame.add_paragraph()
            para.text = f"•  {convert_numbers(pt, is_rtl)}"
            para.font.size = PptxPt(18)
            para.font.color.rgb = WHITE
            para.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
            para.space_after = PptxPt(14)
            if is_rtl: para._pPr.set('rtl', '1')
            
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
        with st.spinner("داتایێن زانستی دهێنە دارشتن و فایل دهێنە دیزاینکرن... تکایە چەند چرکەیان بگرە"):
            try:
                content = generate_academic_content(topic, language, pages_count, student_name, department, teacher_name, api_key)
                
                # پاراستن د Session State دا دا کو دانلۆد نەفەوتیت
                st.session_state["docx_file"] = build_docx(content, student_name, department, teacher_name, is_rtl_lang).getvalue()
                st.session_state["pptx_file"] = build_pptx(content, student_name, department, teacher_name, is_rtl_lang).getvalue()
                st.session_state["topic_name"] = topic
                st.session_state["generated"] = True
                st.success("✅ ڕاپۆرت و سمینار ب سەرکەفتیانە و ب شێوازێ ستاندارد ئامادە بوون!")
            except Exception as e:
                st.error(f"کێشەیەک ڕویدا: {e}")

# بەشێ دانلۆدکرنێ (جودا و سەربەخۆ، بێ خەوش کار دکەت)
if st.session_state.get("generated", False):
    st.markdown("### 📥 فایلێن خو داونلۆد بکە:")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            label="📄 داگرتنا فایلا Word (.docx) ب ڕێکخستنا RTL",
            data=st.session_state["docx_file"],
            file_name=f"{st.session_state['topic_name']}_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    with col_d2:
        st.download_button(
            label="📊 داگرتنا فایلا PowerPoint (.pptx) دیزاینا مۆدێرن",
            data=st.session_state["pptx_file"],
            file_name=f"{st.session_state['topic_name']}_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
