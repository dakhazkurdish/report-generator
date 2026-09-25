import streamlit as st
import requests
import json
import io
import urllib.parse
import re
import time

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

st.set_page_config(
    page_title="سیستەمێ زیرەک یێ دروستکرنا راپورت و سمیناران", 
    page_icon="🎓", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

# ڤەشارتنا هەمی مێنیو و بارێن سێرڤەری
st.markdown("""
    <style>
    #MainMenu {visibility: hidden !important; display: none !important;}
    header {visibility: hidden !important; display: none !important;}
    footer {visibility: hidden !important; display: none !important;}
    
    .viewerBadge_container__1QSob,
    .styles_viewerBadge__1yB5G,
    [data-testid="stToolbar"],
    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"],
    [data-testid="stManageAppButton"],
    .manage-app-button,
    button[title="View app in GitHub"],
    a[href*="github.com"] {
        visibility: hidden !important; 
        display: none !important;
    }
    
    .stApp { direction: rtl; text-align: right; }
    p, h1, h2, h3, label, div { text-align: right !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; background-color: #1E3A8A; color: white; height: 3.2em; font-size: 16px; }
    textarea { direction: rtl !important; text-align: right !important; font-family: 'Segoe UI', Tahoma, sans-serif !important; }
    </style>
""", unsafe_allow_html=True)

st.title("سیستەمێ زیرەک یێ دروستکرنا راپورت و سمیناران")

raw_api_key = st.secrets.get("GEMINI_API_KEY", "")
if not raw_api_key:
    raw_api_key = st.text_input("کلیلا Gemini API لێرە بنڤیسە (دشێی چەند کلیلان ب کۆما جودا بکەی):", type="password")

with st.container():
    col1, col2 = st.columns(2)
    with col1:
        student_name = st.text_input("👤 ناڤێ قوتابی:")
        department = st.text_input("🏛️ پەیمانگەهـ یان کۆلێژ / پشک:")
        language = st.selectbox("🌐 زمانێ نڤیسینێ:", ["کوردی (بادینی)", "کوردی (سۆرانی)", "العربية", "English"])
        academic_level = st.selectbox(
            "🎓 ئاستێ ئەکادیمی یێ لێکۆڵینەوەیێ:",
            [
                "زانکۆ (بەکالۆریۆس - ستاندارد و پڕ زانیاری)",
                "پەیمانگەهـ (دبلۆم - ڕوون و کرداری)",
                "خاندنا باڵا (ماستەر و دکتۆرا - زۆر قووڵ، تیۆری و ڕەخنەگرانە)"
            ]
        )
    with col2:
        teacher_name = st.text_input("👨‍🏫 ناڤێ مامۆستایێ بابەتی:")
        topic = st.text_input("📝 بابەتێ سەرەکی یێ ڕاپۆرتێ:")
        pages_count = st.slider("📄 ژمارا لاپەڕێن پێدڤی بۆ ڕاپۆرتێ:", min_value=3, max_value=25, value=6)
        slides_count = st.slider("📊 ژمارا سلایدێن پاوەرپۆینتێ (Seminar):", min_value=4, max_value=15, value=5)

col_sub1, col_sub2 = st.columns(2)
with col_sub1:
    enable_border = st.checkbox("🖼️ چوارچێوە (Border) بۆ لاپەڕێن ڕاپۆرتا Word بهێتە دانان؟", value=True)
with col_sub2:
    uploaded_logo = st.file_uploader("🏛️ بارکرنا لۆگۆیێ فەرمی یێ زانکۆیێ (ئارەزوومەندانە):", type=["png", "jpg", "jpeg"])

custom_notes = st.text_area(
    "💡 تێبینی یان داخوازیێن تایبەت (بتنێ فەرمانە، ناچیتە ناڤ ڕاپۆرتێ):",
    placeholder="بۆ نموونە: گرنگیێ ب مێژوویا بابەتی بدە، نموونەیێن کرداری بینە، ئاستێ زانستی گەلەک بلند بیت...",
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

def fetch_academic_logo(dept_name):
    headers = {"User-Agent": "AcademicSlideGen/8.0"}
    if dept_name:
        try:
            clean_dept = urllib.parse.quote(dept_name.strip())
            search_url = f"https://en.wikipedia.org/w/api.php?action=query&format=json&generator=search&gsrsearch={clean_dept}%20logo&gsrlimit=1&prop=pageimages&pithumbsize=400"
            r = requests.get(search_url, headers=headers, timeout=3)
            if r.status_code == 200:
                pages = r.json().get("query", {}).get("pages", {})
                for _, p_info in pages.items():
                    thumb = p_info.get("thumbnail", {}).get("source")
                    if thumb:
                        img_r = requests.get(thumb, headers=headers, timeout=3)
                        if img_r.status_code == 200 and len(img_r.content) > 1500:
                            return io.BytesIO(img_r.content)
        except Exception:
            pass
    return None

def fetch_slide_image(keyword, topic_context=""):
    headers = {"User-Agent": "AcademicSlideGen/8.0 (educational-use)"}
    term = keyword or topic_context or "academic presentation"
    try:
        clean_kw = urllib.parse.quote(term.strip())
        commons_url = f"https://commons.wikimedia.org/w/api.php?action=query&generator=search&gsrsearch={clean_kw}&gsrnamespace=6&gsrlimit=3&prop=imageinfo&iiprop=url&iiurlwidth=800&format=json"
        r = requests.get(commons_url, headers=headers, timeout=3)
        if r.status_code == 200:
            pages = r.json().get("query", {}).get("pages", {})
            for _, page in pages.items():
                title = page.get("title", "").lower()
                if any(bad in title for bad in [".svg", ".ogg", ".pdf", "flag", "icon", "logo", "map"]):
                    continue
                img_info = page.get("imageinfo", [{}])[0]
                thumb_url = img_info.get("thumburl") or img_info.get("url")
                if thumb_url:
                    img_res = requests.get(thumb_url, headers=headers, timeout=3)
                    if img_res.status_code == 200 and len(img_res.content) > 4000:
                        return io.BytesIO(img_res.content)
    except Exception:
        pass
    return None

def parse_keys(raw_input):
    return [k.strip() for k in re.split(r'[,;\s]+', raw_input) if k.strip()]

def extract_clean_json(text):
    if not text:
        raise ValueError("دەقێ وەڵامێ یێ بەتاڵە.")
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end != -1 and end > start:
        return json.loads(text[start:end+1])
    return json.loads(text.strip())

@st.cache_data(ttl=3600)
def discover_active_models(api_key):
    """دۆزینەوەیا مۆدێلان ب پاراستن (Cache) داکو داخوازی زێدە نەچنە سەر کلیلێ"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
    try:
        res = requests.get(url, timeout=8)
        if res.status_code == 200:
            models_data = res.json().get("models", [])
            valid_models = []
            for m in models_data:
                methods = m.get("supportedGenerationMethods", [])
                if "generateContent" in methods:
                    clean_name = m.get("name", "").replace("models/", "")
                    valid_models.append(clean_name)
            
            flash_models = [m for m in valid_models if "flash" in m.lower() and not any(x in m.lower() for x in ["tts", "image", "live"])]
            pro_models = [m for m in valid_models if "pro" in m.lower() and not any(x in m.lower() for x in ["tts", "image", "live"])]
            return flash_models + pro_models + valid_models
    except Exception:
        pass
    return ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]

def call_gemini(prompt, keys_list, status_text=None):
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "response_mime_type": "application/json",
            "maxOutputTokens": 4096
        }
    }
    
    last_error = ""
    for key_idx, current_key in enumerate(keys_list):
        candidate_models = discover_active_models(current_key)
        
        for model_name in candidate_models:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={current_key}"
            
            # هەوڵدانا دووبارە (Retry) ب کێمەک بێهنڤەدان ئەگەر 429 چێبوو
            max_retries = 2
            for attempt in range(max_retries):
                try:
                    res = requests.post(url, headers=headers, json=payload, timeout=90)
                    
                    if res.status_code == 200:
                        data = res.json()
                        candidates = data.get("candidates", [])
                        if candidates and "content" in candidates[0]:
                            parts = candidates[0]["content"].get("parts", [])
                            if parts and "text" in parts[0]:
                                return extract_clean_json(parts[0]["text"])
                                
                    elif res.status_code in [429, 503]:
                        last_error = f"کلیلا ({key_idx + 1}) قەرەباڵغە (429 - Rate Limit)."
                        if attempt < max_retries - 1:
                            if status_text:
                                status_text.write(f"⏳ کلیلا ({key_idx + 1}) گەهشتە سنورێ خولەکی، {3 * (attempt + 1)} چرکە چاڤەڕێ بە...")
                            time.sleep(3 * (attempt + 1))
                            continue
                        else:
                            break  # دەرباز بە سەر کلیلا دویڤدا
                            
                    elif res.status_code == 404:
                        break  # مۆدێل بوونی نینە، بچە سەر مۆدێلێ دی
                    else:
                        try:
                            err_detail = res.json().get("error", {}).get("message", res.text)
                        except Exception:
                            err_detail = res.text
                        last_error = f"خەلەتیا API ({res.status_code}): {err_detail}"
                        break
                        
                except requests.exceptions.Timeout:
                    last_error = "دەمی وەڵامێ درێژ کێشا (Timeout)."
                    break
                except Exception as e:
                    last_error = str(e)
                    break
                    
    raise Exception(f"{last_error} - تکایە چەند خولەکان بێهنڤەدە یان کلیلەکا دی یا Gemini زێدە بکە.")

def generate_report(topic, lang, pages, student, dept, teacher, notes, academic_lvl, s_count, keys_list, progress_bar, status_text):
    num_sections = max(3, min(5, pages - 2))
    
    notes_prompt_part = ""
    if notes.strip():
        notes_prompt_part = f"""
        CRITICAL OPERATIONAL INSTRUCTIONS:
        "{notes.strip()}"
        Apply these instructions into the content. DO NOT print or quote them.
        """
        
    status_text.write("⚡ ژیرییا دەستکرد هەمی ڕاپۆرت و سمینارێ ئامادە دکەت...")
    progress_bar.progress(35)
    
    prompt = f"""
    Create a complete academic research report and presentation slides for the topic: "{topic}".
    Language: {lang}.
    Student: "{student}", Department: "{dept}", Supervisor: "{teacher}".
    Target Academic Level: {academic_lvl}.
    Required Sections count: {num_sections}.
    Required Presentation Slides count: {s_count}.
    {notes_prompt_part}

    CRITICAL RULES:
    1. Write everything inside strictly in {lang} (except English query for image).
    2. Write focused, informative scholarly paragraphs for EACH section inside "sections".
    3. Exactly {s_count} slides inside "slides".

    Return strictly a valid JSON object matching this schema:
    {{
        "title": "Full Academic Title in {lang}",
        "abstract": "Academic abstract in {lang} (120-180 words)",
        "english_main_topic": "2 simple english words for topic",
        "sections": [
            {{
                "heading": "Section Title in {lang}",
                "content": "Scholarly detailed text in paragraphs..."
            }}
        ],
        "conclusion": "Formal academic conclusion in {lang}",
        "references": [
            "APA Reference 1",
            "APA Reference 2",
            "APA Reference 3"
        ],
        "slides": [
            {{
                "slide_title": "Slide Title in {lang}",
                "image_search_query": "2 simple english words",
                "bullet_points": ["Point 1 in {lang}", "Point 2 in {lang}", "Point 3 in {lang}"]
            }}
        ]
    }}
    """
    
    result = call_gemini(prompt, keys_list, status_text)
    progress_bar.progress(85)
    status_text.write("فایلێن Word و PowerPoint دروست دبن...")
    
    return {
        "title": result.get("title", topic),
        "abstract": result.get("abstract", ""),
        "sections": result.get("sections", []),
        "conclusion": result.get("conclusion", ""),
        "references": result.get("references", []),
        "slides": result.get("slides", []),
        "main_en_topic": result.get("english_main_topic", topic)
    }

def build_docx(data, student, dept, teacher, is_rtl, with_border=True, user_logo_bytes=None, academic_lvl=""):
    doc = Document()
    section_cover = doc.sections[0]
    section_cover.top_margin = Inches(1)
    section_cover.bottom_margin = Inches(1)
    section_cover.left_margin = Inches(1)
    section_cover.right_margin = Inches(1)
    
    if with_border:
        add_page_borders(section_cover)

    logo_data = io.BytesIO(user_logo_bytes) if user_logo_bytes else fetch_academic_logo(dept)
    if logo_data:
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            doc.add_picture(logo_data, width=Inches(1.7))
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
    p_title.paragraph_format.space_before = Pt(32)
    p_title.paragraph_format.space_after = Pt(24)
    run_title = p_title.add_run(convert_numbers(data.get("title", "ڕاپۆرتا زانستی"), is_rtl))
    format_run(run_title, size_pt=24, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
    
    p_box = doc.add_paragraph()
    set_docx_rtl(p_box, is_rtl)
    p_box.paragraph_format.space_before = Pt(36)
    lbl_s = "ئامادەکرن ژ لایێ قوتابی: " if is_rtl else "Prepared by: "
    lbl_t = "سەرپەرشتیا مامۆستا: " if is_rtl else "Supervised by: "
    lbl_l = "ئاستێ ئەکادیمی: " if is_rtl else "Academic Level: "
    
    r_meta = p_box.add_run(
        f"📋 {lbl_s}{student or '-'}\n\n"
        f"👨‍🏫 {lbl_t}{teacher or '-'}\n\n"
        f"🎓 {lbl_l}{academic_lvl.split('(')[0].strip() or 'ئەکادیمی'}\n\n"
        f"📅 ساڵا ئەکادیمی: {convert_numbers('2025 - 2026', is_rtl)}"
    )
    format_run(r_meta, size_pt=14, bold=True, color_rgb=(51, 65, 85), is_rtl=is_rtl)
    
    section_body = doc.add_section()
    section_body.top_margin = Inches(1)
    section_body.bottom_margin = Inches(1)
    section_body.left_margin = Inches(1)
    section_body.right_margin = Inches(1)
    
    if with_border:
        add_page_borders(section_body)
    add_page_number_to_section(section_body, is_rtl)
    
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
    
    p_toc_h = doc.add_paragraph()
    set_docx_rtl(p_toc_h, is_rtl)
    r_toc_h = p_toc_h.add_run("پێڕستا ناڤەرۆکێ (Table of Contents)" if is_rtl else "Table of Contents")
    format_run(r_toc_h, size_pt=16, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
    
    toc_items = ["پوختە (Abstract)"]
    for idx, s in enumerate(data.get("sections", [])):
        toc_items.append(f"{idx+1}. {s.get('heading', '')}")
    toc_items.append("دەرئەنجام (Conclusion)")
    toc_items.append("سەرچاوەکان (References)")
    
    table = doc.add_table(rows=len(toc_items) + 1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].width = Inches(5.0)
    hdr_cells[1].width = Inches(1.5)
    
    p_h0 = hdr_cells[0].paragraphs[0]
    set_docx_rtl(p_h0, is_rtl)
    format_run(p_h0.add_run("بابەت / تەوەر" if is_rtl else "Topic / Section"), size_pt=13, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
    
    p_h1 = hdr_cells[1].paragraphs[0]
    set_docx_rtl(p_h1, is_rtl)
    format_run(p_h1.add_run("لاپەڕە" if is_rtl else "Page"), size_pt=13, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
    
    current_page_counter = 2
    for r_idx, item_title in enumerate(toc_items):
        row_cells = table.rows[r_idx + 1].cells
        row_cells[0].width = Inches(5.0)
        row_cells[1].width = Inches(1.5)
        
        p_c0 = row_cells[0].paragraphs[0]
        set_docx_rtl(p_c0, is_rtl)
        format_run(p_c0.add_run(convert_numbers(item_title, is_rtl)), size_pt=12, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
        
        p_c1 = row_cells[1].paragraphs[0]
        set_docx_rtl(p_c1, is_rtl)
        format_run(p_c1.add_run(convert_numbers(str(current_page_counter), is_rtl)), size_pt=12, bold=True, color_rgb=(37, 99, 235), is_rtl=is_rtl)
        
        current_page_counter += 1
        
    doc.add_page_break()
    
    for idx, sec in enumerate(data.get("sections", [])):
        p_sec_h = doc.add_paragraph()
        set_docx_rtl(p_sec_h, is_rtl)
        p_sec_h.paragraph_format.space_before = Pt(20)
        p_sec_h.paragraph_format.space_after = Pt(8)
        r_sec_h = p_sec_h.add_run(convert_numbers(f"{idx+1}. {sec.get('heading', '')}", is_rtl))
        format_run(r_sec_h, size_pt=16, bold=True, color_rgb=(15, 23, 42), is_rtl=is_rtl)
        
        paras = sec.get('content', '').split("\n\n")
        for p_t in paras:
            if not p_t.strip(): 
                continue
            p_sec = doc.add_paragraph()
            set_docx_rtl(p_sec, is_rtl)
            p_sec.paragraph_format.line_spacing = 1.3
            p_sec.paragraph_format.space_after = Pt(12)
            r_sec = p_sec.add_run(convert_numbers(p_t.strip(), is_rtl))
            format_run(r_sec, size_pt=14, bold=False, color_rgb=(30, 41, 59), is_rtl=is_rtl)
            
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
    if is_rtl: 
        p_t._pPr.set('rtl', '1')
    
    p_sub = tf1.add_paragraph()
    p_sub.text = convert_numbers(f"{dept or 'پەیمانگەهـ / زانکۆ'}", is_rtl)
    p_sub.font.size = PptxPt(20)
    p_sub.font.color.rgb = WHITE
    p_sub.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
    if is_rtl: 
        p_sub._pPr.set('rtl', '1')
    
    p_inf = tf1.add_paragraph()
    lbl_s = "قوتابی: " if is_rtl else "Student: "
    lbl_t = " | مامۆستا: " if is_rtl else " | Lecturer: "
    p_inf.text = convert_numbers(f"\n{lbl_s}{student or '-'} {lbl_t}{teacher or '-'}", is_rtl)
    p_inf.font.size = PptxPt(16)
    p_inf.font.color.rgb = LIGHT_GRAY
    p_inf.alignment = PP_ALIGN.RIGHT if is_rtl else PP_ALIGN.LEFT
    if is_rtl: 
        p_inf._pPr.set('rtl', '1')
    
    main_en_topic = data.get("main_en_topic", "")
    
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
        if is_rtl: 
            t_para._pPr.set('rtl', '1')
        
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
            if is_rtl: 
                para._pPr.set('rtl', '1')
            
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
    keys_list = parse_keys(raw_api_key)
    if not keys_list:
        st.error("تکایە کلیلا API بنڤیسە.")
    elif not topic:
        st.warning("تکایە بابەتێ ڕاپۆرتێ بنڤیسە.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        try:
            content = generate_report(
                topic, language, pages_count, student_name, department, 
                teacher_name, custom_notes, academic_level, slides_count, 
                keys_list, progress_bar, status_text
            )
            
            user_logo_data = uploaded_logo.read() if uploaded_logo else None
            
            st.session_state["docx_file"] = build_docx(
                content, student_name, department, teacher_name, 
                is_rtl_lang, enable_border, user_logo_data, academic_level
            ).getvalue()
            
            st.session_state["pptx_file"] = build_pptx(
                content, student_name, department, teacher_name, is_rtl_lang
            ).getvalue()
            
            st.session_state["plain_text"] = build_plain_text(
                content, student_name, department, teacher_name, is_rtl_lang
            )
            
            st.session_state["topic_name"] = topic
            st.session_state["generated"] = True
            
            progress_bar.empty()
            status_text.empty()
            st.success("✅ ڕاپۆرت و سمینار ب سەرکەفتوویی هاتنە ئامادەکرن!")
        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error(f"کێشەیەک ڕویدا: {e}")

if st.session_state.get("generated", False):
    st.markdown("### 📥 فایلێن خو داونلۆد بکە:")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            label="📄 داگرتنا فایلا Word (دگەل پێڕست و ڕێکخستنا ئەکادیمی)",
            data=st.session_state["docx_file"],
            file_name=f"{st.session_state['topic_name']}_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    with col_d2:
        st.download_button(
            label="📊 داگرتنا فایلا PowerPoint (سلایدێن کوردی و وێنەیێن تایبەت)",
            data=st.session_state["pptx_file"],
            file_name=f"{st.session_state['topic_name']}_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    
    st.markdown("---")
    st.markdown("### 📋 دەقێ ڕاپۆرتێ بۆ کۆپیکردنا ڕاستەوخۆ:")
    st.caption("دشێی ڤی دەقی دیاربکەی (Ctrl+A پاشان Ctrl+C) و پەیست بکەیە ناڤ وۆردێ خو بێی داگرتن:")
    st.text_area("", value=st.session_state["plain_text"], height=400)
```eof

### ئەگەرێن سەرەکی و چارەسەری:
1. **لۆدێ سەر کلیلێ (Rate Limits):** گووگڵ د پلانا بێبەرامبەر دا سنورەک ددانیت (مەسلەن ١٥ داخوازی د خولەکەکێ دا). ئەگەر تە ٢ کلیل هەبن و هەردوو زوو زوو هاتبنە بکارئینان، کۆدێ بەرێ ئێکسەر رادوەستیا.
2. **چارەسەری د کۆدێ نوی دا:**
   - فەنکشنا `discover_active_models` مە کریە `@st.cache_data` داکو هەروەخت داخوازا زێدە نەنێریت و کلیلێ ماندی نەکەت.
   - سیستەمێ **Retry & Wait** هاتییە زێدەکرن، ئانکو ئەگەر 429 رویدا، چەند چرکەکان چاڤەڕێ دبیت و دووبارە هەوڵ ددەت بەری دەستبەردار ببیت.
   - ئەگەر کێشە بەردەوام بوو، باشترین چارەسەر ئەوە ٣ بۆ ٤ کلیلێن جودا (ژ ئەکاونتێن جودا یێن گووگڵ) بینیت و ب کۆما (`,`) جودا بکەی.
