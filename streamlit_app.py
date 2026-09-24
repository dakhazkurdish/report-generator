import streamlit as st
import json
import os
import requests
import io
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt

# ==========================================
# دیزاینا لاپەڕێ ماڵپەری
# ==========================================
st.set_page_config(
    page_title="دروستکەرێ ڕاپۆرت و سمیناران",
    page_icon="🎓",
    layout="centered"
)

# ڕێکخستنا شێوازێ نڤیسینێ و ئاراستەیێ کوردی (RTL)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;600;700;800&display=swap');
    html, body, [class*="css"] {
        font-family: 'Vazirmatn', sans-serif;
        direction: rtl;
        text-align: right;
    }
    .stTextInput > div > div > input, .stSelectbox > div > div > div {
        text-align: right;
        direction: rtl;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🎓 پلاتفۆڕمێ دروستکرنا ڕاپۆرت و سمیناران")
st.write("ناڤێ بابەتێ خو بنڤیسە دا کو د خولەکەکێ دا ڕاپۆرتا Word و سمینارا PowerPoint ب شێوەیەکێ ئەکادیمی بو تە بەرهەڤ ببیت.")

# ==========================================
# پەیوەندی ب Gemini API
# کلیلا تە د ناڤ سێرڤەری دا دپارێزیت (Secrets)
# ==========================================
def get_api_key():
    # دەمێ ل سەر Streamlit Cloud بەلاڤ دکەی، ل بەشێ Secrets دادمەزرێنی
    if "GEMINI_API_KEY" in st.secrets:
        return st.secrets["GEMINI_API_KEY"]
    return os.environ.get("GEMINI_API_KEY", "")

def call_gemini(api_key, prompt):
   url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.3}
    }
    res = requests.post(url, headers=headers, json=payload, timeout=60)
    if res.status_code == 200:
        return res.json()['candidates'][0]['content']['parts'][0]['text']
    else:
        raise Exception(f"API Error: {res.text}")

# ==========================================
# دروستکرنا Word د ناڤ بیرگەهێ دا (Memory Buffer)
# ==========================================
def generate_docx(data):
    doc = Document()
    
    # لاپەڕێ ناڤونیشانی
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = p_title.add_run(data.get("title", "Academic Report"))
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(26)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(31, 78, 121)
    
    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if data.get("student_name"):
        p_info.add_run(f"ئامادەکرن: {data['student_name']}\n").font.size = Pt(14)
    if data.get("department"):
        p_info.add_run(f"پشک: {data['department']}\n").font.size = Pt(13)
    if data.get("supervisor"):
        p_info.add_run(f"سەرپەرشتیار: {data['supervisor']}\n").font.size = Pt(13)
        
    doc.add_page_break()
    
    # پێشەکی
    h1 = doc.add_heading("1. پێشەکی (Introduction)", level=1)
    h1.runs[0].font.color.rgb = RGBColor(31, 78, 121)
    p_intro = doc.add_paragraph(data.get("introduction", ""))
    p_intro.paragraph_format.line_spacing = 1.25
    p_intro.paragraph_format.space_after = Pt(12)
    
    # تەوەرێن سەرەکی
    sections = data.get("sections", [])
    for idx, sec in enumerate(sections, start=2):
        h = doc.add_heading(f"{idx}. {sec.get('heading', '')}", level=1)
        h.runs[0].font.color.rgb = RGBColor(31, 78, 121)
        p = doc.add_paragraph(sec.get("content", ""))
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.space_after = Pt(12)
        
    # دەرئەنجام
    h_conc = doc.add_heading(f"{len(sections) + 2}. دەرئەنجام (Conclusion)", level=1)
    h_conc.runs[0].font.color.rgb = RGBColor(31, 78, 121)
    doc.add_paragraph(data.get("conclusion", ""))
    
    # سەرچاوە
    h_ref = doc.add_heading(f"{len(sections) + 3}. سەرچاوە (References - APA)", level=1)
    h_ref.runs[0].font.color.rgb = RGBColor(31, 78, 121)
    for r in data.get("references", []):
        doc.add_paragraph(r, style='List Bullet')
        
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# دروستکرنا PowerPoint د ناڤ بیرگەهێ دا
# ==========================================
def generate_pptx(data):
    prs = Presentation()
    prs.slide_width = PptInches(13.33)
    prs.slide_height = PptInches(7.5)
    
    # Title Slide
    s0 = prs.slides.add_slide(prs.slide_layouts[0])
    s0.shapes.title.text = data.get("title", "Presentation")
    s0.placeholders[1].text = f"ئامادەکرن: {data.get('student_name', '')}\n{data.get('department', '')}"
    
    # Content Slides
    for item in data.get("slides", []):
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = item.get("slide_title", "")
        tf = s.placeholders[1].text_frame
        pts = item.get("bullet_points", [])
        if pts:
            tf.text = pts[0]
            for p in pts[1:]:
                p_elem = tf.add_paragraph()
                p_elem.text = p
                
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# فۆڕما زانیاریێن قوتابی
# ==========================================
with st.form("academic_form"):
    topic = st.text_input("بابەتێ سەرەکی یێ ڕاپۆرتێ (Topic) *", placeholder="بۆ نموونە: Artificial Intelligence in Healthcare")
    
    col1, col2 = st.columns(2)
    with col1:
        student_name = st.text_input("ناڤێ قوتابی", placeholder="ناڤێ خو بنڤیسە")
    with col2:
        dept = st.text_input("پشک / کولێژ", placeholder="بۆ نموونە: پشکا زانستێن کۆمپیۆتەری")
        
    supervisor = st.text_input("ناڤێ مامۆستا / سەرپەرشتیار (ئارەزوومەندانە)")
    lang = st.selectbox("زمانێ ناڤەرۆکا ڕاپۆرتێ", ["English", "Kurdish", "Arabic"])
    
    submit_btn = st.form_submit_button("🚀 دروستکرنا فایلان")

# ==========================================
# کردارا دروستکرن و بەخشینا فایلان
# ==========================================
if submit_btn:
    api_key = get_api_key()
    if not api_key:
        st.error("کلیلێ API نەهاتیە دیتن! پێدڤییە د بەشێ Streamlit Secrets دا بهێتە زێدەکرن.")
    elif not topic.strip():
        st.warning("تکایە ناڤێ بابەتێ بنڤیسە.")
    else:
        with st.spinner("ژیریا دەستکرد خەریکە سەرچاوە و دەقێ ئەکادیمی ئامادە دکەت..."):
            prompt = f"""
            You are an expert academic research assistant. Create a comprehensive, formal, and citation-backed report and presentation content for the topic: "{topic}".
            Target Language: {lang}.
            Student details: Name: {student_name}, Department: {dept}, Supervisor: {supervisor}.

            Return ONLY a raw valid JSON object with NO markdown formatting, NO ```json backticks:
            {{
                "title": "Concise Academic Title",
                "student_name": "{student_name}",
                "department": "{dept}",
                "supervisor": "{supervisor}",
                "introduction": "Comprehensive academic introduction explaining context, importance, and objectives (at least 200 words).",
                "sections": [
                    {{"heading": "Title of Chapter/Section 1", "content": "In-depth academic discussion with technical details and analysis (at least 250 words)."}},
                    {{"heading": "Title of Chapter/Section 2", "content": "Further in-depth analysis, methodology, or comparison (at least 250 words)."}},
                    {{"heading": "Title of Chapter/Section 3", "content": "Challenges, future outlook, and implementations (at least 200 words)."}}
                ],
                "conclusion": "Thorough summary of findings and final academic takeaways (at least 150 words).",
                "references": [
                    "Author, A. (Year). Title of paper. Journal Name, Vol(Issue), pages.",
                    "Author, B. (Year). Title of book or paper. Publisher/Conference."
                ],
                "slides": [
                    {{"slide_title": "Overview & Objectives", "bullet_points": ["Key purpose of research", "Main research question", "Scope of study"]}},
                    {{"slide_title": "Key Theoretical Concepts", "bullet_points": ["Core principle definition", "Critical mechanisms", "Industry standards"]}},
                    {{"slide_title": "Findings & Analysis", "bullet_points": ["Primary quantitative/qualitative result", "Comparative performance", "Practical benefits"]}},
                    {{"slide_title": "Future Scope & Recommendations", "bullet_points": ["Scalability and adoption", "Emerging challenges", "Strategic takeaway"]}}
                ]
            }}
            """
            try:
                raw_text = call_gemini(api_key, prompt)
                clean_text = raw_text.replace("```json", "").replace("```", "").strip()
                data = json.loads(clean_text)
                
                # چێکرنا فایلان د مێمۆری دا
                docx_file = generate_docx(data)
                pptx_file = generate_pptx(data)
                
                st.success("✅ پیرۆزە! فایلێن تە ب سەرکەفتی ئامادە بوون.")
                
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.download_button(
                        label="📄 داگرتنا ڕاپۆرتا Word (.docx)",
                        data=docx_file,
                        file_name=f"{topic[:20]}_Report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                with col_d2:
                    st.download_button(
                        label="📊 داگرتنا سمینارا PowerPoint (.pptx)",
                        data=pptx_file,
                        file_name=f"{topic[:20]}_Seminar.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            except Exception as e:
                st.error(f"کێشەیەک ڕویدا: {e}")
