import streamlit as st
import requests
import json
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
import io

st.set_page_config(page_title="دروستکەرێ ڕاپۆرت و سمیناران", page_icon="🎓", layout="centered")

st.markdown("""
    <style>
    .stApp { direction: rtl; text-align: right; }
    p, h1, h2, h3, label { text-align: right !important; }
    </style>
""", unsafe_allow_html=True)

st.title("🎓 سیستەمێ زیرەکێ دروستکرنا ڕاپۆرت و سمیناران")
st.write("ناڤێ بابەتی بنڤیسە دا کو ب شێوەیەکێ ستاندارد و ئەکادیمی ڕاپۆرت و پاوەرپۆینت بۆ تە بهێنە دروستکرن.")

# کۆنترۆڵکرنا کلیلێ (ئەگەر د Secrets دا هەبیت یان ل دەستپێکێ لێبدەت)
api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    api_key = st.text_input("کلیلا Gemini API لێرە دابنێ:", type="password")

topic = st.text_input("بابەتێ ڕاپۆرتێ چییە؟ (بۆ نموونە: زیرەکییا دەستکرد د نوشداریدا)")
language = st.selectbox("زمانێ نڤیسینێ هەڵبژێرە:", ["کوردی (بادینی)", "کوردی (سۆرانی)", "English", "العربية"])
pages_count = st.slider("ژمارا لاپەڕێن پێدڤی بۆ ڕاپۆرتێ:", min_value=2, max_value=8, value=3)

def generate_academic_content(topic, lang, pages, key):
    # بتنێ مۆدێلێن نوو و کارا یێن فەرمی
    candidate_models = ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-3.5-flash"]
    headers = {"Content-Type": "application/json"}
    
    prompt = f"""
    You are an academic researcher. Generate a complete academic report and presentation slides on the topic: "{topic}".
    Language to write in: {lang}.
    Length: Approximately {pages} pages of content.

    Return the result STRICTLY as a valid JSON object without markdown fences, with these exact keys:
    {{
        "title": "Title of the research",
        "introduction": "Detailed academic introduction",
        "sections": [
            {{"heading": "Section Heading", "content": "Detailed academic body content"}}
        ],
        "conclusion": "Academic conclusion summary",
        "references": ["Ref 1 in APA format", "Ref 2 in APA format", "Ref 3 in APA format"],
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
            response = requests.post(url, headers=headers, json=payload, timeout=60)
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

def create_docx(data):
    doc = Document()
    doc.add_heading(data.get("title", "ڕاپۆرت"), level=0)
    
    doc.add_heading("پێشەکی (Introduction)", level=1)
    doc.add_paragraph(data.get("introduction", ""))
    
    for sec in data.get("sections", []):
        doc.add_heading(sec.get("heading", ""), level=1)
        doc.add_paragraph(sec.get("content", ""))
        
    doc.add_heading("دەرئەنجام (Conclusion)", level=1)
    doc.add_paragraph(data.get("conclusion", ""))
    
    doc.add_heading("سەرچاوەکان (References)", level=1)
    for ref in data.get("references", []):
        doc.add_paragraph(f"• {ref}")
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def create_pptx(data):
    prs = Presentation()
    
    # Title Slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = data.get("title", "پریزێنتەیشن")
    slide.placeholders[1].text = "ئامادەکرییە ژ لایێ سیستەمێ ئەکادیمی یێ زیرەک"
    
    # Content Slides
    bullet_slide_layout = prs.slide_layouts[1]
    for s_data in data.get("slides", []):
        slide = prs.slides.add_slide(bullet_slide_layout)
        slide.shapes.title.text = s_data.get("slide_title", "")
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame
        tf.clear()
        for pt in s_data.get("bullet_points", []):
            p = tf.add_paragraph()
            p.text = pt
            p.level = 0
            
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

if st.button("🚀 دروستکرنا ڕاپۆرت و سمینارێ", type="primary"):
    if not api_key:
        st.error("تکایە دەستپێکێ کلیلا API بنڤیسە یان د بەشێ سێرڤەری دا دابنێ.")
    elif not topic:
        st.warning("تکایە ناڤێ بابەتێ ڕاپۆرتێ بنڤیسە.")
    else:
        with st.spinner("داتایێن زانستی دهێنە کۆمکرن و فایل بەرهەڤ دبن... چەند چرکەیان بگرە"):
            try:
                data = generate_academic_content(topic, language, pages_count, api_key)
                
                docx_file = create_docx(data)
                pptx_file = create_pptx(data)
                
                st.success("✅ ب سەرکەفتیانە ڕاپۆرت و سمینار هاتنە دروستکرن!")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📄 داگرتنا فایلا Word (.docx)",
                        data=docx_file,
                        file_name=f"{topic}_report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                with col2:
                    st.download_button(
                        label="📊 داگرتنا فایلا PowerPoint (.pptx)",
                        data=pptx_file,
                        file_name=f"{topic}_presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            except Exception as e:
                st.error(f"کێشەیەک ڕویدا: {e}")
