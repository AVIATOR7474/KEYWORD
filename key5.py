import streamlit as st
import docx
from docx import Document
import io
import html
import re

st.set_page_config(page_title="SEO Word Document Optimizer", layout="wide")

def main():
    st.title("📝 SEO Word Document Optimizer")
    uploaded_file = st.file_uploader("📤 Upload .docx", type=["docx"])
    primary_keywords = st.text_area("Primary Keywords (one per line)", height=100)
    secondary_keywords = st.text_area("Secondary Keywords (one per line)", height=100)
    sensitivity = st.slider("Heading Sensitivity", 1, 10, 5) # هذا المتغير لم يعد له تأثير مباشر في هذا المنطق

    if uploaded_file and (primary_keywords.strip() or secondary_keywords.strip()):
        if st.button("🔄 Process Document"):
            process_document(uploaded_file, primary_keywords, secondary_keywords, sensitivity)

def process_document(uploaded_file, primary_keywords, secondary_keywords, sensitivity):
    with st.spinner("Processing..."):
        try:
            doc = Document(uploaded_file)
            primary_kw_list = [kw.strip() for kw in primary_keywords.split("\n") if kw.strip()]
            secondary_kw_list = [kw.strip() for kw in secondary_keywords.split("\n") if kw.strip()]
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            html_content = docx_to_html(doc, primary_kw_list, secondary_kw_list)
            
            st.download_button("📥 Download DOCX", buffer.getvalue(), "SEO_Optimized.docx")
            st.download_button("🌐 Download HTML", html_content.encode('utf-8'), "SEO_Optimized.html")
            st.code(html_content, language="html")
        except Exception as e:
            st.error(f"Error: {e}")

def docx_to_html(doc, primary_keywords, secondary_keywords):
    html_content = []
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        processed_text = html.escape(text) # ابدأ بتهريب النص بالكامل لتجنب مشاكل HTML
        
        # معالجة الكلمات البحثية الرئيسية
        for pk in primary_keywords:
            # استخدم re.sub للبحث عن الكلمة بالكامل وتغليفها
            # r'\b' يضمن مطابقة الكلمة بالكامل (حدود الكلمة)
            processed_text = re.sub(r'\b' + re.escape(pk) + r'\b', f'<span class="primary-keyword">{pk}</span>', processed_text, flags=re.IGNORECASE)
        
        # معالجة الكلمات البحثية الثانوية
        for sk in secondary_keywords:
            processed_text = re.sub(r'\b' + re.escape(sk) + r'\b', f'<span class="secondary-keyword">{sk}</span>', processed_text, flags=re.IGNORECASE)
        
        # دائمًا ضع النص داخل فقرة <p>
        html_content.append(f"<p dir='rtl'>{processed_text}</p>")
    
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>SEO Optimized Document</title>
    <style>
        /* أنماط اختيارية لتمييز الكلمات البحثية داخل النص */
        .primary-keyword {{
            font-weight: bold;
            color: #2E86AB; /* لون أزرق مميز */
            background-color: #E0F2F7; /* خلفية فاتحة */
            padding: 2px 4px;
            border-radius: 3px;
        }}
        .secondary-keyword {{
            font-weight: bold;
            color: #A23B72; /* لون بنفسجي مميز */
            background-color: #F7E0ED; /* خلفية فاتحة */
            padding: 2px 4px;
            border-radius: 3px;
        }}
        p {{
            margin-bottom: 1rem;
            line-height: 1.6;
        }}
    </style>
</head>
<body>
{''.join(html_content)}
</body>
</html>"""

if __name__ == "__main__":
    main()
