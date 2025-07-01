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
        
        # ابدأ بتهريب النص بالكامل لتجنب مشاكل HTML
        processed_text = html.escape(text) 
        
        # معالجة الكلمات البحثية الرئيسية
        # يجب أن تكون الكلمات البحثية الرئيسية ذات أولوية أعلى
        for pk in primary_keywords:
            # استخدم re.sub للبحث عن الكلمة بالكامل وتغليفها بـ <h2>
            # r'\b' يضمن مطابقة الكلمة بالكامل (حدود الكلمة)
            # flags=re.IGNORECASE يجعل المطابقة غير حساسة لحالة الأحرف
            processed_text = re.sub(r'\b' + re.escape(pk) + r'\b', f'<h2 dir="rtl" style="display:inline;">{pk}</h2>', processed_text, flags=re.IGNORECASE)
        
        # معالجة الكلمات البحثية الثانوية
        # تأكد من أن الكلمات الثانوية لا تتداخل مع الكلمات الرئيسية التي تم تحويلها بالفعل
        for sk in secondary_keywords:
            # تأكد من عدم استبدال الكلمات التي أصبحت جزءًا من <h2> بالفعل
            # هذا النمط يضمن أننا لا نطابق داخل علامات HTML الموجودة
            processed_text = re.sub(r'(?<!<h[23][^>]*?>)\b' + re.escape(sk) + r'\b(?!</h[23]>)', f'<h3 dir="rtl" style="display:inline;">{sk}</h3>', processed_text, flags=re.IGNORECASE)
        
        # دائمًا ضع النص الناتج داخل فقرة <p>
        html_content.append(f"<p dir='rtl'>{processed_text}</p>")
    
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>SEO Optimized Document</title>
    <style>
        /* أنماط اختيارية لجعل العناوين تظهر في نفس السطر */
        h2 {{
            display: inline; /* يجعل h2 يظهر في نفس السطر */
            font-size: 1.5rem; /* حجم الخط */
            margin: 0; /* إزالة الهوامش الافتراضية */
            padding: 0; /* إزالة الحشوة الافتراضية */
            color: #2E86AB; /* لون اختياري */
        }}
        h3 {{
            display: inline; /* يجعل h3 يظهر في نفس السطر */
            font-size: 1.2rem; /* حجم الخط */
            margin: 0; /* إزالة الهوامش الافتراضية */
            padding: 0; /* إزالة الحشوة الافتراضية */
            color: #A23B72; /* لون اختياري */
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
