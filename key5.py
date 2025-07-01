import streamlit as st
import docx
from docx import Document
import io
import html
import re
import random 

st.set_page_config(page_title="SEO Word Document Optimizer", layout="wide")

def main():
    st.title("📝 SEO Word Document Optimizer")
    uploaded_file = st.file_uploader("📤 Upload .docx", type=["docx"])
    primary_keywords = st.text_area("Primary Keywords (one per line)", height=100)
    secondary_keywords = st.text_area("Secondary Keywords (one per line)", height=100)
    sensitivity = st.slider("Heading Sensitivity", 1, 10, 5) 

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
    
    # عدادات لتتبع عدد مرات ظهور كل كلمة بحثية كعنوان
    primary_kw_counts = {kw: {'h2': 0, 'h3': 0} for kw in primary_keywords}
    secondary_kw_counts = {kw: {'h2': 0, 'h3': 0} for kw in secondary_keywords}

    MIN_HEADINGS = 3
    MAX_HEADINGS = 6

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # ابدأ بتهريب النص بالكامل
        processed_text = html.escape(text) 
        
        # قائمة لتتبع الكلمات التي تم تحويلها في هذه الفقرة لتجنب التكرار المباشر
        converted_in_this_paragraph = set()

        # معالجة الكلمات البحثية الرئيسية أولاً
        # الأولوية: H2 ثم H3
        for pk in primary_keywords:
            if pk in converted_in_this_paragraph: 
                continue

            # حاول تحويلها إلى H2
            if primary_kw_counts[pk]['h2'] < MAX_HEADINGS:
                # استخدم تعبير نمطي للبحث عن الكلمة ككلمة كاملة
                # وتأكد أنها ليست داخل علامة HTML موجودة بالفعل
                # هذا التعبير النمطي يبحث عن الكلمة فقط إذا لم تكن محاطة بـ <...>
                # ولكن الأسهل هو الاعتماد على ترتيب المعالجة
                
                # سنقوم بالاستبدال مباشرة، وبما أننا نستخدم html.escape() في البداية،
                # فإن الكلمات لن تكون داخل علامات HTML بعد.
                # بعد أول استبدال، ستصبح الكلمة داخل <h2...> أو <h3...>
                # ولن يتم مطابقتها مرة أخرى بواسطة re.sub للكلمات الأخرى.
                
                # البحث عن الكلمة فقط إذا لم تكن جزءًا من علامة HTML
                # هذا النمط يطابق الكلمة إذا لم تكن مسبوقة بـ < أو متبوعة بـ >
                # ولكن هذا قد يكون معقدًا. الأبسط هو الاعتماد على ترتيب المعالجة.

                # الطريقة الأبسط: استبدل الكلمة إذا وجدت.
                # بما أننا نستخدم html.escape() في البداية، فالنص "نظيف".
                # بعد الاستبدال الأول، ستصبح الكلمة جزءًا من HTML ولن يتم مطابقتها مرة أخرى.
                
                # البحث عن الكلمة ككلمة كاملة
                pattern = r'\b' + re.escape(pk) + r'\b'
                if re.search(pattern, processed_text, flags=re.IGNORECASE):
                    processed_text = re.sub(
                        pattern, 
                        f'<h2 dir="rtl" style="display:inline;">{pk}</h2>', 
                        processed_text, 
                        count=1, 
                        flags=re.IGNORECASE
                    )
                    primary_kw_counts[pk]['h2'] += 1
                    converted_in_this_paragraph.add(pk)
                    continue 
            
            # حاول تحويلها إلى H3 إذا لم يتم تحويلها إلى H2
            if primary_kw_counts[pk]['h3'] < MAX_HEADINGS:
                pattern = r'\b' + re.escape(pk) + r'\b'
                if re.search(pattern, processed_text, flags=re.IGNORECASE):
                    processed_text = re.sub(
                        pattern, 
                        f'<h3 dir="rtl" style="display:inline;">{pk}</h3>', 
                        processed_text, 
                        count=1, 
                        flags=re.IGNORECASE
                    )
                    primary_kw_counts[pk]['h3'] += 1
                    converted_in_this_paragraph.add(pk)

        # معالجة الكلمات البحثية الثانوية
        # الأولوية: H3 ثم H2
        for sk in secondary_keywords:
            if sk in converted_in_this_paragraph:
                continue

            # حاول تحويلها إلى H3
            if secondary_kw_counts[sk]['h3'] < MAX_HEADINGS:
                pattern = r'\b' + re.escape(sk) + r'\b'
                if re.search(pattern, processed_text, flags=re.IGNORECASE):
                    processed_text = re.sub(
                        pattern, 
                        f'<h3 dir="rtl" style="display:inline;">{sk}</h3>', 
                        processed_text, 
                        count=1, 
                        flags=re.IGNORECASE
                    )
                    secondary_kw_counts[sk]['h3'] += 1
                    converted_in_this_paragraph.add(sk)
                    continue

            # حاول تحويلها إلى H2 إذا لم يتم تحويلها إلى H3
            if secondary_kw_counts[sk]['h2'] < MAX_HEADINGS:
                pattern = r'\b' + re.escape(sk) + r'\b'
                if re.search(pattern, processed_text, flags=re.IGNORECASE):
                    processed_text = re.sub(
                        pattern, 
                        f'<h2 dir="rtl" style="display:inline;">{sk}</h2>', 
                        processed_text, 
                        count=1, 
                        flags=re.IGNORECASE
                    )
                    secondary_kw_counts[sk]['h2'] += 1
                    converted_in_this_paragraph.add(sk)

        html_content.append(f"<p dir='rtl'>{processed_text}</p>")
    
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>SEO Optimized Document</title>
    <style>
        h2 {{
            display: inline;
            font-size: 1.5rem;
            margin: 0;
            padding: 0;
            color: #2E86AB; 
            font-weight: bold;
        }}
        h3 {{
            display: inline;
            font-size: 1.2rem;
            margin: 0;
            padding: 0;
            color: #A23B72; 
            font-weight: bold;
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
