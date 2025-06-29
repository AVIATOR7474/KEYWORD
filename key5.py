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
            st.download_button("🌐 Download HTML", html_content, "SEO_Optimized.html")
            st.code(html_content, language="html")
        except Exception as e:
            st.error(f"Error: {e}")

def docx_to_html(doc, primary_keywords, secondary_keywords):
    html_content = []
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # Process primary keywords first
        processed_text = text
        for pk in primary_keywords:
            if pk in processed_text:
                # Replace primary keyword with H1 tag
                processed_text = processed_text.replace(pk, f"<h1 dir='rtl'>{pk}</h1>")
        
        # Then process secondary keywords if no primary keywords were found
        if "<h1" not in processed_text:
            for sk in secondary_keywords:
                if sk in processed_text:
                    # Replace secondary keyword with H2 or H3 tag
                    word_count = len(sk.split())
                    if word_count < 3:  # Adjust this threshold as needed
                        processed_text = processed_text.replace(sk, f"<h2 dir='rtl'>{sk}</h2>")
                    else:
                        processed_text = processed_text.replace(sk, f"<h3 dir='rtl'>{sk}</h3>")
        
        # If no keywords were found, use paragraph
        if "<h1" not in processed_text and "<h2" not in processed_text and "<h3" not in processed_text:
            processed_text = f"<p dir='rtl'>{html.escape(processed_text)}</p>"
        else:
            # Escape the remaining text that's not in headings
            parts = re.split(r'(<h[123].*?</h[123]>)', processed_text)
            processed_text = ""
            for part in parts:
                if part.startswith("<h"):
                    processed_text += part
                else:
                    processed_text += html.escape(part)
        
        html_content.append(processed_text)
    
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>SEO Optimized Document</title>
</head>
<body>
{''.join(html_content)}
</body>
</html>"""

if __name__ == "__main__":
    main()