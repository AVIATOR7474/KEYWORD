import streamlit as st
import docx
from docx import Document
import io
import html
import re
from collections import defaultdict

st.set_page_config(page_title="SEO Word Document Optimizer", layout="wide")

def main():
    st.title("📝 SEO Word Document Optimizer")
    uploaded_file = st.file_uploader("📤 Upload .docx", type=["docx"])
    primary_keywords = st.text_area("Primary Keywords (one per line)", height=100)
    secondary_keywords = st.text_area("Secondary Keywords (one per line)", height=100)
    
    if uploaded_file and (primary_keywords.strip() or secondary_keywords.strip()):
        if st.button("🔄 Process Document"):
            process_document(uploaded_file, primary_keywords, secondary_keywords)

def process_document(uploaded_file, primary_keywords, secondary_keywords):
    with st.spinner("Processing..."):
        try:
            doc = Document(uploaded_file)
            primary_kw_list = [kw.strip() for kw in primary_keywords.split("\n") if kw.strip()]
            secondary_kw_list = [kw.strip() for kw in secondary_keywords.split("\n") if kw.strip()]
            
            # Process document
            optimized_doc = optimize_document(doc, primary_kw_list, secondary_kw_list)
            
            # Save DOCX
            buffer = io.BytesIO()
            optimized_doc.save(buffer)
            buffer.seek(0)
            
            # Generate HTML
            html_content = docx_to_html(optimized_doc)
            
            st.success("Document processed successfully!")
            st.download_button("📥 Download DOCX", buffer.getvalue(), "SEO_Optimized.docx")
            st.download_button("🌐 Download HTML", html_content, "SEO_Optimized.html")
            
            with st.expander("HTML Preview"):
                st.code(html_content, language="html")
                
        except Exception as e:
            st.error(f"Error: {e}")

def optimize_document(doc, primary_keywords, secondary_keywords):
    """Optimize the document by applying SEO headings only to specified keywords"""
    optimized_doc = Document()
    keyword_usage = defaultdict(int)
    
    # Define usage limits
    usage_limits = {
        'primary': {'h1': 1, 'h2': 3, 'h3': 2},
        'secondary': {'h2': 1, 'h3': 2}
    }
    
    for para in doc.paragraphs:
        original_text = para.text.strip()
        if not original_text:
            continue
        
        # Check for primary keywords first
        keyword_found = None
        heading_level = None
        
        for kw in primary_keywords:
            if kw.lower() in original_text.lower():
                # Check if we haven't exceeded usage limits
                if keyword_usage.get(f"{kw}_h1", 0) < usage_limits['primary']['h1']:
                    heading_level = 1
                elif keyword_usage.get(f"{kw}_h2", 0) < usage_limits['primary']['h2']:
                    heading_level = 2
                elif keyword_usage.get(f"{kw}_h3", 0) < usage_limits['primary']['h3']:
                    heading_level = 3
                
                if heading_level:
                    keyword_found = kw
                    keyword_usage[f"{kw}_h{heading_level}"] = keyword_usage.get(f"{kw}_h{heading_level}", 0) + 1
                    break
        
        # Check for secondary keywords if no primary keyword found
        if not keyword_found:
            for kw in secondary_keywords:
                if kw.lower() in original_text.lower():
                    # Check if we haven't exceeded usage limits
                    if keyword_usage.get(f"{kw}_h2", 0) < usage_limits['secondary']['h2']:
                        heading_level = 2
                    elif keyword_usage.get(f"{kw}_h3", 0) < usage_limits['secondary']['h3']:
                        heading_level = 3
                    
                    if heading_level:
                        keyword_found = kw
                        keyword_usage[f"{kw}_h{heading_level}"] = keyword_usage.get(f"{kw}_h{heading_level}", 0) + 1
                        break
        
        # Add to document with appropriate heading or as normal paragraph
        if keyword_found and heading_level:
            # Extract just the keyword (not the whole paragraph)
            optimized_doc.add_heading(keyword_found, level=heading_level)
            # Add the original paragraph as normal text
            optimized_doc.add_paragraph(original_text)
        else:
            optimized_doc.add_paragraph(original_text)
    
    return optimized_doc

def docx_to_html(doc):
    """Convert document to RTL HTML with proper headings"""
    html_content = []
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        if para.style.name.startswith('Heading'):
            level = int(para.style.name.split()[-1])
            html_content.append(f"<h{level} dir='rtl'>{html.escape(text)}</h{level}>")
        else:
            html_content.append(f"<p dir='rtl'>{html.escape(text)}</p>")
    
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
