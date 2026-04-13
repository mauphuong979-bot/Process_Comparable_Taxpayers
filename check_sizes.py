from docx import Document
from docx.shared import Pt

def check_style_sizes(file_path):
    doc = Document(file_path)
    found_styles = {}
    
    for para in doc.paragraphs:
        style_name = para.style.name
        if style_name not in found_styles and "Section_Level" in style_name:
            # Check first run font size
            size = None
            if para.runs:
                for run in para.runs:
                    if run.font.size:
                        size = run.font.size.pt
                        break
            # If not in run, it might be in style. We'll just report what we found in runs
            found_styles[style_name] = size
            print(f"Style: {style_name} | Sample Text: {para.text[:30]!r} | Run Font Size: {size}")

if __name__ == "__main__":
    check_style_sizes(r"d:\AI\Python\GitHub\Fix Comparable Taxpayers\Demo_2025 avg Report.docx")
