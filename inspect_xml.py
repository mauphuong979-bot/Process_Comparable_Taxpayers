from docx import Document
from docx.oxml.ns import qn

def inspect_format(file_path):
    doc = Document(file_path)
    para = None
    for p in doc.paragraphs:
        if "Bulk Rejections Summary" in p.text:
            para = p
            break
            
    if para:
        print(f"Para text: {para.text}")
        print(f"Style: {para.style.name}")
        # Show pPr XML
        pPr = para._element.get_or_add_pPr()
        print(f"XML pPr: {para._element.xml}")
        
        # Check for outline level in XML
        outline_lvl = pPr.xpath('./w:outlineLvl')
        if outline_lvl:
            print(f"Outline Level found in XML: {outline_lvl[0].get(qn('w:val'))}")
        else:
            print("No Outline Level found in XML")

if __name__ == "__main__":
    inspect_format(r"d:\AI\Python\GitHub\Fix Comparable Taxpayers\Demo_2025 avg Report.docx")
