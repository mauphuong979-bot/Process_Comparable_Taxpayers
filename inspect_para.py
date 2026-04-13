from docx import Document

def inspect_specific_para(file_path, target_text):
    doc = Document(file_path)
    for i, para in enumerate(doc.paragraphs):
        if target_text in para.text:
            print(f"--- Paragraph {i} ---")
            print(f"Text: {para.text!r}")
            print(f"Paragraph Style Name: {para.style.name}")
            print(f"Paragraph Style Type: {para.style.type}")
            
            for j, run in enumerate(para.runs):
                print(f"  Run {j}:")
                print(f"    Text: {run.text!r}")
                if run.style:
                    print(f"    Character Style: {run.style.name}")
                print(f"    Font Bold: {run.font.bold}")

if __name__ == "__main__":
    inspect_specific_para(r"d:\AI\Python\GitHub\Fix Comparable Taxpayers\Demo_2025 avg Report.docx", "Bulk Rejections Summary")
