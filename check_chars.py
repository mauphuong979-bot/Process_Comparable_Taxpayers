from docx import Document

def inspect_chars(file_path):
    doc = Document(file_path)
    for para in doc.paragraphs:
        if "Arm" in para.text and "Length" in para.text:
            print(f"Text: {para.text!r}")
            print(f"Char codes: {[ord(c) for c in para.text]}")
            break

if __name__ == "__main__":
    inspect_chars(r"d:\AI\Python\GitHub\Fix Comparable Taxpayers\Demo_2025 avg Report.docx")
