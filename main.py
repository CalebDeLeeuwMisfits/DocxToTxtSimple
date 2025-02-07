import os
from docx import Document

def convert_docx_to_txt(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.docx'):
            docx_path = os.path.join(input_folder, filename)
            doc = Document(docx_path)
            text = "\n".join([para.text for para in doc.paragraphs])
            txt_filename = f"{os.path.splitext(filename)[0]}.txt"
            output_path = os.path.join(output_folder, txt_filename)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text)
            print(f"Converted {filename} to {txt_filename}")

if __name__ == "__main__":
    input_folder = r"C:\Users\user\Documents" #change to path where pdfs are
    output_folder = os.path.join(input_folder, "Text")
    convert_docx_to_txt(input_folder, output_folder)
