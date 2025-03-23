import os
from docx import Document

def convert_zip_to_word_only(input_folder, selected_files, title, author, date, work_dir):
    doc = Document()
    doc.add_heading(title, 0)
    doc.add_paragraph(f"{author} – {date}")
    doc.add_paragraph("------------------------------")

    for md_file in selected_files:
        page_name = os.path.splitext(md_file)[0]
        md_path = os.path.join(input_folder, md_file)
        doc.add_heading(page_name.replace('-', ' '), level=1)
        with open(md_path, "r", encoding="utf-8") as f:
            doc.add_paragraph(f.read())

    docx_output = os.path.join(work_dir, "Notion_Document.docx")
    doc.save(docx_output)
    return docx_output