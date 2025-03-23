import streamlit as st
import zipfile
import os
import shutil
import uuid
import datetime
from logic.convert import convert_zip_to_word_only

st.set_page_config(page_title="Notion ➜ Word Converter", page_icon="📝")

st.image("static/logo.png", width=100)
st.title("Notion ➜ Word Converter")
st.markdown("""Glissez votre **export Notion (.zip)** ici, personnalisez la couverture,
choisissez les pages à inclure, et téléchargez un **fichier Word (.docx)** prêt à convertir en PDF.

💡 Astuce : ouvrez le fichier dans Word ou Google Docs pour exporter en PDF.""")

uploaded_file = st.file_uploader("Déposez votre fichier .zip", type="zip")

title = st.text_input("Titre du document", f"Export Notion – {datetime.date.today().isoformat()}")
author = st.text_input("Auteur", "Laurent Lefebvre")
custom_date = st.text_input("Date", datetime.date.today().strftime("%d/%m/%Y"))

if uploaded_file:
    with st.spinner("Analyse du fichier..."):
        session_id = str(uuid.uuid4())
        work_dir = f"temp/{session_id}"
        os.makedirs(work_dir, exist_ok=True)
        zip_path = os.path.join(work_dir, "notion.zip")

        with open(zip_path, "wb") as f:
            f.write(uploaded_file.read())

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(f"{work_dir}/extracted")

        md_files = [f for f in os.listdir(f"{work_dir}/extracted") if f.endswith(".md")]
        page_selection = st.multiselect("Sélectionnez les pages à inclure :", md_files, default=md_files)

        if st.button("📝 Générer le document Word"):
            docx_path = convert_zip_to_word_only(
                f"{work_dir}/extracted", page_selection, title, author, custom_date, work_dir
            )

            with open(docx_path, "rb") as f_docx:
                with open(docx_path, "rb") as f_docx:
                with open(docx_path, "rb") as f_docx:
                
with open(docx_path, "rb") as f_docx:
    st.download_button("📥 Télécharger le Word", f_docx, "Notion_Document.docx")

    # Aperçu HTML simple du document Word
    from docx import Document
    doc = Document(docx_path)
    st.subheader("Aperçu du contenu Word")
    for para in doc.paragraphs:
        st.markdown(f"<p>{para.text}</p>", unsafe_allow_html=True)


            shutil.rmtree(work_dir)
