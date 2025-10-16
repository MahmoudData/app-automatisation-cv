import streamlit as st
from cv_process import read_cv, extract_info_from_cv, fill_word_template_with_lists
import os
from PIL import Image
import tempfile

st.set_page_config(page_title="Automatisation CV", page_icon="📄")

logo = Image.open("parlym_logo.png")

st.image(logo, width=300)

st.title("Traitement Automatique des CV")

# Sélection de la langue
langue = st.selectbox("Choisissez la langue de génération", ["fr", "en"], format_func=lambda x: "Français" if x == "fr" else "Anglais")

# Choix du template selon la langue
if langue == "en":
    template_path = "template_cv_p_en.docx"
else:
    template_path = "template_cv_p.docx"

# Sélection du fichier CV
uploaded_cv = st.file_uploader("Téléchargez le fichier CV (PDF ou Word)", type=["docx", "pdf"])

# Bouton pour lancer le traitement
if uploaded_cv is not None and template_path:
    st.write(f"**Fichier sélectionné :** {uploaded_cv.name}")

    # Bouton pour générer le fichier
    if st.button("Lancer le traitement"):
        output_path = None  # Initialiser output_path avant le bloc try
        
        with st.spinner("Traitement en cours..."):
            try:
                # Lire le contenu du fichier uploadé
                file_content = uploaded_cv.read()
                file_name = uploaded_cv.name
                
                # Extraire le texte du CV
                cv_content = read_cv(file_content=file_content, file_name=file_name)
                
                if cv_content and not cv_content.startswith("Type de fichier non pris en charge"):

                    extracted_info = extract_info_from_cv(cv_content, language=langue)

                    output_path = f"{uploaded_cv.name.split('.')[0]}_parlym.docx"

                    fill_word_template_with_lists(template_path, output_path, extracted_info, language=langue)

                    st.success(f"Fichier généré avec succès : {output_path}")

                    # bouton pour télécharger le fichier généré
                    with open(output_path, "rb") as result_file:
                        st.download_button(
                            label="Télécharger le fichier généré",
                            data=result_file,
                            file_name=output_path,
                        )
                else:
                    st.error(f"Erreur lors de la lecture du fichier: {cv_content}")
                    
            except Exception as e:
                st.error(f"Une erreur s'est produite : {str(e)}")
            finally:
                # Nettoyer le fichier de sortie si il existe
                if output_path and os.path.exists(output_path):
                    os.remove(output_path)
