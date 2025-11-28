import streamlit as st
from cv_process import extract_info_from_cv, fill_word_template_with_lists, extract_text_from_file, generate_dc_filename
import os
from PIL import Image
import tempfile
from pathlib import Path
from cv_process_inter import extract_info_from_cv_inter, generate_cv_inter_filename
from generator_inter import fill_cv_inter_template

st.set_page_config(page_title="Automatisation CV", page_icon="📄")

logo = Image.open("parlym_logo.png")

st.image(logo, width=300)

st.title("Traitement Automatique des CV")

# Sélecteur du type de CV
type_cv = st.selectbox("Type de DC", ["DC", "DC INTER"], format_func=lambda x: "DC" if x == "DC" else "DC INTER")

# Sélection de la langue
langue = st.selectbox("Choisissez la langue de génération", ["fr", "en"], format_func=lambda x: "Français" if x == "fr" else "Anglais")

# Choix du template selon le type de CV et la langue
if type_cv == "DC INTER":
    if langue == "en":
        template_path = "template_dc_inter_en.docx"
    else:
        template_path = "template_dc_inter.docx"
else:
    if langue == "en":
        template_path = "template_cv_p_en.docx"
    else:
        template_path = "template_cv_p.docx"

# Sélection du fichier CV
uploaded_cv = st.file_uploader("Téléchargez le fichier CV (PDF ou Word)", type=["docx", "pdf"])

def save_uploaded_file(uploaded_file) -> str:
    """
    Sauvegarde le fichier uploadé temporairement.
    
    Args:
        uploaded_file: Fichier uploadé par l'utilisateur
    
    Returns:
        Chemin vers le fichier sauvegardé
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded_file.name).suffix) as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        return tmp_file.name

# Bouton pour lancer le traitement
if uploaded_cv is not None and template_path:
    st.write(f"**Fichier sélectionné :** {uploaded_cv.name}")

    # Bouton pour générer le fichier
    if st.button("Lancer le traitement"):
        output_path = None  # Initialiser output_path avant le bloc try
        cv_temp_path = None  # Pour stocker le chemin temporaire
        
        with st.spinner("Traitement en cours..."):
            try:
                # Sauvegarder le fichier uploadé temporairement
                cv_temp_path = save_uploaded_file(uploaded_cv)
                
                # Extraire le texte du CV en utilisant extract_text_from_file
                cv_content = extract_text_from_file(cv_temp_path)
                
                if cv_content:
                    if type_cv == "DC INTER":
                        # Extraction IA et génération pour CV INTER
                        extracted_info = extract_info_from_cv_inter(cv_content, language=langue)
                        output_path = generate_cv_inter_filename(extracted_info)
                        fill_cv_inter_template(template_path, output_path, extracted_info, language=langue)
                    else:
                        # Logique DC standard
                        extracted_info = extract_info_from_cv(cv_content, language=langue)
                        output_path = generate_dc_filename(extracted_info)
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
                    st.error("Erreur : Impossible d'extraire le texte du fichier")
                    
            except ValueError as ve:
                st.error(f"Erreur de format : {str(ve)}")
            except Exception as e:
                st.error(f"Une erreur s'est produite : {str(e)}")
            finally:
                # Nettoyer le fichier temporaire uploadé
                if cv_temp_path and os.path.exists(cv_temp_path):
                    try:
                        os.unlink(cv_temp_path)
                    except OSError:
                        pass
                
                # Nettoyer le fichier de sortie si il existe
                if output_path and os.path.exists(output_path):
                    os.remove(output_path)