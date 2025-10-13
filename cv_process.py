
from docx import Document
import openai
import json
import os
from PyPDF2 import PdfReader
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT  # Pour aligner les paragraphes
from docx.shared import Inches  # Pour définir les positions en pouces
import docx2txt
import streamlit as st
import tempfile
import re
from datetime import datetime


openai.api_key = st.secrets["OPENAI_API_KEY"]


def extract_text_from_docx(file_content=None, file_path=None):
    """Extrait le texte d'un fichier DOCX."""
    try:
        if file_content:
            # Si on a le contenu du fichier (pour uploaded files)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                tmp_file.write(file_content)
                tmp_file.flush()
                
                # Méthode 1: Utiliser docx2txt
                try:
                    text = docx2txt.process(tmp_file.name)
                    os.unlink(tmp_file.name)
                    return text
                except Exception:
                    # Fallback avec python-docx
                    doc = Document(tmp_file.name)
                    text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
                    os.unlink(tmp_file.name)
                    return text
        else:
            # Si on a un chemin de fichier
            try:
                text = docx2txt.process(file_path)
                return text
            except Exception:
                # Fallback avec python-docx
                doc = Document(file_path)
                text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
                return text
    except Exception as e:
        st.error(f"Erreur lors de l'extraction du texte DOCX: {str(e)}")
        return None


def convert_word_to_pdf(docx_path):
    """
    OBSOLÈTE: Cette fonction n'est plus utilisée sur Streamlit Cloud.
    La conversion directe DOCX->PDF n'est pas supportée sans pywin32.
    """
    st.warning("⚠️ Conversion PDF non disponible sur cette plateforme. Le fichier DOCX sera traité directement.")
    return None


def extract_text_from_pdf(pdf_path):
    """Extrait le texte d'un fichier PDF."""
    reader = PdfReader(pdf_path)
    cv_text = "".join(page.extract_text() or "" for page in reader.pages)
    return cv_text.strip()


def read_cv(file_path=None, file_content=None, file_name=None):
    """
    Lit un CV en format .docx ou .pdf.
    
    Arguments:
        file_path (str): Chemin vers le fichier (pour fichiers locaux)
        file_content (bytes): Contenu du fichier (pour uploaded files)
        file_name (str): Nom du fichier pour déterminer l'extension
    """
    if file_path:
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        if ext == ".docx":
            return extract_text_from_docx(file_path=file_path)
        elif ext == ".pdf":
            return extract_text_from_pdf(file_path)
        else:
            return "Type de fichier non pris en charge. Veuillez fournir un fichier .docx ou .pdf."
    
    elif file_content and file_name:
        ext = os.path.splitext(file_name.lower())[1]
        
        if ext == ".docx":
            return extract_text_from_docx(file_content=file_content)
        elif ext == ".pdf":
            # Pour les PDF uploadés
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                tmp_file.write(file_content)
                tmp_file.flush()
                tmp_file_path = tmp_file.name

            # Ouvre le fichier PDF et extrait le texte
            with open(tmp_file_path, "rb") as f:
                text = extract_text_from_pdf(f)

            os.unlink(tmp_file_path)
            return text
        else:
            return "Type de fichier non pris en charge. Veuillez fournir un fichier .docx ou .pdf."
    
    else:
        return "Paramètres insuffisants pour lire le fichier."
    
def generate_trigramme(prenom, nom):
    """
    Génère le trigramme : première lettre du prénom + deux premières consonnes du nom.
    Exemple : Cédric GOBERT -> CGB
    """
    prenom = prenom.strip().upper() if prenom else ""
    nom = nom.strip().upper() if nom else ""
    first_letter = prenom[0] if prenom else ""
    consonnes = re.sub(r'[AEIOUY]', '', nom)
    trigramme = first_letter + consonnes[:2]
    return trigramme
    

def extract_info_from_cv(cv_text):
    """
    Extrait des informations structurées à partir d'un texte de CV en utilisant l'API OpenAI.
    
    Arguments :
        cv_text (str) : Contenu textuel du CV.

    Retourne :
        dict : Un dictionnaire JSON contenant les informations extraites.
    """

    # Définition du schéma pour Function Calling (sans TRI)
    function_schema = {
        "name": "extract_cv_info",
        "parameters": {
            "type": "object",
            "properties": {
                "PRENOM": {"type": "string", "description": "prénom"},
                "NOM": {"type": "string", "description": "nom"},
                "INTITULE_DU_POSTE": {"type": "string", "description": "L'intitulé du poste recherché."},
                "EXPERTISE": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Les activités et compétences spécifiques (par exemple, Etude de constructibilité, Résolution des problématiques, Leadership)."
                },
                "SECTEUR": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Les domaines principaux d'expertise (par exemple, Bâtiment, Industrie, Oil & Gas)."
                },
                "METHODOLOGIE": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Les méthodologies et outils maîtrisés (par exemple, Pack office, MS Project, Naviswork)."
                },
                "HABILITATION": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Les habilitations professionnelles spécifiques (par exemple, GIES 1/2, Elf Gabon HS3)."
                },
                "Projets effectués": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "CLIENT_NOM": {"type": "string", "description": "Nom du client."},
                            "DATE_DEBUT": {"type": "string", "description": "Date de début du projet."},
                            "DATE_FIN": {"type": "string", "description": "Date de fin du projet."},
                            "INTITULE_POSTE": {"type": "string", "description": "Intitulé du poste occupé."},
                            "INTITULE_PROJET": {"type": "string", "description": "Intitulé du projet réalisé."},
                            "DETAILS_PROJET": {"type": "string", "description": "Informations supplémentaires tel que le budget, les effectifs et les heures sans accident"},
                            "REALISATION": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Réalisations principales du projet."
                            }
                        },
                        "required": ["CLIENT_NOM", "DATE_DEBUT", "DATE_FIN", "INTITULE_POSTE", "INTITULE_PROJET", "REALISATION"]
                    }
                },
                "Diplômes": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "ANNEE_DIPLOME": {"type": "string", "description": "Année d'obtention du diplôme."},
                            "INTITULE_DIPLOME": {"type": "string", "description": "Intitulé complet du diplôme obtenu."}
                        },
                        "required": ["ANNEE_DIPLOME", "INTITULE_DIPLOME"]
                    }
                },
                "Langues": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "LANGUE": {"type": "string", "description": "Nom de la langue parlée."},
                            "NIVEAU": {"type": "string", "description": "Niveau de maîtrise de la langue (exemple : Courant, Intermédiaire)."}
                        },
                        "required": ["LANGUE", "NIVEAU"]
                    }
                },
                "Formations complémentaires": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "ANNEE_FORMATION": {"type": "string", "description": "Année de la formation complémentaire."},
                            "INTITULE_FORMATION": {"type": "string", "description": "Intitulé complet de la formation complémentaire."}
                        },
                        "required": ["ANNEE_FORMATION", "INTITULE_FORMATION"]
                    }
                }
            },
            "required": [
                "INTITULE_DU_POSTE", "EXPERTISE", "SECTEUR", "METHODOLOGIE", "HABILITATION", "Projets effectués", "Diplômes", "Langues", "Formations complémentaires"
            ]
        }
    }

    # Appel à l'API OpenAI avec Function Calling
    response = openai.chat.completions.create(
        model="gpt-5",
        messages=[
            {"role": "system", "content": "Tu es un assistant qui aide à extraire les informations des CV."},
            {"role": "user", "content": cv_text}
        ],
        functions=[function_schema],
        function_call={"name": "extract_cv_info"}
    )

    # Accéder directement aux arguments sous forme de chaîne JSON
    arguments = response.choices[0].message.function_call.arguments

    # Convertir la chaîne JSON en dictionnaire Python
    info = json.loads(arguments)

    # Générer le trigramme localement
    prenom = info.get("PRENOM", "")
    nom = info.get("NOM", "")
    info["TRI"] = generate_trigramme(prenom, nom)

    # Extraire l'âge via regex sur le texte du CV
    age_match = re.search(r'(\d{2})\s*ans', cv_text, re.IGNORECASE)
    if age_match:
        age = int(age_match.group(1))
        current_year = datetime.now().year
        annee_naissance = current_year - age
        info["ANNEE"] = annee_naissance
    else:
        info["ANNEE"] = ""

    # Extraire le téléphone via regex sur le texte du CV
    tel_match = re.search(r'(\d{2}(?:[\s\.-]?\d{2}){4})', cv_text)
    if tel_match:
        info["TELEPHONE"] = tel_match.group(1)
    else:
        info["TELEPHONE"] = ""

    # Extraire l'email via regex sur le texte du CV
    email_match = re.search(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', cv_text)
    if email_match:
        info["EMAIL"] = email_match.group(0)
    else:
        info["EMAIL"] = ""

    return info


def fill_word_template_with_lists(template_path, output_path, data):
    """
    Remplit un modèle Word avec des données (y compris dans l'en-tête),
    en remplaçant les placeholders et en appliquant les styles nécessaires.

    Arguments :
        template_path (str) : Chemin vers le modèle Word.
        output_path (str) : Chemin vers le fichier Word généré.
        data (dict) : Données à insérer dans le fichier Word.
    """
    doc = Document(template_path)

    # --- 🔹 1. Gestion des en-têtes et pieds de page ---
    for section in doc.sections:
        header = section.header
        footer = section.footer

        # En-tête
        for paragraph in header.paragraphs:
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))

        # Pied de page (optionnel, même logique)
        for paragraph in footer.paragraphs:
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))

    # --- 🔹 2. Corps du document ---
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"  # Placeholder au format {{KEY}}

            # --- Projets effectués ---
            if key == "Projets effectués" and isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""
                    for projet in value:
                        client_nom = projet.get('CLIENT_NOM', 'Non spécifié')
                        dates = f"{projet.get('DATE_DEBUT', 'N/A')} - {projet.get('DATE_FIN', 'N/A')}"
                        client_date_line = f"{client_nom}\t{dates}"

                        client_date_paragraph = paragraph.insert_paragraph_before(client_date_line)
                        client_date_paragraph.style = "italique gras"

                        tab_stops = client_date_paragraph.paragraph_format.tab_stops
                        tab_stop = tab_stops.add_tab_stop(Inches(6.5))
                        tab_stop.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                        post_paragraph = paragraph.insert_paragraph_before(projet.get('INTITULE_POSTE', 'Non spécifié'))
                        post_paragraph.style = paragraph.style

                        paragraph.insert_paragraph_before("")

                        project_paragraph = paragraph.insert_paragraph_before(projet.get('INTITULE_PROJET', 'Non spécifié'))
                        project_paragraph.style = paragraph.style
                        project_paragraph.runs[0].bold = True

                        details_projet = projet.get('DETAILS_PROJET', '').strip()
                        if details_projet:
                            project_paragraph = paragraph.insert_paragraph_before(details_projet)
                            project_paragraph.style = paragraph.style

                        paragraph.insert_paragraph_before("")

                        realizations = projet.get('REALISATION', [])
                        if realizations:
                            realizations_paragraph = paragraph.insert_paragraph_before("Réalisations :")
                            realizations_paragraph.style = paragraph.style
                            realizations_paragraph.runs[0].bold = True

                            for realization in realizations:
                                realization = realization.strip()
                                if realization:
                                    realization_paragraph = paragraph.insert_paragraph_before(realization)
                                    realization_paragraph.style = "Liste à puces1"

                            paragraph.insert_paragraph_before("")

            # --- Diplômes ---
            elif key == "Diplômes" and isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""
                    for diplome in value:
                        diploma_line = f"{diplome.get('ANNEE_DIPLOME', 'N/A')}    {diplome.get('INTITULE_DIPLOME', 'Non spécifié')}"
                        diploma_paragraph = paragraph.insert_paragraph_before(diploma_line)
                        diploma_paragraph.style = paragraph.style

            # --- Langues ---
            elif key == "Langues" and isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""
                    for langue in value:
                        language_line = f"{langue.get('LANGUE', 'Non spécifié')}    {langue.get('NIVEAU', 'Non spécifié')}"
                        language_paragraph = paragraph.insert_paragraph_before(language_line)
                        language_paragraph.style = paragraph.style

            # --- Formations complémentaires ---
            elif key == "Formations complémentaires" and isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""
                    for formation in value:
                        formation_line = f"{formation.get('ANNEE_FORMATION', 'N/A')}    {formation.get('INTITULE_FORMATION', 'Non spécifié')}"
                        formation_paragraph = paragraph.insert_paragraph_before(formation_line)
                        formation_paragraph.style = paragraph.style

            # --- Listes génériques ---
            elif isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""
                    for item in value:
                        list_paragraph = paragraph.insert_paragraph_before(str(item))
                        list_paragraph.style = paragraph.style

            # --- Valeurs simples ---
            elif placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, str(value))

    # --- 🔹 3. Sauvegarde du fichier final ---
    doc.save(output_path)
