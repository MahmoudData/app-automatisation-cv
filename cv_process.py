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
                text = extract_text_from_pdf(tmp_file.name)
                os.unlink(tmp_file.name)
                return text
        else:
            return "Type de fichier non pris en charge. Veuillez fournir un fichier .docx ou .pdf."
    
    else:
        return "Paramètres insuffisants pour lire le fichier."
    

def extract_info_from_cv(cv_text):
    """
    Extrait des informations structurées à partir d'un texte de CV en utilisant l'API OpenAI.
    
    Arguments :
        cv_text (str) : Contenu textuel du CV.

    Retourne :
        dict : Un dictionnaire JSON contenant les informations extraites.
    """
    # Définition du schéma pour Function Calling
    function_schema = {
        "name": "extract_cv_info",
        "parameters": {
            "type": "object",
            "properties": {
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
    return json.loads(arguments)



def fill_word_template_with_lists(template_path, output_path, data):
    """
    Remplit un modèle Word avec des données, en remplaçant les placeholders
    et en appliquant un style spécifique si nécessaire.

    Arguments :
        template_path (str) : Chemin vers le modèle Word.
        output_path (str) : Chemin vers le fichier Word généré.
        data (dict) : Données à insérer dans le fichier Word.
    """
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"  # Placeholder au format {{KEY}}

            # Gestion des projets effectués
            if key == "Projets effectués" and isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""  

                    for projet in value:  
                        client_nom = projet.get('CLIENT_NOM', 'Non spécifié')
                        dates = f"{projet.get('DATE_DEBUT', 'N/A')} - {projet.get('DATE_FIN', 'N/A')}"

                        # Ajouter le texte avec tabulation
                        client_date_line = f"{client_nom}\t{dates}"
                        client_date_paragraph = paragraph.insert_paragraph_before(client_date_line)
                        client_date_paragraph.style = "italique gras" 

                        # Ajouter un tab stop aligné à droite
                        tab_stops = client_date_paragraph.paragraph_format.tab_stops
                        tab_stop = tab_stops.add_tab_stop(Inches(6.5)) 
                        tab_stop.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  

                        # Ajout du titre de poste
                        post_paragraph = paragraph.insert_paragraph_before(f"{projet.get('INTITULE_POSTE', 'Non spécifié')}")
                        post_paragraph.style = paragraph.style  

                        # Saut de ligne après le poste
                        paragraph.insert_paragraph_before("")  

                        # Ajout du projet en gras
                        project_paragraph = paragraph.insert_paragraph_before(f"{projet.get('INTITULE_PROJET', 'Non spécifié')}")
                        project_paragraph.style = paragraph.style  
                        project_paragraph.runs[0].bold = True  

                        # Ajout de details du projet
                        details_projet = projet.get('DETAILS_PROJET', '').strip()
                        if details_projet:  
                            project_paragraph = paragraph.insert_paragraph_before(details_projet)
                            project_paragraph.style = paragraph.style 

                        # Saut de ligne après le projet
                        paragraph.insert_paragraph_before("")  

                        # Récupérer les réalisations
                        realizations = projet.get('REALISATION', [])

                        # Vérifier si des réalisations existent avant d'ajouter quoi que ce soit
                        if realizations:  # Si le champ n'est pas vide
                            # Ajouter le titre "Réalisations :" en gras avant les réalisations
                            realizations_paragraph = paragraph.insert_paragraph_before("Réalisations :")
                            realizations_paragraph.style = paragraph.style  
                            realizations_paragraph.runs[0].bold = True  

                            # Ajouter les réalisations sous forme de bullet points
                            for realization in realizations:
                                realization = realization.strip()  
                                if realization:  
                                    realization_paragraph = paragraph.insert_paragraph_before(realization)
                                    realization_paragraph.style = "Liste à puces1"  
                            paragraph.insert_paragraph_before("") 

            # Gestion des diplômes
            elif key == "Diplômes" and isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""
                    for diplome in value:
                        diploma_line = (
                            f"{diplome.get('ANNEE_DIPLOME', 'N/A')}    {diplome.get('INTITULE_DIPLOME', 'Non spécifié')}"
                        )
                        diploma_paragraph = paragraph.insert_paragraph_before(diploma_line)
                        diploma_paragraph.style = paragraph.style

            # Gestion des langues
            elif key == "Langues" and isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""
                    for langue in value:
                        language_line = (
                            f"{langue.get('LANGUE', 'Non spécifié')}    {langue.get('NIVEAU', 'Non spécifié')}"
                        )
                        language_paragraph = paragraph.insert_paragraph_before(language_line)
                        language_paragraph.style = paragraph.style 

            # Gestion des formations complémentaires
            elif key == "Formations complémentaires" and isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""
                    for formation in value:
                        formation_line = (
                            f"{formation.get('ANNEE_FORMATION', 'N/A')}    {formation.get('INTITULE_FORMATION', 'Non spécifié')}"
                        )
                        formation_paragraph = paragraph.insert_paragraph_before(formation_line)
                        formation_paragraph.style = paragraph.style

            # Gestion des listes pour d'autres sections
            elif isinstance(value, list):
                if placeholder in paragraph.text:
                    paragraph.text = ""  # Efface le placeholder
                    for item in value:
                        list_paragraph = paragraph.insert_paragraph_before(str(item))
                        list_paragraph.style = paragraph.style  

            # Gestion des textes simples
            elif placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, str(value))

    # Sauvegarder le fichier Word rempli
    doc.save(output_path)
   