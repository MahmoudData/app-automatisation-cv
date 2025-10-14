
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
from pydantic import BaseModel, Field
from typing import List
from openai import OpenAI


openai.api_key = st.secrets["OPENAI_API_KEY"]

client = OpenAI()


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
    # Enlever les espaces pour les noms composés
    nom_sans_espace = nom.replace(" ", "")
    first_letter = prenom[0] if prenom else ""
    consonnes = re.sub(r'[AEIOUY]', '', nom_sans_espace)
    trigramme = first_letter + consonnes[:2]
    return trigramme
    

class Projet(BaseModel):
    CLIENT_NOM: str = Field(..., description="Nom du client.")
    DATE_DEBUT: str = Field(..., description="Date de début du projet au format MM/AAAA.")
    DATE_FIN: str = Field(..., description="Date de fin du projet au format MM/AAAA.")
    INTITULE_POSTE: str = Field(..., description="Intitulé du poste occupé.")
    INTITULE_PROJET: str = Field(..., description="Intitulé du projet réalisé.")
    DETAILS_PROJET: str = Field(None, description="Informations supplémentaires tel que le budget, les effectifs et les heures sans accident")
    REALISATION: List[str] = Field(..., description="Réalisations principales du projet.")


class Diplome(BaseModel):
    ANNEE_DIPLOME: str = Field(..., description="Année d'obtention du diplôme.")
    INTITULE_DIPLOME: str = Field(..., description="Intitulé complet du diplôme obtenu.")


class Langue(BaseModel):
    LANGUE: str = Field(..., description="Nom de la langue parlée.")
    NIVEAU: str = Field(..., description="Niveau de maîtrise de la langue (exemple : Courant, Intermédiaire).")


class FormationComplementaire(BaseModel):
    ANNEE_FORMATION: str = Field(..., description="Année de la formation complémentaire.")
    INTITULE_FORMATION: str = Field(..., description="Intitulé complet de la formation complémentaire.")


class CVInfo(BaseModel):
    PRENOM: str = Field(..., description="prénom")
    NOM: str = Field(..., description="nom")
    EMAIL: str = Field(..., description="Adresse email")
    INTITULE_DU_POSTE: str = Field(..., description="L'intitulé du poste recherché.")
    EXPERTISE: List[str] = Field(..., description="Les activités et compétences spécifiques (par exemple, Etude de constructibilité, Résolution des problématiques, Leadership).")
    SECTEUR: List[str] = Field(..., description="Les domaines principaux d'expertise (par exemple, Bâtiment, Industrie, Oil & Gas).")
    METHODOLOGIE: List[str] = Field(..., description="Les méthodologies et outils maîtrisés (par exemple, Pack office, MS Project, Naviswork).")
    HABILITATION: List[str] = Field(..., description="Les habilitations professionnelles spécifiques (par exemple, GIES 1/2, Elf Gabon HS3).")
    Projets_effectués: List[Projet] = Field(..., description="Liste des projets effectués avec les détails de chaque mission.")
    Diplômes: List[Diplome] = Field(..., description="Liste des diplômes obtenus.")
    Langues: List[Langue] = Field(..., description="Langues parlées avec leur niveau de maîtrise.")
    Formations_complémentaires: List[FormationComplementaire] = Field(..., description="Formations complémentaires suivies.")


def extract_info_from_cv(cv_text: str) -> CVInfo:
    """
    Extrait des informations structurées à partir d'un texte de CV en utilisant l'API OpenAI.
    
    Arguments :
        cv_text (str) : Contenu textuel du CV.

    Retourne :
        CVInfo : Un objet Pydantic contenant les informations extraites.
    """
    completion = client.chat.completions.parse(
        model="gpt-5",
        messages=[
            {"role": "system", "content": "Tu es un assistant qui aide à extraire les informations des CV."},
            {"role": "user", "content": cv_text},
        ],
        response_format=CVInfo,  
    )

    parsed: CVInfo = completion.choices[0].message.parsed
    info = parsed.model_dump()

    # Générer le trigramme localement
    prenom = info.get("PRENOM", "")
    nom = info.get("NOM", "")
    info["NOM"] = nom.upper()  # Forcer le nom en majuscule
    info["TRI"] = generate_trigramme(prenom, nom)

    # Extraire l'âge via regex sur le texte du CV
    age_match = re.search(r"(\d{2})\s*ans(?!\s*d['’]?\s*exp)", cv_text, re.IGNORECASE)
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
        # Nettoyer le numéro pour enlever espaces, tirets, points
        raw_tel = tel_match.group(1)
        digits = re.sub(r'[^0-9]', '', raw_tel)
        # Reformater en XX.XX.XX.XX.XX
        if len(digits) == 10:
            info["TELEPHONE"] = '.'.join([digits[i:i+2] for i in range(0, 10, 2)])
        else:
            info["TELEPHONE"] = raw_tel
    else:
        info["TELEPHONE"] = ""

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
            if key == "Projets_effectués" and isinstance(value, list):
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
            elif key == "Formations_complémentaires" and isinstance(value, list):
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
