from pydantic import BaseModel, Field
from typing import List, Optional
from openai import OpenAI
import re
import streamlit as st
import openai
import os

openai.api_key = st.secrets["OPENAI_API_KEY"]

client = OpenAI()

# Modèles pour le CV format International

class Diplome(BaseModel):
    """Diplôme obtenu avec année, intitulé, établissement et ville"""
    ANNEE: str = Field(..., description="Année d'obtention du diplôme.")
    DIPLOME: str = Field(..., description="Intitulé complet du diplôme obtenu.")
    ECOLE: str = Field(..., description="Nom de l'établissement ou école.")
    VILLE: str = Field(..., description="Ville de l'établissement.")


class Projet(BaseModel):
    """
    Projet unifié utilisé à la fois pour le tableau des expériences et la section projets.
    Chaque projet correspond à une ligne du tableau ET une entrée dans la section projets.
    """
    CLIENT_NOM: str = Field(..., description="Nom du client du projet (ex: SOFREGAZ, TPAO, GAZ de FRANCE).")
    DATES: str = Field(..., description="Dates du projet au format : AAAA ou AAAA-AAAA ou MM/AAAA - MM/AAAA.")
    ENTREPRISE: str = Field(..., description="Nom de l'entreprise pour laquelle la personne travaillait (ex: PROCESS SYSTEMS, AIR LIQUIDE Engineering).")
    INTITULE_POSTE: str = Field(..., description="Intitulé précis du poste occupé sur ce projet (ex: Chef de Projet).")
    INTITULE_PROJET: str = Field(..., description="Intitulé complet du projet réalisé.")
    REALISATION: List[str] = Field(..., description="Liste des réalisations principales du projet.")


class CompetenceLangue(BaseModel):
    """Niveau de maîtrise d'une langue (lue, écrite, parlée)"""
    LUE: str = Field(..., description="Niveau de lecture (ex: Maternelle, Courant, Scolaire ou niveau indiqué).")
    ECRITE: str = Field(..., description="Niveau d'écriture (ex: Maternelle, Courant, Scolaire ou niveau indiqué).")
    PARLEE: str = Field(..., description="Niveau oral (ex: Maternelle, Courant, Scolaire ou niveau indiqué).")


class Langue(BaseModel):
    """Langue parlée avec niveaux de compétence"""
    LANGUE: str = Field(..., description="Nom de la langue (ex: Français, Anglais, Espagnol). Laiiser vide si non spécifié.")
    COMPETENCE: CompetenceLangue = Field(..., description="Niveaux de compétence pour cette langue.")


class CVInterInfo(BaseModel):
    """
    Modèle principal pour le CV format International/Consultant.
    Correspond à la template CV_INTER_TEMPLATE_1.docx
    """
    PRENOM: str = Field(..., description="Prénom du consultant.")
    NOM: str = Field(..., description="Nom du consultant.")
    POSTE: str = Field(..., description="Intitulé du poste recherché ou actuel.")
    NOM_EMPLOYE: str = Field(..., description="Nom de l'employé tel qu'enregistré officiellement.")
    EDUCATION: str = Field(..., description="Niveau d'éducation le plus élevé atteint. (ex: Ingénieur des Mines, Master en génie civil)")
    NATIONALITE: str = Field(..., description="Nationalité du consultant.")
    DATE_NAISSANCE: str = Field(..., description="Date de naissance au format JJ/MM/AAAA.")
    AFFILIATIONS: Optional[str] = Field(None, description="Affiliation à des associations ou groupements professionnels.")
    ATTRIBUTIONS: Optional[str] = Field(None, description="Attributions spécifiques ou responsabilités particulières.")
    QUALIFICATIONS: List[str] = Field(..., description="Liste des qualifications professionnelles principales.")
    DIPLOMES: List[Diplome] = Field(..., description="Liste des diplômes obtenus avec détails.")
    CERTIFICATIONS: List[str] = Field(..., description="Liste des certifications et habilitations professionnelles.")
    PROJETS: List[Projet] = Field(..., description="Liste unifiée des projets. Chaque projet sera utilisé pour générer à la fois une ligne dans le tableau des expériences ET une entrée dans la section projets détaillés.")
    LANGUES: List[Langue] = Field(..., description="Langues maîtrisées avec niveaux de compétence (lue, écrite, parlée).")
    PAYS: str = Field(...,description="Pays dans lesquels le candidat a travaillé durant les dix dernières années. Format : texte, éléments séparés par des virgules et 'et' avant le dernier élément.")


def calculer_nombre_annees(experiences):
    """
    Calcule le nombre d'années d'expérience professionnelle totale à partir des expériences.
    Prend le min et le max des années trouvées dans PERIODE_DATES et retourne la différence
    """
    annees = []
    for exp in experiences:
        periode = exp.get("PERIODE_DATES", "")
        annees += [int(a) for a in re.findall(r"\b(19\d{2}|20\d{2})\b", periode)]
    if annees:
        return max(annees) - min(annees) 
    return 0

    
def generate_cv_inter_filename(infos: dict) -> str:
    nom = infos.get("NOM")
    prenom = infos.get("PRENOM")
    if nom and prenom:
        return f"DC INTER {nom} {prenom}.docx"
    else:
        return "DC INTER.docx"
    

def extract_info_from_cv_inter(cv_text: str, language: str = "fr") -> CVInterInfo:
    """
    Extrait des informations structurées à partir d'un texte de CV en utilisant l'API OpenAI.
    
    Arguments :
        cv_text (str) : Contenu textuel du CV.

    Retourne :
        CVInfo : Un objet Pydantic contenant les informations extraites.
    """
    system_prompt = {
        "fr": "Tu es un assistant qui aide à extraire les informations des CV.",
        "en": "You are an assistant that helps extract information from resumes. Extract the required fields in english."
    }

    system_prompt = system_prompt.get(language, system_prompt["fr"])
    
    completion = client.chat.completions.parse(
        model="gpt-5",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": cv_text},
        ],
        response_format=CVInterInfo,  
    )

    parsed: CVInterInfo = completion.choices[0].message.parsed
    info = parsed.model_dump()

    # Mettre en majuscule les clés importantes
    for key in ["NOM", "NOM_EMPLOYE", "EDUCATION", "AFFILIATIONS", "NATIONALITE", "ATTRIBUTIONS"]:
        if key in info and isinstance(info[key], str):
            info[key] = info[key].upper()

    # Calculer NOMBRE_ANNEES à partir des expériences
    nb_annees = calculer_nombre_annees(info.get("EXPERIENCES", []))
    if nb_annees == 1:
        info["NOMBRE_ANNEES"] = "1 AN"
    elif nb_annees > 1:
        info["NOMBRE_ANNEES"] = f"{nb_annees} ANS"
    else:
        info["NOMBRE_ANNEES"] = ""

    # Dupliquer POSTE vers POSTE_1 et POSTE_2
    if "POSTE" in info:
        info["POSTE_1"] = info["POSTE"]
        info["POSTE_2"] = info["POSTE"]
    
    info["POSTE_2"] = info["POSTE_2"].upper()  

    return info

