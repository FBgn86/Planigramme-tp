import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import re
from docx import Document
import PyPDF2
from odf.opendocument import load
from odf.text import P
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuration de la page
st.set_page_config(page_title="Générateur de Planigramme TP", layout="wide")
st.title("📋 Générateur Automatique de Planigramme TP")

# Initialisation du session_state pour le bouton réinitialiser
if 'fichiers_uploades' not in st.session_state:
    st.session_state.fichiers_uploades = []
if 'techniques_detectees' not in st.session_state:
    st.session_state.techniques_detectees = []

# Bouton Réinitialiser en haut de page
if st.session_state.fichiers_uploades or st.session_state.techniques_detectees:
    if st.button("🔄 Réinitialiser tout", type="secondary"):
        st.session_state.fichiers_uploades = []
        st.session_state.techniques_detectees = []
        st.rerun()

# Listes de techniques par catégorie
TECHNIQUES_SALES = [
    "éplucher", "laver", "émincer", "hacher", "tailler", "couper", "râper",
    "monder", "épépiner", "dénoyauter", "écailler", "vider", "nettoyer poisson",
    "lever filets", "désosser", "dépouiller", "écosser", "égrainer", "battre viande"
]

TECHNIQUES_PROPRES = [
    "nettoyer plan de travail", "désinfecter", "ranger", "organiser poste",
    "préparer matériel", "vérifier température frigo", "contrôler DLC",
    "étiqueter", "filmer", "stocker", "trier déchets", "nettoyer ustensiles"
]

TECHNIQUES_CUISSON = [
    "cuire au four", "frire", "sauter", "griller", "rôtir", "braiser", "mijoter",
    "pocher", "cuire vapeur", "sous vide", "blanchir", "confire", "dorer",
    "réchauffer", "maintenir au chaud", "cuisson basse température"
]

def detecter_categorie(technique):
    """Détecte la catégorie d'une technique"""
    tech_lower = technique.lower()
    for t in TECHNIQUES_SALES:
        if t in tech_lower:
            return "sale", "🟢"
    for t in TECHNIQUES_PROPRES:
        if t in tech_lower:
            return "propre", "🔵"
    for t in TECHNIQUES_CUISSON:
        if t in tech_lower:
            return "cuisson", "🔴"
    return "autre", "⚪"

def extraire_texte_fichier(fichier):
    """Extrait le texte selon le format du fichier"""
    nom = fichier.name.lower()

    try:
        if nom.endswith('.docx'):
            doc = Document(fichier)
            return "\n".join([p.text for p in doc.paragraphs])

        elif nom.endswith('.pdf'):
            pdf = PyPDF2.PdfReader(fichier)
            return "\n".join([page.extract_text() for page in pdf.pages])

        elif nom.endswith('.odt'):
            doc = load(fichier)
            textes = []
            for p in doc.getElementsByType(P):
                textes.append(str(p))
            return "\n".join(textes)

        elif nom.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(fichier)
            return "\n".join(df.astype(str).values.flatten())

        elif nom.endswith('.idt'):
            return fichier.read().decode('utf-8', errors='ignore')

        else:
            return fichier.read().decode('utf-8', errors='ignore')

    except Exception as e:
        st.error(f"Erreur lecture {fichier.name}: {str(e)}")
        return ""

def extraire_techniques(texte):
    """Extrait les techniques du texte"""
    lignes = texte.split('\n')
    techniques = []

    for ligne in lignes:
        ligne = ligne.strip()
        if len(ligne) > 3 and not ligne.startswith('#'):
            # Nettoyage basique
            ligne = re.sub(r'^\d+[\.\)]\s*', '', ligne)  # Enlever numérotation
            ligne = re.sub(r'^[-•*]\s*', '', ligne)      # Enlever puces
            if ligne:
                techniques.append(ligne)

    return techniques

def planifier_techniques(techniques):
    """Crée le planning avec horaires"""
    lignes = []
    heure_debut = datetime.now().replace(hour=8, minute=0, second=0, microsecond=0)

    for i, tech in enumerate(techniques):
        categorie, emoji = detecter_categorie(tech)
        duree = 30  # minutes par défaut

        # Ajustement durée selon catégorie
        if categorie == "cuisson":
            duree = 45
        elif categorie == "sale":
            duree = 20

        heure_fin = heure_debut + timedelta(minutes=duree)

        lignes.append({
            'heure_debut': heure_debut,
            'heure_fin': heure_fin,
            'technique': tech,
            'categorie': categorie,
            'emoji': emoji,
            'surveillance': categorie == "cuisson"
        })

        heure_debut = heure_fin + timedelta(minutes=5)  # Pause 5 min

    return lignes

def creer_excel_colore(lignes):
    """Crée le fichier Excel avec une seule feuille colorée"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Planigramme"

    # En-têtes
    headers = ["Heure début", "Heure fin", "Technique", "Catégorie", "Type"]
    ws.append(headers)

    # Styles
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    couleurs = {
        "sale": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),      # Vert
        "propre": PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),    # Bleu
        "cuisson": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),   # Rouge
        "autre": PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")      # Gris
    }

    # Remplissage des données
    for ligne in lignes:
        ws.append([
            ligne['heure_debut'].strftime("%H:%M"),
            ligne['heure_fin'].strftime("%H:%M"),
            f"{ligne['emoji']} {ligne['technique']}",
            ligne['categorie'].capitalize(),
            "Surveillance" if ligne['surveillance'] else "Effectif"
        ])

        # Coloration de la ligne
        row_num = ws.max_row
        fill = couleurs.get(ligne['categorie'], couleurs["autre"])
        for col in range(1, 6):
            ws.cell(row=row_num, column=col).fill = fill

    # Style en-tête
    for col in range(1, 6):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Ajustement colonnes
    for col in range(1, 6):
        ws.column_dimensions[get_column_letter(col)].width = 25

    # Sauvegarde en mémoire
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Interface utilisateur
st.subheader("📁 Upload des fiches techniques")
fichiers = st.file_uploader(
    "Glissez vos fichiers ici (DOCX, PDF, ODT, XLSX, IDT)",
    accept_multiple_files=True,
    type=['docx', 'pdf', 'odt', 'xlsx', 'xls', 'idt']
)

if fichiers:
    st.session_state.fichiers_uploades = fichiers
    st.success(f"✅ {len(fichiers)} fichier(s) uploadé(s)")

    # Traitement
    toutes_techniques = []
    for fichier in fichiers:
        with st.spinner(f"Traitement de {fichier.name}..."):
            texte = extraire_texte_fichier(fichier)
            techniques = extraire_techniques(texte)
            toutes_techniques.extend(techniques)

    st.session_state.techniques_detectees = toutes_techniques

    if st.button("🚀 Générer le planigramme", type="primary"):
        if toutes_techniques:
            lignes_plan = planifier_techniques(toutes_techniques)
            excel_file = creer_excel_colore(lignes_plan)

            st.success("✅ Planigramme généré avec succès !")

            # Aperçu
            st.subheader("Aperçu du planning")
            df_preview = pd.DataFrame([{
                'Début': l['heure_debut'].strftime("%H:%M"),
                'Fin': l['heure_fin'].strftime("%H:%M"),
                'Technique': f"{l['emoji']} {l['technique']}",
                'Catégorie': l['categorie'].capitalize(),
                'Type': 'Surveillance' if l['surveillance'] else 'Effectif'
            } for l in lignes_plan])
            st.dataframe(df_preview, use_container_width=True)

            # Téléchargement
            st.download_button(
                label="📥 Télécharger le planigramme Excel",
                data=excel_file,
                file_name="planigramme_tp.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ Aucune technique détectée.")
else:
    st.info("👆 Veuillez uploader vos fiches techniques pour commencer")
