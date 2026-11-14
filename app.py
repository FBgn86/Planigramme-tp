import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import re
import zipfile
import xml.etree.ElementTree as ET
from docx import Document
import PyPDF2
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# Configuration de la page
st.set_page_config(page_title="Générateur de Planigramme TP", layout="wide")
st.title("📋 Générateur Automatique de Planigramme TP")

# Initialisation du session_state
if 'fichiers_uploades' not in st.session_state:
    st.session_state.fichiers_uploades = []
if 'techniques_detectees' not in st.session_state:
    st.session_state.techniques_detectees = []

# Bouton Réinitialiser
if st.session_state.fichiers_uploades:
    if st.button("🔄 Réinitialiser tout", type="secondary"):
        st.session_state.fichiers_uploades = []
        st.session_state.techniques_detectees = []
        st.rerun()

# Listes de techniques
TECHNIQUES_SALES = ["éplucher", "laver", "émincer", "hacher", "tailler", "couper", "râper", "monder", "épépiner", "dénoyauter"]
TECHNIQUES_PROPRES = ["nettoyer", "désinfecter", "ranger", "organiser", "préparer", "étiqueter", "stocker"]
TECHNIQUES_CUISSON = ["cuire", "frire", "sauter", "griller", "rôtir", "braiser", "mijoter", "pocher", "blanchir", "dorer"]

def detecter_categorie(technique):
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

def extraire_texte_odt(fichier):
    """Extrait le texte d'un fichier ODT (format ZIP/XML)"""
    try:
        # ODT est un ZIP contenant content.xml
        with zipfile.ZipFile(fichier) as z:
            with z.open('content.xml') as xml_file:
                tree = ET.parse(xml_file)
                root = tree.getroot()

                # Namespace ODF
                ns = {'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'}

                textes = []
                for p in root.iter('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p'):
                    texte = ''.join(t.text or '' for t in p.iter('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}span'))
                    if not texte:
                        texte = ''.join(t.text or '' for t in p.iter() if t.text)
                    if texte.strip():
                        textes.append(texte.strip())

                return "\n".join(textes)
    except Exception as e:
        st.error(f"Erreur ODT: {str(e)}")
        return ""

def extraire_texte_fichier(fichier):
    """Extrait le texte selon le format"""
    nom = fichier.name.lower()

    try:
        if nom.endswith('.docx'):
            doc = Document(fichier)
            return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

        elif nom.endswith('.pdf'):
            pdf = PyPDF2.PdfReader(fichier)
            return "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        elif nom.endswith('.odt'):
            return extraire_texte_odt(fichier)

        elif nom.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(fichier)
            return "\n".join(df.astype(str).values.flatten())

        elif nom.endswith('.idt'):
            return fichier.read().decode('utf-8', errors='ignore')

        else:
            return fichier.read().decode('utf-8', errors='ignore')

    except Exception as e:
        st.error(f"❌ Erreur {fichier.name}: {str(e)}")
        return ""

def extraire_techniques(texte):
    """Extrait les lignes de techniques"""
    lignes = texte.split('\n')
    techniques = []
    for ligne in lignes:
        ligne = ligne.strip()
        if len(ligne) > 3:
            ligne = re.sub(r'^\d+[\.\)]\s*', '', ligne)
            ligne = re.sub(r'^[-•*]\s*', '', ligne)
            if ligne and not ligne.startswith('#'):
                techniques.append(ligne)
    return techniques

def planifier_techniques(techniques):
    """Crée le planning"""
    lignes = []
    heure = datetime.now().replace(hour=8, minute=0, second=0, microsecond=0)

    for tech in techniques:
        categorie, emoji = detecter_categorie(tech)
        duree = 45 if categorie == "cuisson" else 20 if categorie == "sale" else 15
        heure_fin = heure + timedelta(minutes=duree)

        lignes.append({
            'debut': heure.strftime("%H:%M"),
            'fin': heure_fin.strftime("%H:%M"),
            'technique': f"{emoji} {tech}",
            'categorie': categorie,
            'type': 'Surveillance' if categorie == "cuisson" else 'Effectif'
        })
        heure = heure_fin + timedelta(minutes=5)

    return lignes

def creer_excel(lignes):
    """Crée l'Excel coloré"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Planigramme"

    headers = ["Heure début", "Heure fin", "Technique", "Catégorie", "Type"]
    ws.append(headers)

    couleurs = {
        "sale": "C6EFCE",      # Vert
        "propre": "BDD7EE",    # Bleu
        "cuisson": "FFC7CE",   # Rouge
        "autre": "F2F2F2"      # Gris
    }

    for l in lignes:
        ws.append([l['debut'], l['fin'], l['technique'], l['categorie'].capitalize(), l['type']])
        row = ws.max_row
        fill = PatternFill(start_color=couleurs.get(l['categorie'], "F2F2F2"),
                          end_color=couleurs.get(l['categorie'], "F2F2F2"),
                          fill_type="solid")
        for col in range(1, 6):
            ws.cell(row=row, column=col).fill = fill

    # Style header
    for col in range(1, 6):
        c = ws.cell(row=1, column=col)
        c.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        c.font = Font(color="FFFFFF", bold=True)

    for col in range(1, 6):
        ws.column_dimensions[get_column_letter(col)].width = 25

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Interface
st.subheader("📁 Upload des fiches techniques")
fichiers = st.file_uploader(
    "Formats acceptés: DOCX, PDF, ODT, XLSX, IDT",
    accept_multiple_files=True,
    type=['docx', 'pdf', 'odt', 'xlsx', 'xls', 'idt']
)

if fichiers:
    st.session_state.fichiers_uploades = fichiers
    st.success(f"✅ {len(fichiers)} fichier(s) chargé(s)")

    toutes_techniques = []
    for f in fichiers:
        with st.spinner(f"📖 Lecture de {f.name}..."):
            texte = extraire_texte_fichier(f)
            if texte:
                techs = extraire_techniques(texte)
                st.info(f"{f.name}: {len(techs)} technique(s) trouvée(s)")
                toutes_techniques.extend(techs)
            else:
                st.warning(f"⚠️ {f.name}: aucun texte extrait")

    st.session_state.techniques_detectees = toutes_techniques

    if st.button("🚀 Générer le planigramme", type="primary"):
        if toutes_techniques:
            plan = planifier_techniques(toutes_techniques)
            excel = creer_excel(plan)

            st.success("✅ Planigramme généré !")

            # Aperçu
            df = pd.DataFrame(plan)
            st.dataframe(df, use_container_width=True)

            st.download_button(
                "📥 Télécharger Excel",
                excel,
                "planigramme_tp.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ Aucune technique détectée. Vérifiez le contenu de vos fichiers.")
else:
    st.info("👆 Uploadez vos fiches techniques (ODT, DOCX, PDF...)")
